"""
Lightweight metadata store for every SPIR that's been processed: what was
uploaded, what came out, and when. Files themselves live on disk under
DATA_DIR/files/<job_id>/; this table just indexes them so they're browsable
and re-downloadable later, including after a restart.

Uses plain sqlite3 (no extra dependency) so this runs anywhere Python does.
If usage grows past a single small team, swapping DB_PATH for a Postgres
connection string is a contained change -- nothing above this module needs
to know which database is behind it.
"""
import sqlite3
import json
import os
import datetime
import contextlib

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data')
FILES_DIR = os.path.join(DATA_DIR, 'files')
DB_PATH = os.path.join(DATA_DIR, 'bom_tool.db')

os.makedirs(FILES_DIR, exist_ok=True)


def get_conn():
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA busy_timeout = 30000')
    return conn


@contextlib.contextmanager
def immediate_transaction():
    """A write transaction that takes SQLite's write lock up front (BEGIN
    IMMEDIATE) instead of on the first write statement -- used everywhere a
    'check if a row exists, else insert it' sequence must be atomic across
    concurrent processes (Part Number -> Material Temp Number allocation,
    SPIR sequence numbers). SQLite serializes writers file-wide, so a second
    caller's BEGIN IMMEDIATE simply blocks (up to the busy_timeout above)
    until the first one commits, then sees its committed row -- no lost
    updates and no duplicate numbers.

    isolation_level=None puts the connection in autocommit mode so this
    module's own explicit BEGIN IMMEDIATE / COMMIT take effect directly,
    instead of being pre-empted by sqlite3's default implicit transaction
    handling."""
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.isolation_level = None
    conn.execute('PRAGMA busy_timeout = 30000')
    conn.execute('BEGIN IMMEDIATE')
    try:
        yield conn
        conn.execute('COMMIT')
    except Exception:
        conn.execute('ROLLBACK')
        raise
    finally:
        conn.close()


def _next_counter(conn, name: str, first_value: int) -> int:
    """Atomically allocates the next value from a named sequence, creating
    it at `first_value` the first time it's used. Caller must already hold
    the write lock (see immediate_transaction) -- this alone is NOT safe
    against concurrent callers."""
    row = conn.execute('SELECT next_value FROM counters WHERE name = ?', (name,)).fetchone()
    if row is None:
        conn.execute('INSERT INTO counters (name, next_value) VALUES (?, ?)', (name, first_value + 1))
        return first_value
    value = row['next_value']
    conn.execute('UPDATE counters SET next_value = next_value + 1 WHERE name = ?', (name,))
    return value


def init_db():
    with get_conn() as conn:
        conn.execute('''
            CREATE TABLE IF NOT EXISTS jobs (
                job_id           TEXT PRIMARY KEY,
                created_at       TEXT NOT NULL,
                spir_no          TEXT,
                original_filename TEXT,
                sheets_json      TEXT,
                tags_json        TEXT,
                input_file       TEXT,
                extraction_file  TEXT,
                output_file      TEXT,
                zip_file         TEXT,
                status           TEXT DEFAULT 'ok',
                error_message    TEXT
            )
        ''')
        conn.execute('CREATE INDEX IF NOT EXISTS idx_jobs_spir ON jobs(spir_no)')
        conn.execute('CREATE INDEX IF NOT EXISTS idx_jobs_created ON jobs(created_at)')

        # job_type distinguishes a full run (extraction + SAP OUTPUT, from
        # the main page) from an extraction-only run (from the dedicated
        # Extraction page) so each page's history view can show just its
        # own kind of job. Guarded ALTER -- no-op on an already-migrated DB.
        existing_cols = {row['name'] for row in conn.execute('PRAGMA table_info(jobs)')}
        if 'job_type' not in existing_cols:
            conn.execute("ALTER TABLE jobs ADD COLUMN job_type TEXT NOT NULL DEFAULT 'full'")
        conn.execute('CREATE INDEX IF NOT EXISTS idx_jobs_type ON jobs(job_type)')

        # rows_count/spares_count back the History table's ROWS/SPARES
        # columns -- captured once at process time (same counting logic
        # build_extraction itself uses: one equipment row per tag, plus
        # one row per flagged spare item) rather than re-opening the
        # saved .xlsx on every history listing.
        if 'rows_count' not in existing_cols:
            conn.execute('ALTER TABLE jobs ADD COLUMN rows_count INTEGER')
        if 'spares_count' not in existing_cols:
            conn.execute('ALTER TABLE jobs ADD COLUMN spares_count INTEGER')

        # Settings page's Profile card 'Extractions' stat: which logged-in
        # user created this job. Jobs made before this column existed have
        # no owner (NULL) -- the count only reflects jobs made from here on.
        if 'username' not in existing_cols:
            conn.execute('ALTER TABLE jobs ADD COLUMN username TEXT')
        conn.execute('CREATE INDEX IF NOT EXISTS idx_jobs_username ON jobs(username)')

        # spir_seq: the sequential "SPIR 1", "SPIR 2", ... label for this
        # job, distinct from spir_no (the real SPIR document number parsed
        # from the file). Allocated once, atomically, per successfully
        # parsed upload -- see next_spir_seq(). NULL for jobs from before
        # this column existed, or that failed before a number was assigned.
        if 'spir_seq' not in existing_cols:
            conn.execute('ALTER TABLE jobs ADD COLUMN spir_seq INTEGER')
        conn.execute('CREATE UNIQUE INDEX IF NOT EXISTS idx_jobs_spir_seq ON jobs(spir_seq) WHERE spir_seq IS NOT NULL')

        # Named atomic sequences (see _next_counter/immediate_transaction):
        # 'spir' for spir_seq, plus one per Material Temp Number series
        # ('equipment_material_temp', 'spare_material_temp').
        conn.execute('''
            CREATE TABLE IF NOT EXISTS counters (
                name        TEXT PRIMARY KEY,
                next_value  INTEGER NOT NULL
            )
        ''')

        # Central Part Number master: a Part Number's Material Temp Number
        # is assigned once -- the first time that part number is seen
        # anywhere in the system -- and reused forever after. part_number
        # is the NORMALIZED key (trimmed/uppercased, see rules.normalize_
        # part_value); part_number_display keeps the first-seen raw value
        # for showing back to a user. part_type separates the equipment
        # series (keyed by Model Number, 40001+) from the spare series
        # (keyed by Manufacturer's Part Number, 500001+) -- the two ranges
        # never overlap, so material_temp_number is UNIQUE across both.
        conn.execute('''
            CREATE TABLE IF NOT EXISTS part_number_master (
                id                     INTEGER PRIMARY KEY AUTOINCREMENT,
                part_number            TEXT NOT NULL,
                part_number_display    TEXT NOT NULL,
                part_type              TEXT NOT NULL,
                material_temp_number   INTEGER NOT NULL,
                created_at             TEXT NOT NULL,
                updated_at             TEXT NOT NULL,
                UNIQUE(part_number, part_type),
                UNIQUE(material_temp_number)
            )
        ''')
        conn.execute('CREATE INDEX IF NOT EXISTS idx_part_master_mtn ON part_number_master(material_temp_number)')

        # Every (SPIR job, Part Number) pairing ever seen -- the audit
        # trail that answers "which SPIRs used this part". Never
        # overwritten/deleted when a part number repeats in a later SPIR;
        # a new row is added instead, all pointing at the same master row.
        conn.execute('''
            CREATE TABLE IF NOT EXISTS spir_part_numbers (
                id                     INTEGER PRIMARY KEY AUTOINCREMENT,
                job_id                 TEXT NOT NULL REFERENCES jobs(job_id),
                part_number_master_id  INTEGER NOT NULL REFERENCES part_number_master(id),
                part_type              TEXT NOT NULL,
                tag                    TEXT,
                material_temp_number   INTEGER NOT NULL,
                created_at             TEXT NOT NULL,
                UNIQUE(job_id, part_number_master_id)
            )
        ''')
        conn.execute('CREATE INDEX IF NOT EXISTS idx_spir_pn_job ON spir_part_numbers(job_id)')
        conn.execute('CREATE INDEX IF NOT EXISTS idx_spir_pn_master ON spir_part_numbers(part_number_master_id)')

        # One row per BOM_WORKING row written to a job's OUTPUT file --
        # the extracted record the Parts Master page searches and returns
        # (Part Number, tag, SAP Material Number, category, description,
        # plant, manufacturer, country). Filled by sap_output.build_sap_
        # output for new jobs, and from the saved OUTPUT file for jobs made
        # before this table existed (sap_output.backfill_part_output_rows).
        conn.execute('''
            CREATE TABLE IF NOT EXISTS part_output_rows (
                id                        INTEGER PRIMARY KEY AUTOINCREMENT,
                job_id                    TEXT NOT NULL REFERENCES jobs(job_id),
                material_temp_number      INTEGER NOT NULL,
                tag                       TEXT,
                sap_material_number       TEXT,
                material_category         TEXT,
                new_description           TEXT,
                maintenance_plant         TEXT,
                manufacturer_name         TEXT,
                manufacturer_country_name TEXT
            )
        ''')
        conn.execute('CREATE INDEX IF NOT EXISTS idx_part_output_job ON part_output_rows(job_id)')
        conn.execute('CREATE INDEX IF NOT EXISTS idx_part_output_mtn ON part_output_rows(material_temp_number)')
        # part_number (Manufacturers Part Number) / part_type added later:
        # rows stored before they existed lack both, so they're cleared and
        # re-read from each job's OUTPUT file by the backfill.
        output_cols = {row['name'] for row in conn.execute('PRAGMA table_info(part_output_rows)')}
        if 'part_number' not in output_cols:
            conn.execute('ALTER TABLE part_output_rows ADD COLUMN part_number TEXT')
            conn.execute('ALTER TABLE part_output_rows ADD COLUMN part_type TEXT')
            conn.execute('DELETE FROM part_output_rows')


def next_spir_seq() -> int:
    """Allocates the next sequential 'SPIR N' number, atomically -- two
    uploads processed at the same instant by different requests/processes
    still get distinct, gapless-except-for-failures numbers (see
    immediate_transaction)."""
    with immediate_transaction() as conn:
        return _next_counter(conn, 'spir', 1)


_MTN_SERIES = {
    'equipment': ('equipment_material_temp', 40001),
    'spare': ('spare_material_temp', 500001),
}


def get_or_create_material_temp_number(part_number: str, part_type: str, display_value: str = None):
    """The central rule: if `part_number` (already normalized by the
    caller) has been seen before for this part_type, return its existing
    Material Temp Number unchanged. Otherwise allocate the next number in
    that type's series and record the new Part Number -> Material Temp
    Number mapping permanently. Runs under immediate_transaction so two
    concurrent callers with the SAME new part number can never both mint a
    number -- the second one blocks until the first commits, then finds
    the row the first one just inserted.

    Returns (material_temp_number: int, is_new: bool, master_id: int).
    """
    if part_type not in _MTN_SERIES:
        raise ValueError(f'unknown part_type {part_type!r}')
    counter_name, first_value = _MTN_SERIES[part_type]

    with immediate_transaction() as conn:
        row = conn.execute(
            'SELECT id, material_temp_number FROM part_number_master WHERE part_number = ? AND part_type = ?',
            (part_number, part_type)
        ).fetchone()
        if row is not None:
            return row['material_temp_number'], False, row['id']

        mtn = _next_counter(conn, counter_name, first_value)
        now = datetime.datetime.utcnow().isoformat()
        cur = conn.execute('''
            INSERT INTO part_number_master
                (part_number, part_number_display, part_type, material_temp_number, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (part_number, display_value or part_number, part_type, mtn, now, now))
        return mtn, True, cur.lastrowid


def record_spir_part_number(job_id: str, part_number_master_id: int, part_type: str, tag, material_temp_number):
    """Links a job (SPIR) to a Part Number it used -- the audit trail
    behind 'which SPIRs used this part'. Idempotent per (job, part number):
    if the same part number is flagged under several tags within one job,
    only the first link is kept (see the UNIQUE constraint) -- the job/part
    pairing itself doesn't repeat, only the row that established it."""
    with get_conn() as conn:
        conn.execute('''
            INSERT OR IGNORE INTO spir_part_numbers
                (job_id, part_number_master_id, part_type, tag, material_temp_number, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (job_id, part_number_master_id, part_type, tag, material_temp_number,
              datetime.datetime.utcnow().isoformat()))


PART_OUTPUT_COLUMNS = ('part_number', 'part_type', 'tag', 'sap_material_number', 'material_category', 'new_description',
                       'maintenance_plant', 'manufacturer_name', 'manufacturer_country_name')


def record_part_output_rows(job_id: str, rows: list):
    """Stores one job's OUTPUT rows' searchable attributes (see the
    part_output_rows table). Each row is a dict with 'material_temp_number'
    plus PART_OUTPUT_COLUMNS; replaces anything already stored for the job
    so re-recording the same job never duplicates rows."""
    cols = ('material_temp_number',) + PART_OUTPUT_COLUMNS
    with immediate_transaction() as conn:
        conn.execute('DELETE FROM part_output_rows WHERE job_id = ?', (job_id,))
        conn.executemany(
            f'INSERT INTO part_output_rows (job_id, {", ".join(cols)}) VALUES (?{", ?" * len(cols)})',
            [(job_id, *[r.get(c) for c in cols]) for r in rows])


def jobs_missing_part_output_rows():
    """Jobs with an OUTPUT file but no part_output_rows yet (made before
    that table existed) -- (job_id, output_file) pairs for
    sap_output.backfill_part_output_rows to fill in from the saved file."""
    with get_conn() as conn:
        return [(r['job_id'], r['output_file']) for r in conn.execute('''
            SELECT j.job_id, j.output_file FROM jobs j
            WHERE j.output_file IS NOT NULL
              AND NOT EXISTS (SELECT 1 FROM part_output_rows r WHERE r.job_id = j.job_id)
        ''')]


# Parts Master search field -> the stored column(s) it matches (r =
# part_output_rows, j = jobs), each compared with LIKE ? ESCAPE '!' against
# wildcard_to_like's pattern.
PART_SEARCH_FIELDS = {
    'part_number': ['r.part_number'],
    'tag': ['r.tag'],
    'sap_material_number': ['r.sap_material_number'],
    'material_temp_number': ['CAST(r.material_temp_number AS TEXT)'],
    'material_category': ['r.material_category'],
    'new_description': ['r.new_description'],
    'maintenance_plant': ['r.maintenance_plant'],
    'manufacturer_name': ['r.manufacturer_name'],
    'manufacturer_country_name': ['r.manufacturer_country_name'],
    'spir': ['j.spir_no', "('SPIR ' || j.spir_seq)"],
}


def wildcard_to_like(query: str) -> str:
    """Turns a Parts Master search value into a LIKE pattern (used with
    ESCAPE '!'): `*` is the only wildcard, so `ABC` is an exact match,
    `ABC*` starts with, `*ABC` ends with and `*ABC*` contains. LIKE's own
    `%` and `_` (and the `!` escape character) are escaped so they match
    literally. SQLite's LIKE is case-insensitive for ASCII, so matches are too."""
    escaped = query.strip().replace('!', '!!').replace('%', '!%').replace('_', '!_')
    return escaped.replace('*', '%')


def search_part_master(query: str, limit: int = 50, field: str = None):
    """Parts Master search over every extracted OUTPUT record
    (part_output_rows): `field` (one of PART_SEARCH_FIELDS) picks the
    stored column to match, None matches any of them. `query` may use `*`
    wildcards (see wildcard_to_like); without one it must match the whole
    value. Returns one entry per matching record -- newest job first, then
    in the record's OUTPUT row order -- with the History page's details of
    the job it came from (file name, SPIR No, tag count, processed time,
    and its Extraction/OUTPUT file names for /api/download)."""
    q = wildcard_to_like(query)
    if field:
        conditions = PART_SEARCH_FIELDS[field]
    else:
        conditions = [c for cs in PART_SEARCH_FIELDS.values() for c in cs]
    record_cols = ('material_temp_number',) + PART_OUTPUT_COLUMNS
    with get_conn() as conn:
        rows = conn.execute(f"""
            SELECT {', '.join('r.' + c for c in record_cols)},
                   j.job_id, j.spir_no, j.spir_seq, j.original_filename, j.created_at,
                   j.tags_json, j.extraction_file, j.output_file
            FROM part_output_rows r
            JOIN jobs j ON j.job_id = r.job_id
            WHERE {' OR '.join(f"{c} LIKE ? ESCAPE '!'" for c in conditions)}
            ORDER BY j.created_at DESC, r.id
            LIMIT ?
        """, (*[q] * len(conditions), limit)).fetchall()
    return [{
        **{c: r[c] for c in record_cols},
        'job_id': r['job_id'],
        'spir_no': r['spir_no'],
        'spir_label': f"SPIR {r['spir_seq']}" if r['spir_seq'] else None,
        'filename': r['original_filename'],
        'extracted_at': r['created_at'],
        'tags_count': len(json.loads(r['tags_json'] or '[]')),
        'extraction_file': r['extraction_file'],
        'output_file': r['output_file'],
    } for r in rows]


def spir_part_numbers_for_job(job_id: str):
    """The 'SPIR No. | Part Number | Material Temp Number' rows for one
    job -- used to show what a single upload resolved to right after
    processing."""
    with get_conn() as conn:
        rows = conn.execute('''
            SELECT sp.tag, sp.part_type, sp.material_temp_number, m.part_number_display
            FROM spir_part_numbers sp
            JOIN part_number_master m ON m.id = sp.part_number_master_id
            WHERE sp.job_id = ?
            ORDER BY sp.id
        ''', (job_id,)).fetchall()
    return [dict(r) for r in rows]


def job_dir(job_id: str) -> str:
    d = os.path.join(FILES_DIR, job_id)
    os.makedirs(d, exist_ok=True)
    return d


def insert_job(job_id, spir_no, original_filename, sheets, tags,
               input_file, extraction_file, output_file, zip_file,
               status='ok', error_message=None, job_type='full',
               rows_count=None, spares_count=None, username=None, spir_seq=None):
    with get_conn() as conn:
        conn.execute('''
            INSERT INTO jobs (job_id, created_at, spir_no, original_filename,
                               sheets_json, tags_json, input_file, extraction_file,
                               output_file, zip_file, status, error_message, job_type,
                               rows_count, spares_count, username, spir_seq)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (job_id, datetime.datetime.utcnow().isoformat(), spir_no, original_filename,
              json.dumps(sheets or []), json.dumps(tags or []),
              input_file, extraction_file, output_file, zip_file, status, error_message, job_type,
              rows_count, spares_count, username, spir_seq))


def count_jobs_by_username(username: str) -> int:
    with get_conn() as conn:
        row = conn.execute('SELECT COUNT(*) AS n FROM jobs WHERE username = ? AND status = ?',
                            (username, 'ok')).fetchone()
    return row['n'] if row else 0


def list_jobs(limit=50, search=None, job_type=None):
    q = 'SELECT * FROM jobs'
    where = []
    params = []
    if search:
        where.append('(spir_no LIKE ? OR original_filename LIKE ?)')
        params += [f'%{search}%', f'%{search}%']
    if job_type:
        where.append('job_type = ?')
        params.append(job_type)
    if where:
        q += ' WHERE ' + ' AND '.join(where)
    q += ' ORDER BY created_at DESC LIMIT ?'
    params.append(limit)
    with get_conn() as conn:
        rows = conn.execute(q, params).fetchall()
    return [_row_to_dict(r) for r in rows]


def get_job(job_id: str):
    with get_conn() as conn:
        row = conn.execute('SELECT * FROM jobs WHERE job_id = ?', (job_id,)).fetchone()
    return _row_to_dict(row) if row else None


def delete_jobs(job_ids: list) -> list:
    """Removes each job's DB row and its on-disk files directory. Returns
    the job_ids that actually existed and were deleted (silently skips
    ones that don't -- e.g. already deleted by someone else)."""
    import shutil
    deleted = []
    with get_conn() as conn:
        for job_id in job_ids:
            row = conn.execute('SELECT job_id FROM jobs WHERE job_id = ?', (job_id,)).fetchone()
            if not row:
                continue
            conn.execute('DELETE FROM jobs WHERE job_id = ?', (job_id,))
            conn.execute('DELETE FROM part_output_rows WHERE job_id = ?', (job_id,))
            deleted.append(job_id)
    for job_id in deleted:
        shutil.rmtree(os.path.join(FILES_DIR, job_id), ignore_errors=True)
    return deleted


def _row_to_dict(row):
    d = dict(row)
    d['sheets'] = json.loads(d.pop('sheets_json') or '[]')
    d['tags'] = json.loads(d.pop('tags_json') or '[]')
    return d
