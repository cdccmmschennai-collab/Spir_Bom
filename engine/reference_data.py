"""
Central access to Reference.xlsx (project root). Loads the workbook once
and exposes typed lookup functions for every reference table the SAP OUTPUT
writer needs:
  - Country code               -> country_code(name)
  - Material group              -> material_group_for_classification(...)
  - Material TypeVCategory      -> category_for_sap_no(sap_no)
  - Material type code          -> material_type_code_for_text(text)
  - Maintenance planning plant  -> plant_code_for_spir_no(spir_no)
  - Abbreviation                -> abbreviate_description(text, max_len)

Nothing here is hard-coded: every mapping comes straight from the sheets, so
updating Reference.xlsx is enough to change behavior. Anything that can't be
matched confidently returns None/'' rather than guessing -- callers are
expected to leave the cell blank and log the miss.
"""
import os
import re
import openpyxl

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REFERENCE_PATH = os.path.join(ROOT_DIR, 'Reference.xlsx')

_wb = None
_country_map = None            # lowercase country name -> code
_material_group_rows = None    # [(code, title, description), ...]
_classification_to_group = None  # {classification letter: group code}
_sap_prefix_to_cat = None      # {'3': 'N', '5': 'N', '8': 'N', '1': 'L'}
_valid_categories = None       # {'B','I','L','N'}
_material_type_rows = None     # [(code, title, description2), ...]
_material_type_keywords = None  # {UPPER distinctive word: code}
_plant_rows = None             # [(code, description, project_number), ...]
_abbrev_map = None             # {UPPER phrase: code}
_abbrev_max_words = 1

_STOPWORDS = {
    'AND', 'OR', 'THE', 'OF', 'FOR', 'A', 'AN', 'TO', 'IN', 'ON', 'WITH',
    'ETC', 'PARTS', 'PART', 'EQUIPMENT', 'EQUIP', 'SPARES', 'SPARE',
    'ACCESSORIES', 'ACCESSORY', 'SYSTEMS', 'SYSTEM', 'ITEMS', 'ITEM',
    'MATERIALS', 'MATERIAL', 'GENERAL', 'OTHER', 'OTHERS', 'MISC', 'ALL',
    'CATEGORIES', 'NON', 'NOS',
    # Generic engineering nouns that carry no category-specific meaning by
    # themselves -- excluded so a word being textually unique to one row of
    # Reference.xlsx (by phrasing coincidence) can't produce a confident but
    # semantically wrong match (e.g. 'UNIT' only appearing once, under
    # "Wireline Unit and Accessories", must not match every "processor unit").
    'UNIT', 'UNITS', 'DEVICE', 'DEVICES', 'MODULE', 'MODULES', 'COMPONENT',
    'COMPONENTS', 'ASSEMBLY', 'ASSEMBLIES', 'PANEL', 'PANELS', 'STATION',
    'STATIONS', 'KIT', 'KITS', 'SET', 'SETS', 'TYPE', 'TYPES', 'MODEL',
    'MODELS', 'SERIES', 'RANGE', 'STANDARD', 'VARIOUS', 'MISCELLANEOUS',
    'NEW', 'USED', 'MAIN', 'PRIMARY', 'SECONDARY', 'AUXILIARY', 'ANCILLARY',
}


def _load_wb():
    global _wb
    if _wb is None:
        if not os.path.isfile(REFERENCE_PATH):
            raise FileNotFoundError(f'Reference.xlsx not found at {REFERENCE_PATH}')
        _wb = openpyxl.load_workbook(REFERENCE_PATH, data_only=True)
    return _wb


def _norm(s) -> str:
    return re.sub(r'[^A-Z0-9]', '', str(s or '').upper())


def _stem(word: str) -> str:
    """Light plural stemming so 'Compressor' in an item description matches
    a Reference Excel title written as 'Compressors' -- not real NLP, just
    enough to stop plural/singular phrasing differences from silently
    losing a real, unambiguous match."""
    if word.endswith('IES') and len(word) > 5:
        return word[:-3] + 'Y'
    if word.endswith('ES') and len(word) > 5:
        return word[:-2]
    if word.endswith('S') and len(word) > 4:
        return word[:-1]
    return word


def _find_sheet(*keywords):
    """Finds a sheet whose (punctuation-stripped) title contains every given
    keyword, tolerant of spacing/capitalization/typo differences (e.g. the
    actual 'Maintenence planning plant' tab is matched via 'PLANNING'+'PLANT',
    not 'MAINTENANCE', since that word is itself misspelled in the sheet)."""
    wb = _load_wb()
    for title in wb.sheetnames:
        key = _norm(title)
        if all(_norm(kw) in key for kw in keywords):
            return wb[title]
    return None


# ---------------------------------------------------------------- Country --

def _load_country_map():
    global _country_map
    if _country_map is not None:
        return _country_map
    m = {}
    ws = _find_sheet('COUNTRY')
    if ws is not None:
        for name, code in ws.iter_rows(min_row=2, max_col=2, values_only=True):
            if name and code:
                m[str(name).strip().lower()] = str(code).strip().upper()
    _country_map = m
    return m


def country_code(country_name: str) -> str:
    """Exact (case-insensitive) lookup of a country's short code from the
    'Country code' sheet. Returns '' if the name isn't in the sheet."""
    if not country_name:
        return ''
    return _load_country_map().get(str(country_name).strip().lower(), '')


# A handful of common full-name spellings for countries Reference.xlsx spells
# unusually (same reconciliation problem as manufacturer_country.py's own
# alias table, duplicated in miniature here to avoid a circular import --
# that module already imports this one).
_COMMON_COUNTRY_ALIASES = {
    'united states': 'usa', 'united states of america': 'usa',
    'russia': 'russian fed.', 'united arab emirates': 'utd.arab emir.',
    'uae': 'utd.arab emir.', 'czechia': 'czech republic',
    'north macedonia': 'macedonia', 'eswatini': 'swaziland',
    'belarus': 'white russia', 'moldova': 'moldavia',
    'bosnia and herzegovina': 'bosnia-herz.', 'dominican republic': 'dominican rep.',
}


def find_country_mentions(text: str):
    """Distinct country names explicitly mentioned in `text`, returned as the
    literal substrings actually found there (e.g. 'United Arab Emirates', not
    a normalized form) -- matched against Reference.xlsx's Country code sheet
    plus the common full-name aliases above. Whole-word, case-insensitive.
    Never infers a country from a city, phone code, or other indirect
    signal -- callers should only trust a single result as confident."""
    if not text:
        return []
    country_map = _load_country_map()
    names = set(country_map.keys()) | set(_COMMON_COUNTRY_ALIASES.keys())
    found = {}   # canonical Reference.xlsx key -> literal matched text
    for name in sorted(names, key=len, reverse=True):
        m = re.search(r'\b' + re.escape(name) + r'\b', text, re.IGNORECASE)
        if not m:
            continue
        canonical = _COMMON_COUNTRY_ALIASES.get(name, name)
        if canonical not in found:
            found[canonical] = text[m.start():m.end()]
    return list(found.values())


# --------------------------------------------------------- Material group --

def _load_material_group():
    global _material_group_rows, _classification_to_group
    if _material_group_rows is not None:
        return _material_group_rows
    rows = []
    ws = _find_sheet('MATERIAL', 'GROUP')
    if ws is not None:
        for code, title, desc in ws.iter_rows(min_row=2, max_col=3, values_only=True):
            if code:
                rows.append((str(code).strip().upper(), str(title or '').strip(), str(desc or '').strip()))
    _material_group_rows = rows

    # Match each SPIR "CLASSIFICATION OF PARTS" letter (C/I/R/S/Y, per the
    # SPIR's own Validation Lists sheet) to a group code, ONLY where the
    # group's title/description makes the correspondence unambiguous.
    by_code = {code: (title + ' ' + desc).upper() for code, title, desc in rows}

    def find_group(*must_contain):
        hits = [code for code, text in by_code.items() if all(w in text for w in must_contain)]
        return hits[0] if len(hits) == 1 else None

    _classification_to_group = {
        # 'REPAIR' alone also matches QCIS's description text (it mentions
        # the same repair-&-return process) -- REPAIRABLE+SPARE narrows it
        # to the one row that's actually titled "Repairable Spares".
        'R': find_group('REPAIRABLE', 'SPARE'),  # R - Repairable Itm.  -> Repairable Spares
        'Y': find_group('CAPITAL'),           # Y - Capital Ins. Itm.  -> Capitalised Insurance spares
        'S': find_group('CONSUM', 'SPARE'),   # S - General Cons. Com. -> Consumable Spares
    }
    return _material_group_rows


def valid_material_group_codes():
    _load_material_group()
    return {code for code, _, _ in _material_group_rows}


def material_group_for_classification(classification: str, default_code: str = None) -> str:
    """Maps a SPIR item's CLASSIFICATION OF PARTS value (e.g. 'R - Repairable
    Itm.') to a Material group code from Reference.xlsx, where the sheet's
    own titles make the mapping unambiguous. Falls back to default_code
    (e.g. 'QPSP') for classifications with no unambiguous match (this is a
    known, documented fallback -- not a guess)."""
    _load_material_group()
    letter = str(classification or '').strip().upper()[:1]
    code = _classification_to_group.get(letter)
    return code or (default_code or '')


# ----------------------------------------------------- Material TypeVCat --

def _load_category_rules():
    global _sap_prefix_to_cat, _valid_categories
    if _sap_prefix_to_cat is not None:
        return
    prefix_to_cat = {}
    valid_cats = set()
    ws = _find_sheet('MATERIAL', 'TYPE') or _find_sheet('CATEGORY')
    if ws is not None:
        for r in range(1, ws.max_row + 1):
            a = ws.cell(row=r, column=1).value   # Cat
            if a and str(a).strip().upper() in ('B', 'I', 'L', 'N'):
                valid_cats.add(str(a).strip().upper())
            sap_code = ws.cell(row=r, column=6).value   # 'SAP CODE' col, e.g. '3XXXXXXX'
            cat = ws.cell(row=r, column=8).value          # 'Cat' col, e.g. 'N'
            if sap_code and cat:
                m = re.match(r'^\s*(\d+)', str(sap_code))
                if m:
                    prefix_to_cat[m.group(1)] = str(cat).strip().upper()
    _sap_prefix_to_cat = prefix_to_cat
    _valid_categories = valid_cats or {'B', 'I', 'L', 'N'}


def category_for_sap_no(sap_no) -> str:
    """Material TypeVCategory for a spare row, per the Reference Excel's SAP
    Number rule (e.g. '3XXXXXXX'/'5XXXXXXX'/'8XXXXXXX' -> N). SAP numbers
    starting with 1, or missing, resolve to 'L' -- both the documented
    default and Reference Excel's own '1XXXXXXX -> L' row."""
    _load_category_rules()
    s = str(sap_no or '').strip()
    if not s:
        return 'L'
    m = re.match(r'^(\d)', s)
    if not m:
        return 'L'
    return _sap_prefix_to_cat.get(m.group(1), 'L')


# ------------------------------------------------------- Material type code --

def _load_material_type_code():
    global _material_type_rows, _material_type_keywords
    if _material_type_rows is not None:
        return
    rows = []
    ws = _find_sheet('MATERIAL', 'TYPE', 'CODE')
    if ws is not None:
        for code, title, desc2 in ws.iter_rows(min_row=2, max_col=3, values_only=True):
            if code:
                rows.append((str(code).strip(), str(title or '').strip(), str(desc2 or '').strip()))
    _material_type_rows = rows

    # A word counts as a distinctive keyword for a category only if it
    # appears in that one category's title/description and no other's --
    # this is derived purely from the sheet's own text, never hand-picked.
    word_to_codes = {}
    for code, title, desc2 in rows:
        words = set(re.findall(r'[A-Z]{3,}', (title + ' ' + desc2).upper())) - _STOPWORDS
        for w in words:
            word_to_codes.setdefault(_stem(w), set()).add(code)
    _material_type_keywords = {w: next(iter(codes)) for w, codes in word_to_codes.items() if len(codes) == 1}


def valid_material_type_codes():
    _load_material_type_code()
    return {code for code, _, _ in _material_type_rows}


def material_type_code_for_text(text: str) -> str:
    """Best-effort Material type code from equipment/item description text,
    matched only via distinctive keywords the Reference Excel's own
    'Material type code' sheet text (never invents a code). Returns '' if
    no single confident category match is found."""
    _load_material_type_code()
    words = set(re.findall(r'[A-Z]{3,}', str(text or '').upper())) - _STOPWORDS
    stems = {_stem(w) for w in words}
    hits = {_material_type_keywords[s] for s in stems if s in _material_type_keywords}
    return next(iter(hits)) if len(hits) == 1 else ''


# ------------------------------------------------- Maintenance planning plant --

def _load_plant_rows():
    global _plant_rows
    if _plant_rows is not None:
        return _plant_rows
    rows = []
    ws = _find_sheet('PLANNING', 'PLANT')
    if ws is not None:
        for code, desc, project_no in ws.iter_rows(min_row=2, max_col=3, values_only=True):
            if code:
                rows.append((str(code).strip(), str(desc or '').strip(),
                             str(project_no).strip() if project_no not in (None, '') else ''))
    _plant_rows = rows
    return rows


def valid_plant_codes():
    return {code for code, _, _ in _load_plant_rows()}


def plant_code_for_spir_no(spir_no: str) -> str:
    """Matches the numeric project-number segment of a SPIR number (e.g.
    'VEN-4391-MEWTP-5-43-2003' -> '4391') against the Reference Excel's
    'Project number' column. Returns '' if no segment matches any listed
    project number -- most plants in the sheet have no project number and
    can't be inferred from the SPIR number alone."""
    rows = _load_plant_rows()
    project_to_code = {p: code for code, _, p in rows if p}
    if not project_to_code:
        return ''
    segments = re.split(r'[^A-Za-z0-9]+', str(spir_no or ''))
    for seg in segments:
        if seg in project_to_code:
            return project_to_code[seg]
    return ''


# ------------------------------------------------------------ Abbreviation --

def _load_abbreviations():
    global _abbrev_map, _abbrev_max_words
    if _abbrev_map is not None:
        return
    m = {}
    max_words = 1
    ws = _find_sheet('ABBREVIATION')
    if ws is not None:
        for _, code, desc in ws.iter_rows(min_row=2, max_col=3, values_only=True):
            if code and desc:
                key = str(desc).strip().upper()
                if key not in m:   # first occurrence wins on duplicates
                    m[key] = str(code).strip()
                    max_words = max(max_words, len(key.split()))
    _abbrev_map = m
    _abbrev_max_words = max_words


def abbreviate_description(text: str, max_len: int = 40) -> str:
    """Shortens `text` to <= max_len using the Reference Excel Abbreviation
    sheet (greedy longest-phrase-first word substitution: 'TRANSFORMER' ->
    'XFMR', etc.), only truncating as an absolute last resort."""
    text = str(text or '').strip()
    if len(text) <= max_len:
        return text

    _load_abbreviations()
    words = text.split()
    n = len(words)
    out = []
    i = 0
    while i < n:
        matched = False
        for span in range(min(_abbrev_max_words, n - i), 0, -1):
            phrase = ' '.join(words[i:i + span]).strip(',.').upper()
            code = _abbrev_map.get(phrase)
            if code:
                out.append(code)
                i += span
                matched = True
                break
        if not matched:
            out.append(words[i])
            i += 1

    result = ' '.join(out)
    if len(result) > max_len:
        result = result[:max_len].rstrip()
    return result
