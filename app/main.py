import os
import sys
import time
import logging
import shutil
import threading
import uuid
import zipfile
import traceback
from concurrent.futures import ThreadPoolExecutor

from fastapi import FastAPI, UploadFile, File, HTTPException, Query, Request, Response, Depends
from fastapi.responses import FileResponse, RedirectResponse, JSONResponse
from starlette.background import BackgroundTask
from pydantic import BaseModel

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from engine.parser import parse_spir
from engine.workbook_loader import SUPPORTED_EXTENSIONS
from engine.extraction import build_extraction
from engine.sap_output import build_sap_output, backfill_part_output_rows
from engine.combine import consolidate_jobs
from engine import db
from engine import auth

# Parser diagnostics (e.g. which sheets were detected as SPIR sheets) go to
# the server log / `docker compose logs app`, never to the user.
logging.basicConfig(level=logging.WARNING, format='%(levelname)s %(name)s: %(message)s')
logging.getLogger('engine').setLevel(logging.INFO)
logger = logging.getLogger('bom_tool')

app = FastAPI(title='BOM Tool')


@app.exception_handler(Exception)
async def _unhandled_error(request: Request, exc: Exception):
    """Any unexpected server error still answers in JSON (the pages always
    expect JSON), with the details kept in the server log."""
    logger.exception('Unhandled error on %s %s', request.method, request.url.path)
    return JSONResponse({'detail': 'Server error while handling this request. Please try again; '
                                   'if it keeps happening, contact the administrator.'},
                        status_code=500)
db.init_db()
backfill_part_output_rows()
auth.init_auth_db()

static_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'static')

# Behind HTTPS (the Docker/Caddy deployment, see DEPLOY.md) set COOKIE_SECURE=true so the
# login cookie is only ever sent over HTTPS. Off by default for plain-HTTP local runs.
COOKIE_SECURE = os.environ.get('COOKIE_SECURE', '').strip().lower() in ('1', 'true', 'yes')


def require_login(request: Request) -> str:
    username = auth.get_session_user(request.cookies.get(auth.SESSION_COOKIE))
    if not username:
        raise HTTPException(401, 'Not logged in.')
    return username


def _count_rows_spares(parsed: dict):
    """(rows_count, spares_count) for the History table's ROWS/SPARES
    columns, counted the exact same way build_extraction itself writes
    rows: one equipment row per tag, plus one row per (tag, flagged
    item) pair across every sheet -- so this always matches the real
    Extraction file's data row count without having to reopen it."""
    spares_count = 0
    for tag in parsed['tag_order']:
        for sn in parsed['sheet_names']:
            for it in parsed['sheets'][sn]['items']:
                if tag in it['flags']:
                    spares_count += 1
    rows_count = len(parsed['tag_order']) + spares_count
    return rows_count, spares_count


class LoginRequest(BaseModel):
    username: str
    password: str


class JobIdsRequest(BaseModel):
    job_ids: list[str]


@app.post('/api/login')
def login(body: LoginRequest, response: Response):
    result = auth.verify_login(body.username, body.password)
    if result is None:
        raise HTTPException(401, 'Invalid username or password.')
    token = auth.create_session(result['username'])
    response.set_cookie(auth.SESSION_COOKIE, token, httponly=True, samesite='lax',
                         secure=COOKIE_SECURE, max_age=auth.SESSION_TTL_DAYS * 86400)
    return {'ok': True, **result}


@app.post('/api/logout')
def logout(request: Request, response: Response):
    token = request.cookies.get(auth.SESSION_COOKIE)
    if token:
        auth.delete_session(token)
    response.delete_cookie(auth.SESSION_COOKIE, httponly=True, samesite='lax', secure=COOKIE_SECURE)
    return {'ok': True}


@app.get('/api/me')
def me(username: str = Depends(require_login)):
    return {'username': username}


@app.get('/api/profile')
def profile(username: str = Depends(require_login)):
    """Settings page's Profile card. Real data only: full_name/created_at/
    last_login come from the users table, extractions_count from counting
    this user's own successful jobs -- jobs made before job ownership was
    tracked have no owner and aren't counted."""
    user = auth.get_user(username)
    if not user:
        raise HTTPException(404, 'User not found.')
    user['extractions_count'] = db.count_jobs_by_username(username)
    return user


class ChangePasswordRequest(BaseModel):
    current_password: str
    new_password: str


@app.post('/api/change-password')
def change_password(body: ChangePasswordRequest, username: str = Depends(require_login)):
    if len(body.new_password) < 6:
        raise HTTPException(400, 'New password must be at least 6 characters.')
    ok = auth.change_password(username, body.current_password, body.new_password)
    if not ok:
        raise HTTPException(400, 'Current password is incorrect.')
    return {'ok': True}


# ---------------------------------------------------------------------------
# SPIR processing runs in a background worker, never inside the HTTP request:
# the upload returns a job id at once and the page polls
# /api/process-status/{job_id}. A big or slow file therefore can't hold a
# request open until the reverse proxy times out (which the browser sees as
# an HTML error page instead of JSON), and can't freeze the rest of the app.
# One worker = one file at a time, so two large SPIRs never compete for RAM.
# ---------------------------------------------------------------------------
_executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix='spir')
_jobs = {}             # job_id -> {'status', 'result', 'error', 'username', 'queued_at', 'finished_at'}
_jobs_lock = threading.Lock()
_JOB_KEEP_SECONDS = 6 * 3600


def _set_job(job_id, **fields):
    with _jobs_lock:
        _jobs.setdefault(job_id, {}).update(fields)


def _prune_jobs():
    cutoff = time.time() - _JOB_KEEP_SECONDS
    with _jobs_lock:
        for jid in [j for j, v in _jobs.items() if v.get('finished_at') and v['finished_at'] < cutoff]:
            del _jobs[jid]


def _process_spir(job_id, jdir, filename, username, job_type, make_zip):
    """Parse one saved SPIR and build its Extraction + SAP OUTPUT files (and
    the results zip for the main tool). Returns the page's result dict and
    records the job in history either way."""
    src_path = os.path.join(jdir, filename)
    base = os.path.splitext(filename)[0]
    try:
        parsed = parse_spir(src_path)
        spir_seq = db.next_spir_seq()
        extraction_name = f'{base}_Extraction.xlsx'
        output_name = f'{base}_OUTPUT.xlsx'
        zip_name = f'{base}_results.zip' if make_zip else None

        extraction_path = os.path.join(jdir, extraction_name)
        output_path = os.path.join(jdir, output_name)

        build_extraction(parsed, extraction_path)
        build_sap_output(parsed, output_path, spir_filename=base, job_id=job_id)

        if make_zip:
            with zipfile.ZipFile(os.path.join(jdir, zip_name), 'w') as zf:
                zf.write(extraction_path, extraction_name)
                zf.write(output_path, output_name)

        rows_count, spares_count = _count_rows_spares(parsed)
        db.insert_job(
            job_id=job_id, spir_no=parsed['spir_no'], original_filename=filename,
            sheets=parsed['sheet_names'], tags=parsed['tag_order'],
            input_file=filename, extraction_file=extraction_name,
            output_file=output_name, zip_file=zip_name, status='ok', job_type=job_type,
            rows_count=rows_count, spares_count=spares_count, username=username, spir_seq=spir_seq,
        )

        result = {
            'job_id': job_id, 'spir_no': parsed['spir_no'], 'spir_seq': spir_seq,
            'spir_label': f'SPIR {spir_seq}', 'sheets': parsed['sheet_names'],
            'tags': parsed['tag_order'], 'extraction_file': extraction_name,
            'output_file': output_name,
            'part_numbers': db.spir_part_numbers_for_job(job_id),
        }
        if make_zip:
            result['zip_file'] = zip_name
        return result

    except Exception as e:
        db.insert_job(
            job_id=job_id, spir_no=None, original_filename=filename,
            sheets=[], tags=[], input_file=filename, extraction_file=None,
            output_file=None, zip_file=None, status='error',
            error_message=f'{e}\n{traceback.format_exc()}', job_type=job_type,
        )
        raise


def _run_job(job_id, jdir, filename, username, job_type, make_zip):
    _set_job(job_id, status='processing', started_at=time.time())
    try:
        result = _process_spir(job_id, jdir, filename, username, job_type, make_zip)
        _set_job(job_id, status='done', result=result, finished_at=time.time())
    except MemoryError:
        _set_job(job_id, status='error', finished_at=time.time(),
                 error='This file is too large for the server to process. Remove embedded '
                       'pictures/objects or unused sheets in Excel, save it again and retry.')
    except Exception as e:
        if not isinstance(e, ValueError):   # ValueError = the parser's own "can't read this file"
            logger.exception('Processing job %s (%s) failed', job_id, filename)
        _set_job(job_id, status='error', error=f'Could not process this file: {e}',
                 finished_at=time.time())


def _submit_upload(file: UploadFile, username: str, job_type: str, make_zip: bool):
    if not file.filename or not file.filename.lower().endswith(SUPPORTED_EXTENSIONS):
        raise HTTPException(400, 'Please upload an Excel SPIR file (.xlsx, .xlsm, .xlsb, .xltx or .xls).')
    filename = os.path.basename(file.filename)
    job_id = uuid.uuid4().hex[:12]
    jdir = db.job_dir(job_id)
    with open(os.path.join(jdir, filename), 'wb') as f:
        shutil.copyfileobj(file.file, f)

    _prune_jobs()
    _set_job(job_id, status='queued', username=username, filename=filename, queued_at=time.time())
    _executor.submit(_run_job, job_id, jdir, filename, username, job_type, make_zip)
    return JSONResponse({'job_id': job_id, 'status': 'queued'}, status_code=202)


@app.post('/api/process')
def process(file: UploadFile = File(...), username: str = Depends(require_login)):
    """The main tool's upload: Extraction + SAP OUTPUT + a results zip."""
    return _submit_upload(file, username, job_type='full', make_zip=True)


@app.post('/api/process-extraction')
def process_extraction(file: UploadFile = File(...), username: str = Depends(require_login)):
    """The dedicated Extraction page's upload -- builds both the
    Extraction file and the SAP OUTPUT file (same outputs as the main
    tool's /api/process), but is recorded with job_type='extraction' so
    it shows up on the Extraction page's own history view rather than
    the main tool's."""
    return _submit_upload(file, username, job_type='extraction', make_zip=False)


@app.get('/api/process-status/{job_id}')
def process_status(job_id: str, username: str = Depends(require_login)):
    """Polled by the upload pages until the job is 'done' (with the same
    result the upload used to return directly) or 'error'."""
    with _jobs_lock:
        job = dict(_jobs.get(job_id) or {})
        queued_ahead = sum(1 for v in _jobs.values()
                           if v.get('status') == 'queued' and v.get('queued_at', 0) < job.get('queued_at', 0))
    if not job:
        raise HTTPException(404, 'This processing job is no longer known to the server '
                                 '(the server may have restarted). Please upload the file again.')
    out = {'job_id': job_id, 'status': job['status']}
    if job['status'] == 'queued':
        out['queued_ahead'] = queued_ahead
    elif job['status'] == 'done':
        out['result'] = job['result']
    elif job['status'] == 'error':
        out['error'] = job['error']
    return out


@app.get('/api/download/{job_id}/{filename}')
def download(job_id: str, filename: str, username: str = Depends(require_login)):
    path = os.path.join(db.job_dir(job_id), filename)
    if not os.path.isfile(path):
        raise HTTPException(404, 'File not found.')
    return FileResponse(path, filename=filename, media_type='application/octet-stream')


@app.get('/api/parts/search')
def parts_search(q: str = Query(..., min_length=1),
                       field: str = Query(None, pattern=f"^({'|'.join(db.PART_SEARCH_FIELDS)})$"),
                       username: str = Depends(require_login)):
    """Parts Master search (see db.search_part_master): matches the
    selected `field` (Part Number, Tag, SAP Material Number, Material Temp
    Number, Category, New Description, Plant, Manufacturer Name/Country,
    or SPIR) against every stored extracted OUTPUT record and returns the
    matching records; omitted, it matches any field. `q` may use `*` as a
    wildcard (ABC exact, ABC* starts with, *ABC ends with, *ABC* contains)."""
    return db.search_part_master(q, field=field)


@app.get('/api/history')
def history(search: str = Query(None), limit: int = 50, job_type: str = Query(None),
                   username: str = Depends(require_login)):
    return db.list_jobs(limit=limit, search=search, job_type=job_type)


@app.get('/api/job/{job_id}')
def job_detail(job_id: str, username: str = Depends(require_login)):
    j = db.get_job(job_id)
    if not j:
        raise HTTPException(404, 'Job not found.')
    return j


@app.post('/api/jobs/combine')
def combine_jobs(body: JobIdsRequest, username: str = Depends(require_login)):
    """History page's 'Combine' action: CONSOLIDATES the selected jobs'
    data into one shared table per kind -- a single 'Extraction' sheet
    with every job's rows stacked together, a single 'BOM_WORKING' sheet,
    and a single 'SPIR VS TAG' sheet -- not a separate sheet per job.
    Nothing is re-parsed and no new job/history row is created for the
    merged file itself."""
    job_sources = []
    for job_id in body.job_ids:
        j = db.get_job(job_id)
        if not j or j['status'] != 'ok':
            continue
        extraction_path = os.path.join(db.job_dir(job_id), j['extraction_file']) if j['extraction_file'] else None
        if extraction_path and not os.path.isfile(extraction_path):
            extraction_path = None
        output_path = os.path.join(db.job_dir(job_id), j['output_file']) if j['output_file'] else None
        if output_path and not os.path.isfile(output_path):
            output_path = None
        if extraction_path or output_path:
            job_sources.append({'extraction_path': extraction_path, 'output_path': output_path})

    if not job_sources:
        raise HTTPException(400, 'None of the selected rows have a usable Extraction or OUTPUT file to combine.')

    combined_path = consolidate_jobs(job_sources)
    filename = f'Combined_{len(body.job_ids)}_files.xlsx'
    return FileResponse(
        combined_path, filename=filename, media_type='application/octet-stream',
        background=BackgroundTask(os.remove, combined_path),
    )


@app.post('/api/jobs/delete')
def delete_jobs(body: JobIdsRequest, username: str = Depends(require_login)):
    deleted = db.delete_jobs(body.job_ids)
    return {'deleted': deleted}


@app.get('/login.html')
async def login_page():
    return FileResponse(os.path.join(static_dir, 'login.html'))


@app.get('/theme.js')
async def theme_js():
    """Public (no login needed) -- login.html itself needs to apply a
    previously-saved appearance preference too."""
    return FileResponse(os.path.join(static_dir, 'theme.js'), media_type='application/javascript')


@app.get('/api.js')
async def api_js():
    """Public (no login needed) -- shared fetch/upload helpers, used by
    login.html too."""
    return FileResponse(os.path.join(static_dir, 'api.js'), media_type='application/javascript')


@app.get('/cdc-logo.jpg')
async def cdc_logo():
    """Public (no login needed) -- used as the sidebar brand mark and the
    browser-tab favicon on every page, including login.html."""
    return FileResponse(os.path.join(static_dir, 'cdc-logo.jpg'), media_type='image/jpeg')


@app.get('/bom-log.png')
async def login_background():
    """Public (no login needed) -- login.html's own background image."""
    return FileResponse(os.path.join(static_dir, 'bom-log.png'), media_type='image/png')


def _gated_page(filename: str):
    def handler(request: Request):
        if not auth.get_session_user(request.cookies.get(auth.SESSION_COOKIE)):
            return RedirectResponse('/login.html')
        return FileResponse(os.path.join(static_dir, filename))
    return handler


app.get('/')(_gated_page('extraction.html'))
app.get('/index.html')(_gated_page('index.html'))
app.get('/extraction.html')(_gated_page('extraction.html'))
app.get('/extraction-history.html')(_gated_page('extraction-history.html'))
app.get('/batch-extraction.html')(_gated_page('batch-extraction.html'))
app.get('/parts-master.html')(_gated_page('parts-master.html'))
app.get('/settings.html')(_gated_page('settings.html'))
