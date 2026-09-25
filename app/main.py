import os
import sys
import logging
import shutil
import uuid
import zipfile
import traceback

from fastapi import FastAPI, UploadFile, File, HTTPException, Query, Request, Response, Depends
from fastapi.responses import FileResponse, RedirectResponse
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

app = FastAPI(title='BOM Tool')
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
async def login(body: LoginRequest, response: Response):
    result = auth.verify_login(body.username, body.password)
    if result is None:
        raise HTTPException(401, 'Invalid username or password.')
    token = auth.create_session(result['username'])
    response.set_cookie(auth.SESSION_COOKIE, token, httponly=True, samesite='lax',
                         secure=COOKIE_SECURE, max_age=auth.SESSION_TTL_DAYS * 86400)
    return {'ok': True, **result}


@app.post('/api/logout')
async def logout(request: Request, response: Response):
    token = request.cookies.get(auth.SESSION_COOKIE)
    if token:
        auth.delete_session(token)
    response.delete_cookie(auth.SESSION_COOKIE, httponly=True, samesite='lax', secure=COOKIE_SECURE)
    return {'ok': True}


@app.get('/api/me')
async def me(username: str = Depends(require_login)):
    return {'username': username}


@app.get('/api/profile')
async def profile(username: str = Depends(require_login)):
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
async def change_password(body: ChangePasswordRequest, username: str = Depends(require_login)):
    if len(body.new_password) < 6:
        raise HTTPException(400, 'New password must be at least 6 characters.')
    ok = auth.change_password(username, body.current_password, body.new_password)
    if not ok:
        raise HTTPException(400, 'Current password is incorrect.')
    return {'ok': True}


@app.post('/api/process')
async def process(file: UploadFile = File(...), username: str = Depends(require_login)):
    if not file.filename.lower().endswith(SUPPORTED_EXTENSIONS):
        raise HTTPException(400, 'Please upload an Excel SPIR file (.xlsx, .xlsm, .xlsb, .xltx or .xls).')

    job_id = uuid.uuid4().hex[:12]
    jdir = db.job_dir(job_id)

    src_path = os.path.join(jdir, file.filename)
    with open(src_path, 'wb') as f:
        shutil.copyfileobj(file.file, f)

    base = os.path.splitext(file.filename)[0]

    try:
        parsed = parse_spir(src_path)
        spir_seq = db.next_spir_seq()
        extraction_name = f'{base}_Extraction.xlsx'
        output_name = f'{base}_OUTPUT.xlsx'
        zip_name = f'{base}_results.zip'

        extraction_path = os.path.join(jdir, extraction_name)
        output_path = os.path.join(jdir, output_name)
        zip_path = os.path.join(jdir, zip_name)

        build_extraction(parsed, extraction_path)
        build_sap_output(parsed, output_path, spir_filename=base, job_id=job_id)

        with zipfile.ZipFile(zip_path, 'w') as zf:
            zf.write(extraction_path, extraction_name)
            zf.write(output_path, output_name)

        rows_count, spares_count = _count_rows_spares(parsed)
        db.insert_job(
            job_id=job_id, spir_no=parsed['spir_no'], original_filename=file.filename,
            sheets=parsed['sheet_names'], tags=parsed['tag_order'],
            input_file=file.filename, extraction_file=extraction_name,
            output_file=output_name, zip_file=zip_name, status='ok',
            rows_count=rows_count, spares_count=spares_count, username=username, spir_seq=spir_seq,
        )

        return {
            'job_id': job_id, 'spir_no': parsed['spir_no'], 'spir_seq': spir_seq,
            'spir_label': f'SPIR {spir_seq}', 'sheets': parsed['sheet_names'],
            'tags': parsed['tag_order'], 'extraction_file': extraction_name,
            'output_file': output_name, 'zip_file': zip_name,
            'part_numbers': db.spir_part_numbers_for_job(job_id),
        }

    except Exception as e:
        db.insert_job(
            job_id=job_id, spir_no=None, original_filename=file.filename,
            sheets=[], tags=[], input_file=file.filename, extraction_file=None,
            output_file=None, zip_file=None, status='error',
            error_message=f'{e}\n{traceback.format_exc()}',
        )
        raise HTTPException(422, f'Could not process this file: {e}')


@app.post('/api/process-extraction')
async def process_extraction(file: UploadFile = File(...), username: str = Depends(require_login)):
    """The dedicated Extraction page's upload -- builds both the
    Extraction file and the SAP OUTPUT file (same outputs as the main
    tool's /api/process), but is recorded with job_type='extraction' so
    it shows up on the Extraction page's own history view rather than
    the main tool's."""
    if not file.filename.lower().endswith(SUPPORTED_EXTENSIONS):
        raise HTTPException(400, 'Please upload an Excel SPIR file (.xlsx, .xlsm, .xlsb, .xltx or .xls).')

    job_id = uuid.uuid4().hex[:12]
    jdir = db.job_dir(job_id)

    src_path = os.path.join(jdir, file.filename)
    with open(src_path, 'wb') as f:
        shutil.copyfileobj(file.file, f)

    base = os.path.splitext(file.filename)[0]

    try:
        parsed = parse_spir(src_path)
        spir_seq = db.next_spir_seq()
        extraction_name = f'{base}_Extraction.xlsx'
        output_name = f'{base}_OUTPUT.xlsx'

        extraction_path = os.path.join(jdir, extraction_name)
        output_path = os.path.join(jdir, output_name)

        build_extraction(parsed, extraction_path)
        build_sap_output(parsed, output_path, spir_filename=base, job_id=job_id)

        rows_count, spares_count = _count_rows_spares(parsed)
        db.insert_job(
            job_id=job_id, spir_no=parsed['spir_no'], original_filename=file.filename,
            sheets=parsed['sheet_names'], tags=parsed['tag_order'],
            input_file=file.filename, extraction_file=extraction_name,
            output_file=output_name, zip_file=None, status='ok', job_type='extraction',
            rows_count=rows_count, spares_count=spares_count, username=username, spir_seq=spir_seq,
        )

        return {
            'job_id': job_id, 'spir_no': parsed['spir_no'], 'spir_seq': spir_seq,
            'spir_label': f'SPIR {spir_seq}', 'sheets': parsed['sheet_names'],
            'tags': parsed['tag_order'], 'extraction_file': extraction_name,
            'output_file': output_name,
            'part_numbers': db.spir_part_numbers_for_job(job_id),
        }

    except Exception as e:
        db.insert_job(
            job_id=job_id, spir_no=None, original_filename=file.filename,
            sheets=[], tags=[], input_file=file.filename, extraction_file=None,
            output_file=None, zip_file=None, status='error',
            error_message=f'{e}\n{traceback.format_exc()}', job_type='extraction',
        )
        raise HTTPException(422, f'Could not process this file: {e}')


@app.get('/api/download/{job_id}/{filename}')
async def download(job_id: str, filename: str, username: str = Depends(require_login)):
    path = os.path.join(db.job_dir(job_id), filename)
    if not os.path.isfile(path):
        raise HTTPException(404, 'File not found.')
    return FileResponse(path, filename=filename, media_type='application/octet-stream')


@app.get('/api/parts/search')
async def parts_search(q: str = Query(..., min_length=1),
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
async def history(search: str = Query(None), limit: int = 50, job_type: str = Query(None),
                   username: str = Depends(require_login)):
    return db.list_jobs(limit=limit, search=search, job_type=job_type)


@app.get('/api/job/{job_id}')
async def job_detail(job_id: str, username: str = Depends(require_login)):
    j = db.get_job(job_id)
    if not j:
        raise HTTPException(404, 'Job not found.')
    return j


@app.post('/api/jobs/combine')
async def combine_jobs(body: JobIdsRequest, username: str = Depends(require_login)):
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
async def delete_jobs(body: JobIdsRequest, username: str = Depends(require_login)):
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
    async def handler(request: Request):
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
