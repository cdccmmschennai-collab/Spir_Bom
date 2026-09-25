// Shared fetch/upload helpers for every page.
//
// The server always answers in JSON, but anything in front of it (nginx,
// a CDN) answers errors with an HTML page -- e.g. "413 Request Entity Too
// Large" or "504 Gateway Time-out". Calling res.json() on that page is what
// produced "Unexpected token '<'". readJson() never does that: it turns
// such responses into a readable message instead.

// Keep in sync with nginx's client_max_body_size on the server (3G).
const MAX_UPLOAD_MB = 3072;

function formatSize(mb) {
  return mb >= 1024 ? `${(mb / 1024).toFixed(1).replace(/\.0$/, '')} GB` : `${Math.round(mb)} MB`;
}

class AuthError extends Error {
  constructor() { super('Your session has expired. Please sign in again.'); this.name = 'AuthError'; }
}

function httpErrorMessage(status) {
  if (status === 413) {
    // Refused by the web server in front of the app (nginx's
    // client_max_body_size), whose limit may be lower than MAX_UPLOAD_MB.
    return 'The server refused this file because it is larger than the web server allows ' +
           '(HTTP 413). Ask the administrator to raise the upload limit (nginx ' +
           `client_max_body_size, expected ${formatSize(MAX_UPLOAD_MB)}), or open the file in Excel, ` +
           'delete embedded pictures/objects and unused sheets, save it again and re-upload.';
  }
  if (status === 502 || status === 503) {
    return 'The server is restarting or temporarily unavailable. Please wait a minute and try again.';
  }
  if (status === 504) return 'The server took too long to respond. Please try again.';
  if (status === 401) return new AuthError().message;
  return `Unexpected response from the server (HTTP ${status}). Please reload the page and try again.`;
}

// Parses a JSON response. For a non-JSON response (a proxy's HTML error
// page) it throws an Error with a readable message rather than a JSON
// syntax error. A JSON error body ({detail: ...}) is returned as-is, so
// callers keep checking res.ok / data.detail exactly as before.
async function readJson(res) {
  const type = res.headers.get('content-type') || '';
  if (type.includes('application/json')) {
    try { return await res.json(); } catch (e) { /* fall through */ }
  }
  const err = new Error(httpErrorMessage(res.status));
  err.status = res.status;
  throw err;
}

function formatElapsed(ms) {
  const s = Math.floor(ms / 1000);
  return s < 60 ? `${s}s` : `${Math.floor(s / 60)}m ${String(s % 60).padStart(2, '0')}s`;
}

const _sleep = ms => new Promise(resolve => setTimeout(resolve, ms));

// POSTs `form` like fetch() does, but reports upload progress (fetch can't),
// which matters for multi-GB SPIRs. Resolves with a standard Response.
function _postWithProgress(url, form, onPercent) {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open('POST', url);
    xhr.upload.onprogress = e => { if (e.lengthComputable && onPercent) onPercent(e.loaded / e.total); };
    xhr.onload = () => resolve(new Response(xhr.responseText, {
      status: xhr.status,
      headers: { 'content-type': xhr.getResponseHeader('content-type') || '' },
    }));
    xhr.onerror = xhr.ontimeout = xhr.onabort = () => reject(new Error('network'));
    xhr.send(form);
  });
}


// While a file is still being sent, leaving the page would cancel the upload
// (the browser holds the data until then), so ask before leaving. Once it's
// uploaded the server owns the job and the user can go anywhere.
let _activeUploads = 0;
window.addEventListener('beforeunload', e => {
  if (_activeUploads > 0) { e.preventDefault(); e.returnValue = ''; }
});

// Uploads one SPIR to `url` (/api/process or /api/process-extraction; add
// ?source=batch etc. so the page can find its jobs again) and returns the
// server's job id as soon as the file is saved there. onProgress(text) gets
// "Uploading 42% of 1.2 GB..." lines.
async function uploadSpirFile(url, file, onProgress) {
  const report = text => { if (onProgress) onProgress(text); };
  if (file.size > MAX_UPLOAD_MB * 1024 * 1024) {
    throw new Error(`${file.name} is ${formatSize(file.size / 1048576)}, larger than the ` +
                    `${formatSize(MAX_UPLOAD_MB)} upload limit. Open it in Excel, delete embedded ` +
                    'pictures/objects and unused sheets, save it again and re-upload.');
  }

  report('Uploading... (stay on this page until the upload finishes)');
  const form = new FormData();
  form.append('file', file);
  const sizeText = formatSize(file.size / 1048576);
  let res;
  _activeUploads++;
  try {
    res = await _postWithProgress(url, form, frac =>
      report(frac < 1 ? `Uploading ${Math.floor(frac * 100)}% of ${sizeText}... (stay on this page until the upload finishes)`
                      : 'Upload complete, saving on server...'));
  } catch (e) {
    throw new Error('Could not reach the server (network problem, or the server is restarting). Please try again.');
  } finally {
    _activeUploads--;
  }
  if (res.status === 401) throw new AuthError();
  const job = await readJson(res);
  if (!res.ok) throw new Error(job.detail || httpErrorMessage(res.status));
  return job.job_id;
}

// One status line for a job as the server reports it (see /api/my-jobs).
function jobStatusText(st) {
  const t = formatElapsed((st.elapsed_seconds || 0) * 1000);
  if (st.status === 'queued') {
    return `Waiting for ${st.queued_ahead ? st.queued_ahead + ' other file(s)' : 'another file'} to finish... ${t}`;
  }
  return `Processing... ${t} (you can leave this page; it keeps running)`;
}

// Waits for a server job to finish. Resolves with its result, or rejects with
// a readable Error (an AuthError when the session has expired). Brief outages
// (a proxy error page, a network blip) are retried.
async function waitForJob(jobId, onProgress) {
  const report = text => { if (onProgress) onProgress(text); };
  let failures = 0;
  for (;;) {
    let st;
    try {
      const r = await fetch(`/api/process-status/${jobId}`, { cache: 'no-store' });
      if (r.status === 401) throw new AuthError();
      st = await readJson(r);
      if (!r.ok) {
        const err = new Error(st.detail || httpErrorMessage(r.status));
        err.final = true;   // e.g. 404: the job is gone after a server restart
        throw err;
      }
      failures = 0;
    } catch (e) {
      if (e instanceof AuthError || e.final) throw e;
      if (++failures >= 60) {
        throw new Error('Lost contact with the server while processing. Check the History page ' +
                        'in a few minutes; if the file is not there, upload it again.');
      }
      report('Processing... (reconnecting to server)');
      await _sleep(2000);
      continue;
    }
    if (st.status === 'done') return st.result;
    if (st.status === 'error') throw new Error(st.error || 'Processing failed.');
    report(jobStatusText(st));
    await _sleep(2000);
  }
}

// Upload + wait, for pages that show one file at a time.
async function processSpirFile(url, file, onProgress) {
  return waitForJob(await uploadSpirFile(url, file, onProgress), onProgress);
}

// This user's running/finished jobs started from one page ('extraction',
// 'batch', 'main'), oldest first -- so a page can show them again after the
// user navigated away and came back.
async function listMyJobs(source) {
  const r = await fetch(`/api/my-jobs?source=${encodeURIComponent(source)}`, { cache: 'no-store' });
  if (r.status === 401) throw new AuthError();
  const data = await readJson(r);
  if (!r.ok) throw new Error(data.detail || httpErrorMessage(r.status));
  return data;
}

// Removes finished jobs from a page's list (they stay in History).
async function dismissJobs(jobIds) {
  if (!jobIds.length) return;
  await fetch('/api/my-jobs/dismiss', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ job_ids: jobIds }),
  }).catch(() => {});
}
