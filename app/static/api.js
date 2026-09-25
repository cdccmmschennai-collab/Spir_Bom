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
    return `The file is larger than the server's upload limit (${formatSize(MAX_UPLOAD_MB)}). Open it in Excel, ` +
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

// Uploads one SPIR to `url` (/api/process or /api/process-extraction) and
// waits for the server's background job to finish. onProgress(text) gets a
// short status line ("Uploading...", "Processing... 1m 05s") to display.
// Resolves with the job's result; rejects with a readable Error (or an
// AuthError when the session has expired).
async function processSpirFile(url, file, onProgress) {
  const report = text => { if (onProgress) onProgress(text); };
  if (file.size > MAX_UPLOAD_MB * 1024 * 1024) {
    throw new Error(`${file.name} is ${formatSize(file.size / 1048576)}, larger than the ` +
                    `${formatSize(MAX_UPLOAD_MB)} upload limit. Open it in Excel, delete embedded ` +
                    'pictures/objects and unused sheets, save it again and re-upload.');
  }

  report('Uploading...');
  const form = new FormData();
  form.append('file', file);
  const sizeText = formatSize(file.size / 1048576);
  let res;
  try {
    res = await _postWithProgress(url, form, frac =>
      report(frac < 1 ? `Uploading ${Math.floor(frac * 100)}% of ${sizeText}...`
                      : 'Upload complete, saving on server...'));
  } catch (e) {
    throw new Error('Could not reach the server (network problem, or the server is restarting). Please try again.');
  }
  if (res.status === 401) throw new AuthError();
  const job = await readJson(res);
  if (!res.ok) throw new Error(job.detail || httpErrorMessage(res.status));

  // Poll the background job. Brief outages (a proxy error page, a network
  // blip) are retried; only a definite answer ends the wait.
  const started = Date.now();
  let failures = 0;
  for (;;) {
    await _sleep(2000);
    let st;
    try {
      const r = await fetch(`/api/process-status/${job.job_id}`, { cache: 'no-store' });
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
      report(`Processing... ${formatElapsed(Date.now() - started)} (reconnecting to server)`);
      continue;
    }
    if (st.status === 'done') return st.result;
    if (st.status === 'error') throw new Error(st.error || 'Processing failed.');
    report(st.status === 'queued'
      ? `Waiting for ${st.queued_ahead ? st.queued_ahead + ' other file(s)' : 'another file'} to finish... ${formatElapsed(Date.now() - started)}`
      : `Processing... ${formatElapsed(Date.now() - started)}`);
  }
}
