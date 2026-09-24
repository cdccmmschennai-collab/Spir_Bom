# SPIR BOM Tool

Upload a SPIR (`.xlsx` / `.xlsm`) and get back two files, built with the exact
rules worked out for the ENMC / MEWTP SPIRs:

1. **Extraction file** — flat, SAP-field-coded row-per-part export, tag-major
   grouped, position numbers restarting per equipment.
2. **OUTPUT file** — the 40-column SAP material-master upload structure
   (Equipment/BOM-header + spare rows), with the 40000/500000 Material Temp
   Number series, Mfr-Part-Number dedup, duplicate-value highlighting, and
   live QAR conversion.

## Project layout

```
bom-tool/
  engine/
    parser.py       # reads any SPIR file (auto-detects 1 or many sheets)
    rules.py         # the shared business rules (SPF codes, numbering, dedup)
    fx.py             # live currency -> QAR lookup, with a peg fallback
    extraction.py    # writes the Extraction-format file
    sap_output.py    # writes the SAP OUTPUT-format file
    db.py             # SQLite metadata store (every job, input + output files)
  app/
    main.py          # FastAPI backend (upload -> process -> download -> history)
    static/index.html  # single-page upload UI + history table
  data/               # created on first run: bom_tool.db + files/<job_id>/...
  requirements.txt
```

## Run it locally

```bash
pip install -r requirements.txt
uvicorn app.main:app --host 0.0.0.0 --port 8001
```

Then open **http://localhost:8001** in a browser, drop in a SPIR file, and
download the two results (or a zip of both).

## Storage & history

Every upload is kept, not just the latest one:

- **Files**: the original input SPIR, both generated outputs, and the zip are
  all saved under `data/files/<job_id>/` — this folder is what you'd back up
  or move to shared storage.
- **Metadata**: `data/bom_tool.db` (SQLite) records every job — SPIR number,
  sheets, tags, timestamp, and whether it succeeded or errored (with the
  error message, if it failed).
- The **History** table on the web page lists every past run and lets you
  re-download any file from it, including after the app has been restarted.
- `GET /api/history?search=...` filters by SPIR number or filename;
  `GET /api/job/{job_id}` returns full detail for one run.

If this later needs to be shared across a team rather than run locally, the
`data/` folder is the one thing that needs to live on shared/persistent
storage (a mounted volume, network share, or swapping `db.py` to point at a
real Postgres instance instead of SQLite — the rest of the app doesn't care
which database is behind `engine/db.py`).

## Extending the rules

All of the business logic lives in `engine/rules.py` and inside
`engine/extraction.py` / `engine/sap_output.py` — nothing is hardcoded to a
specific SPIR number or project. If a future SPIR needs a different rule
(e.g. a different numbering start point, a different SPF code format), that's
a one-line change in `rules.py`, not a rewrite.

## Known assumptions worth double-checking on new SPIRs

- **Row/column layout**: assumes the standard SPIR template — tag names in
  row 1 (cols C:F), model/serial/qty in rows 4/6/7, item rows starting at
  row 8, SPIR number in `Y1`, equipment description in `X2`, manufacturer in
  `Y3`, supplier in `W4`, vendor contact block in `K33`.
- **Material Temp Numbers** (40000s / 500000s series) are internal
  placeholders, not real SAP numbers — they still need real assignment by
  QatarEnergy's SAP master-data team before upload.
- **FX conversion** uses a free public rate API with a same-session cache;
  swap the URL in `engine/fx.py` if your org has a preferred FX data source.

## Deploying beyond your own machine

For users to open it at an HTTPS link, run it with Docker on a server:
`Dockerfile` + `docker-compose.yml` start the app behind Caddy, which handles
HTTPS certificates automatically. **[DEPLOY.md](DEPLOY.md)** is the
step-by-step guide for a Hostinger VPS, including working without a domain
yet, adding one later, managing users and daily backups.
