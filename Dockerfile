# BOM Tool web app (FastAPI). Runs behind Caddy for HTTPS -- see
# docker-compose.yml and DEPLOY.md.
FROM python:3.14-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

WORKDIR /app

COPY requirements.txt .
RUN pip install -r requirements.txt

COPY app/ app/
COPY engine/ engine/
COPY Reference.xlsx manage_users.py ./

# Non-root user. data/ (database, uploads, results, logs) is a volume
# mounted over /app/data at runtime, so it survives rebuilds.
RUN useradd --create-home --uid 1000 bom \
    && mkdir -p /app/data \
    && chown -R bom:bom /app
USER bom

EXPOSE 8001

HEALTHCHECK --interval=30s --timeout=5s --start-period=60s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8001/login.html', timeout=4)"

# One worker only: SQLite and the per-job files under data/ expect a
# single process. --proxy-headers trusts Caddy's X-Forwarded-* headers.
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8001", "--proxy-headers", "--forwarded-allow-ips", "*"]
