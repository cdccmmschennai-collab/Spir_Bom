# Deploying the BOM Tool on a Hostinger VPS (Docker + HTTPS)

Users open the tool at an HTTPS link. It runs as two containers:

- **app**: the BOM Tool (FastAPI, Python 3.14). Not reachable from outside.
- **caddy**: the public web server on ports 80/443. It gets a free HTTPS
  certificate automatically, renews it, and forwards requests to the app.

Everything the tool stores (database, uploaded SPIRs, generated Extraction and
OUTPUT files, logs) is in the `data/` folder on the VPS disk. Rebuilding or
updating the containers never deletes it.

---

## 1. Create the VPS (Hostinger hPanel)

1. Buy a **VPS** plan. KVM 2 (2 vCPU, 8 GB RAM, 100 GB disk) is comfortable. The
   current `data/` folder is about 5 GB and grows with every upload.
2. OS: **Ubuntu 24.04** (plain, or Hostinger's "Ubuntu with Docker" template).
3. Note the VPS **IP address** and the **root password / SSH key**.
4. In hPanel → VPS → **Firewall**: allow inbound **22, 80, 443** (TCP), and
   **443 UDP** if offered.

## 2. Install Docker on the VPS

From your PC: `ssh root@<VPS-IP>`, then (skip if you chose the Docker template):

```bash
curl -fsSL https://get.docker.com | sh
docker --version && docker compose version

# OS firewall too (same ports as hPanel)
ufw allow 22/tcp && ufw allow 80/tcp && ufw allow 443/tcp && ufw allow 443/udp
ufw --force enable
```

## 3. Copy the project to the VPS

On **your Windows PC**, in PowerShell, from the `bom-tool` folder:

```powershell
# Package the code (no venv, no data)
tar -czf bom-tool.tgz --exclude=venv --exclude=data --exclude=__pycache__ --exclude=.claude .
scp bom-tool.tgz root@<VPS-IP>:/opt/
```

On the **VPS**:

```bash
mkdir -p /opt/bom-tool && tar -xzf /opt/bom-tool.tgz -C /opt/bom-tool && rm /opt/bom-tool.tgz
cd /opt/bom-tool
```

### Optional: bring your existing data (users, history, Parts Master records)

Stop the local server first, so the database isn't being written while you copy
it. Then, on your PC:

```powershell
tar -czf bom-data.tgz data
scp bom-data.tgz root@<VPS-IP>:/opt/bom-tool/
```

On the VPS:

```bash
cd /opt/bom-tool && tar -xzf bom-data.tgz && rm bom-data.tgz
```

Skip this to start empty. Then create the users as in step 6.

### Either way, give the app container ownership of `data/`

The app runs as a non-root user (uid 1000):

```bash
mkdir -p /opt/bom-tool/data && chown -R 1000:1000 /opt/bom-tool/data
```

## 4. Set the address and start

```bash
cd /opt/bom-tool
cp .env.example .env
nano .env
```

Set `SITE_ADDRESS`:

- **No domain yet:** use the VPS IP with dashes + `.sslip.io`. For IP
  `72.60.1.23`, that's `SITE_ADDRESS=72-60-1-23.sslip.io`. sslip.io is a free
  public DNS service that points that name at your IP, so a real, trusted
  HTTPS certificate works without buying a domain.
- **With a domain:** see step 7.

Start it:

```bash
docker compose up -d --build
docker compose ps           # both containers "running", app "healthy"
docker compose logs -f      # Ctrl+C to stop watching
```

Open **`https://<SITE_ADDRESS>`** in a browser. The login page appears with a valid
padlock. The first visit can take up to a minute while the certificate is issued.

## 5. Everyday commands (run in `/opt/bom-tool`)

| Task | Command |
|---|---|
| See status | `docker compose ps` |
| App logs | `docker compose logs -f app` |
| Restart | `docker compose restart` |
| Stop / start | `docker compose down` / `docker compose up -d` |
| Update after code changes | copy the new code as in step 3 (code only), then `docker compose up -d --build` |

The containers restart automatically after a crash or a VPS reboot.

## 6. Managing users

```bash
docker compose exec app python manage_users.py list
docker compose exec app python manage_users.py add <username> <password>
docker compose exec app python manage_users.py remove <username>
```

To import a roster file, copy it into `data/` first, then:

```bash
docker compose exec app python manage_users.py import-members data/roster.csv
```

## 7. Adding a domain later

1. Buy the domain (Hostinger or anywhere). In its DNS settings, add an **A record**,
   e.g. `bom` → `<VPS-IP>`.
2. Wait until `ping bom.yourcompany.com` shows the VPS IP.
3. On the VPS, set `SITE_ADDRESS=bom.yourcompany.com` in `.env`, then run
   `docker compose up -d`.

Caddy gets the new certificate automatically. Users switch to the new link; data and
logins are unchanged. The login cookie is tied to the address, so users sign in once
more on the new link.

## 8. Daily backup

`data/` holds everything. This makes a consistent copy of the SQLite database
(safe while the app is running) and archives the whole folder every night at 02:00,
keeping the last 7:

```bash
mkdir -p /opt/bom-backups
cat > /opt/bom-backup.sh <<'EOF'
#!/bin/sh
set -e
cd /opt/bom-tool
docker compose exec -T app python -c "import sqlite3; s=sqlite3.connect('data/bom_tool.db'); d=sqlite3.connect('data/bom_tool.backup.db'); s.backup(d); d.close()"
tar -czf /opt/bom-backups/bom-data-$(date +%F).tgz --exclude=data/bom_tool.db data
ls -1t /opt/bom-backups/bom-data-*.tgz | tail -n +8 | xargs -r rm
EOF
chmod +x /opt/bom-backup.sh
(crontab -l 2>/dev/null; echo "0 2 * * * /opt/bom-backup.sh") | crontab -
```

Keep a copy somewhere other than the VPS too, for example by downloading one now and
then with `scp root@<VPS-IP>:/opt/bom-backups/<file> .`.

**To restore:**
1. `docker compose down`
2. Extract the archive into `/opt/bom-tool`.
3. Rename `data/bom_tool.backup.db` to `data/bom_tool.db`.
4. `chown -R 1000:1000 data`
5. `docker compose up -d`

## Notes

- The VPS needs outbound internet access, which Hostinger allows by default. The tool
  fetches currency rates (exchangerate-api.com) and looks up manufacturer countries
  online.
- The app runs as **one** process on purpose: SQLite and the per-job files expect a
  single writer. That comfortably handles a team uploading and searching SPIRs.
- Running locally without Docker still works exactly as before (see README.md).
