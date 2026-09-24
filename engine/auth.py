"""
Username/password auth for the BOM Tool, backed by the same SQLite database
as job history (see engine/db.py). Passwords are stored salted + hashed
(PBKDF2-HMAC-SHA256, no extra dependency needed) -- never in plain text.
Sessions are random opaque tokens handed out as an httponly cookie and
checked against a sessions table -- no JWT/expiry-claim parsing to get
wrong.

Accounts are managed with manage_users.py, not through the web UI itself --
this is a small internal tool, not a public signup product.
"""
import re
import hashlib
import hmac
import secrets
import datetime

from .db import get_conn

SESSION_COOKIE = 'bom_session'
SESSION_TTL_DAYS = 14
_PBKDF2_ITERATIONS = 200_000
# No 0/O/1/l/I -- a generated password a person has to retype from memory
# or a screenshot shouldn't hinge on telling those apart.
_PASSWORD_ALPHABET = 'ABCDEFGHJKMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789'


def init_auth_db():
    with get_conn() as conn:
        conn.execute('''
            CREATE TABLE IF NOT EXISTS users (
                username        TEXT PRIMARY KEY,
                salt            TEXT NOT NULL,
                password_hash   TEXT NOT NULL,
                created_at      TEXT NOT NULL
            )
        ''')
        conn.execute('''
            CREATE TABLE IF NOT EXISTS sessions (
                token           TEXT PRIMARY KEY,
                username        TEXT NOT NULL,
                created_at      TEXT NOT NULL,
                expires_at      TEXT NOT NULL
            )
        ''')
        # Added for the 41-member bulk-import flow: full_name is the
        # source name a username was generated from; activated=0 means
        # the row's password_hash is still the shared bootstrap password
        # (see create_member/verify_login) rather than one only this
        # person knows. Guarded ALTER so this is a no-op on a DB that's
        # already been migrated.
        existing_cols = {row['name'] for row in conn.execute('PRAGMA table_info(users)')}
        if 'full_name' not in existing_cols:
            conn.execute('ALTER TABLE users ADD COLUMN full_name TEXT')
        if 'activated' not in existing_cols:
            conn.execute('ALTER TABLE users ADD COLUMN activated INTEGER NOT NULL DEFAULT 1')
        if 'activated_at' not in existing_cols:
            conn.execute('ALTER TABLE users ADD COLUMN activated_at TEXT')
        # Settings page's Profile card: last successful login, updated on
        # every verify_login() success (not just the first-login rotation).
        if 'last_login' not in existing_cols:
            conn.execute('ALTER TABLE users ADD COLUMN last_login TEXT')


def _hash_password(password: str, salt: str) -> str:
    return hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt.encode('utf-8'),
                                _PBKDF2_ITERATIONS).hex()


def create_user(username: str, password: str):
    username = username.strip()
    salt = secrets.token_hex(16)
    password_hash = _hash_password(password, salt)
    with get_conn() as conn:
        conn.execute('''
            INSERT INTO users (username, salt, password_hash, created_at) VALUES (?, ?, ?, ?)
            ON CONFLICT(username) DO UPDATE SET salt = excluded.salt, password_hash = excluded.password_hash
        ''', (username, salt, password_hash, datetime.datetime.utcnow().isoformat()))


def remove_user(username: str):
    with get_conn() as conn:
        conn.execute('DELETE FROM users WHERE username = ?', (username.strip(),))


def list_users():
    with get_conn() as conn:
        rows = conn.execute(
            'SELECT username, full_name, activated, created_at FROM users ORDER BY username'
        ).fetchall()
    return [dict(r) for r in rows]


def get_user(username: str):
    """Settings page's Profile card data -- returns None if the account
    doesn't exist (shouldn't happen for an already-authenticated caller,
    but callers should still handle it rather than assume)."""
    with get_conn() as conn:
        row = conn.execute(
            'SELECT username, full_name, created_at, last_login FROM users WHERE username = ?',
            (username.strip(),)
        ).fetchone()
    return dict(row) if row else None


def change_password(username: str, current_password: str, new_password: str) -> bool:
    """Settings page's 'Update Password'. Returns False (and changes
    nothing) if `current_password` doesn't match this account's real
    current password -- never trusts the caller's claimed identity alone."""
    with get_conn() as conn:
        row = conn.execute('SELECT salt, password_hash FROM users WHERE username = ?',
                            (username.strip(),)).fetchone()
        if not row:
            return False
        if not hmac.compare_digest(_hash_password(current_password, row['salt']), row['password_hash']):
            return False
        new_salt = secrets.token_hex(16)
        new_hash = _hash_password(new_password, new_salt)
        conn.execute('UPDATE users SET salt = ?, password_hash = ? WHERE username = ?',
                      (new_salt, new_hash, username.strip()))
    return True


def verify_user(username: str, password: str) -> bool:
    with get_conn() as conn:
        row = conn.execute('SELECT salt, password_hash FROM users WHERE username = ?',
                            (username.strip(),)).fetchone()
    if not row:
        # Still hash something so a login attempt against a nonexistent
        # username takes about as long as a real one -- avoids leaking
        # which usernames exist via response-time differences.
        _hash_password(password, secrets.token_hex(16))
        return False
    return hmac.compare_digest(_hash_password(password, row['salt']), row['password_hash'])


def generate_password(length: int = 10) -> str:
    """A random password for a member's first login to replace the shared
    bootstrap one with -- drawn from _PASSWORD_ALPHABET (no visually
    ambiguous characters)."""
    return ''.join(secrets.choice(_PASSWORD_ALPHABET) for _ in range(length))


def slugify_username(full_name: str, conn=None) -> str:
    """'John Smith' -> 'johnsmith': lowercase, letters+digits only, no
    separator (the format chosen for the 41-member roster). Collision-safe
    -- if that slug is already taken by a DIFFERENT full name, appends 2,
    3, ... until free, so two people who'd otherwise slugify to the same
    username both still get a working, unique account."""
    base = re.sub(r'[^a-z0-9]', '', full_name.strip().lower()) or 'member'

    def _taken(conn, candidate):
        row = conn.execute('SELECT full_name FROM users WHERE username = ?', (candidate,)).fetchone()
        return row is not None and row['full_name'] != full_name.strip()

    if conn is not None:
        candidate, n = base, 2
        while _taken(conn, candidate):
            candidate = f'{base}{n}'
            n += 1
        return candidate
    with get_conn() as c:
        candidate, n = base, 2
        while _taken(c, candidate):
            candidate = f'{base}{n}'
            n += 1
        return candidate


def normalize_username_input(s: str) -> str:
    """Applies the exact same normalization used to generate a username
    (see slugify_username) to whatever a member types at login, so 'John
    Smith', 'JOHNSMITH', 'john  smith', 'John-Smith' all resolve to the
    one stored username 'johnsmith' -- forgiving of case/spacing/
    punctuation without needing fuzzy matching."""
    return re.sub(r'[^a-z0-9]', '', (s or '').strip().lower())


def create_member(full_name: str, initial_password: str) -> str:
    """Bulk-import path (see manage_users.py import-members): creates an
    UNACTIVATED account -- its password_hash is the shared bootstrap
    password everyone starts with, not one unique to this person yet.
    verify_login() rotates it to a real unique password on first
    successful login. Returns the generated username."""
    full_name = full_name.strip()
    with get_conn() as conn:
        username = slugify_username(full_name, conn=conn)
        salt = secrets.token_hex(16)
        password_hash = _hash_password(initial_password, salt)
        conn.execute('''
            INSERT INTO users (username, salt, password_hash, created_at, full_name, activated)
            VALUES (?, ?, ?, ?, ?, 0)
        ''', (username, salt, password_hash, datetime.datetime.utcnow().isoformat(), full_name))
    return username


def verify_login(username_input: str, password: str):
    """The one login check used for every account, whichever way it was
    created. Returns None on any failure (bad username or bad password --
    never distinguished, to avoid leaking which usernames exist).
    On success:
      - an already-activated account (the normal case, and every account
        made with create_user()) -> {'username', 'first_login': False}
      - an unactivated bulk-imported account, on its first successful
        login using the shared bootstrap password -> generates a brand
        new password unique to this person, stores it (replacing the
        shared one -- other members' rows are untouched), marks the
        account activated, and returns {'username', 'first_login': True,
        'new_password': <shown once, never stored in plain text>}.
    """
    key = normalize_username_input(username_input)
    with get_conn() as conn:
        row = conn.execute('SELECT username, salt, password_hash, activated FROM users WHERE username = ?',
                            (key,)).fetchone()
        if not row:
            _hash_password(password, secrets.token_hex(16))
            return None
        if not hmac.compare_digest(_hash_password(password, row['salt']), row['password_hash']):
            return None

        now = datetime.datetime.utcnow().isoformat()
        if not row['activated']:
            new_password = generate_password()
            new_salt = secrets.token_hex(16)
            new_hash = _hash_password(new_password, new_salt)
            conn.execute('''
                UPDATE users SET salt = ?, password_hash = ?, activated = 1, activated_at = ?, last_login = ?
                WHERE username = ?
            ''', (new_salt, new_hash, now, now, row['username']))
            return {'username': row['username'], 'first_login': True, 'new_password': new_password}

        conn.execute('UPDATE users SET last_login = ? WHERE username = ?', (now, row['username']))
        return {'username': row['username'], 'first_login': False}


def create_session(username: str) -> str:
    token = secrets.token_urlsafe(32)
    now = datetime.datetime.utcnow()
    expires = now + datetime.timedelta(days=SESSION_TTL_DAYS)
    with get_conn() as conn:
        conn.execute('INSERT INTO sessions (token, username, created_at, expires_at) VALUES (?, ?, ?, ?)',
                      (token, username, now.isoformat(), expires.isoformat()))
    return token


def get_session_user(token: str):
    if not token:
        return None
    with get_conn() as conn:
        row = conn.execute('SELECT username, expires_at FROM sessions WHERE token = ?', (token,)).fetchone()
    if not row:
        return None
    if datetime.datetime.fromisoformat(row['expires_at']) < datetime.datetime.utcnow():
        delete_session(token)
        return None
    return row['username']


def delete_session(token: str):
    with get_conn() as conn:
        conn.execute('DELETE FROM sessions WHERE token = ?', (token,))
