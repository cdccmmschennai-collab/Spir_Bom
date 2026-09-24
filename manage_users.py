"""
Manage BOM Tool login accounts (there's no self-signup -- this is the only
way to add/remove one).

Usage:
    python manage_users.py add <username> <password>
    python manage_users.py remove <username>
    python manage_users.py list
    python manage_users.py import-members <roster.txt or roster.csv> [initial_password]

import-members accepts either a plain text file, one full name per line
(blank lines and lines starting with '#' are skipped):

    # group roster
    John Smith
    Jane Doe

...or a CSV with a NAME column (any other columns, like an S.no index, are
ignored):

    S.no,NAME
    1,John Smith
    2,Jane Doe

For each name, generates a username (lowercase, letters+digits only, no
separator -- 'John Smith' -> 'johnsmith'; a colliding name gets 'johnsmith2',
etc.) and creates the account with the given initial_password (or
DEFAULT_INITIAL_PASSWORD if omitted). That password works only until the
member's first successful login, at which point the system generates a new
password unique to them and shows it once -- see engine/auth.verify_login.
Re-running this against the same file is safe: a full name already in the
database is skipped, so nobody's already-rotated password gets reset.
"""
import sys
import os
import csv

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from engine import auth

DEFAULT_INITIAL_PASSWORD = 'CDC@2026'


def _existing_full_names():
    return {u['full_name'] for u in auth.list_users() if u['full_name']}


def _read_names(path: str):
    """Plain text (one name per line, '#' comments skipped) or a CSV with
    a NAME column (any other columns, e.g. an S.no index, are ignored) --
    detected by whether the first line contains a comma and the word
    'NAME', never by file extension alone (a .csv someone renamed to .txt,
    or vice versa, still parses correctly)."""
    with open(path, encoding='utf-8-sig') as f:
        lines = f.read().splitlines()
    if lines and ',' in lines[0] and 'NAME' in lines[0].upper():
        reader = csv.DictReader(lines)
        name_field = next((fn for fn in (reader.fieldnames or [])
                            if fn and fn.strip().upper() == 'NAME'), None)
        if name_field:
            return [row[name_field].strip() for row in reader if (row.get(name_field) or '').strip()]
    return [ln.strip() for ln in lines if ln.strip() and not ln.strip().startswith('#')]


def import_members(path: str, initial_password: str):
    if not os.path.isfile(path):
        print(f"File not found: {path}")
        return
    already = _existing_full_names()
    created = []
    skipped = []
    for name in _read_names(path):
        if name in already:
            skipped.append(name)
            continue
        username = auth.create_member(name, initial_password)
        created.append((name, username))
        already.add(name)

    if created:
        print(f"Created {len(created)} account(s):")
        for name, username in created:
            print(f"  {name:30s} -> {username}")
    if skipped:
        print(f"Skipped {len(skipped)} already-existing name(s): {', '.join(skipped)}")
    if not created and not skipped:
        print('No names found in file.')
    if created:
        print(f"\nShared initial password for new accounts: {initial_password}")
        print("Each member's first login replaces it with a password only they know.")


def main():
    auth.init_auth_db()
    args = sys.argv[1:]

    if len(args) == 3 and args[0] == 'add':
        _, username, password = args
        auth.create_user(username, password)
        print(f"User '{username}' created/updated.")
    elif len(args) == 2 and args[0] == 'remove':
        auth.remove_user(args[1])
        print(f"User '{args[1]}' removed.")
    elif len(args) == 1 and args[0] == 'list':
        users = auth.list_users()
        if not users:
            print('No users yet.')
        for u in users:
            status = 'activated' if u['activated'] else 'pending first login'
            name = f" ({u['full_name']})" if u['full_name'] else ''
            print(f"{u['username']}{name}  [{status}]  created {u['created_at']}")
    elif len(args) in (2, 3) and args[0] == 'import-members':
        path = args[1]
        initial_password = args[2] if len(args) == 3 else DEFAULT_INITIAL_PASSWORD
        import_members(path, initial_password)
    else:
        print(__doc__)


if __name__ == '__main__':
    main()
