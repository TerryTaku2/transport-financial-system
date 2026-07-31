#!/usr/bin/env python3
"""Register a new spoke site (a local-server PC) and print its API key once.

Run this against the HUB's database (e.g. via Render's shell, or locally
against the same DATABASE_URL Render uses) whenever a new local-server PC
needs to sync. The plaintext key is never stored anywhere — only its hash
— so copy it into that site's .env as SYNC_API_KEY immediately; it can't
be recovered later, only revoked and reissued.

Examples:
    python provision_sync_site.py site-nairobi-01
    python provision_sync_site.py site-nairobi-01 "Nairobi Depot"
"""

import argparse
import secrets
import sys

from werkzeug.security import generate_password_hash

from app import app, db, SyncSite


def provision_site(site_id: str, display_name: str | None) -> str:
    with app.app_context():
        if SyncSite.query.filter_by(site_id=site_id).first():
            raise ValueError(f'Site "{site_id}" is already registered.')

        api_key = secrets.token_urlsafe(32)
        site = SyncSite(
            site_id=site_id,
            api_key_hash=generate_password_hash(api_key),
            display_name=display_name or site_id,
            is_active=True,
        )
        db.session.add(site)
        db.session.commit()
        return api_key


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Register a new spoke sync site')
    parser.add_argument('site_id', help='Unique short id for the site, e.g. site-nairobi-01')
    parser.add_argument('display_name', nargs='?', default=None, help='Human-readable name (optional)')
    args = parser.parse_args()

    try:
        key = provision_site(args.site_id, args.display_name)
    except Exception as exc:  # pragma: no cover - CLI error handling
        print(f'Error: {exc}', file=sys.stderr)
        sys.exit(1)

    print(f'Site "{args.site_id}" registered.')
    print(f'SYNC_API_KEY={key}')
    print('Copy this into that site\'s .env now — it will not be shown again.')
