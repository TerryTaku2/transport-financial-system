#!/usr/bin/env python3
"""One-off script: reset the 'admin' user's password. Run once, then delete."""
import secrets

from app import app, db, User

NEW_PASSWORD = secrets.token_urlsafe(12)

with app.app_context():
    admin = User.query.filter_by(username='admin').first()
    if not admin:
        raise SystemExit("No user with username 'admin' found.")
    admin.set_password(NEW_PASSWORD)
    db.session.commit()
    print(f"Password reset for '{admin.username}'. New password: {NEW_PASSWORD}")
    print("Log in and change this password immediately.")
