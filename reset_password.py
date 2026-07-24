#!/usr/bin/env python3
"""Reset a user's password in the deployed Transport ERP database.

Examples:
    python reset_password.py MyNewPassword123! 
    python reset_password.py --username admin MyNewPassword123!
    NEW_PASSWORD=MyNewPassword123! python reset_password.py
"""

import argparse
import os
import sys

from app import app, db, User


def reset_password(username: str, new_password: str) -> None:
    if not new_password or len(new_password) < 6:
        raise ValueError('Password must be at least 6 characters long.')

    with app.app_context():
        user = User.query.filter_by(username=username).first()
        if not user:
            raise ValueError(f'User "{username}" was not found.')

        user.set_password(new_password)
        db.session.commit()
        print(f'Password updated successfully for user: {username}')
        print(f'New password: {new_password}')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Reset a user password')
    parser.add_argument('password', nargs='?', default=os.environ.get('NEW_PASSWORD'), help='New password to set')
    parser.add_argument('--username', default='admin', help='Username to update (default: admin)')
    args = parser.parse_args()

    try:
        reset_password(args.username, args.password)
    except Exception as exc:  # pragma: no cover - CLI error handling
        print(f'Error: {exc}', file=sys.stderr)
        sys.exit(1)
