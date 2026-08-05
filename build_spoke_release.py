#!/usr/bin/env python3
"""Builds a spoke .exe and packages it as a release .zip ready to upload
on the hub's Spoke Releases page (see SPOKE_SETUP.md, "Publishing an
update").

Usage:
    venv\\Scripts\\python.exe build_spoke_release.py 2026.08.05

Steps:
  1. Writes the given version string to VERSION (read by the running app
     as APP_VERSION, and compared against the hub's published version by
     check_for_spoke_update).
  2. Runs PyInstaller against spoke_build.spec.
  3. Copies VERSION and the launcher scripts into dist\\TransportERP\\ —
     these live at the top level next to TransportERP.exe, not inside
     _internal\\, since check_for_spoke_update reads VERSION from
     os.path.dirname(sys.executable), and the launcher has to run before
     the app (and _internal\\) even exists.
  4. Zips dist\\TransportERP\\ into <version>.zip in this repo's root.

The printed .zip is what gets uploaded as-is on Admin -> Spoke Releases —
the version string there must match what was passed here.
"""
import hashlib
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent


def main():
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(1)
    version = sys.argv[1]

    (REPO_ROOT / 'VERSION').write_text(version, encoding='utf-8')
    print(f'Wrote VERSION = {version}')

    python = REPO_ROOT / 'venv' / 'Scripts' / 'python.exe'
    if not python.exists():
        python = Path(sys.executable)
    subprocess.run([str(python), '-m', 'PyInstaller', 'spoke_build.spec', '--noconfirm'],
                   cwd=REPO_ROOT, check=True)

    dist_dir = REPO_ROOT / 'dist' / 'TransportERP'
    for name in ('VERSION', 'launcher.ps1', 'launcher.bat'):
        shutil.copy2(REPO_ROOT / name, dist_dir / name)
    print(f'Copied VERSION + launcher scripts into {dist_dir}')

    zip_path = REPO_ROOT / f'{version}.zip'
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for path in dist_dir.rglob('*'):
            if path.is_file():
                zf.write(path, path.relative_to(dist_dir))

    sha256 = hashlib.sha256(zip_path.read_bytes()).hexdigest()
    size_mb = zip_path.stat().st_size / 1_048_576
    print()
    print(f'Release package: {zip_path}  ({size_mb:.1f} MB)')
    print(f'sha256: {sha256}')
    print()
    print(f'Next: on the hub, go to Admin -> Spoke Releases and publish version "{version}" with this .zip.')
    print('(The sha256 above is just for your own verification — the hub recomputes it on upload.)')


if __name__ == '__main__':
    main()
