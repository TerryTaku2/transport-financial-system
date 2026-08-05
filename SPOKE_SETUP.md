# Setting up a spoke PC (local server at a site)

This turns a Windows PC at a site into a full local copy of the system that
works offline and syncs with the central server (the Render deployment,
"the hub") whenever it has internet. No Python install needed — this uses
the packaged `TransportERP.exe`.

## 1. Register the site on the hub

From a browser, log in to the hub as an admin and go to **Admin → Sync
Sites → Register New Site**. Pick a `site_id` (lowercase, hyphens, e.g.
`site-nairobi-01`) and a display name, then submit.

You'll land on a page showing a `.env` block with a generated API key.
**Copy it now** — it's shown once and only its hash is stored; if you
navigate away without copying it, revoke that site entry and register a
new one.

## 2. Get the app onto the spoke PC

Copy the whole `dist\TransportERP` folder (built from `spoke_build.spec`)
onto the spoke PC — a USB drive or network share is fine. Put it somewhere
permanent, e.g. `C:\TransportERP\`. The folder contains `TransportERP.exe`,
an `_internal` folder it depends on, `launcher.bat`/`launcher.ps1`, and a
`VERSION` file — keep them all together.

To rebuild this folder yourself (e.g. after an app update), on the dev
machine:
```
venv\Scripts\python.exe build_spoke_release.py <version>
```
Output lands in `dist\TransportERP\` (this also produces a `<version>.zip`
ready to publish — see "Keeping the spoke up to date automatically"
below, the usual way to get an update onto an already-deployed spoke
without touching it by hand).

## 3. Configure the spoke

In `C:\TransportERP\`, create a file named `.env` (no filename before the
dot) containing:

```
FLASK_ENV=production
SECRET_KEY=<generate a random 32+ character string, unique per site>
ADMIN_PASSWORD=<a password for this site's first login>

SITE_ID=site-nairobi-01
SYNC_ENABLED=true
SYNC_HUB_URL=https://transport-financial-system.onrender.com
SYNC_API_KEY=<the key from step 1>
SYNC_INTERVAL_SECONDS=60
```

`FLASK_ENV=production` matters, not just cosmetically — in development
mode Flask's auto-reloader forks a worker process that survives killing
the main one, leaving an orphaned server behind on every restart.

Do **not** set `DATABASE_URL`. Left unset, the app stores its SQLite
database (`transport_erp.db`) right next to `TransportERP.exe`, which is
what makes data survive a restart. Setting `DATABASE_URL` yourself only
makes sense if you want the file somewhere else — use an absolute path if
so (e.g. `sqlite:///C:/TransportERP/data/transport_erp.db`), never a bare
relative one.

## 4. First run

Double-click `launcher.bat` (not `TransportERP.exe` directly — see "Keeping
the spoke up to date automatically" below for why), or from a terminal:
```
C:\TransportERP\launcher.bat
```
First launch creates the database and prints the admin login (username
`admin`, the password from `ADMIN_PASSWORD` above). Open
`http://localhost:5000` in a browser on that PC, log in, and change the
password.

That's localhost-only by default in production mode — fine if only this
one PC uses it. If other devices at the site (other PCs, phones on the
same LAN) need to reach it too, add `HOST=0.0.0.0` to `.env` and connect
via this PC's LAN IP instead. Do that over plain HTTP only on a trusted
local network; for anything wider, put it behind HTTPS the same way the
dev machine's LAN setup already does (Caddy + mkcert) rather than
exposing the raw Flask server.

Leave it running for a minute, then check the hub's **Admin → Sync
Health** page — this site should show as **Online** with a recent last
pull/push time.

## 5. Keep it running unattended (service-wrap)

A double-clicked `.exe` stops when the console window closes or the PC
reboots. For a real deployment, wrap it as a Windows service with
[NSSM](https://nssm.cc/) so it survives both:

```
nssm install TransportERP "C:\TransportERP\launcher.bat"
nssm set TransportERP AppDirectory "C:\TransportERP"
nssm start TransportERP
```

Point NSSM at `launcher.bat`, not `TransportERP.exe` — see the next
section for why. `AppDirectory` matters either way — it's what makes the
service load `.env` and find `transport_erp.db` in the right place, same
as running it manually from that folder.

## 6. Keeping the spoke up to date automatically

Once registered and synced, a spoke checks the hub every few hours for a
newer published build (see `SPOKE_UPDATE_CHECK_SECONDS` in `.env`,
default 6 hours) and, if one exists, downloads and unpacks it in the
background into `_update_staged\` next to the install — this never
interrupts anyone using the app.

The update isn't applied until the spoke's **next restart** (service
restart, reboot, or a person closing and reopening it). That's what
`launcher.bat`/`launcher.ps1` are for: they run *before* `TransportERP.exe`
starts, check for a staged update, and swap the files in — something the
app itself can never safely do to its own running `.exe`/`_internal\*.dll`
on Windows. This is also why Step 4 and the NSSM command above point at
`launcher.bat` instead of the `.exe` directly; if you skip the launcher,
updates will download and stage but never actually apply.

Nothing here touches `.env` or `transport_erp.db` — a release only ever
contains the app's own files, so local config and data always survive an
update untouched.

### Publishing an update

On the dev machine, after making changes:
```
venv\Scripts\python.exe build_spoke_release.py 2026.08.05
```
(pick any version string that sorts/reads clearly — a date works well).
This rebuilds the `.exe`, bundles the launcher scripts and a `VERSION`
file into `dist\TransportERP\`, and zips the whole folder into
`2026.08.05.zip` in the repo root.

Then, on the **hub**, log in as an admin and go to **Admin → Spoke
Releases → Publish a New Release**, enter the same version string, and
upload that `.zip`. Every spoke picks it up on its next check-in — no
need to touch any individual site PC.

If a published version turns out to be broken, publish a fixed one the
same way, or use **Make Latest** on Spoke Releases to point spokes back
at a previous, known-good version — spokes that already staged the bad
one will pick up whichever version is flagged latest on their next check.

## 7. Ongoing

- Local logins, offline reads/writes, and everything else work with no
  internet at all. Sync happens automatically in the background whenever
  the hub is reachable, roughly every `SYNC_INTERVAL_SECONDS`.
- **Sync Conflicts** (hub, admin nav) shows anywhere two sites edited the
  same record while both were offline — last-write-wins, but nothing is
  ever silently discarded; review these periodically.
- **Sync Sites** (hub) lets you revoke a site's key if a PC is
  decommissioned or its key is compromised, without affecting any other
  site.
- If a site PC is replaced, don't reuse its `site_id`/key — revoke the
  old entry and register a fresh one for the new machine.
