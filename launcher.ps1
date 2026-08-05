# Entry point for a spoke install (see SPOKE_SETUP.md) - this is what NSSM
# or a desktop shortcut should point at, NOT TransportERP.exe directly.
#
# TransportERP.exe can download and unpack a newer version of itself into
# _update_staged\<version>\ (see check_for_spoke_update in app.py), but it
# can never safely apply that update to its OWN running files - Windows
# keeps a process's .exe and _internal\*.dll locked while it's executing.
# This script runs BEFORE that process exists at all, so nothing is
# locked yet: it applies any staged update first, then starts the app.
#
# Safe to run with no update pending - it just starts the app.

$ErrorActionPreference = 'Stop'
$installDir = $PSScriptRoot
$marker = Join-Path $installDir 'update_ready.json'
$logFile = Join-Path $installDir 'launcher.log'

function Write-Log($message) {
    "$(Get-Date -Format o)  $message" | Out-File -FilePath $logFile -Append -Encoding utf8
}

if (Test-Path $marker) {
    try {
        $info = Get-Content $marker -Raw | ConvertFrom-Json
        $stagedDir = $info.staged_dir
        $version = $info.version
        if ($stagedDir -and (Test-Path $stagedDir)) {
            Write-Log "Applying staged update $version from $stagedDir"
            # /E recurse, /IS include-same (overwrite even if robocopy thinks
            # a file is unchanged), no /MIR - this only ever overwrites files
            # the release actually contains (the .exe, _internal\*, VERSION),
            # so .env, transport_erp.db and this very log/marker are left
            # alone since the release zip never contains them.
            robocopy $stagedDir $installDir /E /IS /NFL /NDL /NJH /NJS | Out-Null
            Write-Log "Update $version applied"
        } else {
            Write-Log "Update marker found but staged folder missing ($stagedDir) - skipping"
        }
    } catch {
        Write-Log "Failed to apply staged update: $_"
    } finally {
        Remove-Item $marker -Force -ErrorAction SilentlyContinue
        $stagingRoot = Join-Path $installDir '_update_staged'
        if (Test-Path $stagingRoot) {
            Remove-Item $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

$exePath = Join-Path $installDir 'TransportERP.exe'
Write-Log "Starting $exePath"
Start-Process -FilePath $exePath -WorkingDirectory $installDir -NoNewWindow -Wait
