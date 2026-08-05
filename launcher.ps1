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
#
# Whatever fails below is written to launcher.log AND kept on screen with
# a pause, rather than letting the window vanish the instant an error is
# thrown - a console that flashes shut gives whoever's sitting at the spoke
# PC nothing to go on, and nothing to relay back either.

$installDir = $PSScriptRoot
$marker = Join-Path $installDir 'update_ready.json'
$logFile = Join-Path $installDir 'launcher.log'

function Write-Log($message) {
    $line = "$(Get-Date -Format o)  $message"
    $line | Out-File -FilePath $logFile -Append -Encoding utf8
    Write-Output $line
}

function Fail-AndPause($message) {
    Write-Log "ERROR: $message"
    Write-Output ''
    Write-Output 'The spoke app failed to start. Details are above and in launcher.log in this folder.'
    Read-Host 'Press Enter to close this window'
    exit 1
}

try {
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
                # alone since the release zip never contains them. Robocopy's
                # own "0-7 is success" exit codes are informational, not
                # errors - not treated as a failure here.
                robocopy $stagedDir $installDir /E /IS /NFL /NDL /NJH /NJS | Out-Null
                Write-Log "Update $version applied"
            } else {
                Write-Log "Update marker found but staged folder missing ($stagedDir) - skipping"
            }
        } catch {
            # An update that fails to apply should never block a normal
            # start-up - log it and fall through to starting whatever
            # version is already here.
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
    if (-not (Test-Path $exePath)) {
        Fail-AndPause "TransportERP.exe not found at $exePath - is this launcher sitting in the same folder as the app (alongside the _internal folder)?"
    }

    Write-Log "Starting $exePath"
    $proc = Start-Process -FilePath $exePath -WorkingDirectory $installDir -NoNewWindow -PassThru -Wait -ErrorAction Stop
    if ($proc.ExitCode -ne 0) {
        Fail-AndPause "TransportERP.exe exited with code $($proc.ExitCode). Common causes: another program is already using port 5000, or the Microsoft Visual C++ Redistributable isn't installed on this PC."
    }
} catch {
    Fail-AndPause "$_"
}
