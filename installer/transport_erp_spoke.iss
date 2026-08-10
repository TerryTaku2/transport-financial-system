; TransFleet ERP — spoke installer.
; Compile with Inno Setup 6 (free, jrsoftware.org/isinfo.php):
;   1. Rebuild the app first:  venv\Scripts\python.exe -m PyInstaller spoke_build.spec
;   2. Open this file in Inno Setup, or run from a command prompt:
;        "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\transport_erp_spoke.iss
;   Output lands in installer\output\TransFleet-ERP-Setup.exe — that one file is the
;   whole distributable. No Python, no .env editing, no admin rights required to install
;   or run it (installs per-user under %LocalAppData%).
;
; What happens after a user runs the installer:
;   - Files land in %LocalAppData%\TransFleet ERP
;   - Start Menu / optional Desktop / optional "start with Windows" shortcuts are created
;   - The app launches once, opens the user's browser to http://127.0.0.1:5000
;   - The app's own first-run /setup wizard (see app.py's `setup` route) takes it from
;     there — SECRET_KEY is generated automatically, and connecting to a hub (if any) is
;     one form, not a copy-pasted API key. See SPOKE_SETUP.md for the underlying flow
;     this replaces.

#define MyAppName "TransFleet ERP"
#define MyAppVersion "1.0"
#define MyAppPublisher "T-Tech Solutions"
#define MyAppExeName "TransportERP.exe"
#define SourceDir "..\dist\TransportERP"

[Setup]
AppId={{B36F2C6E-6E6E-4C7B-9B8E-6B6F1C6E9A11}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
; Per-user install (no admin rights needed) — the target audience is a
; non-technical user double-clicking a downloaded installer, not an IT
; admin with elevation available.
DefaultDirName={localappdata}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=output
OutputBaseFilename=TransFleet-ERP-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}
; The app's own local database/.env live inside {app} (next to the .exe,
; by app.py's own FROZEN-mode design) — leave them behind on uninstall
; by default so re-running the installer later doesn't lose local data.
; Uninstall.exe still offers to remove everything if the user asks for it.

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional shortcuts:"
Name: "startupicon"; Description: "Start {#MyAppName} automatically when Windows starts"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
Source: "{#SourceDir}\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\_internal\*"; DestDir: "{app}\_internal"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#SourceDir}\launcher.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\launcher.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\VERSION"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\launcher.bat"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{userdesktop}\{#MyAppName}"; Filename: "{app}\launcher.bat"; Tasks: desktopicon
Name: "{userstartup}\{#MyAppName}"; Filename: "{app}\launcher.bat"; Tasks: startupicon

[Run]
; The app opens the browser itself on launch (see app.py's webbrowser.open
; call in FROZEN mode) — this just starts the process; "postinstall" makes
; it the default-checked "Launch now" box on the finish page, same as any
; normal installer.
Filename: "{app}\launcher.bat"; Description: "Launch {#MyAppName} now"; Flags: postinstall nowait skipifsilent
