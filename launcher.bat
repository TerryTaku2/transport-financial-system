@echo off
REM Double-click entry point for a spoke install - see launcher.ps1 for
REM what this actually does (apply a staged update if one is waiting,
REM then start TransportERP.exe). Point NSSM at this file too, not at
REM TransportERP.exe directly - see SPOKE_SETUP.md.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0launcher.ps1"
