@echo off
"%~dp0PagePilot.Manager.exe" enable --scope machine --replace-existing --root "%~dp0."
exit /b %errorlevel%
