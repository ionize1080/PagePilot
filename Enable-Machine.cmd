@echo off
"%~dp0PagePilot.Manager.exe" enable --scope machine --root "%~dp0."
exit /b %errorlevel%
