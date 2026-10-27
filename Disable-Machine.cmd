@echo off
"%~dp0PagePilot.Manager.exe" disable --scope machine --root "%~dp0."
exit /b %errorlevel%
