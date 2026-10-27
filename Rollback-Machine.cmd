@echo off
"%~dp0PagePilot.Manager.exe" rollback --scope machine --root "%~dp0."
exit /b %errorlevel%
