@echo off
"%~dp0PagePilot.Manager.exe" ui --scope machine --root "%~dp0."
exit /b %errorlevel%
