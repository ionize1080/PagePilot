@echo off
"%~dp0PagePilot.Manager.exe" disable --root "%~dp0."
exit /b %errorlevel%
