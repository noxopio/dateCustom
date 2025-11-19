@echo off
REM Wrapper to run setDate.ps1 from the batch file location and forward all args
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setDate.ps1" %*
exit /b %ERRORLEVEL%
