@echo off
setlocal

set "PROJECT_DIR=%~dp0.."
cd /d "%PROJECT_DIR%"

python -m src.main
exit /b %ERRORLEVEL%
