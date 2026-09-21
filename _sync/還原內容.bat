@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set PY=python
where python >nul 2>nul
if errorlevel 1 set PY=py
echo.
echo   Restoring article content. This takes about 50 minutes.
echo   You can close this window anytime; it resumes where it stopped.
echo   Do NOT press any key while it runs.
echo.
%PY% -u restore.py
echo.
pause
