@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set PY=python
where python >nul 2>nul
if errorlevel 1 set PY=py
echo.
echo   Pushing article content to HackMD.
echo   With a list file it pushes only those; without one, the whole table.
echo   You can close this window anytime; it resumes where it stopped.
echo.
%PY% -u push_content.py
echo.
pause
