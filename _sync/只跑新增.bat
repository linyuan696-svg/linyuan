@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8

set PY=python
where python >nul 2>nul
if errorlevel 1 set PY=py

if "%HACKMD_API_TOKEN%"=="" (
  echo [X] 還沒設定 Token，請先雙擊「首次設定.bat」。
  echo.
  pause
  exit /b 1
)

echo ==========================================
echo   只跑新增的那幾篇
echo ==========================================
echo.
echo   只處理動作欄是「新增」的列，其餘全部略過
echo   完成後照常重建目錄。
echo.
echo   新寫了一篇、不想把整表重推一遍時用這個。
echo.
echo   執行前請先把 Excel 存檔並關掉。
echo.
pause
echo.

%PY% sync_hackmd.py --new-only
set RC=%ERRORLEVEL%
echo.
if "%RC%"=="0" (
  echo   新增完成。
) else (
  echo   出錯了 ^(代碼 %RC%^)，請截圖給我。
)
echo.
pause
