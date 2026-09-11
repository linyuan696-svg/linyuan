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
echo   只重建目錄
echo ==========================================
echo.
echo   不處理動作欄，只依照表格目前的列順序
echo   重新產生四本目錄筆記。
echo.
echo   調整過文章順序之後用這個。
echo.
echo   執行前請先把 Excel 存檔並關掉。
echo.
pause
echo.

%PY% sync_hackmd.py --index-only
set RC=%ERRORLEVEL%
echo.
if "%RC%"=="0" (
  echo   目錄重建完成。
) else (
  echo   出錯了 ^(代碼 %RC%^)，請截圖給我。
)
echo.
pause
