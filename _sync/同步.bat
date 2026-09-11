@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8

set PY=python
where python >nul 2>nul
if errorlevel 1 set PY=py

if "%HACKMD_API_TOKEN%"=="" (
  echo [X] 還沒設定 Token。
  echo     請先雙擊同一個資料夾裡的「首次設定.bat」。
  echo.
  pause
  exit /b 1
)

echo ==========================================
echo   HackMD 同步
echo ==========================================
echo.
echo   會依照 HackMD同步對照表.xlsx 的「動作」欄
echo   處理新增 / 更新 / 刪除，然後重建四本目錄。
echo.
echo   執行前請先把 Excel 存檔並關掉。
echo.
pause
echo.

%PY% sync_hackmd.py
set RC=%ERRORLEVEL%
echo.
if "%RC%"=="0" (
  echo ==========================================
  echo   跑完了。請打開對照表確認：
  echo   新增的列網址欄已填好、動作欄已清空。
  echo ==========================================
) else (
  echo ==========================================
  echo   中途出錯了 ^(代碼 %RC%^)。
  echo   請把上面整段訊息截圖給我。
  echo ==========================================
)
echo.
pause
