@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set PY=python
where python >nul 2>nul
if errorlevel 1 set PY=py

echo ==========================================
echo   HackMD 同步 - 首次設定
echo   這個只要做一次，做完就不用再開了
echo ==========================================
echo.

%PY% --version >nul 2>nul
if errorlevel 1 (
  echo [X] 找不到 Python。請先安裝 Python 再執行這個檔。
  echo.
  pause
  exit /b 1
)
echo [1/2] 安裝需要的套件...
%PY% -m pip install --quiet --upgrade requests openpyxl
if errorlevel 1 (
  echo [X] 套件安裝失敗，請把上面的訊息截圖給我。
  echo.
  pause
  exit /b 1
)
echo      完成。
echo.

echo [2/2] 設定 HackMD API Token
echo.
echo      到 HackMD 網站 ^> 右上角頭像 ^> Settings ^> API
echo      按 Create API token，把產生的那串字複製起來。
echo.
set /p TK=貼上 Token 後按 Enter： 
if "%TK%"=="" (
  echo.
  echo [X] 沒有輸入東西，取消設定。
  echo.
  pause
  exit /b 1
)
setx HACKMD_API_TOKEN "%TK%" >nul
if errorlevel 1 (
  echo [X] 設定失敗。
  echo.
  pause
  exit /b 1
)
echo.
echo ==========================================
echo   設定完成，Token 已經永久記住了。
echo   以後只要雙擊「同步.bat」就好。
echo ==========================================
echo.
pause
