@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

set "APP_PYTHON=py -3.13"
%APP_PYTHON% -c "import sys" >nul 2>&1
if errorlevel 1 set "APP_PYTHON=py -3"
%APP_PYTHON% -c "import sys" >nul 2>&1
if errorlevel 1 set "APP_PYTHON=python"
%APP_PYTHON% -c "import sys" >nul 2>&1
if errorlevel 1 (
    echo [錯誤] 找不到 Python，請先安裝 Python 3.11～3.13。
    pause
    exit /b 1
)

if not exist ".venv\Scripts\pythonw.exe" (
    echo 正在建立虛擬環境……
    %APP_PYTHON% -m venv .venv
    if errorlevel 1 goto :failed
)

echo 正在確認必要套件……
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :failed
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto :failed

start "" ".venv\Scripts\pythonw.exe" "WordMerge_V0.1.0.pyw"
exit /b 0

:failed
echo.
echo [錯誤] 環境建立或套件安裝失敗。
pause
exit /b 1
