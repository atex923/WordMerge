@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

set "APP_VERSION=0.1.1"
set "APP_SCRIPT=WordMerge_V0.1.1.pyw"
set "BUILD_PYTHON=py -3.13"

%BUILD_PYTHON% -c "import sys" >nul 2>&1
if errorlevel 1 set "BUILD_PYTHON=py -3"
%BUILD_PYTHON% -c "import sys" >nul 2>&1
if errorlevel 1 set "BUILD_PYTHON=python"
%BUILD_PYTHON% -c "import sys" >nul 2>&1
if errorlevel 1 (
    echo [錯誤] 找不到 Python 3.11～3.13。
    pause
    exit /b 1
)

%BUILD_PYTHON% -m pip install -r requirements-build.txt
if errorlevel 1 goto :failed

if not exist "Nuitka_Output" mkdir "Nuitka_Output"

%BUILD_PYTHON% -m nuitka ^
  --mode=standalone ^
  --windows-console-mode=disable ^
  --enable-plugin=tk-inter ^
  --include-package=tkinterdnd2 ^
  --include-package=docx ^
  --include-package=docxcompose ^
  --include-package=lxml ^
  --assume-yes-for-downloads ^
  --output-dir="Nuitka_Output\Standalone" ^
  --output-filename="WordMerge_V0.1.1.exe" ^
  --product-name="Word組合工" ^
  --file-description="多份 Word 文件合併工具" ^
  --file-version=%APP_VERSION%.0 ^
  --product-version=%APP_VERSION%.0 ^
  --report="Nuitka_Output\WordMerge_Standalone_Report.xml" ^
  "%APP_SCRIPT%"

if errorlevel 1 goto :failed
echo.
echo Standalone 建置完成：Nuitka_Output\Standalone
pause
exit /b 0

:failed
echo.
echo [錯誤] Standalone 建置失敗，回傳碼 %ERRORLEVEL%。
pause
exit /b 1
