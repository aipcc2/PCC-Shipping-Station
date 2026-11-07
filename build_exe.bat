@echo off
setlocal
title PCC Shipping App - Build EXE (single file)
color 0A
echo =============================================
echo   PCC Shipping App - Build single-file EXE
echo =============================================
echo.

REM Work from the folder this .bat lives in.
pushd "%~dp0"

REM -- CHECK PYTHON ---------------------------------------------------------
py -3.12 --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python 3.12 not found ^(the "py -3.12" launcher^).
    echo Install Python 3.12 from https://python.org/downloads
    echo Make sure to check "Add Python to PATH" during install.
    popd
    pause
    exit /b 1
)
echo [OK] Python found:
py -3.12 --version
echo.

REM -- INSTALL / UPDATE PACKAGES -------------------------------------------
echo Installing / updating pyinstaller, pillow, python-docx, openpyxl, python-barcode, reportlab...
py -3.12 -m pip install --upgrade pyinstaller pillow python-docx openpyxl python-barcode reportlab --quiet --no-warn-script-location
if %errorlevel% neq 0 (
    echo [ERROR] pip install failed.
    popd
    pause
    exit /b 1
)
echo [OK] Packages ready.
echo.

REM -- CLOSE ANY RUNNING INSTANCE -----------------------------------------
REM A running exe locks the file, so PyInstaller can't overwrite it (WinError 5).
REM Close it first; harmless if it isn't running. Short pause lets antivirus
REM release its scan lock before we delete/rebuild.
echo Closing any running PCC Shipping App (so the exe isn't locked)...
taskkill /f /im "PCC Shipping App.exe" >nul 2>&1
timeout /t 2 /nobreak >nul
echo.

REM -- CLEAN OLD BUILD -----------------------------------------------------
echo Cleaning old build output...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
echo.

REM -- BUILD FROM THE SPEC (single file) ----------------------------------
echo Building "PCC Shipping App.exe" - please wait 60-120 seconds...
echo.
py -3.12 -m PyInstaller --noconfirm --clean "PCC Shipping App.spec"
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Build failed.
    popd
    pause
    exit /b 1
)

echo.
echo =============================================
echo   SUCCESS - build complete
echo =============================================
echo.
echo   Your single-file app:
echo     %~dp0dist\PCC Shipping App.exe
echo.
echo -- TO DEPLOY -------------------------------------------------
echo   Copy that ONE file wherever you need it (a station's local
echo   disk, or the network share). Nothing else goes with it.
echo.
echo   ("build" is just PyInstaller scratch - safe to ignore/delete.)
echo.
popd
pause
endlocal
