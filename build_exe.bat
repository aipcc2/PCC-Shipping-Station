@echo off
title PCC Shipping App - Build EXE
color 0A
echo =============================================
echo  PCC Shipping App - Build EXE
echo =============================================
echo.

REM ── CHECK PYTHON ──────────────────────────────────────────
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found.
    echo Install from https://python.org/downloads
    echo Make sure to check "Add Python to PATH" during install.
    pause
    exit /b 1
)
echo [OK] Python found:
python --version
echo.

REM ── INSTALL PACKAGES ──────────────────────────────────────
echo Installing pyinstaller, pillow, python-docx, openpyxl, python-barcode...
py -3.12 -m pip uninstall -y qrcode >nul 2>&1
py -3.12 -m pip install pyinstaller pillow python-docx openpyxl python-barcode --quiet --no-warn-script-location
if %errorlevel% neq 0 (
    echo [ERROR] pip install failed.
    pause
    exit /b 1
)
echo [OK] Packages ready.
echo.

REM ── CLEAR STALE BUILD CACHES ──────────────────────────────
pushd "%~dp0"
if exist "build" rmdir /s /q "build"
if exist "PCC Shipping_NEW.spec" del /q "PCC Shipping_NEW.spec"

REM ── BUILD TO A TEMP NAME ──────────────────────────────────
REM Building to "PCC Shipping_NEW.exe" instead of overwriting the live
REM "PCC Shipping.exe" lets the build succeed even if operators currently
REM have the app running on their machines (which locks the file on the
REM network share).
echo Building PCC Shipping_NEW.exe - please wait 60-90 seconds...
echo.
py -3.12 -m PyInstaller --onefile --windowed --name "PCC Shipping_NEW" --add-data "logo.png;." --collect-all openpyxl --collect-all barcode --collect-all PIL --collect-all docx "%~dp0shipping_app.py"

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Build failed.
    popd
    pause
    exit /b 1
)

REM ── SWAP NEW EXE INTO PLACE ───────────────────────────────
echo.
echo Build complete. Attempting to swap into PCC Shipping.exe...

REM Move the existing PCC Shipping.exe aside (if any operator has it open,
REM this will fail — we report partial success and let the admin try the
REM rename later when the app isn't in use).
if exist "dist\PCC Shipping.exe" (
    move /y "dist\PCC Shipping.exe" "dist\PCC Shipping_OLD.exe" >nul 2>&1
    if errorlevel 1 (
        echo.
        echo =============================================
        echo  PARTIAL SUCCESS - Manual swap needed
        echo =============================================
        echo.
        echo The new EXE was built successfully:
        echo   %~dp0dist\PCC Shipping_NEW.exe
        echo.
        echo But the old "PCC Shipping.exe" is currently in use by one or more
        echo operators. Once everyone has closed the app:
        echo.
        echo   1. Delete the old "dist\PCC Shipping.exe"
        echo   2. Rename "dist\PCC Shipping_NEW.exe" to "dist\PCC Shipping.exe"
        echo.
        echo Or just keep using the _NEW file directly.
        echo.
        popd
        pause
        exit /b 0
    )
)

REM Old exe (if any) is now out of the way — move the new one in
move /y "dist\PCC Shipping_NEW.exe" "dist\PCC Shipping.exe" >nul
if errorlevel 1 (
    echo [ERROR] Failed to rename PCC Shipping_NEW.exe to PCC Shipping.exe
    popd
    pause
    exit /b 1
)

REM Clean up the old backup
if exist "dist\PCC Shipping_OLD.exe" del /q "dist\PCC Shipping_OLD.exe"

echo.
echo =============================================
echo  SUCCESS - EXE is ready
echo =============================================
echo.
echo   %~dp0dist\PCC Shipping.exe
echo.
echo Next steps:
echo   1. Copy "PCC Shipping.exe" to the Shipping PC desktops
echo   2. Update any desktop shortcuts to point to the new name
echo   3. Double-click - no install needed on shipping PCs
echo.
popd
pause