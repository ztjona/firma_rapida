@echo off
setlocal enabledelayedexpansion
echo ============================================================
echo   Firma Rapida - Build standalone executable
echo ============================================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found in PATH.
    pause & exit /b 1
)

echo [1/3] Installing build dependencies...
pip install pyinstaller pyhanko PyMuPDF Pillow cryptography qrcode docx2pdf pyhanko-certvalidator --upgrade --quiet
if errorlevel 1 (
    echo ERROR: Failed to install dependencies.
    pause & exit /b 1
)

echo.
echo [2/3] Building executable with PyInstaller...
echo       (This may take several minutes)
echo.

:: Clean previous build artifacts
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist

pyinstaller firma_rapida.spec --noconfirm
if errorlevel 1 (
    echo.
    echo ERROR: PyInstaller build failed.
    echo Check the output above for details.
    pause & exit /b 1
)

echo.
echo [3/3] Done!
echo.
echo   Output: dist\FirmaRapida.exe
echo.

:: Show file size
for %%F in (dist\FirmaRapida.exe) do (
    set /a SIZE_MB=%%~zF / 1048576
    echo   Size: !SIZE_MB! MB
)

echo.
echo ============================================================
echo   FirmaRapida.exe is ready to distribute.
echo   No installation required for end users.
echo ============================================================
echo.
pause
