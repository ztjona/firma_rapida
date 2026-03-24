@echo off
echo ============================================================
echo   Firma Rapida - Instalacion de dependencias
echo ============================================================
echo.

:: Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python no esta instalado o no esta en el PATH.
    echo Descargue Python desde https://www.python.org/downloads/
    echo Asegurese de marcar "Add Python to PATH" al instalar.
    pause
    exit /b 1
)

echo Instalando dependencias principales...
echo.

pip install pyhanko PyMuPDF Pillow cryptography qrcode --upgrade

if errorlevel 1 (
    echo.
    echo ERROR: Hubo un problema al instalar las dependencias.
    echo Intente ejecutar como administrador o use:
    echo   pip install pyhanko PyMuPDF Pillow cryptography --user
    pause
    exit /b 1
)

echo.
echo Instalando soporte para Word (.docx)...
echo (Requiere Microsoft Word instalado para funcionar)
echo.

pip install docx2pdf --upgrade

if errorlevel 1 (
    echo.
    echo AVISO: No se pudo instalar docx2pdf.
    echo La firma de PDFs funcionara normalmente.
    echo Para firmar desde Word, instale manualmente: pip install docx2pdf
    echo O use LibreOffice como alternativa.
)

echo.
echo ============================================================
echo   Instalacion completada!
echo   Ejecute "ejecutar.bat" para iniciar Firma Rapida.
echo   Ahora puede abrir PDF y Word (.docx) para firmar.
echo ============================================================
pause
