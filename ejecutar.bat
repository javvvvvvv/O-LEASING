@echo off
chcp 65001 >nul
title O-Leasing - Servidor
cd /d "%~dp0"

if not exist ".venv\Scripts\activate.bat" (
    echo [ERROR] No se encontro el entorno virtual ".venv".
    echo Ejecuta primero "instalar.bat" para preparar el sistema.
    echo.
    pause
    exit /b 1
)

call ".venv\Scripts\activate.bat"

python -m streamlit --version >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Streamlit no esta instalado correctamente.
    echo Ejecuta de nuevo "instalar.bat".
    echo.
    pause
    exit /b 1
)

echo ============================================================
echo   O-LEASING - Iniciando servidor
echo ============================================================
echo.
echo   Acceso en este equipo:      http://localhost:8501
for /f "tokens=2 delims=:" %%A in ('ipconfig ^| findstr /R /C:"IPv4"') do (
    echo   Acceso desde otros equipos: http://%%A:8501 ^(quitando espacios^)
)
echo.
echo   NO CIERRES ESTA VENTANA mientras el sistema este en uso.
echo   Para detener el servidor, cierra esta ventana o presiona Ctrl+C.
echo ============================================================
echo.

start "" http://localhost:8501
python -m streamlit run app.py --server.port 8501 --server.address 0.0.0.0 --server.headless true

echo.
echo El servidor se detuvo o hubo un error. Revisa el mensaje de arriba.
pause
