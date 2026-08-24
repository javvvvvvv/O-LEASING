@echo off
chcp 65001 >nul
title O-Leasing - Servidor
cd /d "%~dp0"

set "PYTHON_EXE=%~dp0.venv\Scripts\python.exe"

if not exist "%PYTHON_EXE%" (
    echo [ERROR] No se encontro el entorno virtual en "%~dp0.venv".
    echo Ejecuta primero "instalar.bat" para preparar el sistema.
    echo.
    pause
    exit /b 1
)

"%PYTHON_EXE%" -m streamlit --version >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Streamlit no esta instalado correctamente.
    echo Ejecuta de nuevo "instalar.bat".
    echo.
    pause
    exit /b 1
)

"%PYTHON_EXE%" scripts\generate_cert.py >nul 2>nul

echo ============================================================
echo   O-LEASING - Servidor Seguro (HTTPS / SSL)
echo ============================================================
echo.
echo   Acceso seguro en este equipo:   https://localhost:8501
for /f "tokens=2 delims=:" %%A in ('ipconfig ^| findstr /R /C:"IPv4"') do (
    echo   Acceso seguro otros equipos:    https://%%A:8501
)
echo.
echo   [NOTA] Al ser certificado local propio, tu navegador podria
echo   mostrar "Advertencia de privacidad / No seguro". Solo haz clic
echo   en "Configuracion avanzada" y luego "Continuar a localhost".
echo.
echo   NO CIERRES ESTA VENTANA mientras el sistema este en uso.
echo   Para detener el servidor, cierra esta ventana o presiona Ctrl+C.
echo ============================================================
echo.

start "" https://localhost:8501
"%PYTHON_EXE%" -m streamlit run app.py --server.port 8501 --server.address 0.0.0.0 --server.sslCertFile certs/cert.pem --server.sslKeyFile certs/key.pem --server.headless true

echo.
echo El servidor se detuvo o hubo un error. Revisa el mensaje de arriba.
pause

