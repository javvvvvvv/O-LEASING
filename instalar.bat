@echo off
chcp 65001 >nul
title Instalando O-Leasing - Dependencias
color 0A
echo ============================================================
echo   INSTALADOR - O-LEASING
echo ============================================================
echo.
cd /d "%~dp0"

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] No se encontro Python instalado en este servidor.
    echo.
    echo Instala Python 3.10 o superior desde https://www.python.org/downloads/
    echo IMPORTANTE: durante la instalacion marca la casilla
    echo             "Add python.exe to PATH"
    echo.
    echo Despues de instalar Python, vuelve a ejecutar este archivo.
    echo.
    pause
    exit /b 1
)

echo [OK] Python encontrado:
python --version
echo.

echo Creando entorno virtual aislado en ".venv" ...
echo (esto evita conflictos con otras versiones de Python o librerias
echo  que ya existan en este servidor)
python -m venv ".venv"
if %errorlevel% neq 0 (
    echo [ERROR] No se pudo crear el entorno virtual.
    pause
    exit /b 1
)

echo.
echo Activando entorno virtual...
call ".venv\Scripts\activate.bat"

echo.
echo Actualizando pip...
python -m pip install --upgrade pip >nul

echo.
echo Instalando librerias necesarias (esto puede tardar varios minutos)...
python -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Fallo la instalacion de una o mas librerias.
    echo Revisa tu conexion a internet e intenta de nuevo.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   INSTALACION COMPLETADA CORRECTAMENTE
echo ============================================================
echo   Ahora ejecuta "ejecutar.bat" para iniciar el sistema.
echo ============================================================
echo.
pause
