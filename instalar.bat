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
    echo [ERROR] No se encontro Python instalado en este servidor/equipo.
    echo.
    echo Instala Python 3.10 o superior desde https://www.python.org/downloads/
    echo IMPORTANTE: durante la instalacion marca la casilla
    echo             "Add python.exe to PATH"
    echo.
    pause
    exit /b 1
)

echo [OK] Python encontrado:
python --version
echo.

if not exist "%~dp0.venv\Scripts\python.exe" (
    echo Creando entorno virtual aislado en ".venv" ...
    python -m venv "%~dp0.venv"
    if %errorlevel% neq 0 (
        echo [ERROR] No se pudo crear el entorno virtual.
        pause
        exit /b 1
    )
)

set "PYTHON_EXE=%~dp0.venv\Scripts\python.exe"

echo.
echo Actualizando pip...
"%PYTHON_EXE%" -m pip install --upgrade pip

echo.
echo Instalando librerias necesarias desde requirements.txt...
"%PYTHON_EXE%" -m pip install -r requirements.txt
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
