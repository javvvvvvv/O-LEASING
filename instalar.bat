@echo off
chcp 65001 >nul
title Parche O-Leasing
echo.
echo === PARCHE O-LEASING ===
echo.
echo Este script copia app.py y core\cartera_contable.py
echo a la carpeta donde esta este INSTALAR.bat
echo.
echo IMPORTANTE: pon este ZIP extraido DENTRO de:
echo   C:\O-Leasing\punto y aparte\V10.01\o-leasing\
echo o copia a mano los archivos a esa ruta.
echo.

set "DEST=%~dp0"
if not exist "%DEST%app.py" (
  echo ERROR: no encuentro app.py junto a este bat.
  pause
  exit /b 1
)

if not exist "%DEST%core" mkdir "%DEST%core"

echo Copiando app.py ...
copy /Y "%DEST%app.py" "%DEST%app.py" >nul

if exist "%DEST%core\cartera_contable.py" (
  echo core\cartera_contable.py ya esta en destino.
) else (
  if exist "%~dp0core\cartera_contable.py" (
    copy /Y "%~dp0core\cartera_contable.py" "%DEST%core\cartera_contable.py"
  )
)

echo Borrando __pycache__ ...
if exist "%DEST%core\__pycache__" rd /s /q "%DEST%core\__pycache__"
if exist "%DEST%__pycache__" rd /s /q "%DEST%__pycache__"

echo.
echo Listo. Cierra O-Leasing y vuelve a abrirlo.
echo.
pause
