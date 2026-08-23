@echo off
chcp 65001 >nul
title O-Leasing - Construyendo app de escritorio
color 0B
cd /d "%~dp0"

echo ============================================================
echo   CONSTRUCTOR DE APP DE ESCRITORIO - O-LEASING
echo ============================================================
echo.
echo   Este proceso convierte el sistema en un programa .exe
echo   normal de Windows: con su icono, su ventana propia,
echo   sin consola negra y sin necesidad de navegador.
echo.
echo   Tarda varios minutos la primera vez. NO cierres la ventana.
echo ============================================================
echo.

if not exist ".venv\Scripts\activate.bat" (
    echo [ERROR] No se encontro el entorno virtual ".venv".
    echo Ejecuta primero "instalar.bat".
    echo.
    pause
    exit /b 1
)
call ".venv\Scripts\activate.bat"

echo Verificando herramientas de construccion...
python -m pip show pyinstaller >nul 2>nul
if %errorlevel% neq 0 (
    echo Instalando dependencias de escritorio faltantes...
    python -m pip install pywebview streamlit-desktop-app pyinstaller
)

echo.
echo Limpiando construcciones anteriores...
if exist "build" rmdir /s /q "build"
if exist "dist\O-Leasing" rmdir /s /q "dist\O-Leasing"
if exist "O-Leasing.spec" del /q "O-Leasing.spec"
if exist "dist_src" rmdir /s /q "dist_src"

echo.
echo Preparando la copia de app.py que se entrega (sin comentarios ni
echo notas internas de desarrollo)...
mkdir "dist_src" >nul 2>nul
python dev\preparar_release.py app.py dist_src\app.py
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] No se pudo preparar app.py para entrega. Revisa el mensaje
    echo de arriba - la construccion se detiene aqui A PROPOSITO para
    echo nunca entregar una version que no se pudo verificar.
    pause
    exit /b 1
)

echo.
echo Compilando el ejecutable (esto es lo que tarda varios minutos)...
echo.
python -m PyInstaller launcher.py ^
    --name "O-Leasing" ^
    --icon "assets\icono.ico" ^
    --onedir --windowed --noconfirm --clean ^
    --contents-directory . ^
    --add-data "dist_src\app.py;." ^
    --collect-all streamlit ^
    --copy-metadata streamlit ^
    --collect-all pywebview ^
    --collect-all streamlit_desktop_app ^
    --collect-all pyzipper ^
    --hidden-import numpy_financial ^
    --hidden-import pptx ^
    --hidden-import dateutil.relativedelta ^
    --hidden-import PIL._tkinter_finder

if not exist "dist\O-Leasing" (
    echo.
    echo [ERROR] La construccion no genero la carpeta esperada.
    echo Revisa los mensajes de arriba para ver el detalle del error
    echo ^(suele ser un error de PyInstaller mostrado en rojo^).
    pause
    exit /b 1
)

echo.
echo Copiando datos, configuracion e imagenes al programa final...
xcopy /e /i /y "data" "dist\O-Leasing\data" >nul
copy /y "leasing.db" "dist\O-Leasing\leasing.db" >nul
if exist ".streamlit" xcopy /e /i /y ".streamlit" "dist\O-Leasing\.streamlit" >nul
if exist "assets" xcopy /e /i /y "assets" "dist\O-Leasing\assets" >nul

echo.
echo ============================================================
echo   LISTO
echo ============================================================
echo   Tu app de escritorio esta en:
echo   dist\O-Leasing\O-Leasing.exe
echo.
echo   Pruebala haciendo doble clic ahi mismo.
echo.
echo   Si no abre o se cierra sola: vuelve a hacerle doble clic,
echo   espera 5 segundos, y si aparece un mensaje de error revisa
echo   el archivo "error_al_iniciar.txt" que se crea junto al .exe.
echo   Ese archivo dice EXACTAMENTE que fallo.
echo.
echo   Para entregarla a un cliente como un INSTALADOR real
echo   (con icono, menu inicio y desinstalador), sigue el
echo   paso 2 descrito en LEEME_INSTALADOR.txt usando Inno Setup.
echo.
echo   PROTECCION DEL CODIGO: app.py se entrega sin comentarios ni notas
echo   internas (ve dev\preparar_release.py). Los demas modulos
echo   (models, reports, core, finanzas.py) ya viajan compilados dentro
echo   del propio .exe por como funciona PyInstaller. Para una proteccion
echo   mas fuerte (nombres de variables disfrazados, bytecode cifrado),
echo   la herramienta hecha para esto es PyArmor: pip install pyarmor,
echo   luego "pyarmor gen dist_src\app.py -O dist_src_ofuscado" y usar
echo   esa carpeta en vez de dist_src arriba. No es un paso obligatorio.
echo ============================================================
echo.
pause
