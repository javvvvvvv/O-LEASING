; Instalador de Windows para O-Leasing, hecho con Inno Setup (gratis).
; Genera "ERP_Leasing_Pro_Setup.exe": el cliente lo descarga, elige carpeta
; y queda instalado con su propio desinstalador en "Agregar o quitar
; programas", como cualquier programa normal.
;
; Requisito previo: corre primero "construir_app_escritorio.bat" para
; generar la carpeta dist\O-Leasing. Este script empaqueta esa carpeta,
; no el código fuente.
;
; Cómo usarlo:
;  1) Instala Inno Setup (una sola vez): https://jrsoftware.org/isdl.php
;  2) Abre este archivo con Inno Setup.
;  3) Presiona "Compile" (Ctrl+F9).
;  4) El instalador queda en la carpeta "instalador_salida".

#define MyAppName "O-Leasing"
#define MyAppVersion "6.0"
#define MyAppPublisher "Javier Illán"
#define MyAppExeName "O-Leasing.exe"

[Setup]
AppId={{8F1B2C4A-3E9D-4B6A-9C2F-ERP-LEASING-PRO}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
; Se instala en una carpeta propia del disco (no en "Archivos de
; programa") para que el programa pueda guardar y modificar su
; base de datos (leasing.db) sin necesitar permisos de administrador
; cada vez que alguien guarda información.
DefaultDirName={sd}\O-Leasing
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=instalador_salida
OutputBaseFilename=ERP_Leasing_Pro_Setup
SetupIconFile=assets\icono.ico
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear un acceso directo en el Escritorio"; GroupDescription: "Accesos directos:"; Flags: unchecked

[Files]
; Copia TODO lo que generó construir_app_escritorio.bat (el .exe,
; sus librerías internas, la carpeta data, la base de datos y la
; configuración de tema), manteniendo la estructura de carpetas.
Source: "dist\O-Leasing\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{group}\Desinstalar {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir {#MyAppName} ahora"; Flags: nowait postinstall skipifsilent
