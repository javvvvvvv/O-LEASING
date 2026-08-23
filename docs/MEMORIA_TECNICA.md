# Memoria técnica — O-Leasing

Documento de referencia funcional y arquitectónica, mantenido para trazabilidad
del desarrollo (útil para un eventual registro de software ante INDAUTOR México).

## 1. Descripción funcional

O-Leasing es un sistema de control de cartera de arrendamiento (leasing)
financiero. Permite:

- Dar de alta y administrar contratos de arrendamiento (inversión, renta,
  residual, plazo, tasa implícita).
- Generar tablas de amortización mensuales y proyecciones de intereses,
  rentas y residual.
- Calcular indicadores de rentabilidad por contrato (margen, TIR, punto de
  equilibrio).
- Registrar y calcular pólizas contables (inicial, parcialidades,
  comisión, ajustes por eventos especiales como siniestros), incluyendo
  reglas de póliza personalizadas por concepto (alta, edición y
  activación/desactivación de reglas propias).
- Conciliar facturas CFDI (XML del SAT) contra lo esperado según la tabla
  de amortización, detectando diferencias y duplicados.
- Exportar reportes en PDF, Excel y PowerPoint (estados de cuenta,
  pólizas, reportes de conciliación).
- Operar en modalidad multiempresa: cada empresa tiene su propia base de
  datos SQLite y configuración (logo, catálogo de cuentas, reglas).
- Respaldo y restauración automática de la base de datos, con pantallas
  de error amigables (sin traceback técnico) y una pantalla dedicada
  para el caso de base de datos dañada.

## 2. Arquitectura actual

Autoría original: Javier Illán. Stack: Python + Streamlit.

```
o-leasing/
├── app.py            # UI (Streamlit) + rutas + lógica de negocio + acceso a datos
├── finanzas.py        # CORE: fórmulas financieras puras (sin Streamlit)
├── launcher.py         # Punto de entrada del ejecutable de escritorio
├── test_finanzas.py    # Pruebas unitarias del core financiero
├── data/                # Bases SQLite por empresa + respaldos (no va a git)
├── assets/              # Ícono, logo de marca (SVG)
└── .streamlit/config.toml
```

`finanzas.py` ya está correctamente separado como capa **core**: solo
fórmulas (amortización, tasa implícita, TIR, rentabilidad), sin
dependencia de Streamlit ni de la base de datos, cubierto por
`test_finanzas.py`. Esto permite corregir una fórmula en un solo lugar y
que se refleje en toda la app (reportes, pólizas, dashboard).

`app.py` concentra hoy las capas de **rutas/UI** y **negocio** en un solo
archivo (~7000 líneas). Es una app de escritorio madura y en producción,
con manejo defensivo de errores (try/except en operaciones críticas,
pantallas de error amigables propias — `_pantalla_error_amigable`,
`_pantalla_bd_danada` —, `showErrorDetails = "none"` para no exponer
trazas técnicas al usuario final, y respaldo automático antes de
operaciones destructivas).

La capa **models** ya se empezó a separar: `models/db.py` contiene el
motor de conexión SQLite (`_engine`) y las utilidades de respaldo/
restauración (`bd_esta_sana`, `crear_respaldo_automatico`,
`listar_respaldos_automaticos`, `restaurar_respaldo_automatico`,
`carpeta_respaldos_auto`). Se movieron primero porque ya recibían
`db_path` como parámetro explícito — cero acoplamiento con
`st.session_state` o con "cuál es la empresa activa" —, así que el
riesgo de moverlas fue mínimo (copiar/pegar, sin tocar lógica). Probado
de forma aislada (sin depender de `app.py` ni de Streamlit corriendo) y
con las 22 pruebas de `finanzas.py` + arranque real del servidor después
del cambio.

`app.py` sigue teniendo `get_db()` (que sí sabe cuál es la empresa
activa) y el resto de accesos a datos (`init_db`, `cargar_catalogo`,
`get_cfg`/`set_cfg`, CRUD de contratos, facturas, etc.), pendientes de
mover en la siguiente iteración del punto 3 de abajo.

## 3. Ruta de migración a arquitectura por capas (recomendada, no aplicada aún)

Dado que `app.py` es un sistema financiero en operación, la separación de
capas se plantea como refactor **incremental y con pruebas** (ya en
marcha, ver sección 2), no como una reescritura de una sola vez (mayor
riesgo de introducir errores en cálculos financieros o en la conciliación
de facturas). Orden sugerido:

1. ~~`core/` — mover funciones puras de cálculo~~ ya cumplido por
   `finanzas.py` desde el diseño original.
2. `models/` — **en marcha.** `models/db.py` ya tiene el motor de conexión
   y respaldos. Falta mover: `init_db`, `init_facturas_db`,
   `cargar_catalogo`/`guardar_catalogo`, `get_cfg`/`set_cfg`, y el CRUD de
   contratos/facturas/anotaciones/eventos especiales — todas reciben o
   pueden recibir una conexión explícita, así que el patrón a seguir es
   el mismo que ya se usó en `models/db.py`.
3. `ui/` — dejar en `app.py` (o dividir por página) solo el renderizado
   Streamlit, llamando a `core/` y `models/`.

## 4. Identidad de marca

O-Leasing es una sub-marca de **Orange**. En toda salida visible del
sistema (interfaz y exportables) debe quedar identificada tanto la
sub-marca (O-Leasing) como la marca matriz (Orange), sin que el usuario
tenga que buscarla.

- `assets/orange-solo-logo.svg` / `.png` — isotipo (hoja/círculo naranja),
  compartido entre Orange y O-Leasing. Se usa como sello en el sidebar
  cuando la empresa activa no tiene logo propio cargado, y es la base del
  ícono de la aplicación (`icono.ico`).
- `assets/o-leasing-logo.svg` / `.png` — logotipo completo con el nombre
  "O-LEASING". Aparece en el encabezado del panel lateral y en el
  encabezado de todo PDF y Excel generado por el sistema.
- `assets/orange-crew-logo.svg` / `.png` — logotipo completo de la marca
  matriz "ORANGE". Aparece en el pie del sidebar ("Una app de Orange") y
  en el pie de página de todo PDF y Excel generado, junto al crédito de
  O-Leasing.
- `assets/icono.ico` — ícono de la ventana/pestaña del navegador y del
  ejecutable de escritorio, regenerado a partir del isotipo real de la
  marca (antes era un aro naranja genérico sin relación con el logo).

Dónde queda la marca visible hoy en el sistema:

| Salida | O-Leasing | Orange |
|---|---|---|
| Sidebar (panel lateral) | Logotipo en el encabezado | Logo + texto en el pie |
| Ícono de ventana / .exe | Isotipo compartido | — |
| PDF (Estado de Cuenta, Póliza) | Logotipo en encabezado | Logo en pie de página |
| Excel (todas las hojas con título de empresa) | Logotipo en encabezado | Crédito en texto del pie |
| Exportación CONTPAQi (.txt de importación contable) | — | — |

La exportación a CONTPAQi es intencionalmente la única salida sin logo:
es un archivo de datos plano que otro sistema contable importa
directamente, no un documento para leer — insertar imágenes ahí
rompería el formato que CONTPAQi espera.

## 5. Preparación para despliegue web

La app corre hoy como aplicación de escritorio (PyInstaller +
`streamlit-desktop-app`) con SQLite en archivos locales por empresa. Antes
de publicarla en internet, hay tres puntos a resolver (no son cambios de
código menores, requieren decisión de producto):

1. **Persistencia**: en la mayoría de plataformas de hosting web (p. ej.
   Streamlit Community Cloud) el sistema de archivos es efímero — los
   `.db` locales se pierden en cada reinicio/despliegue. Se necesita una
   base de datos externa (Postgres, Turso, etc.) antes de tener usuarios
   reales en producción.
2. **Autenticación**: hoy no hay login; cualquiera con la URL vería todas
   las empresas. Un despliegue público necesita autenticación por usuario
   antes de exponerse.
3. **Multiempresa concurrente**: el diseño actual de "una empresa activa
   en `session_state`" funciona para un usuario local; con varios usuarios
   simultáneos en la nube hay que revisar el aislamiento de sesión.

## 5.1 Autenticación y roles (multi-usuario en el mismo equipo)

Se agregó `models/auth.py`: capa de datos para usuarios, separada de las
bases de datos de cada empresa (`data/usuarios.db`), para que borrar o
restaurar la base de una empresa nunca afecte las cuentas de acceso.

- Contraseñas con PBKDF2-HMAC-SHA256 (200,000 iteraciones) + salt aleatorio
  por usuario, solo con la librería estándar de Python (sin dependencias
  nuevas).
- Tres roles fijos: `admin` (todo, incluida Configuración), `captura`
  (alta/edición de contratos, sin Configuración) y `lectura` (consulta,
  bloqueado de las pantallas de alta/edición/baja).
- Flujo de arranque en `app.py`: `splash` (video de bienvenida en pantalla
  completa) → `login` (o alta del primer administrador si aún no hay
  usuarios) → `app`. Cada etapa corta la ejecución con `st.stop()` para que
  el resto del sistema (empresas, catálogo, menú) no corra hasta que haya
  sesión iniciada.
- Pendiente para una siguiente iteración: permisos a nivel de página más
  finos dentro de cada rol (hoy `lectura` bloquea pantalla completa, no
  campo por campo) y bitácora de qué usuario hizo cada cambio.

## 6. Control de versiones

Commits siguiendo Conventional Commits, por ejemplo:
`git commit -m "feat(ui): agrega logo de marca al sidebar"`
