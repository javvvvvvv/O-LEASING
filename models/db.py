# ============================================================================
# PROPIEDAD INTELECTUAL Y LICENCIA COMERCIAL CERRADA
# ============================================================================
# Autor Legal y Titular de Derechos: JAVIER ILLAN GONZALEZ
# Organización: ORANGE CREW
# Contacto: ILLANJAVIER9@GMAIL.COM
#
# ADVERTENCIA LEGAL (MÉXICO Y GLOBAL):
# Este código fuente y su arquitectura son propiedad intelectual exclusiva de
# JAVIER ILLAN GONZALEZ. Queda estrictamente prohibida su reproducción,
# distribución, modificación, ingeniería inversa, copia o uso comercial sin la
# autorización expresa y por escrito del autor. Obra protegida conforme a la
# Ley Federal del Derecho de Autor y tratados internacionales aplicables.
# ============================================================================
"""
models/db.py — Capa de acceso a datos (SQLite) de O-Leasing.

Primer bloque movido fuera de app.py como parte de la migración a
arquitectura por capas (ver docs/MEMORIA_TECNICA.md, sección 3). Reglas
de esta capa:

- Todo aquí recibe `db_path` (o una conexión) como parámetro explícito.
  Nada de esta capa sabe qué es "la empresa activa" ni lee
  `st.session_state` — eso es una decisión de UI, no de datos.
- Se conserva `@st.cache_resource` en `_engine` porque es cacheo de
  rendimiento (una sola conexión por archivo .db durante la sesión del
  proceso), no lógica de negocio; quitarlo reabriría la conexión en cada
  rerun de Streamlit.
- Movido por copiar/pegar textual desde app.py (mismo comportamiento,
  mismas firmas), para que el riesgo de este primer paso del refactor
  sea mínimo. La siguiente función que se mueva aquí debe seguir el
  mismo patrón: recibir `db_path`/`conn` explícito, no leer session_state.
"""
import os
import glob
import shutil
import sqlite3
from datetime import datetime

import streamlit as st

try:
    import pyzipper
    _PYZIPPER_DISPONIBLE = True
except ImportError:
    _PYZIPPER_DISPONIBLE = False


def carpeta_respaldos_auto(db_path: str) -> str:
    carpeta = os.path.join(os.path.dirname(os.path.abspath(db_path)), "respaldos_auto")
    os.makedirs(carpeta, exist_ok=True)
    return carpeta


def bd_esta_sana(db_path: str) -> bool:
    """PRAGMA integrity_check en una conexión aparte y de solo lectura,
    para no interferir con la conexión principal que usa el resto de la app."""
    try:
        conn = sqlite3.connect(f"file:{os.path.abspath(db_path)}?mode=ro", uri=True)
        r = conn.execute("PRAGMA integrity_check").fetchone()
        conn.close()
        return bool(r) and r[0] == "ok"
    except Exception:
        return False


def crear_respaldo_automatico(db_path: str, max_respaldos: int = 8, password: str | None = None):
    """Se llama una vez por sesión, ANTES de abrir la base de datos principal.
    Si el archivo actual está sano, guarda una copia con fecha/hora y borra
    los respaldos automáticos más viejos que excedan max_respaldos. Si el
    archivo YA está dañado, no se sobre-escribe ningún respaldo bueno.

    Si se da `password`, el respaldo se guarda como un .zip cifrado con
    AES-256 (vía pyzipper) en vez de una copia plana del .db — así, si
    alguien copia la carpeta de respaldos (una laptop robada, una carpeta
    de la nube mal compartida, un USB perdido), no puede abrir la base de
    datos de tus clientes sin la contraseña. Si pyzipper no está instalado
    o no se da contraseña, se guarda igual que siempre, sin cifrar — nunca
    se deja de hacer el respaldo por falta de esa librería."""
    if not os.path.exists(db_path):
        return
    if not bd_esta_sana(db_path):
        return
    carpeta = carpeta_respaldos_auto(db_path)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_base = os.path.basename(db_path)

    if password and _PYZIPPER_DISPONIBLE:
        destino = os.path.join(carpeta, f"{nombre_base}.{ts}.bak.zip")
        try:
            with pyzipper.AESZipFile(
                destino, "w", compression=pyzipper.ZIP_LZMA, encryption=pyzipper.WZ_AES
            ) as zf:
                zf.setpassword(password.encode("utf-8"))
                zf.write(db_path, arcname=nombre_base)
        except Exception:
            return
    else:
        destino = os.path.join(carpeta, f"{nombre_base}.{ts}.bak")
        try:
            shutil.copy2(db_path, destino)
        except Exception:
            return

    # Se limpian por separado los respaldos planos y los cifrados —
    # cambiar de "sin contraseña" a "con contraseña" (o viceversa) no debe
    # hacer que de repente se borren de golpe los 8 buenos del otro tipo.
    for patron in (f"{nombre_base}.*.bak", f"{nombre_base}.*.bak.zip"):
        respaldos = sorted(glob.glob(os.path.join(carpeta, patron)))
        for viejo in respaldos[:-max_respaldos]:
            try: os.remove(viejo)
            except Exception: pass


def listar_respaldos_automaticos(db_path: str):
    carpeta = carpeta_respaldos_auto(db_path)
    nombre_base = os.path.basename(db_path)
    archivos = sorted(
        glob.glob(os.path.join(carpeta, f"{nombre_base}.*.bak"))
        + glob.glob(os.path.join(carpeta, f"{nombre_base}.*.bak.zip")),
        reverse=True,
    )
    return [(a, datetime.fromtimestamp(os.path.getmtime(a)), a.endswith(".zip")) for a in archivos]


def restaurar_respaldo_automatico(db_path: str, archivo_respaldo: str, password: str | None = None):
    """Guarda el .db dañado a un lado (para forense/soporte) y pone en su
    lugar el respaldo elegido. Si el respaldo es un .zip cifrado, lo
    descifra primero con `password` — si la contraseña no abre el
    archivo, se avisa con un error claro en vez de dejar la base de datos
    a medio restaurar."""
    if os.path.exists(db_path):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        try:
            shutil.move(db_path, db_path + f".DAÑADO_{ts}.bak")
        except Exception:
            pass
    for ext in ("", "-wal", "-shm"):
        p = db_path + ext
        if os.path.exists(p) and ext:
            try: os.remove(p)
            except Exception: pass

    if archivo_respaldo.endswith(".zip"):
        if not _PYZIPPER_DISPONIBLE:
            raise RuntimeError("Este respaldo está cifrado pero la librería pyzipper no está instalada.")
        try:
            with pyzipper.AESZipFile(archivo_respaldo, "r") as zf:
                zf.setpassword((password or "").encode("utf-8"))
                nombre_interno = zf.namelist()[0]
                with zf.open(nombre_interno) as origen, open(db_path, "wb") as destino:
                    shutil.copyfileobj(origen, destino)
        except RuntimeError:
            raise RuntimeError("La contraseña no coincide con la de este respaldo.")
    else:
        shutil.copy2(archivo_respaldo, db_path)


@st.cache_resource
def _engine(db_path: str):
    conn = sqlite3.connect(db_path, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    # DELETE en vez de WAL: ver nota extensa en el módulo original — mucho
    # más robusto en carpetas sincronizadas por la nube o con antivirus
    # agresivo, que es la causa más común de "database disk image is
    # malformed" en instalaciones de escritorio de un solo usuario como ésta.
    conn.execute("PRAGMA journal_mode=DELETE")
    conn.execute("PRAGMA synchronous=FULL")
    return conn
