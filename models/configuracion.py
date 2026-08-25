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
models/configuracion.py — Configuración de la empresa activa y catálogo de
cuentas contables.

Cuarta pieza movida fuera de app.py en el refactor por capas (ver
docs/PLAN_REFACTOR_CAPAS.md). Mismo patrón de inyección que reports/:
este módulo NO sabe cómo conectarse a la base de datos de la empresa
activa (eso depende de en qué empresa esté parado el usuario en ese
momento, algo que solo app.py conoce vía st.session_state). En vez de
importar Streamlit aquí, app.py registra una sola vez, al arrancar,
una función `get_conn` que regresa la conexión correcta.
"""
import json

_get_conn = None


def set_resolver_conexion(fn) -> None:
    """app.py llama esto una sola vez al arrancar: `fn` debe ser una
    función sin argumentos que regrese la conexión sqlite3 de la
    empresa activa (típicamente su propio `get_db`)."""
    global _get_conn
    _get_conn = fn


def _conn():
    import streamlit as st
    from models.db import _engine
    db_path = st.session_state.get('_active_db_path')
    if not db_path:
        # Fallback if not yet initialized
        return None
    return _engine(db_path)


def cargar_catalogo() -> dict:
    conn = _conn()
    if not conn: return {}
    rows = conn.execute("SELECT clave,cuenta,nombre,activa FROM catalogo_cuentas").fetchall()
    return {
        r['clave']: {
            'cuenta': r['cuenta'], 'nombre': r['nombre'],
            'activa': bool(r['activa'] if r['activa'] is not None else 1),
        }
        for r in rows
    }


def guardar_catalogo(cat: dict) -> None:
    conn = _conn()
    for k, d in cat.items():
        # Si el usuario deja la cuenta en blanco en el admin (Cuentas →
        # Macro), se marca automáticamente como inactiva: no debe
        # generarse en el archivo de Contpaqi ni usarse en ninguna
        # póliza mientras siga vacía.
        cuenta_limpia = (d.get('cuenta') or '').strip()
        activa = 1 if (cuenta_limpia and d.get('activa', True)) else 0
        conn.execute(
            "UPDATE catalogo_cuentas SET cuenta=?,nombre=?,activa=? WHERE clave=?",
            (cuenta_limpia, d['nombre'], activa, k),
        )
    conn.commit()


def get_cfg(key: str, default=None):
    r = _conn().execute("SELECT valor FROM configuracion WHERE clave=?", (key,)).fetchone()
    return r['valor'] if r else default


def set_cfg(key: str, value) -> None:
    conn = _conn()
    conn.execute("INSERT OR REPLACE INTO configuracion VALUES(?,?)", (key, value))
    conn.commit()


def get_config_codificacion_cuentas() -> dict:
    """Cada empresa puede tener un esquema distinto de codificación de
    cuentas (algunas usan 17 dígitos con contratos de 4-4, otras usan
    esquemas más cortos con contratos numerados distinto) — esto vive en
    la configuración de la empresa activa, no es fijo para todo el sistema."""
    default = {'digitos_base': 7, 'digitos_contrato': 8, 'sufijo': '00', 'solo_anexo': False}
    raw = get_cfg('codif_cuentas_json')
    if raw:
        try:
            d = json.loads(raw)
            default.update({k: v for k, v in d.items() if k in default})
        except Exception:
            pass
    return default


def set_config_codificacion_cuentas(cfg: dict) -> None:
    set_cfg('codif_cuentas_json', json.dumps(cfg, ensure_ascii=False))


def cta_sg(base: str, id_c: str, cfg: dict = None) -> str:
    """Arma el número de cuenta específico de un contrato (cuenta base +
    identificador del contrato + sufijo), con el ancho de dígitos que cada
    empresa haya configurado en 'Cuentas (Macro) → Codificación de
    Cuentas'. Antes esto era fijo (7+8+2=17 dígitos); ahora se adapta."""
    cfg = cfg or get_config_codificacion_cuentas()
    base_digits = base.replace('-', '')
    dbase = int(cfg.get('digitos_base', 7))
    base_digits = base_digits.ljust(dbase, '0')[:dbase] if dbase > 0 else base_digits

    origen = id_c.split('-')[-1] if cfg.get('solo_anexo') else id_c.replace('-', '')
    origen_sin_ceros = origen.lstrip('0') or '0'
    dcontr = int(cfg.get('digitos_contrato', 8))
    id_fmt = origen_sin_ceros.zfill(dcontr)[-dcontr:] if dcontr > 0 else ''

    return f"{base_digits}{id_fmt}{cfg.get('sufijo','')}"
