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
#
# O-Leasing v6
# Control de cartera de arrendamiento (leasing) + conciliación de facturas
# CFDI contra la tabla de amortización.
#
# Autor: Javier Illán
# Contacto para dudas de esta app: illanjavier9@gmail.com
import streamlit as st
from ui.components import sfig, explain, estado_vacio, sz, titled_chart, titled_table
import pandas as pd
from core.cfdi import clasificar_concepto, normalizar_contrato, normalizar_folio, _PAT_CONTRATO, _PAT_MES
from core.cotizador import simular_cotizacion, generar_pdf_cotizacion
from reports.cierre_mensual_pdf import generar_pdf_cierre_mensual
import numpy as np
import numpy_financial as npf
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
import sqlite3, io, os, base64, json, sys, shutil, glob, traceback
import re
import xml.etree.ElementTree as ET
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors as rl_colors
from PIL import Image
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
except ImportError:
    pass

from models.auth import (
    get_auth_db_path, init_auth_db, hay_usuarios, crear_usuario,
    verificar_login, listar_usuarios, cambiar_estado_usuario,
    cambiar_rol_usuario, resetear_password, ROLES, ROL_LABELS,
    crear_sesion, validar_sesion, cerrar_sesion,
)

def _asegurar_tabla_pantallas(db_path):
    """Crea usuario_pantallas si no existe (compatible con auth.py viejo)."""
    import sqlite3 as _sq
    conn = _sq.connect(db_path)
    try:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS usuario_pantallas (
                user_id INTEGER NOT NULL,
                pantalla TEXT NOT NULL,
                PRIMARY KEY (user_id, pantalla)
            )
        """)
        conn.commit()
    finally:
        conn.close()


def _obtener_pantallas_usuario_local(db_path, user_id):
    import sqlite3 as _sq
    _asegurar_tabla_pantallas(db_path)
    conn = _sq.connect(db_path)
    try:
        rows = conn.execute(
            "SELECT pantalla FROM usuario_pantallas WHERE user_id=? ORDER BY pantalla",
            (int(user_id),),
        ).fetchall()
        if not rows:
            return None
        return [r[0] for r in rows]
    finally:
        conn.close()


def _guardar_pantallas_usuario_local(db_path, user_id, pantallas):
    import sqlite3 as _sq
    _asegurar_tabla_pantallas(db_path)
    conn = _sq.connect(db_path)
    try:
        conn.execute("DELETE FROM usuario_pantallas WHERE user_id=?", (int(user_id),))
        n = 0
        for p in pantallas or []:
            p = (p or "").strip()
            if not p:
                continue
            conn.execute(
                "INSERT OR IGNORE INTO usuario_pantallas (user_id, pantalla) VALUES (?,?)",
                (int(user_id), p),
            )
            n += 1
        conn.commit()
        return True, f"Se guardaron {n} pantalla(s) para el usuario."
    except Exception as e:
        return False, str(e)
    finally:
        conn.close()


try:
    from models.auth import obtener_pantallas_usuario as _op_auth
    from models.auth import guardar_pantallas_usuario as _gp_auth
    # Probar que realmente existen y no son stubs
    if not callable(_op_auth) or not callable(_gp_auth):
        raise ImportError("stubs")
    def obtener_pantallas_usuario(db_path, user_id):
        try:
            return _op_auth(db_path, user_id)
        except Exception:
            return _obtener_pantallas_usuario_local(db_path, user_id)
    def guardar_pantallas_usuario(db_path, user_id, pantallas):
        try:
            ok, msg = _gp_auth(db_path, user_id, pantallas)
            # Si el auth viejo devuelve el mensaje de falta archivo, usar local
            if not ok and msg and "Falta models" in str(msg):
                return _guardar_pantallas_usuario_local(db_path, user_id, pantallas)
            # Si falla por tabla inexistente, crear y reintentar local
            if not ok and msg and ("no such table" in str(msg).lower() or "usuario_pantallas" in str(msg).lower()):
                return _guardar_pantallas_usuario_local(db_path, user_id, pantallas)
            return ok, msg
        except Exception:
            return _guardar_pantallas_usuario_local(db_path, user_id, pantallas)
except ImportError:
    obtener_pantallas_usuario = _obtener_pantallas_usuario_local
    guardar_pantallas_usuario = _guardar_pantallas_usuario_local
from models.configuracion import (
    cargar_catalogo, guardar_catalogo, get_cfg, set_cfg,
    get_config_codificacion_cuentas, set_config_codificacion_cuentas, cta_sg,
    set_resolver_conexion as _set_resolver_cfg,
)
from reports.excel import formatear_hoja_excel, excel_con_formato
from reports.pdf import _grafica_amort_para_pdf, _pie_pdf_marca, pdf_estado_cuenta, pdf_poliza

# reports/excel.py y reports/pdf.py no saben qué es "la empresa activa" —
# se les inyecta aquí, una sola vez, cómo resolver el nombre para el
# título de sus documentos (ver docstring de cada módulo).
# Igual que reports/: models/configuracion.py no sabe conectarse a la BD
# de la empresa activa (eso depende de st.session_state) — se le inyecta
# aquí get_db, que ya sabe resolver la empresa activa.
_set_resolver_cfg(lambda: get_db())

# Carpeta real de la aplicación: funciona igual corriendo como script (.py)
# o ya empaquetada como ejecutable de escritorio (.exe), y sin importar
# desde qué carpeta se haya lanzado (doble clic, acceso directo, etc.).
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR    = os.path.join(BASE_DIR, "data")
MASTER_FILE = os.path.join(DATA_DIR, "empresas.json")
DEFAULT_DB  = os.path.join(BASE_DIR, "leasing.db")

# Populate session state early
try:
    get_db_path()
    st.session_state['_nombre_empresa'] = get_cfg('nombre_empresa', get_empresa_actual().get('nombre', 'Mi Empresa'))
except:
    pass

os.makedirs(DATA_DIR, exist_ok=True)

def load_empresas():
    if not os.path.exists(MASTER_FILE):
        d = [{"id":"default","nombre":"Orange Leasing","db_path":DEFAULT_DB}]
        save_empresas(d); return d
    with open(MASTER_FILE,"r",encoding="utf-8") as f: return json.load(f)

def save_empresas(lst):
    with open(MASTER_FILE,"w",encoding="utf-8") as f:
        json.dump(lst, f, ensure_ascii=False, indent=2)

def get_empresas_permitidas(auth_user=None):
    if auth_user is None:
        auth_user = st.session_state.get('auth_user')
    todas = load_empresas()
    if not auth_user: return todas
    rol = auth_user.get('rol', 'lectura')
    if rol in ('admin', 'super_usuario'): return todas
    grupo_id = auth_user.get('grupo_id')
    if not grupo_id: return []
    try:
        from models.auth import obtener_empresas_de_grupo
        permitidas_ids = obtener_empresas_de_grupo(AUTH_DB, grupo_id)
        return [e for e in todas if e['id'] in permitidas_ids]
    except Exception:
        return []

def get_empresa_actual():
    emps = get_empresas_permitidas()
    if not emps: return {"id":"none", "nombre":"Sin acceso (contacte al admin)", "db_path":":memory:"}
    eid  = st.session_state.get("empresa_id", emps[0]["id"])
    for e in emps:
        if e["id"] == eid: return e
    return emps[0]

_PAT_ID_CONTRATO_VALIDO = re.compile(r'^\d{4}-\d{4}$')

def get_db_path():
    p = get_empresa_actual().get("db_path", DEFAULT_DB)
    # Ojo: empresas.json puede guardar rutas relativas (ej. "data/empresa2.db").
    # Las anclamos a BASE_DIR para que jalen igual sin importar desde dónde
    # se abra el programa.
    p = p if os.path.isabs(p) else os.path.join(BASE_DIR, p)
    import streamlit as st
    st.session_state["_active_db_path"] = p
    return p

# Estas claves de session_state NO deben sobrevivir un cambio de empresa.
# Cada empresa numera sus contratos por su cuenta (ej. "0635-0003"), así que
# el mismo ID puede existir en dos empresas apuntando a clientes distintos.
# Sin este limpiado, "Estado de Cuenta" (y otras pantallas con selector)
# se quedaba mostrando el contrato de la empresa anterior. Bug ya resuelto,
# pero dejo la lista aquí por si se agregan más selectores.
_CLAVES_SELECCION_CONTRATO = [
    "ec_sel", "ec_contrato", "amort_sel_contrato", "ib_sel", "rentab_csel",
    "reporte_cliente_sel", "sel_masiva", "anot_new_contrato",
    "fexcl_mora_ind", "fexcl_evt_reg", "fexcl_mora_cli",
]

def limpiar_seleccion_contrato():
    for k in list(st.session_state.keys()):
        if k in _CLAVES_SELECCION_CONTRATO or k.startswith("dbi_"):
            del st.session_state[k]

def hoy_ref() -> date:
    """Fecha que el sistema usa como "hoy" para todos los cálculos (meses
    transcurridos, próximo pago, punto de equilibrio, vencimientos, etc.).
    Normalmente es la fecha real, pero el selector 'Fecha de análisis' del
    sidebar la puede cambiar para ver un cierre pasado. Se cambia en un solo
    lugar para que toda la app quede sincronizada."""
    v = st.session_state.get('fecha_analisis')
    return v if v else date.today()

def viendo_fecha_pasada() -> bool:
    return hoy_ref() != date.today()

# TASA_IVA reservada para cuando se reactiven cálculos de IVA en el sistema
# (por ahora, a petición explícita, ninguna pantalla calcula ni muestra IVA).
# TASA_IVA = 0.16

_ICONO_APP = os.path.join(BASE_DIR, "assets", "icono.ico")
st.set_page_config(page_title="O-Leasing", layout="wide",
                    page_icon=_ICONO_APP if os.path.exists(_ICONO_APP) else None)

_LOGO_SOLO_PATH     = os.path.join(BASE_DIR, "assets", "orange-solo-logo.svg")
_LOGO_WORDMARK_PATH = os.path.join(BASE_DIR, "assets", "o-leasing-logo.svg")
# Variante con el texto en color claro, para usar sobre el fondo oscuro del
# menú lateral — la original tiene el texto "O-LEASING" casi negro (#1A1A1A)
# y se pensó para fondos claros (la tarjeta de login, la barra superior);
# puesta tal cual sobre el sidebar oscuro, el texto prácticamente desaparece.
_LOGO_WORDMARK_CLARO_PATH = os.path.join(BASE_DIR, "assets", "o-leasing-logo-claro.svg")
_LOGO_OLEASING_PNG  = os.path.join(BASE_DIR, "assets", "o-leasing-logo.png")
_LOGO_ORANGE_PNG    = os.path.join(BASE_DIR, "assets", "orange-crew-logo.png")


def _svg_b64(ruta: str) -> str | None:
    """Lee un SVG de /assets y lo devuelve codificado en base64 para
    incrustarlo como <img> en HTML de Streamlit. None si el archivo no
    existe (evita romper el sidebar si falta un asset)."""
    try:
        with open(ruta, "rb") as f:
            return base64.b64encode(f.read()).decode("ascii")
    except Exception:
        return None


_LOGO_SOLO_B64          = _svg_b64(_LOGO_SOLO_PATH)
_LOGO_WORDMARK_B64      = _svg_b64(_LOGO_WORDMARK_PATH)
_LOGO_WORDMARK_CLARO_B64 = _svg_b64(_LOGO_WORDMARK_CLARO_PATH) or _LOGO_WORDMARK_B64
_LOGO_ORANGE_B64    = _svg_b64(os.path.join(BASE_DIR, "assets", "orange-crew-logo.svg"))

C = dict(primary="#1E5C4F", secondary="#3B7C6E", accent="#B3261E",
         success="#1C7A4D", warning="#96660C", info="#3E6FA6",
         gold="#8A6D2F", light="#E1EEEC", purple="#6E5A9C", teal="#1C7A4D",
         brand="#1E5C4F", brand2="#3B7C6E")
PAL = ["#1E5C4F","#1C7A4D","#96660C","#B3261E","#3E6FA6",
       "#6E5A9C","#8A6D2F","#B4607E","#5C7A3E","#3B7C6E"]

# Paleta para gráficas — tonos suaves de la misma paleta del tema (verde
# contable, azul pizarra, ámbar, ladrillo), para que las gráficas no se vean
# "pegadas" encima del resto del diseño.
PAL_PASTEL = ["#5E9587","#7C93AD","#D9A75C","#C4796C","#7FA8C9",
              "#9C8AA5","#B49A5E","#C79AAA","#84AD79","#6E86AE"]
C_PASTEL = dict(primary="#5E9587", secondary="#7C93AD", accent="#C4796C",
                success="#5E9587", warning="#D9A75C", info="#7FA8C9",
                gold="#B49A5E", purple="#9C8AA5", teal="#5E9587")
px.defaults.color_discrete_sequence = PAL_PASTEL
px.defaults.color_continuous_scale = ["#F3F4F6", "#9BC4B8", "#1E5C4F"]

# Colormaps propios para las tablas con degradado (pandas Styler), en vez
# de los de matplotlib de fábrica que no combinaban con la paleta de la app.
import matplotlib
matplotlib.use('Agg')  # backend sin ventana — el servidor no tiene pantalla
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap
CMAP_INDIGO = LinearSegmentedColormap.from_list("ledger_indigo", ["#161B2200", "#2F81F766", "#2f81f7"])
CMAP_CORAL  = LinearSegmentedColormap.from_list("ledger_coral",  ["#161B2200", "#DA363366", "#da3633"])
CMAP_TEAL   = LinearSegmentedColormap.from_list("ledger_teal",   ["#161B2200", "#23863666", "#238636"])

from ui.theme import inyectar_css
inyectar_css()

# ============================================================================
# BIENVENIDA (video de arranque) Y LOGIN — pantallas independientes
# ============================================================================
# Requisito: al abrir el programa, primero se ve el video de bienvenida en
# pantalla completa; encima aparece el login. Cada pantalla se maneja por
# separado (splash -> login -> app), sin mezclarse con el resto del sistema:
# mientras no se resuelva el login, NINGUNA otra parte de app.py corre
# (empresas, catálogo, menú, etc.), tanto por claridad como por seguridad.
AUTH_DB = get_auth_db_path(DATA_DIR)
init_auth_db(AUTH_DB)
try:
    _asegurar_tabla_pantallas(AUTH_DB)
except Exception:
    pass

_VIDEO_BIENVENIDA = os.path.join(BASE_DIR, "assets", "bienvenida.mp4")
_SPLASH_DURACION_DEFAULT = 10.0  # respaldo si el MP4 no se puede leer
_SPLASH_DURACION_MIN = 3.0   # cota de seguridad: si el cálculo de duración
_SPLASH_DURACION_MAX = 20.0  # sale mal (archivo raro, futuro cambio de video),
                              # nunca deja al login esperando de más ni aparece
                              # antes de tiempo — mejor un valor razonable que
                              # uno absurdo que deje la pantalla "colgada".


@st.cache_resource
def _video_duracion_segundos(ruta: str) -> float:
    """Duración real del MP4 leyendo la caja 'mvhd' del contenedor
    (sin depender de ffmpeg/ffprobe, que no viene incluido en el .exe
    empaquetado). Si el archivo no tiene la estructura esperada, o el
    cálculo da algo fuera de un rango razonable, se usa un valor de
    respaldo en vez de arriesgarse a un retraso absurdo."""
    try:
        with open(ruta, "rb") as f:
            data = f.read()

        def _buscar_caja(buf: bytes, nombre: bytes, inicio: int, fin: int):
            pos = inicio
            while pos + 8 <= fin:
                size = int.from_bytes(buf[pos:pos + 4], "big")
                tipo = buf[pos + 4:pos + 8]
                header = 8
                if size == 1:
                    size = int.from_bytes(buf[pos + 8:pos + 16], "big")
                    header = 16
                if size <= 0:
                    break
                if tipo == nombre:
                    return pos + header
                pos += size
            return None

        moov_start = _buscar_caja(data, b"moov", 0, len(data))
        if moov_start is None:
            return _SPLASH_DURACION_DEFAULT
        mvhd_start = _buscar_caja(data, b"mvhd", moov_start, len(data))
        if mvhd_start is None:
            return _SPLASH_DURACION_DEFAULT

        version = data[mvhd_start]
        if version == 1:
            timescale = int.from_bytes(data[mvhd_start + 20:mvhd_start + 24], "big")
            duracion = int.from_bytes(data[mvhd_start + 24:mvhd_start + 32], "big")
        else:
            timescale = int.from_bytes(data[mvhd_start + 12:mvhd_start + 16], "big")
            duracion = int.from_bytes(data[mvhd_start + 16:mvhd_start + 20], "big")
        if timescale <= 0:
            return _SPLASH_DURACION_DEFAULT
        segundos = duracion / timescale
        if segundos < _SPLASH_DURACION_MIN or segundos > _SPLASH_DURACION_MAX:
            return _SPLASH_DURACION_DEFAULT
        return segundos
    except Exception:
        return _SPLASH_DURACION_DEFAULT


# Validar sesion persistente si existe token en query params
if not st.session_state.get('auth_user'):
    _token_persistente = st.query_params.get('st_auth')
    if _token_persistente:
        try:
            _usr_recuperado = validar_sesion(AUTH_DB, _token_persistente)
            if _usr_recuperado:
                st.session_state['auth_user'] = _usr_recuperado
        except Exception:
            pass

# Bienvenida y login: el video se reproduce desenfocado en el fondo
if not st.session_state.get('auth_user'):
    st.markdown(
        '<style>[data-testid="stSidebar"],[data-testid="stHeader"],footer{display:none!important;} '
        '.main .block-container{padding:0!important;max-width:100%!important;}</style>',
        unsafe_allow_html=True,
    )

    with st.container(key='splash_stage'):
        _video_disponible = __import__('os').path.exists(_VIDEO_BIENVENIDA)
        if _video_disponible:
            try:
                st.video(_VIDEO_BIENVENIDA, autoplay=True, muted=True, loop=True)
                st.markdown(
                    """<style>
                    /* Video as fullscreen background */
                    .st-key-splash_stage [data-testid="stVideo"] {
                        position: fixed !important;
                        top: 50% !important;
                        left: 50% !important;
                        min-width: 100% !important;
                        min-height: 100% !important;
                        width: auto !important;
                        height: auto !important;
                        transform: translate(-50%, -50%) scale(1.1) !important;
                        z-index: 0 !important;
                    }
                    .st-key-splash_stage [data-testid="stVideo"] video {
                        object-fit: cover !important;
                        filter: blur(8px) brightness(0.4) !important;
                    }
                    </style>""", unsafe_allow_html=True
                )
            except Exception:
                pass

        st.markdown("<div style='height: 12vh;'></div>", unsafe_allow_html=True)
        col_L, col_C, col_R = st.columns([1, 1.2, 1])
        
        with col_C:
            st.markdown(
                """<style>
                /* Glassmorphism box for login */
                .st-key-login_wrap {
                    position: relative;
                    z-index: 10;
                    background: rgba(15, 15, 15, 0.4) !important;
                    backdrop-filter: blur(15px) !important;
                    -webkit-backdrop-filter: blur(15px) !important;
                    padding: 40px;
                    border-radius: 12px;
                    box-shadow: 0 10px 40px rgba(0, 0, 0, 0.6);
                    border: 1px solid rgba(255, 255, 255, 0.2);
                }
                .st-key-login_wrap h3, .st-key-login_wrap p, .st-key-login_wrap label, .st-key-login_wrap div {
                    text-align: center;
                    color: white !important;
                }
                /* Remove Streamlit form border inside login */
                .st-key-login_wrap [data-testid="stForm"] {
                    border: none !important;
                    background: transparent !important;
                    padding: 0 !important;
                }
                /* Inputs with readable background */
                .st-key-login_wrap div[data-testid="stTextInput"] input {
                    background: rgba(255, 255, 255, 0.95) !important;
                    color: black !important;
                    border-radius: 8px !important;
                    padding: 10px 15px !important;
                    caret-color: black !important;
                }
                .st-key-login_wrap div[data-testid="stTextInput"] {
                    margin-bottom: 15px;
                }
                /* Login button styling */
                .st-key-login_wrap [data-testid="stFormSubmitButton"] button {
                    background: linear-gradient(90deg, #ff8c00, #ff4500) !important;
                    color: white !important;
                    border: none !important;
                    font-weight: bold !important;
                    border-radius: 8px !important;
                    transition: transform 0.2s ease !important;
                }
                .st-key-login_wrap [data-testid="stFormSubmitButton"] button:hover {
                    transform: scale(1.02) !important;
                }
                </style>""", unsafe_allow_html=True
            )
            with st.container(key='login_wrap'):
                _logo_b64 = _svg_b64(_LOGO_WORDMARK_CLARO_PATH)
                if _logo_b64:
                    st.markdown(
                        f'<div style="text-align:center; margin-bottom: 20px;"><img src="data:image/svg+xml;base64,{_logo_b64}" style="width: 60%; max-width: 200px;"></div>',
                        unsafe_allow_html=True
                    )
                
                if not hay_usuarios(AUTH_DB):
                    st.markdown('<h3 style="text-align:center; margin-bottom: 5px;">Bienvenido a O-Leasing</h3>', unsafe_allow_html=True)
                    st.markdown('<p style="text-align:center; color: #ccc!important; margin-bottom: 20px;">Crea tu usuario administrador para comenzar.</p>', unsafe_allow_html=True)
                    with st.form('bootstrap_admin', border=False):
                        _bu = st.text_input('Usuario', placeholder='ej. admin')
                        _bn = st.text_input('Nombre completo', placeholder='Tu nombre')
                        _bp1 = st.text_input('Contraseña', type='password')
                        _bp2 = st.text_input('Confirmar contraseña', type='password')
                        if st.form_submit_button('Crear administrador', use_container_width=True):
                            if _bp1 != _bp2:
                                st.error('Las contraseñas no coinciden.')
                            else:
                                ok, msg = crear_usuario(AUTH_DB, _bu, _bn, _bp1, 'admin')
                                if ok:
                                    st.success(msg + ' Ahora inicia sesión.')
                                    st.rerun()
                                else:
                                    st.error(msg)
                else:
                    st.markdown('<h3 style="text-align:center; margin-bottom: 5px;">Iniciar sesión</h3>', unsafe_allow_html=True)
                    st.markdown('<p style="text-align:center; color: #ccc!important; margin-bottom: 20px;">Sistema de gestión de arrendamiento</p>', unsafe_allow_html=True)
                    with st.form('login_form', border=False):
                        _lu = st.text_input('Usuario', placeholder='Ingresa tu usuario')
                        _lp = st.text_input('Contraseña', type='password', placeholder='Ingresa tu contraseña')
                        if st.form_submit_button('Entrar', use_container_width=True):
                            _user = verificar_login(AUTH_DB, _lu, _lp)
                            if _user:
                                st.session_state['auth_user'] = _user
                                try:
                                    _tok_ses = crear_sesion(AUTH_DB, _user['id'], dias=30)
                                    st.query_params['st_auth'] = _tok_ses
                                except Exception:
                                    pass
                                st.rerun()
                            else:
                                st.error('Usuario o contraseña incorrectos.')
    st.stop()





def facturacion_intereses_rango(fecha_ini: date, fecha_fin: date):
    df = obtener('ACTIVO')
    if df.empty: return pd.DataFrame(), {}
    rows = []; bar = st.progress(0, "Calculando facturación de intereses…")
    cursor = date(fecha_ini.year, fecha_ini.month, 1)
    fin_m  = date(fecha_fin.year,  fecha_fin.month,  1)
    total_meses = min((fin_m.year-cursor.year)*12+(fin_m.month-cursor.month)+1, 60)
    idx = 0
    while cursor <= fin_m and idx < 60:
        fp = pd.Timestamp(cursor); ti=tr=tc=0.0; cnt=0
        for _, row in df.iterrows():
            mc = _mes_en_vigencia_contrato(row, cursor.year, cursor.month, inc_primer_mes=True)
            if mc is None:
                continue
            try:
                dfa, dfr, pl, _t = _amort_tablas_contrato(row)
                ti += float(dfa.iloc[mc - 1]['Interes'])
                tr += float(dfr.iloc[mc - 1]['Interes'])
                tc += float(row.get('Comision_Monto') or 0) / pl if pl else 0.0
                cnt += 1
            except Exception:
                continue
        rows.append({'Mes':cursor.strftime('%Y-%m'),'Mes_Label':cursor.strftime('%b %Y'),
                     'Int_Leasing':round(ti,2),'Int_Residual':round(tr,2),'Amort_Com':round(tc,2),
                     'Solo_Intereses':round(ti+tr,2),'Total_Facturable':round(ti+tr+tc,2),'Contratos':cnt})
        cursor+=relativedelta(months=1); idx+=1
        bar.progress(min(idx/max(total_meses,1),1.0), text=f"Mes {idx}/{total_meses}…")
    bar.empty()
    df2=pd.DataFrame(rows)
    if df2.empty: return df2,{}
    df2['Acumulado']=df2['Total_Facturable'].cumsum()
    mi=df2['Total_Facturable'].idxmax()
    m={'total':df2['Total_Facturable'].sum(),'int_total':df2['Solo_Intereses'].sum(),
       'prom':df2['Total_Facturable'].mean(),'max':df2['Total_Facturable'].max(),
       'max_mes':df2.loc[mi,'Mes_Label'],'meses':len(df2)}
    return df2, m

def mes_pe_capital(row):
    """Mes en que el CAPITAL recuperado (columna 'Capital' de la tabla de
    amortización) alcanza la inversión neta. Dato informativo únicamente:
    en un leasing con valor residual, el residual se recupera aparte, al
    final del contrato, y NO a través de la columna Capital de la
    amortización mes a mes — por eso este número casi nunca se alcanza
    dentro del plazo si el contrato tiene residual, y no debe usarse para
    decidir si un contrato "ya llegó" a su punto de equilibrio (para eso
    ver mes_pe_rentas, abajo)."""
    inv = float(row['Valor_Sin_IVA']) - float(row['Anticipo_Monto'])
    if inv <= 0 or float(row['Mensualidad_Sin_IVA']) <= 0 or int(row['Plazo']) <= 0:
        return None, None
    dfa,_,_,_,_ = calc_amort(round(inv,4), round(float(row['Mensualidad_Sin_IVA']),4),
                             round(float(row['Residual_Monto']),4), int(row['Plazo']),
                             round(float(row['Tasa_Calculada']),8))
    acum = 0.0
    for _,fr in dfa.iterrows():
        acum += fr['Capital']
        if acum >= inv:
            return int(fr['Mes']), dfa
    return None, dfa

def mes_pe_rentas(row):
    """Mes de punto de equilibrio "de caja": cuándo la RENTA acumulada que
    cobras (capital + interés juntos, como realmente entra el dinero) llega
    a igualar la inversión neta. Es el mismo criterio que ya usan las
    gráficas principales de 'Punto de Equilibrio' (columna Mes_PE_Rentas)
    — se usa aquí como LA definición oficial de "¿ya llegó a su punto de
    equilibrio?" en todo el sistema (Dashboard, Estado de Cuenta y el
    reporte), justo para que los tres lugares siempre digan lo mismo."""
    inv = float(row['Valor_Sin_IVA']) - float(row['Anticipo_Monto'])
    renta = float(row['Mensualidad_Sin_IVA'])
    plazo = int(row['Plazo'])
    if inv <= 0 or renta <= 0 or plazo <= 0:
        return None
    acumulado = 0.0
    for mes in range(1, plazo+1):
        acumulado += renta
        if acumulado >= inv:
            return mes
    return None

def _calcular_punto_equilibrio_interna():
    df = obtener('ACTIVO')
    if df.empty: return pd.DataFrame()
    rows=[]
    hoy_pe = hoy_ref()
    for _,row in df.iterrows():
        inv=row['Valor_Sin_IVA']-row['Anticipo_Monto']
        if inv<=0 or row['Mensualidad_Sin_IVA']<=0: continue
        mes_pe_cap,_ = mes_pe_capital(row)
        mes_pe3 = mes_pe_rentas(row)
        g,m,_=rentabilidad(row); t=tir(row)
        fa_pe = row.get('Fecha_Alta')
        meses_transc = 0
        if pd.notnull(fa_pe):
            fa_pe_ts = pd.to_datetime(fa_pe)
            meses_transc = max((hoy_pe.year-fa_pe_ts.year)*12 + (hoy_pe.month-fa_pe_ts.month), 0)
        meses_transc = min(meses_transc, int(row['Plazo']))
        pe_alcanzado_hoy = bool(mes_pe3 is not None and meses_transc >= mes_pe3)
        rows.append({
            'ID_Contrato':row['ID_Contrato'],'Cliente':row['Cliente'],'Vehiculo':row['Vehiculo'],
            'Inversion_Neta':round(inv,2),'Renta':row['Mensualidad_Sin_IVA'],'Plazo':int(row['Plazo']),
            'Mes_PE_Capital':mes_pe_cap if mes_pe_cap else int(row['Plazo']),
            'Mes_PE_Rentas':mes_pe3 if mes_pe3 else int(row['Plazo']),
            'PE_Recuperado': mes_pe_cap is not None,
            'Meses_Transcurridos': meses_transc,
            'PE_Alcanzado_Hoy': pe_alcanzado_hoy,
            'Meses_Para_PE': max((mes_pe3 - meses_transc), 0) if (mes_pe3 is not None and not pe_alcanzado_hoy) else 0,
            'Ganancia':round(g,2),'Margen':round(m,2),'TIR Anual':t,
            'Residual':row['Residual_Monto'],'Tasa_Anual':round(row['Tasa_Calculada']*1200,2)
        })
    df2=pd.DataFrame(rows)
    if not df2.empty:
        df2['Pct_PE']=((df2['Mes_PE_Rentas']/df2['Plazo'])*100).round(1)
    return df2

@st.cache_data(ttl=45, max_entries=64, show_spinner=False)
def _calcular_punto_equilibrio_cached(db_path, _version, _fecha_ref):
    return _calcular_punto_equilibrio_interna()

def calcular_punto_equilibrio():
    # Recorre y recalcula TIR + amortización de CADA contrato activo — el
    # cálculo más pesado del sistema. Se cachea con la misma llave segura
    # que obtener() (empresa activa + versión de datos) para que nunca
    # mezcle carteras de dos empresas ni se quede pegado con números viejos
    # por más de 45 segundos, incluso en el peor de los casos. También se
    # incluye la "Fecha de análisis" en la llave: si no, cambiar esa fecha
    # sin que cambien los datos podía devolver un resultado cacheado que
    # seguía calculado con la fecha anterior hasta por 45 segundos.
    db_path = get_db_path()
    version = _version_datos.get(db_path, 0)
    return _calcular_punto_equilibrio_cached(db_path, version, hoy_ref().isoformat()).copy()

# Base de datos
#
# Nota sobre "database disk image is malformed": ese error significa que el
# .db se dañó a nivel de bytes. Casi siempre pasa por WAL de SQLite (los
# archivos -wal/-shm se desincronizan si la carpeta vive en OneDrive/Drive)
# o por un antivirus escaneando el archivo justo cuando se está escribiendo.
# Por eso aquí NO se usa WAL, se usa el modo clásico — un poco más lento,
# pero mucho más sólido para una sola persona usando la app. Además se
# guarda un respaldo automático al arrancar, y si el archivo se daña la app
# ofrece restaurar el último respaldo con un clic en vez de tronar.

# Capa de datos (SQLite) — ver models/db.py. Nada de esa capa depende de
# Streamlit ni de "empresa activa"; ese contexto lo agrega app.py.
from models.db import (
    carpeta_respaldos_auto, bd_esta_sana, crear_respaldo_automatico,
    listar_respaldos_automaticos, restaurar_respaldo_automatico, _engine,
    _PYZIPPER_DISPONIBLE,
)

def get_db(): return _engine(get_db_path())

def init_facturas_db():
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS facturas (
            uuid            TEXT PRIMARY KEY,
            id_contrato     TEXT,
            fecha_emision   TEXT NOT NULL,
            folio           TEXT,
            subtotal        REAL,
            total           REAL,
            tipo            TEXT,
            periodo         TEXT,
            mes_contrato    INTEGER,
            estatus         TEXT DEFAULT 'PENDIENTE',
            observaciones   TEXT,
            fecha_registro  TEXT DEFAULT CURRENT_TIMESTAMP,
            xml_raw         TEXT,
            FOREIGN KEY (id_contrato) REFERENCES contratos(ID_Contrato)
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS factura_conceptos (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            uuid        TEXT NOT NULL,
            clave       TEXT,
            descripcion TEXT,
            importe     REAL,
            FOREIGN KEY (uuid) REFERENCES facturas(uuid)
        )
    """)
    # Reglas de "contrato facturado con otro número": cuando una arrendataria
    # factura consistentemente con un número de contrato distinto al que
    # tiene en el sistema (typo del cliente, número interno de ellos, etc.),
    # aquí queda guardada la equivalencia para que la próxima vez que llegue
    # ese mismo número mal, el sistema lo resuelva solo — sin volver a dejarla
    # en "Sin contrato" ni pedir que la asignes otra vez a mano.
    c.execute("""
        CREATE TABLE IF NOT EXISTS contrato_alias (
            numero_facturado  TEXT PRIMARY KEY,
            id_contrato_real  TEXT NOT NULL,
            veces_aplicado    INTEGER DEFAULT 0,
            fecha_creacion    TEXT DEFAULT CURRENT_TIMESTAMP,
            fecha_ultimo_uso  TEXT,
            FOREIGN KEY (id_contrato_real) REFERENCES contratos(ID_Contrato)
        )
    """)
    # Registro de auditoría: cada vez que se crea o se corrige una regla de
    # alias queda un renglón aquí, para poder ver el historial completo de
    # qué se aprendió y cuándo (no solo el estado actual de la regla).
    c.execute("""
        CREATE TABLE IF NOT EXISTS contrato_alias_historial (
            id                INTEGER PRIMARY KEY AUTOINCREMENT,
            numero_facturado  TEXT NOT NULL,
            id_contrato_real  TEXT NOT NULL,
            accion            TEXT NOT NULL,
            detalle           TEXT,
            fecha             TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    for idx_sql in [
        "CREATE INDEX IF NOT EXISTS idx_fact_periodo    ON facturas(periodo)",
        "CREATE INDEX IF NOT EXISTS idx_fact_contrato   ON facturas(id_contrato)",
        "CREATE INDEX IF NOT EXISTS idx_fact_estatus    ON facturas(estatus)",
        "CREATE INDEX IF NOT EXISTS idx_fact_tipo       ON facturas(tipo)",
        "CREATE INDEX IF NOT EXISTS idx_fconc_uuid      ON factura_conceptos(uuid)",
        "CREATE INDEX IF NOT EXISTS idx_alias_hist_num  ON contrato_alias_historial(numero_facturado)",
    ]:
        c.execute(idx_sql)
    # Migración: RFC del emisor/receptor (para blindar contra facturas de otra
    # empresa o mezcladas por error) y bandera de cancelación (para que una
    # factura cancelada ante el SAT deje de contar en los totales de control).
    fact_cols = [r[1] for r in c.execute("PRAGMA table_info(facturas)").fetchall()]
    for col, ddl in [
        ("rfc_emisor",   "ALTER TABLE facturas ADD COLUMN rfc_emisor TEXT"),
        ("rfc_receptor", "ALTER TABLE facturas ADD COLUMN rfc_receptor TEXT"),
        ("cancelada",    "ALTER TABLE facturas ADD COLUMN cancelada INTEGER DEFAULT 0"),
        ("fecha_cancelacion", "ALTER TABLE facturas ADD COLUMN fecha_cancelacion TEXT"),
        # El resultado de la comparación (cuánto se esperaba, cuánto venía en
        # el XML, la diferencia) se guarda de forma permanente aquí — antes
        # solo existía un instante, mientras duraba el clic de "Procesar y
        # Conciliar", y se perdía en cuanto cambiabas de pantalla. Así el
        # detalle de CUALQUIER mes o año que ya se procesó queda siempre a
        # la mano, se vuelva a abrir la app hoy o en un año.
        ("esperado",   "ALTER TABLE facturas ADD COLUMN esperado REAL DEFAULT 0"),
        ("facturado",  "ALTER TABLE facturas ADD COLUMN facturado REAL DEFAULT 0"),
        ("diferencia", "ALTER TABLE facturas ADD COLUMN diferencia REAL DEFAULT 0"),
        # Número de contrato tal como venía escrito en el texto del XML,
        # ANTES de aplicar cualquier regla de alias — se conserva aparte de
        # `id_contrato` (a qué contrato real quedó asignada la factura) para
        # poder ver siempre "decía X, se tomó como Y" y para poder aprender
        # la regla la primera vez que se resuelve a mano.
        ("id_contrato_detectado", "ALTER TABLE facturas ADD COLUMN id_contrato_detectado TEXT"),
        ("alias_aplicado", "ALTER TABLE facturas ADD COLUMN alias_aplicado INTEGER DEFAULT 0"),
    ]:
        if col not in fact_cols:
            c.execute(ddl)

    # Migración: O-Leasing dejó de guardar el XML completo de cada CFDI en
    # la base de datos (con miles de facturas eso pesaba mucho sin
    # necesidad — todo lo que de verdad se usa ya vive en columnas propias
    # y en factura_conceptos). Aquí se limpia lo que ya estaba guardado de
    # antes y se recupera el espacio en disco con VACUUM. Es segura de
    # dejar en cada arranque: en cuanto ya no queda ningún xml_raw
    # pendiente, el COUNT sale en 0 y no vuelve a hacer nada.
    if c.execute("SELECT COUNT(*) FROM facturas WHERE xml_raw IS NOT NULL").fetchone()[0]:
        c.execute("UPDATE facturas SET xml_raw = NULL WHERE xml_raw IS NOT NULL")
        conn.commit()
        try:
            conn.execute("VACUUM")
        except Exception:
            pass  # VACUUM es solo para recuperar espacio en disco; si falla, no bloquea el arranque

    # Migración: para poder reclasificar a mano un concepto que el sistema no
    # reconoció (texto distinto al esperado en la póliza) y saber cuáles se
    # tocaron a mano vs. cuáles se clasificaron solas por el texto del XML.
    fconc_cols = [r[1] for r in c.execute("PRAGMA table_info(factura_conceptos)").fetchall()]
    if "manual" not in fconc_cols:
        c.execute("ALTER TABLE factura_conceptos ADD COLUMN manual INTEGER DEFAULT 0")
    if "descripcion" not in fconc_cols:
        c.execute("ALTER TABLE factura_conceptos ADD COLUMN descripcion TEXT")

    # Historial de corridas: cada vez que se procesa un lote de XMLs queda un
    # registro aquí — a diferencia del historial de facturas (que se puede ir
    # sobreescribiendo si re-procesas el mismo período), esto es un renglón
    # nuevo cada vez, para siempre poder ver cuándo se concilió qué y con
    # qué resultado, aunque hayan pasado meses.
    c.execute("""
        CREATE TABLE IF NOT EXISTS conciliacion_corridas (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha_corrida   TEXT NOT NULL,
            periodo         TEXT,
            total_facturas  INTEGER DEFAULT 0,
            conciliadas     INTEGER DEFAULT 0,
            discrepancias   INTEGER DEFAULT 0,
            sin_contrato    INTEGER DEFAULT 0,
            errores         INTEGER DEFAULT 0,
            no_aplica       INTEGER DEFAULT 0,
            rfc_incorrecto  INTEGER DEFAULT 0,
            omitidas        INTEGER DEFAULT 0,
            rfc_configurado TEXT
        )
    """)
    c.execute("CREATE INDEX IF NOT EXISTS idx_corridas_fecha ON conciliacion_corridas(fecha_corrida)")

    # Reglas de póliza personalizadas: para empresas cuya codificación de
    # cuentas o forma de contabilizar no encaja con el generador estándar
    # (pol_inicial/pol_parcialidad/pol_comision) — aquí el usuario dicta,
    # concepto por concepto, qué cuenta se usa y si va de cargo o de abono.
    c.execute("""
        CREATE TABLE IF NOT EXISTS reglas_poliza_custom (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            concepto    TEXT NOT NULL,
            cuenta_base TEXT NOT NULL,
            nombre_cuenta TEXT,
            tipo        TEXT NOT NULL,
            activa      INTEGER DEFAULT 1,
            orden       INTEGER DEFAULT 0
        )
    """)
    conn.commit()

# Única fuente de verdad para el catálogo contable — si agregas una cuenta
# nueva, ponla aquí y ya, se propaga sola a todas las empresas en el próximo
# arranque (antes había que duplicarla en dos diccionarios y era fácil
# olvidar el segundo, por eso a veces "no generaba las cuentas nuevas").
CATALOGO_CUENTAS_DEFAULT = {
    'CXC_CP':                    ('115-00-00', 'CUENTAS POR COBRAR CP'),
    'CXC_LP':                    ('125-00-00', 'CUENTAS POR COBRAR LP'),
    'RESIDUAL':                  ('126-00-00', 'CXC VALOR RESIDUAL'),
    'RESIDUAL_CP':               ('114-01-00', 'CXC VALOR RESIDUAL CP'),
    'BAJA_INV':                  ('117-00-00', 'BAJA INVENTARIO'),
    'INT_CP':                    ('208-00-00', 'INTERESES POR DEVENGAR CP'),
    'INT_LP':                    ('228-00-00', 'INTERESES POR DEVENGAR LP'),
    'ANT_CAP':                   ('118-00-00', 'CXC ANTICIPO A CAPITAL'),
    'PERDIDA_CESION':            ('580-00-00', 'PÉRDIDA POR CESIÓN'),
    'CXC_RESIDUAL_INTERESES':    ('126-00-00', 'CXC VALOR RESIDUAL INTERESES'),
    'CXC_RESIDUAL_INTERESES_CP': ('114-01-00', 'CXC VALOR RESIDUAL INTERESES CP'),
    'ING_RESIDUAL':              ('400-01-04', 'INGRESOS POR VALOR RESIDUAL'),
    'ING_INTERESES':             ('400-01-01', 'INGRESOS POR INTERESES'),
    'CANCELACION_CAPITAL':       ('413-00-00', 'CANCELACIÓN DE CAPITAL'),
    'CANCELACION_ANTICIPO':      ('400-01-01-0990-0000-00', 'CANCELACIÓN ANTICIPO'),
    'COMISION_PASIVO':           ('204-00-00', 'PASIVO POR COMISIÓN APERTURA'),
    'COMISION_GASTO':            ('420-01-00', 'GASTO POR COMISIÓN APERTURA'),
    'AJUSTE_REDONDEO':           ('9999-00-00', 'AJUSTE POR REDONDEO'),
    'PERDIDA_EVENTO_ESPECIAL':   ('585-00-00', 'PÉRDIDA POR SINIESTRO / ROBO / JURÍDICO'),
    'UTILIDADES_ACUMULADAS':     ('350-00-00', 'UTILIDADES ACUMULADAS (AJUSTE EJERC. ANTERIORES)'),
    'CXC_ASEGURADORA':           ('135-00-00', 'CUENTAS POR COBRAR A ASEGURADORA'),
    'SALVAMENTO':                ('118-50-00', 'ACTIVO RECUPERADO / SALVAMENTO'),
    'GASTO_DETERIORO':           ('560-00-00', 'GASTO POR DETERIORO DE CARTERA (PCE)'),
    'ESTIMACION_INCOBRABLES':    ('119-00-00', 'ESTIMACIÓN PARA CUENTAS INCOBRABLES (PCE)'),
}

def init_db():
    conn = get_db(); c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS contratos (
        ID_Contrato TEXT PRIMARY KEY, Cliente TEXT NOT NULL, Vehiculo TEXT,
        Fecha_Alta TEXT, Fecha_Vencimiento TEXT, Valor_Sin_IVA REAL,
        Mensualidad_Sin_IVA REAL, Plazo INTEGER, Comision_Apertura_Pct REAL,
        Comision_Monto REAL, Anticipo_Pct REAL, Anticipo_Monto REAL,
        Residual_Pct REAL, Residual_Monto REAL, Tasa_Calculada REAL,
        Estatus TEXT, Fecha_Baja TEXT, VP_Residual REAL,
        residual_transferred INTEGER DEFAULT 0, Nivel_Morosidad INTEGER DEFAULT 0)""")
    exist = [r[1] for r in c.execute("PRAGMA table_info(contratos)").fetchall()]
    for col,typ,dv in [("residual_transferred","INTEGER","0"),("Nivel_Morosidad","INTEGER","0"),("VP_Residual","REAL","0.0"),
                       ("Fecha_Excl_Poliza","TEXT","NULL"),("Motivo_Excl_Poliza","TEXT","NULL")]:
        if col not in exist: c.execute(f"ALTER TABLE contratos ADD COLUMN {col} {typ} DEFAULT {dv}")
    c.execute("CREATE INDEX IF NOT EXISTS idx_est ON contratos(Estatus)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_fa  ON contratos(Fecha_Alta)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_cli ON contratos(Cliente)")
    for col,typ,dv in [("Tipo_Baja","TEXT","'ANTICIPADA'"),
                       ("Motivo_Baja","TEXT","''")]:
        if col not in exist:
            c.execute(f"ALTER TABLE contratos ADD COLUMN {col} {typ} DEFAULT {dv}")
    c.execute("CREATE TABLE IF NOT EXISTS catalogo_cuentas (clave TEXT PRIMARY KEY, cuenta TEXT NOT NULL, nombre TEXT NOT NULL, activa INTEGER DEFAULT 1)")
    catcol = [r[1] for r in c.execute("PRAGMA table_info(catalogo_cuentas)").fetchall()]
    if "activa" not in catcol:
        c.execute("ALTER TABLE catalogo_cuentas ADD COLUMN activa INTEGER DEFAULT 1")
    existentes = [r[0] for r in c.execute("SELECT clave FROM catalogo_cuentas").fetchall()]
    for k, (cta, nom) in CATALOGO_CUENTAS_DEFAULT.items():
        if k not in existentes:
            c.execute("INSERT INTO catalogo_cuentas (clave,cuenta,nombre,activa) VALUES (?,?,?,1)", (k, cta, nom))
    c.execute("CREATE TABLE IF NOT EXISTS configuracion (clave TEXT PRIMARY KEY, valor TEXT)")
    c.execute("""CREATE TABLE IF NOT EXISTS anotaciones (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ID_Contrato TEXT NOT NULL,
        Fecha TEXT NOT NULL,
        Tipo TEXT,
        Texto TEXT NOT NULL)""")
    c.execute("CREATE INDEX IF NOT EXISTS idx_anot_id ON anotaciones(ID_Contrato)")
    c.execute("""CREATE TABLE IF NOT EXISTS eventos_especiales (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ID_Contrato TEXT NOT NULL,
        Tipo_Evento TEXT NOT NULL,
        Fecha_Evento TEXT NOT NULL,
        Fecha_Registro TEXT NOT NULL,
        Valor_Recuperable REAL DEFAULT 0,
        Estatus_Seguro TEXT DEFAULT 'PENDIENTE',
        Monto_Reclamado REAL DEFAULT 0,
        Monto_Confirmado REAL DEFAULT 0,
        Fecha_Confirmacion TEXT,
        Observaciones TEXT,
        Ajuste_Registrado INTEGER DEFAULT 0)""")
    c.execute("CREATE INDEX IF NOT EXISTS idx_evt_id ON eventos_especiales(ID_Contrato)")
    init_facturas_db()
    conn.commit()

def norm_pct(v): return v*100 if v<=1 else v
# get_config_codificacion_cuentas / set_config_codificacion_cuentas / cta_sg /
# cargar_catalogo / guardar_catalogo / get_cfg / set_cfg viven ahora en
# models/configuracion.py (ver docs/PLAN_REFACTOR_CAPAS.md, fase 2) — se
# importan junto con los demás módulos de models/, arriba del archivo.

# Cálculos financieros
# La fórmula en sí vive en finanzas.py (sin dependencia de Streamlit,
# para poder probarla con pytest). Aquí solo se le agrega el cacheo,
# que es un detalle de rendimiento de la interfaz, no de la fórmula.
from finanzas import (
    calc_amort as _calc_amort_pura,
    calc_res_amort as _calc_res_amort_pura,
    vp_res, calc_tasa, rentabilidad, tir,
)

@st.cache_data(max_entries=4000, ttl=7200)
def calc_amort(inv,renta,residual,plazo,tasa):
    return _calc_amort_pura(inv,renta,residual,plazo,tasa)

@st.cache_data(max_entries=4000, ttl=7200)
def calc_res_amort(vp,tasa,plazo):
    return _calc_res_amort_pura(vp,tasa,plazo)

# ---------------------------------------------------------------------
# Reglas de póliza personalizadas — para empresas cuya forma de
# contabilizar no encaja con el generador estándar (pol_inicial /
# pol_parcialidad / pol_comision). Aquí el usuario dicta, concepto por
# concepto, en qué cuenta va y si es cargo o abono; el sistema solo
# calcula el monto de cada concepto (usando las mismas fórmulas de
# amortización que ya usa toda la app) y arma la póliza con esas reglas.
# ---------------------------------------------------------------------
CONCEPTOS_POLIZA_CUSTOM = {
    'INTERES_MES':        'Interés del mes (leasing)',
    'CAPITAL_MES':        'Capital amortizado del mes',
    'RENTA_TOTAL_MES':    'Renta total del mes (interés + capital)',
    'RESIDUAL_INTERES_MES': 'Interés acumulado del residual del mes',
    'COMISION_MES':       'Comisión de apertura amortizada del mes',
    'ANTICIPO_INICIAL':   'Anticipo (solo en el mes de alta)',
    'RESIDUAL_PACTADO':   'Valor residual pactado (solo en el mes de alta)',
}

def valor_concepto_custom(con: dict, concepto: str, mes: int, anio: int):
    """Misma logica que Tabla Mensual / poliza con primer mes en firma."""
    mc = _mes_en_vigencia_contrato(con, anio, mes, inc_primer_mes=True)
    pl = int(con['Plazo'])
    res = float(con['Residual_Monto'])

    if concepto in ('ANTICIPO_INICIAL', 'RESIDUAL_PACTADO'):
        # Solo mes de alta (mc==1)
        if mc != 1:
            return 0.0
        return float(con['Anticipo_Monto']) if concepto == 'ANTICIPO_INICIAL' else res

    if mc is None:
        return 0.0

    if concepto == 'COMISION_MES':
        com = con.get('Comision_Monto', 0) or 0
        return round(float(com) / pl, 2) if pl else 0.0

    try:
        dfa, dfr, pl, _t = _amort_tablas_contrato(con)
    except Exception:
        return 0.0

    if concepto in ('INTERES_MES', 'CAPITAL_MES', 'RENTA_TOTAL_MES'):
        fila = dfa.iloc[mc - 1]
        if concepto == 'INTERES_MES':
            return round(float(fila['Interes']), 2)
        if concepto == 'CAPITAL_MES':
            return round(float(fila['Capital']), 2)
        return round(float(fila['Interes']) + float(fila['Capital']), 2)

    if concepto == 'RESIDUAL_INTERES_MES':
        return round(float(dfr.iloc[mc - 1]['Interes']), 2)

    return 0.0

def obtener_reglas_poliza_custom(solo_activas: bool = False) -> pd.DataFrame:
    conn = get_db()
    q = "SELECT * FROM reglas_poliza_custom"
    if solo_activas:
        q += " WHERE activa=1"
    q += " ORDER BY orden, id"
    return pd.read_sql_query(q, conn)

def guardar_regla_poliza_custom(concepto, cuenta_base, nombre_cuenta, tipo, regla_id=None):
    conn = get_db()
    if regla_id:
        conn.execute(
            "UPDATE reglas_poliza_custom SET concepto=?, cuenta_base=?, nombre_cuenta=?, tipo=? WHERE id=?",
            (concepto, cuenta_base, nombre_cuenta, tipo, regla_id)
        )
    else:
        conn.execute(
            "INSERT INTO reglas_poliza_custom (concepto, cuenta_base, nombre_cuenta, tipo, activa) VALUES (?,?,?,?,1)",
            (concepto, cuenta_base, nombre_cuenta, tipo)
        )
    conn.commit()

def eliminar_regla_poliza_custom(regla_id):
    conn = get_db()
    conn.execute("DELETE FROM reglas_poliza_custom WHERE id=?", (regla_id,))
    conn.commit()

def activar_regla_poliza_custom(regla_id, activa: bool):
    conn = get_db()
    conn.execute("UPDATE reglas_poliza_custom SET activa=? WHERE id=?", (1 if activa else 0, regla_id))
    conn.commit()

def generar_poliza_custom(contratos_df: pd.DataFrame, mes: int, anio: int) -> pd.DataFrame:
    """Arma la póliza del período usando las reglas personalizadas: una
    línea por contrato x regla activa, con el monto de cargo o abono que
    corresponda. Los conceptos en $0 (no aplican ese mes) se omiten."""
    reglas = obtener_reglas_poliza_custom(solo_activas=True)
    if reglas.empty or contratos_df.empty:
        return pd.DataFrame(columns=['Cuenta','Descripcion','Cargo','Abono','Concepto'])
    cfg_cod = get_config_codificacion_cuentas()
    filas = []
    for _, con in contratos_df.iterrows():
        for _, regla in reglas.iterrows():
            monto = valor_concepto_custom(con, regla['concepto'], mes, anio)
            if abs(monto) < 0.005:
                continue
            cuenta = cta_sg(regla['cuenta_base'], con['ID_Contrato'], cfg=cfg_cod)
            desc_concepto = CONCEPTOS_POLIZA_CUSTOM.get(regla['concepto'], regla['concepto'])
            filas.append({
                'Cuenta': cuenta,
                'Descripcion': f"{regla['nombre_cuenta'] or regla['cuenta_base']} {con['ID_Contrato']}",
                'Cargo': round(monto, 2) if regla['tipo'] == 'CARGO' else 0.0,
                'Abono': round(monto, 2) if regla['tipo'] == 'ABONO' else 0.0,
                'Concepto': f"{desc_concepto} — {con['ID_Contrato']} — {MN[mes-1]} {anio}",
            })
    return pd.DataFrame(filas)

# Pólizas
from core.contpaqi import poliza_txt_contpaqi

def _filas(pol,cat,id_c):
    rows=[]
    for k,desc,cargo,abono,conc in pol:
        # Si la cuenta no está en el catálogo, o el usuario la dejó en blanco
        # / la desactivó en "Cuentas → Macro", ese concepto se omite por
        # completo de la póliza — no se genera con una cuenta vacía.
        if k not in cat: continue
        if not cat[k].get('activa', True): continue
        cta = (cat[k].get('cuenta') or '').strip()
        if not cta: continue
        nom=cat[k]['nombre']
        cfmt=cta.replace('-','') if k=='CANCELACION_ANTICIPO' else cta_sg(cta,id_c)
        rows.append({'Cuenta':cfmt,'Descripcion':f"{nom} {desc}",'Cargo':round(cargo,2),'Abono':round(abono,2),'Concepto':conc})
    df=pd.DataFrame(rows)
    if not df.empty:
        diff=df['Abono'].sum()-df['Cargo'].sum()
        if abs(diff)>0.001:
            c_aj=cat.get('AJUSTE_REDONDEO',{}).get('cuenta','9999-00-00')
            n_aj=cat.get('AJUSTE_REDONDEO',{}).get('nombre','AJUSTE')
            df.loc[len(df)]={'Cuenta':cta_sg(c_aj,id_c),'Descripcion':f"{n_aj} {id_c}",'Cargo':max(diff,0),'Abono':max(-diff,0),'Concepto':"Ajuste redondeo"}
    return df

def pol_inicial(con,cat):
    if con['Fecha_Baja'] and pd.to_datetime(con['Fecha_Baja'])<pd.to_datetime(con['Fecha_Alta']): return pd.DataFrame()
    id_c=con['ID_Contrato']; v=con['Valor_Sin_IVA']; r=con['Mensualidad_Sin_IVA']
    pl=con['Plazo']; t=con['Tasa_Calculada']; res=con['Residual_Monto']; ant=con['Anticipo_Monto']; inv=v-ant
    vp=vp_res(res,t,pl); _,icp,ilp,rcp,rlp=calc_amort(round(inv,4),round(r,4),round(res,4),pl,round(t,8))
    pol=[("ANT_CAP",f"Anticipo {id_c}",ant,0,"Registro anticipo"),
         ("CXC_CP",f"Rentas CP {id_c}",rcp,0,f"Rentas CP ({min(12,pl)}m)"),
         ("CXC_LP",f"Rentas LP {id_c}",rlp,0,f"Rentas LP ({max(0,pl-12)}m)"),
         ("RESIDUAL",f"Residual VP {id_c}",vp,0,"VP residual"),
         ("BAJA_INV",f"Baja activo {id_c}",0,v,"Baja activo leasing"),
         ("INT_CP",f"Int fut CP {id_c}",0,icp,"Int a devengar CP"),
         ("INT_LP",f"Int fut LP {id_c}",0,ilp,"Int a devengar LP")]
    tc=sum(x[2] for x in pol); ta=sum(x[3] for x in pol); diff=ta-tc
    if abs(diff)>0.001: pol.append(("PERDIDA_CESION",f"Cesión {id_c}",max(diff,0),max(-diff,0),"Pérd/Util cesión"))
    return _filas(pol,cat,id_c)

def pol_parcialidad(con,mes,anio,cat,inc=False):
    id_c=con['ID_Contrato']; v=con['Valor_Sin_IVA']; r=con['Mensualidad_Sin_IVA']
    pl=con['Plazo']; t=con['Tasa_Calculada']; res=con['Residual_Monto']; ant=con['Anticipo_Monto']; inv=v-ant
    fa=pd.to_datetime(con['Fecha_Alta']); fp=datetime(anio,mes,1)
    fa_m=datetime(fa.year,fa.month,1); fv=pd.to_datetime(con['Fecha_Vencimiento']); fv_m=datetime(fv.year,fv.month,1)
    if con['Fecha_Baja'] and pd.to_datetime(con['Fecha_Baja'])<fp: return pd.DataFrame()
    if fa_m>fp or fv_m<fp: return pd.DataFrame()
    mt=(anio-fa.year)*12+(mes-fa.month); mc=mt+1
    if mt==0 and not inc: return pd.DataFrame()
    if mc<1 or mc>pl: return pd.DataFrame()
    dfa,_,_,_,_=calc_amort(round(inv,4),round(r,4),round(res,4),pl,round(t,8))
    fi=dfa.iloc[mc-1]; im=fi['Interes']; cap=fi['Capital']; mr=pl-mc
    vpr=vp_res(res,t,pl); dfr=calc_res_amort(round(vpr,4),round(t,8),pl)
    fr=dfr.iloc[mc-1]; ir=fr['Interes']; sr=fr['Saldo_Fin']; rt=con.get('residual_transferred',0); pol=[]
    if mr==12 and not rt:
        pol+=[("RESIDUAL_CP",f"Trasp res {id_c}",sr,0,"Traspaso residual LP-CP"),
              ("RESIDUAL",f"Trasp res {id_c}",0,sr,"Baja residual LP")]
        get_db().execute("UPDATE contratos SET residual_transferred=1 WHERE ID_Contrato=?",(id_c,)); get_db().commit(); _tocar_datos(); rt=1
    if mr>12:
        mf=mc+12
        if mf<=pl:
            ifut=dfa.iloc[mf-1]['Interes']
            pol+=[("INT_LP",f"Trasp int m{mf} {id_c}",ifut,0,"Traspaso int LP-CP"),
                  ("INT_CP",f"Trasp int m{mf} {id_c}",0,ifut,"Reclas int LP-CP")]
    if mc==1 and ant>0:
        pol+=[("CANCELACION_ANTICIPO",f"Cancel ant {id_c}",ant,0,"Cancelación anticipo"),
              ("ANT_CAP",f"Cancel ant {id_c}",0,ant,"Baja anticipo")]
    if mr>12:
        pol+=[("CXC_CP",f"Trasp renta m{mc+12} {id_c}",r,0,"Reclas renta LP-CP"),
              ("CXC_LP",f"Trasp renta m{mc+12} {id_c}",0,r,"Baja renta LP")]
    pol.append(("CXC_CP",f"Cobro renta {id_c}",0,r,f"Cobro renta mes {mc}"))
    pol+=[("INT_CP",f"Int dev {id_c}",im,0,f"Devengo int mes {mc}"),
          ("ING_INTERESES",f"Ing int {id_c}",0,im,"Ingreso intereses"),
          ("CANCELACION_CAPITAL",f"Cancel cap {id_c}",cap,0,f"Amort capital mes {mc}")]
    kr='CXC_RESIDUAL_INTERESES_CP' if (rt or mr<=12) else 'CXC_RESIDUAL_INTERESES'
    pol+=[( kr,f"Int res {id_c}",ir,0,f"Int residual mes {mc}"),
          ("ING_RESIDUAL",f"Ing res {id_c}",0,ir,"Ingreso int residual")]
    return _filas(pol,cat,id_c)

def pol_comision(con,mes,anio,cat,inc_am=False):
    id_c=con['ID_Contrato']; com=con['Comision_Monto']; pl=con['Plazo']
    fa=pd.to_datetime(con['Fecha_Alta']); fam=datetime(fa.year,fa.month,1); fp=datetime(anio,mes,1)
    fv=pd.to_datetime(con['Fecha_Vencimiento']); fvm=datetime(fv.year,fv.month,1)
    if con['Fecha_Baja'] and pd.to_datetime(con['Fecha_Baja'])<fp: return pd.DataFrame()
    if fam>fp or fvm<fp: return pd.DataFrame()
    mt=(anio-fam.year)*12+(mes-fam.month)
    if fam.year==anio and fam.month==mes:
        p=[("COMISION_GASTO",f"Com {id_c}",com,0,"Gasto comisión apertura"),
           ("COMISION_PASIVO",f"Com {id_c}",0,com,"Pasivo comisión apertura")]
        if inc_am:
            a=com/pl
            p+=[("COMISION_PASIVO",f"Amort m1 {id_c}",a,0,"Amort com m1"),
                ("COMISION_GASTO",f"Amort m1 {id_c}",0,a,"Rev gasto com")]
        return _filas(p,cat,id_c)
    na=mt+1 if inc_am else mt
    if na<=0 or na>pl: return pd.DataFrame()
    a=com/pl
    p=[("COMISION_PASIVO",f"Amort m{na} {id_c}",a,0,f"Amort com m{na}/{pl}"),
       ("COMISION_GASTO",f"Amort m{na} {id_c}",0,a,"Rev gasto com")]
    return _filas(p,cat,id_c)

# Eventos especiales: siniestro / robo / jurídico
def saldos_sub_ledger(con, mc):
    pl=int(con['Plazo']); r=con['Mensualidad_Sin_IVA']; t=con['Tasa_Calculada']
    res=con['Residual_Monto']; ant=con['Anticipo_Monto']; v=con['Valor_Sin_IVA']; inv=v-ant; com=con['Comision_Monto']
    mc=max(0,min(mc,pl))
    dfa,_,_,_,_=calc_amort(round(inv,4),round(r,4),round(res,4),pl,round(t,8))
    vpr=vp_res(res,t,pl); dfr=calc_res_amort(round(vpr,4),round(t,8),pl)

    cxc_cp=r*min(12,max(0,pl-mc)); cxc_lp=r*max(0,pl-mc-12)
    int_cp=dfa.iloc[mc:min(pl,mc+12)]['Interes'].sum() if mc<pl else 0.0
    int_lp=dfa.iloc[mc+12:pl]['Interes'].sum() if mc+12<pl else 0.0

    punto=max(0,pl-12)
    if mc<punto: residual=vpr; residual_cp=0.0
    else: residual=0.0; residual_cp=vpr
    lp_end=max(0,min(mc,punto-1)); cp_start=max(0,punto-1)
    cxc_res_int=dfr.iloc[:lp_end]['Interes'].sum() if lp_end>0 else 0.0
    cxc_res_int_cp=dfr.iloc[cp_start:mc]['Interes'].sum() if mc>cp_start else 0.0

    com_pasivo=com*max(0,pl-mc)/pl if pl else 0.0

    return dict(CXC_CP=cxc_cp,CXC_LP=cxc_lp,INT_CP=int_cp,INT_LP=int_lp,
                RESIDUAL=residual,RESIDUAL_CP=residual_cp,
                CXC_RESIDUAL_INTERESES=cxc_res_int,CXC_RESIDUAL_INTERESES_CP=cxc_res_int_cp,
                COMISION_PASIVO=com_pasivo)

def calc_ajuste_evento_especial(con, fecha_evento, valor_recuperable=0.0, monto_seguro_confirmado=0.0, hoy=None):
    hoy=pd.to_datetime(hoy or date.today()); fecha_evento=pd.to_datetime(fecha_evento)
    id_c=con['ID_Contrato']; v=con['Valor_Sin_IVA']; r=con['Mensualidad_Sin_IVA']
    pl=int(con['Plazo']); t=con['Tasa_Calculada']; res=con['Residual_Monto']; ant=con['Anticipo_Monto']; inv=v-ant

    fa=pd.to_datetime(con['Fecha_Alta'])
    dfa,_,_,_,_=calc_amort(round(inv,4),round(r,4),round(res,4),pl,round(t,8))
    vpr=vp_res(res,t,pl); dfr=calc_res_amort(round(vpr,4),round(t,8),pl)

    def meses_transcurridos(desde,hasta): return max(0,(hasta.year-desde.year)*12+(hasta.month-desde.month))
    mes_evento=min(pl,meses_transcurridos(fa,fecha_evento))
    mes_hoy=min(pl,max(mes_evento,meses_transcurridos(fa,hoy)))

    saldo_evento=(dfa.iloc[mes_evento-1]['Saldo']-res if mes_evento>0 else inv-res)+(dfr.iloc[mes_evento-1]['Saldo_Fin'] if mes_evento>0 else vpr)
    saldo_hoy_sistema=(dfa.iloc[mes_hoy-1]['Saldo']-res if mes_hoy>0 else inv-res)+(dfr.iloc[mes_hoy-1]['Saldo_Fin'] if mes_hoy>0 else vpr)
    saldos_hoy=saldos_sub_ledger(con,mes_hoy)

    detalle=[]; interes_ant=capital_ant=interes_res_ant=interes_act=capital_act=interes_res_act=0.0
    for i in range(mes_evento,mes_hoy):
        fmes=fa+relativedelta(months=i+1); es_ant=fmes.year<hoy.year
        ic=dfa.iloc[i]['Interes']; cc=dfa.iloc[i]['Capital']; irr=dfr.iloc[i]['Interes']
        detalle.append({'Mes':i+1,'Fecha':fmes.strftime('%Y-%m'),'Interes_Leasing':ic,'Capital':cc,
                         'Interes_Residual':irr,'Total':ic+cc+irr,'Ejercicio':'Anterior (cerrado)' if es_ant else 'Actual'})
        if es_ant: interes_ant+=ic; capital_ant+=cc; interes_res_ant+=irr
        else: interes_act+=ic; capital_act+=cc; interes_res_act+=irr

    capital_revertir=capital_ant+capital_act
    perdida_total=max(0.0,saldo_evento-valor_recuperable-monto_seguro_confirmado)
    fue_ejerc_anterior=fecha_evento.year<hoy.year

    if fue_ejerc_anterior:
        loss_plug_anterior=perdida_total-capital_revertir+interes_res_ant
        loss_plug_actual=interes_res_act
    else:
        loss_plug_anterior=0.0
        loss_plug_actual=perdida_total-capital_revertir+interes_res_act

    return dict(mes_evento=mes_evento,mes_hoy=mes_hoy,saldo_evento=saldo_evento,saldo_hoy_sistema=saldo_hoy_sistema,
        saldos_hoy=saldos_hoy,es_retroactivo=mes_hoy>mes_evento,detalle=pd.DataFrame(detalle),
        capital_revertir=capital_revertir,interes_revertir=interes_ant+interes_act,
        interes_res_revertir=interes_res_ant+interes_res_act,
        loss_plug_anterior=loss_plug_anterior,loss_plug_actual=loss_plug_actual,
        valor_recuperable=valor_recuperable,monto_seguro_confirmado=monto_seguro_confirmado,perdida_total=perdida_total)

def pol_ajuste_evento(con,calc,cat,tipo_evento):
    id_c=con['ID_Contrato']; pol=[]; sl=calc['saldos_hoy']
    motivo_txt={'SINIESTRO':'Siniestro','ROBO':'Robo','JURIDICO':'Proceso jurídico'}.get(tipo_evento,tipo_evento)

    if sl['CXC_CP']>0.001: pol.append(("CXC_CP",f"Baja {motivo_txt} {id_c}",0,sl['CXC_CP'],"Cancela saldo CxC corto plazo"))
    if sl['CXC_LP']>0.001: pol.append(("CXC_LP",f"Baja {motivo_txt} {id_c}",0,sl['CXC_LP'],"Cancela saldo CxC largo plazo"))
    if sl['RESIDUAL']>0.001: pol.append(("RESIDUAL",f"Baja {motivo_txt} {id_c}",0,sl['RESIDUAL'],"Cancela VP residual LP"))
    if sl['RESIDUAL_CP']>0.001: pol.append(("RESIDUAL_CP",f"Baja {motivo_txt} {id_c}",0,sl['RESIDUAL_CP'],"Cancela VP residual CP"))
    if sl['CXC_RESIDUAL_INTERESES']>0.001: pol.append(("CXC_RESIDUAL_INTERESES",f"Baja {motivo_txt} {id_c}",0,sl['CXC_RESIDUAL_INTERESES'],"Cancela interés residual acumulado LP"))
    if sl['CXC_RESIDUAL_INTERESES_CP']>0.001: pol.append(("CXC_RESIDUAL_INTERESES_CP",f"Baja {motivo_txt} {id_c}",0,sl['CXC_RESIDUAL_INTERESES_CP'],"Cancela interés residual acumulado CP"))
    if sl['INT_CP']>0.001: pol.append(("INT_CP",f"Baja {motivo_txt} {id_c}",sl['INT_CP'],0,"Cancela interés por devengar CP"))
    if sl['INT_LP']>0.001: pol.append(("INT_LP",f"Baja {motivo_txt} {id_c}",sl['INT_LP'],0,"Cancela interés por devengar LP"))

    if sl['COMISION_PASIVO']>0.001:
        pol.append(("COMISION_PASIVO",f"Cierre comisión {id_c}",sl['COMISION_PASIVO'],0,"Cancela pasivo por comisión pendiente"))
        pol.append(("COMISION_GASTO",f"Cierre comisión {id_c}",0,sl['COMISION_PASIVO'],"Cancela gasto diferido de comisión"))

    lpa=calc['loss_plug_anterior']
    if lpa>0.001: pol.append(("UTILIDADES_ACUMULADAS",f"Ajuste {motivo_txt} {id_c} (ejerc. ant.)",lpa,0,"Corrección de error de ejercicios anteriores (NIF B-1 / IAS 8)"))
    elif lpa<-0.001: pol.append(("UTILIDADES_ACUMULADAS",f"Ajuste {motivo_txt} {id_c} (ejerc. ant.)",0,-lpa,"Corrección de error de ejercicios anteriores (NIF B-1 / IAS 8)"))
    lpc=calc['loss_plug_actual']
    if lpc>0.001: pol.append(("PERDIDA_EVENTO_ESPECIAL",f"Pérdida {motivo_txt} {id_c}",lpc,0,f"Pérdida por {motivo_txt.lower()} — ejercicio actual"))
    elif lpc<-0.001: pol.append(("PERDIDA_EVENTO_ESPECIAL",f"Recuperación {motivo_txt} {id_c}",0,-lpc,"Recuperación neta — ejercicio actual"))

    if calc['valor_recuperable']>0.001: pol.append(("SALVAMENTO",f"Salvamento {id_c}",calc['valor_recuperable'],0,"Activo recuperado / valor de salvamento"))
    if calc['monto_seguro_confirmado']>0.001: pol.append(("CXC_ASEGURADORA",f"CxC aseguradora {id_c}",calc['monto_seguro_confirmado'],0,"Reembolso confirmado por la aseguradora"))

    return _filas(pol,cat,id_c)

# CRUD eventos especiales (siniestro / robo / jurídico)
def registrar_evento_especial(id_c,tipo_evento,fecha_evento,valor_recuperable=0.0,monto_reclamado=0.0,obs=""):
    conn=get_db()
    conn.execute("""INSERT INTO eventos_especiales
        (ID_Contrato,Tipo_Evento,Fecha_Evento,Fecha_Registro,Valor_Recuperable,Estatus_Seguro,Monto_Reclamado,Observaciones)
        VALUES (?,?,?,?,?,?,'PENDIENTE',?,?)""",
        (id_c,tipo_evento,fecha_evento,date.today().isoformat(),valor_recuperable,monto_reclamado,obs))
    conn.commit()
    motivo_txt={'SINIESTRO':'Siniestro','ROBO':'Robo','JURIDICO':'Proceso jurídico'}.get(tipo_evento,tipo_evento)
    agregar_anotacion(id_c,f"Evento especial registrado: {motivo_txt} (fecha real del evento: {fecha_evento}). "
                            f"Reembolso de seguro: pendiente de confirmación.","Ajuste contable")

def obtener_eventos_especiales(id_c=None,solo_pendientes=False):
    conn=get_db(); q="SELECT * FROM eventos_especiales"; p=[]; cond=[]
    if id_c: cond.append("ID_Contrato=?"); p.append(id_c)
    if solo_pendientes: cond.append("Estatus_Seguro IN ('PENDIENTE','RECHAZADO')")
    if cond: q+=" WHERE "+" AND ".join(cond)
    q+=" ORDER BY Fecha_Evento DESC"
    return pd.read_sql_query(q,conn,params=p)

def actualizar_estatus_seguro(evento_id,estatus,monto_confirmado=0.0,fecha_confirmacion=None):
    conn=get_db()
    conn.execute("UPDATE eventos_especiales SET Estatus_Seguro=?,Monto_Confirmado=?,Fecha_Confirmacion=? WHERE id=?",
                 (estatus,monto_confirmado,fecha_confirmacion,evento_id))
    conn.commit()
    row=conn.execute("SELECT ID_Contrato,Tipo_Evento FROM eventos_especiales WHERE id=?",(evento_id,)).fetchone()
    if row:
        txt=f"Estatus de seguro actualizado a {estatus}"
        if monto_confirmado: txt+=f" — monto confirmado ${monto_confirmado:,.2f}"
        agregar_anotacion(row['ID_Contrato'],f"{txt} ({row['Tipo_Evento']}).","Ajuste contable")

def marcar_ajuste_registrado(evento_id):
    conn=get_db(); conn.execute("UPDATE eventos_especiales SET Ajuste_Registrado=1 WHERE id=?",(evento_id,)); conn.commit()

# Altas, bajas y consultas de contratos
def renumerar_contrato(id_viejo: str, id_nuevo: str):
    """Cambia el número de un contrato (su llave primaria en la base de
    datos) de forma segura: re-apunta facturas, anotaciones y eventos
    especiales al nuevo número, todo en una sola transacción, para que
    nada se quede huérfano apuntando al número viejo."""
    id_viejo = str(id_viejo).strip()
    id_nuevo = str(id_nuevo).strip()
    if not id_nuevo:
        return False, "El número de contrato no puede quedar en blanco."
    if not _PAT_ID_CONTRATO_VALIDO.match(id_nuevo):
        return False, "El número debe tener el formato NNNN-NNNN (solo dígitos), por ejemplo 0521-0001."
    if id_viejo == id_nuevo:
        return True, "Sin cambios — es el mismo número."
    conn = get_db()
    existe = conn.execute("SELECT 1 FROM contratos WHERE ID_Contrato=?", (id_nuevo,)).fetchone()
    if existe:
        return False, f"Ya existe un contrato con el número {id_nuevo} — elige otro."
    try:
        conn.execute("UPDATE contratos SET ID_Contrato=? WHERE ID_Contrato=?", (id_nuevo, id_viejo))
        conn.execute("UPDATE facturas SET id_contrato=? WHERE id_contrato=?", (id_nuevo, id_viejo))
        conn.execute("UPDATE anotaciones SET ID_Contrato=? WHERE ID_Contrato=?", (id_nuevo, id_viejo))
        conn.execute("UPDATE eventos_especiales SET ID_Contrato=? WHERE ID_Contrato=?", (id_nuevo, id_viejo))
        conn.commit()
        _tocar_datos()
        return True, f"Contrato renumerado de {id_viejo} a {id_nuevo}."
    except Exception as e:
        conn.rollback()
        return False, f"No se pudo renumerar: {e}"

def eliminar_contrato_completo(id_c: str):
    """Elimina un contrato Y todo lo que le pertenece en otras tablas
    (facturas y sus conceptos, anotaciones, eventos especiales) — un
    'eliminar permanentemente' que antes solo borraba el renglón de
    contratos y dejaba huérfanos en las demás tablas, lo cual podía hacer
    que el contrato 'siguiera apareciendo' en pantallas que arman su lista
    de contratos a partir de esas otras tablas en vez de solo 'contratos'."""
    conn = get_db()
    try:
        uuids_factura = [r[0] for r in conn.execute(
            "SELECT uuid FROM facturas WHERE id_contrato=?", (id_c,)
        ).fetchall()]
        if uuids_factura:
            ph = ','.join('?' * len(uuids_factura))
            conn.execute(f"DELETE FROM factura_conceptos WHERE uuid IN ({ph})", uuids_factura)
        conn.execute("DELETE FROM facturas WHERE id_contrato=?", (id_c,))
        conn.execute("DELETE FROM anotaciones WHERE ID_Contrato=?", (id_c,))
        conn.execute("DELETE FROM eventos_especiales WHERE ID_Contrato=?", (id_c,))
        cur = conn.execute("DELETE FROM contratos WHERE ID_Contrato=?", (id_c,))
        borrado = cur.rowcount > 0
        conn.commit()
        _tocar_datos()
        return borrado
    except Exception as e:
        conn.rollback()
        raise

def guardar(c,sobreescribir=True):
    conn=get_db(); cur=conn.cursor()
    existe=cur.execute("SELECT 1 FROM contratos WHERE ID_Contrato=?",(c['ID_Contrato'],)).fetchone()
    if existe and not sobreescribir: return False
    v=c['Valor_Sin_IVA']; ant=c['Anticipo_Monto']; inv=v-ant
    r=c['Mensualidad_Sin_IVA']; res=c['Residual_Monto']; pl=c['Plazo']
    t=calc_tasa(pl,r,inv,res)
    if t is None:
        st.warning(
            f"No se pudo calcular la tasa del contrato **{c.get('ID_Contrato','(nuevo)')}** "
            "con los datos capturados (revisa plazo, renta, valor, anticipo y residual: "
            "deben formar un flujo financiero válido). Se guardó temporalmente con **Tasa = 0%**, "
            "lo cual haría ver sus intereses y rentabilidad en cero. Corrige y vuelve a guardar este contrato."
        )
        t = 0.0
    elif t < 0 or t > 0.15:
        # Rango de negocio razonable para leasing en México: una tasa mensual
        # implícita negativa (el arrendador pierde dinero por diseño) o mayor
        # a 15% mensual (~430% anual) casi siempre delata un error de captura
        # (un cero de más/de menos en el valor, el plazo o la renta), no un
        # contrato real. No se bloquea el guardado — puede ser intencional
        # en un caso legítimo poco común — pero se avisa para que se revise.
        st.warning(
            f"La tasa calculada para **{c.get('ID_Contrato','(nuevo)')}** es de "
            f"**{t*100:.2f}% mensual** ({(((1+t)**12)-1)*100:,.1f}% anual efectivo), "
            "fuera del rango típico de un contrato de leasing. Esto casi siempre indica "
            "un dato mal capturado (valor, anticipo, plazo o renta). Revisa el contrato; "
            "se guardó tal cual lo capturaste."
        )
    vpr=vp_res(res,t,pl)
    tipo_baja = c.get('Tipo_Baja', '')
    if c.get('Fecha_Baja') and c.get('Estatus') == 'BAJA':
        fb = pd.to_datetime(c['Fecha_Baja'])
        fv = pd.to_datetime(c['Fecha_Vencimiento'])
        if pd.isnull(fb):
            tipo_baja = tipo_baja or 'SIN_FECHA'
        elif fb < fv - pd.Timedelta(days=30):
            tipo_baja = tipo_baja or 'ANTICIPADA'
        else:
            tipo_baja = tipo_baja or 'NATURAL'
    cur.execute("""INSERT OR REPLACE INTO contratos
        (ID_Contrato,Cliente,Vehiculo,Fecha_Alta,Fecha_Vencimiento,Valor_Sin_IVA,Mensualidad_Sin_IVA,Plazo,
         Comision_Apertura_Pct,Comision_Monto,Anticipo_Pct,Anticipo_Monto,Residual_Pct,Residual_Monto,
         Tasa_Calculada,Estatus,Fecha_Baja,VP_Residual,residual_transferred,Nivel_Morosidad,
         Tipo_Baja,Motivo_Baja)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (c['ID_Contrato'],c['Cliente'],c['Vehiculo'],c['Fecha_Alta'].isoformat(),c['Fecha_Vencimiento'].isoformat(),
         v,r,pl,c['Comision_Apertura_Pct'],c['Comision_Monto'],c['Anticipo_Pct'],ant,c['Residual_Pct'],res,t,
         c['Estatus'],c['Fecha_Baja'].isoformat() if c['Fecha_Baja'] else None,vpr,
         c.get('residual_transferred',0),c.get('Nivel_Morosidad',0),
         tipo_baja, c.get('Motivo_Baja','')))
    conn.commit(); _tocar_datos(); return True

# Cacheo de obtener() — antes releía toda la tabla en cada clic y con
# carteras grandes se sentía lento. Se cachea, con dos seguros para nunca
# mostrar datos viejos: 1) _tocar_datos() invalida el caché apenas se
# guarda/edita/borra algo, y 2) por si acaso, el caché igual expira solo
# a los 45s. La clave incluye la ruta de la base activa para que al
# cambiar de empresa nunca se mezclen datos de una con otra.
_version_datos = {}

def _tocar_datos(db_path=None):
    db_path = db_path or get_db_path()
    _version_datos[db_path] = _version_datos.get(db_path, 0) + 1

@st.cache_data(ttl=45, max_entries=64, show_spinner=False)
def _obtener_cached(db_path, estatus, _version):
    conn=_engine(db_path); q="SELECT * FROM contratos"; p=[]
    if estatus: q+=" WHERE Estatus=?"; p.append(estatus)
    df=pd.read_sql_query(q,conn,params=p)
    if not df.empty:
        df['Fecha_Alta']=pd.to_datetime(df['Fecha_Alta'], format='mixed', errors='coerce')
        df['Fecha_Vencimiento']=pd.to_datetime(df['Fecha_Vencimiento'], format='mixed', errors='coerce')
        df['Fecha_Baja']=pd.to_datetime(df['Fecha_Baja'], format='mixed', errors='coerce')
        for col,val in [('residual_transferred',0),('Nivel_Morosidad',0),('VP_Residual',0.0)]:
            df[col]=df[col].fillna(val) if col in df.columns else val
        for col in ['Tipo_Baja','Motivo_Baja']:
            if col not in df.columns: df[col]=''
            else: df[col]=df[col].fillna('')
        for col in ['Valor_Sin_IVA','Mensualidad_Sin_IVA','Residual_Monto','Anticipo_Monto',
                    'Comision_Monto','Tasa_Calculada','Plazo','VP_Residual']:
            df[col]=pd.to_numeric(df[col],errors='coerce').fillna(0)
        mask_baja = (df['Estatus']=='BAJA') & (df['Tipo_Baja']=='')
        if mask_baja.any():
            def _tipo(r):
                if pd.isnull(r['Fecha_Baja']): return 'SIN_FECHA'
                return 'NATURAL' if r['Fecha_Baja'] >= r['Fecha_Vencimiento'] - pd.Timedelta(days=30) else 'ANTICIPADA'
            df.loc[mask_baja,'Tipo_Baja']=df[mask_baja].apply(_tipo,axis=1)
    return df

def obtener(estatus=None):
    db_path = get_db_path()
    version = _version_datos.get(db_path, 0)
    return _obtener_cached(db_path, estatus, version).copy()

MOTIVO_VENC_AUTO = "Vencimiento natural del plazo (automático)"

def procesar_vencimientos_naturales():
    conn = get_db()
    hoy = date.today().isoformat()
    pendientes = conn.execute(
        "SELECT ID_Contrato, Fecha_Vencimiento FROM contratos "
        "WHERE Estatus='ACTIVO' AND Fecha_Vencimiento<=?", (hoy,)
    ).fetchall()
    if not pendientes:
        return 0
    for r in pendientes:
        conn.execute(
            """UPDATE contratos
               SET Estatus='BAJA', Fecha_Baja=?, Tipo_Baja='NATURAL', Motivo_Baja=?
               WHERE ID_Contrato=?""",
            (r['Fecha_Vencimiento'], MOTIVO_VENC_AUTO, r['ID_Contrato'])
        )
    conn.commit()
    if pendientes: _tocar_datos()
    return len(pendientes)

# Anotaciones de contratos (para control contable)
def agregar_anotacion(id_c, texto, tipo="General"):
    conn=get_db()
    conn.execute("INSERT INTO anotaciones (ID_Contrato,Fecha,Tipo,Texto) VALUES (?,?,?,?)",
                 (id_c, datetime.now().strftime('%Y-%m-%d %H:%M'), tipo, texto.strip()))
    conn.commit()

def obtener_anotaciones(id_c=None):
    conn=get_db()
    if id_c:
        return pd.read_sql_query(
            "SELECT * FROM anotaciones WHERE ID_Contrato=? ORDER BY Fecha DESC",conn,params=(id_c,))
    return pd.read_sql_query("SELECT * FROM anotaciones ORDER BY Fecha DESC",conn)

def eliminar_anotacion(id_):
    conn=get_db(); conn.execute("DELETE FROM anotaciones WHERE id=?",(id_,)); conn.commit()

# Analytics
# rentabilidad() y tir() ahora viven en finanzas.py (importadas arriba).

def proy_residual(df):
    hoy=hoy_ref(); df2=df[df['Fecha_Vencimiento'].dt.date>hoy].copy()
    if df2.empty: return pd.DataFrame()
    df2['Mes']=df2['Fecha_Vencimiento'].dt.to_period('M')
    r=df2.groupby('Mes')['Residual_Monto'].sum().reset_index(); r['Mes']=r['Mes'].astype(str)
    perms=[(hoy+relativedelta(months=i)).strftime('%Y-%m') for i in range(1,25)]
    return r[r['Mes'].isin(perms)]

def proy_rentas(df,meses=24):
    hoy=hoy_ref(); rows=[]
    for i in range(1,meses+1):
        fm=pd.Timestamp(hoy+relativedelta(months=i))
        tot=df[(df['Fecha_Alta']<=fm)&(df['Fecha_Vencimiento']>=fm)]['Mensualidad_Sin_IVA'].sum()
        rows.append({'Mes':fm.strftime('%Y-%m'),'Rentas':tot})
    return pd.DataFrame(rows)

def calc_int_mes(mes,anio):
    """Intereses del mes — misma logica que Tabla Mensual / Dashboard / poliza con primer mes en firma.

    Universo: todos los contratos (obtener), no solo ACTIVO, para cuadrar con Tabla Mensual.
    """
    df = obtener()
    if df.empty:
        return pd.DataFrame(), {}
    mes = int(mes)
    anio = int(anio)
    rows = []
    for _, row in df.iterrows():
        mc = _mes_en_vigencia_contrato(row, anio, mes, inc_primer_mes=True)
        if mc is None:
            continue
        try:
            dfa, dfr, pl, t = _amort_tablas_contrato(row)
            fa_ = dfa.iloc[mc - 1]
            fr_ = dfr.iloc[mc - 1]
            rows.append({
                'ID_Contrato': row['ID_Contrato'],
                'Cliente': row['Cliente'],
                'Vehiculo': row.get('Vehiculo', ''),
                'Mes_Cont': mc,
                'Plazo': pl,
                'Renta': float(row['Mensualidad_Sin_IVA']),
                'Int_Leasing': round(float(fa_['Interes']), 2),
                'Capital': round(float(fa_['Capital']), 2),
                'Saldo_Cap': round(float(fa_['Saldo']), 2),
                'Int_Residual': round(float(fr_['Interes']), 2),
                'Amort_Com': round(float(row.get('Comision_Monto') or 0) / pl, 2) if pl else 0.0,
                'Tasa_Anual_Pct': round(float(row['Tasa_Calculada']) * 1200, 4),
            })
        except Exception:
            continue
    if not rows:
        return pd.DataFrame(), {}
    df2 = pd.DataFrame(rows)
    df2['Total_Int'] = df2['Int_Leasing'] + df2['Int_Residual']
    m = {
        'il': float(df2['Int_Leasing'].sum()),
        'ir': float(df2['Int_Residual'].sum()),
        'tot': float(df2['Total_Int'].sum()),
        'cap': float(df2['Capital'].sum()),
        'com': float(df2['Amort_Com'].sum()),
        'n': len(df2),
        'top': df2.groupby('Cliente')['Total_Int'].sum().idxmax() if len(df2) > 0 else 'N/A',
    }
    return df2.sort_values('Total_Int', ascending=False), m


def proy_intereses(meses=12):
    df=obtener('ACTIVO')
    if df.empty: return pd.DataFrame()
    hoy=hoy_ref(); rows=[]; bar=st.progress(0,"Proyectando intereses…")
    for i in range(meses):
        fd=hoy+relativedelta(months=i+1); mp,ap=fd.month,fd.year; fp=pd.Timestamp(ap,mp,1)
        ti=tr=0.0; cnt=0
        for _,row in df.iterrows():
            mc = _mes_en_vigencia_contrato(row, ap, mp, inc_primer_mes=True)
            if mc is None:
                continue
            try:
                dfa, dfr, pl, _t = _amort_tablas_contrato(row)
                ti += float(dfa.iloc[mc - 1]['Interes'])
                tr += float(dfr.iloc[mc - 1]['Interes'])
                cnt += 1
            except Exception:
                continue
        rows.append({'Mes':fd.strftime('%Y-%m'),'Label':fd.strftime('%b %Y'),'IL':round(ti,2),'IR':round(tr,2),'Total':round(ti+tr,2),'N':cnt})
        bar.progress((i+1)/meses,text=f"Mes {i+1}/{meses}…")
    bar.empty(); return pd.DataFrame(rows)

def tabla_rentas_mensuales(anio):
    """Rentas por mes con la misma vigencia que Tabla Mensual / poliza."""
    df = obtener('ACTIVO')
    if df.empty:
        return pd.DataFrame()
    MN = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
    data = []
    for _, row in df.iterrows():
        fila = {'ID_Contrato': row['ID_Contrato'], 'Cliente': row['Cliente']}
        for i, n in enumerate(MN, 1):
            mc = _mes_en_vigencia_contrato(row, anio, i, inc_primer_mes=True)
            fila[n] = float(row['Mensualidad_Sin_IVA']) if mc is not None else 0.0
        data.append(fila)
    return pd.DataFrame(data).set_index('ID_Contrato')


def _mes_en_vigencia_contrato(row, anio, mes, inc_primer_mes=True):
    """Misma regla que pol_parcialidad con Incluir mes de alta = True (primer mes en firma).

    Incluye el mes si coincide con poliza parcialidad:
      fa_m <= fm <= fv_m (por mes calendario)
      si BAJA: no posterior al mes de Fecha_Baja
      mc = mt+1 entre 1 y Plazo; mt==0 solo si inc_primer_mes
    Devuelve mc (1-based) o None.
    """
    try:
        fa = pd.to_datetime(row['Fecha_Alta'])
        fv = pd.to_datetime(row['Fecha_Vencimiento'])
    except Exception:
        return None
    if pd.isna(fa) or pd.isna(fv):
        return None
    anio = int(anio)
    mes = int(mes)
    fm = pd.Timestamp(anio, mes, 1)
    fa_m = pd.Timestamp(int(fa.year), int(fa.month), 1)
    fv_m = pd.Timestamp(int(fv.year), int(fv.month), 1)
    if fa_m > fm or fv_m < fm:
        return None
    es_baja = str(row.get('Estatus', '') or '').upper() == 'BAJA'
    if es_baja and pd.notna(row.get('Fecha_Baja')):
        try:
            fb = pd.to_datetime(row['Fecha_Baja'])
            if not pd.isna(fb):
                fb_m = pd.Timestamp(int(fb.year), int(fb.month), 1)
                if fb_m < fm:
                    return None
        except Exception:
            pass
    mt = (anio - int(fa.year)) * 12 + (mes - int(fa.month))
    if mt == 0 and not inc_primer_mes:
        return None
    mc = mt + 1
    try:
        pl = int(row['Plazo'])
    except Exception:
        return None
    if mc < 1 or mc > pl:
        return None
    return mc


def _amort_tablas_contrato(row):
    """dfa (leasing) y dfr (residual) — mismos inputs que polizas."""
    pl = int(row['Plazo'])
    t = round(float(row['Tasa_Calculada']), 8)
    inv = float(row['Valor_Sin_IVA']) - float(row['Anticipo_Monto'])
    dfa, _, _, _, _ = calc_amort(
        round(inv, 4),
        round(float(row['Mensualidad_Sin_IVA']), 4),
        round(float(row['Residual_Monto']), 4),
        pl, t,
    )
    try:
        vpr = float(row['VP_Residual'])
    except Exception:
        vpr = 0.0
    if not vpr or vpr <= 0:
        vpr = float(vp_res(float(row['Residual_Monto']), t, pl))
    dfr = calc_res_amort(round(vpr, 4), t, pl)
    return dfa, dfr, pl, t


@st.cache_data(ttl=60, show_spinner=False)
def _tabla_mensual_conceptos_cached(db_path, anio, _version):
    df = _obtener_cached(db_path, None, _version)
    if df.empty:
        return None, None, None, None, "Sin contratos registrados"
    MN = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
    anio = int(anio)
    ids = df['ID_Contrato'].tolist()
    di = pd.DataFrame(index=ids, columns=range(1, 13), dtype=float)
    dr = di.copy()
    dc = di.copy()
    ds = di.copy()
    for _, row in df.iterrows():
        id_c = row['ID_Contrato']
        try:
            dfa, dfr, pl, _t = _amort_tablas_contrato(row)
        except Exception:
            continue
        try:
            com_mes = round(float(row.get('Comision_Monto') or 0) / pl, 2) if pl else 0.0
        except Exception:
            com_mes = 0.0
        for mes in range(1, 13):
            mc = _mes_en_vigencia_contrato(row, anio, mes, inc_primer_mes=True)
            if mc is None:
                continue
            try:
                di.at[id_c, mes] = round(float(dfa.iloc[mc - 1]['Interes']), 2)
                dr.at[id_c, mes] = round(float(dfr.iloc[mc - 1]['Interes']), 2)
                dc.at[id_c, mes] = com_mes
                ds.at[id_c, mes] = round(float(dfr.iloc[mc - 1]['Saldo_Fin']), 2)
            except Exception:
                continue
    for d in [di, dr, dc, ds]:
        d.columns = MN
        d.dropna(how='all', inplace=True)
    return di, dr, dc, ds, None


def tabla_mensual_conceptos(anio):
    """Fuente de verdad unica: intereses leasing, residual, comision, saldo residual."""
    db_path = get_db_path()
    version = _version_datos.get(db_path, 0)
    return _tabla_mensual_conceptos_cached(db_path, int(anio), version)



def totales_intereses_desde_tabla_mensual(anio, mes):
    """Suma columnas de tabla_mensual_conceptos (verdad = Excel Tabla Mensual).

    Incluye:
      - Intereses leasing (di) — Excel int_208
      - Intereses residual (dr)
      - Amort. comision por apertura (dc) — Excel amort_comision
      - Saldo residual activo (ds)

    Del mes = columna del mes. Acumulado = suma enero..mes.
    """
    MN = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
    mes = int(mes)
    anio = int(anio)
    out = {
        'interes_leasing_mes': 0.0,
        'interes_residual_mes': 0.0,
        'comision_mes': 0.0,
        'saldo_residual_mes': 0.0,
        'interes_leasing_ytd': 0.0,
        'interes_residual_ytd': 0.0,
        'comision_ytd': 0.0,
        'por_mes_leasing': {},
        'por_mes_residual': {},
        'por_mes_comision': {},
        'n_contratos_mes': 0,
        'anio': anio,
        'mes': mes,
        'error': None,
    }
    di, dr, dc, ds, err = tabla_mensual_conceptos(anio)
    if err:
        out['error'] = err
        return out
    if di is None or di.empty:
        out['error'] = 'Sin datos en tabla mensual'
        return out
    for m_idx, nombre in enumerate(MN, start=1):
        if nombre in di.columns:
            out['por_mes_leasing'][m_idx] = round(float(di[nombre].sum(skipna=True)), 2)
        if dr is not None and not dr.empty and nombre in dr.columns:
            out['por_mes_residual'][m_idx] = round(float(dr[nombre].sum(skipna=True)), 2)
        if dc is not None and not dc.empty and nombre in dc.columns:
            out['por_mes_comision'][m_idx] = round(float(dc[nombre].sum(skipna=True)), 2)
    nombre_mes = MN[mes - 1]
    if nombre_mes in di.columns:
        out['interes_leasing_mes'] = round(float(di[nombre_mes].sum(skipna=True)), 2)
        out['n_contratos_mes'] = int(di[nombre_mes].notna().sum())
    if dr is not None and not dr.empty and nombre_mes in dr.columns:
        out['interes_residual_mes'] = round(float(dr[nombre_mes].sum(skipna=True)), 2)
    if dc is not None and not dc.empty and nombre_mes in dc.columns:
        out['comision_mes'] = round(float(dc[nombre_mes].sum(skipna=True)), 2)
    if ds is not None and not ds.empty and nombre_mes in ds.columns:
        out['saldo_residual_mes'] = round(float(ds[nombre_mes].sum(skipna=True)), 2)
    out['interes_leasing_ytd'] = round(sum(out['por_mes_leasing'].get(m, 0.0) for m in range(1, mes + 1)), 2)
    out['interes_residual_ytd'] = round(sum(out['por_mes_residual'].get(m, 0.0) for m in range(1, mes + 1)), 2)
    out['comision_ytd'] = round(sum(out['por_mes_comision'].get(m, 0.0) for m in range(1, mes + 1)), 2)
    return out


def exportar_maestro_completo(anio, di, dcap, drenta, dr, dc, ds, total_cap, total_int, total_renta):
    import io
    import pandas as pd
    from reports.excel import excel_con_formato
    
    # Crear un dataframe para el resumen ejecutivo
    MN = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
    df_resumen = pd.DataFrame({
        'Mes': MN,
        'Capital Amortizado': total_cap.values,
        'Intereses Devengados': total_int.values,
        'Renta Neta (Flujo Total)': total_renta.values
    })
    
    # Preparamos las hojas completas
    # Formateamos un poco los de data quitando index si lo tienen
    skip_currency = ['ID_Contrato', 'Cliente', 'Vehiculo', 'Estatus', 'Fecha_Alta', 'Plazo', 'Fecha_Baja', 'Tasa_Anual_%']
    
    def pre(df):
        return df.reset_index().rename(columns={'index': 'ID_Contrato'})
        
    d1 = pre(dcap)
    d2 = pre(di)
    d3 = pre(drenta)
    d4 = pre(dr)
    d5 = pre(dc)
    d6 = pre(ds)
    
    curr_cols_1 = [c for c in d1.columns if c not in skip_currency]
    
    hojas = {
        'Resumen Ejecutivo': df_resumen,
        'Capital Leasing': d1,
        'Intereses Leasing': d2,
        'Renta Neta': d3,
        'Int. Residual': d4,
        'Amort. Comision': d5,
        'Saldo Residual': d6
    }
    
    # Diccionarios de formato
    c_cols = {
        'Resumen Ejecutivo': ['Capital Amortizado', 'Intereses Devengados', 'Renta Neta (Flujo Total)'],
        'Capital Leasing': curr_cols_1,
        'Intereses Leasing': curr_cols_1,
        'Renta Neta': curr_cols_1,
        'Int. Residual': curr_cols_1,
        'Amort. Comision': curr_cols_1,
        'Saldo Residual': curr_cols_1
    }
    
    p_cols = {
        'Capital Leasing': ['Tasa_Anual_%'],
        'Intereses Leasing': ['Tasa_Anual_%'],
        'Renta Neta': ['Tasa_Anual_%'],
        'Int. Residual': ['Tasa_Anual_%'],
        'Amort. Comision': ['Tasa_Anual_%'],
        'Saldo Residual': ['Tasa_Anual_%']
    }
    
    # Llamar al generador base (que ya le pone logo y colores corporativos a la tabla)
    buf = excel_con_formato(hojas, currency_cols=c_cols, pct_cols=p_cols)
    
    # AHORA VAMOS A ABRIR ESE BUFFER CON OPENPYXL PARA METER LA GRÁFICA
    import openpyxl
    from openpyxl.chart import BarChart, LineChart, Reference, Series
    from openpyxl.chart.axis import DateAxis
    
    buf.seek(0)
    wb = openpyxl.load_workbook(buf)
    
    # Obtener hoja de resumen
    ws = wb['Resumen Ejecutivo']
    
    # Crear gráfica apilada (Capital + Intereses)
    from openpyxl.chart import BarChart, Reference
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Proyección de Flujo por Mes"
    chart.y_axis.title = "Ingresos ($)"
    chart.x_axis.title = "Mes"
    chart.grouping = "clustered"
    chart.gapWidth = 150
    chart.overlap = 0
    
    data = Reference(ws, min_col=2, min_row=3, max_col=4, max_row=15)
    cats = Reference(ws, min_col=1, min_row=4, max_row=15)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    
    ws.add_chart(chart, "F4")
    chart.width = 25
    chart.height = 14
    
    # Guardar en un nuevo buffer
    out_buf = io.BytesIO()
    wb.save(out_buf)
    out_buf.seek(0)
    return out_buf


def reporte_maestro_saldos(anio):
    import numpy as np
    df=obtener()
    if df.empty: return None,None,None,None,"Sin contratos registrados"
    MN=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
    ids=df['ID_Contrato'].tolist()
    d_cap = pd.DataFrame(index=ids,columns=range(1,13),dtype=float)
    d_int = d_cap.copy()
    d_tot = d_cap.copy()
    d_res = d_cap.copy()
    bar=st.progress(0,"Calculando Reporte de Saldos...")
    tot=len(df)
    for i,(_,row) in enumerate(df.iterrows()):
        id_c=row['ID_Contrato']
        try:
            dfa, dfr, pl, _t = _amort_tablas_contrato(row)
        except Exception:
            continue
        interes_array = dfa['Interes'].values
        rem_int = np.zeros(pl)
        for idx in range(pl):
            rem_int[idx] = np.sum(interes_array[idx+1:])
            
        for mes in range(1,13):
            mc = _mes_en_vigencia_contrato(row, anio, mes, inc_primer_mes=True)
            if mc is None:
                # Fuera de vigencia: si ya empezo el contrato, saldos en 0; si aun no, vacio
                try:
                    fa_chk = pd.to_datetime(row['Fecha_Alta'])
                    fm = pd.Timestamp(anio, mes, 1)
                    if fm >= pd.Timestamp(int(fa_chk.year), int(fa_chk.month), 1):
                        d_cap.at[id_c,mes] = 0; d_int.at[id_c,mes] = 0; d_tot.at[id_c,mes] = 0; d_res.at[id_c,mes] = 0
                except Exception:
                    pass
                continue
            idx = mc - 1
            try:
                saldo_cap = float(dfa.iloc[idx]['Saldo'])
                saldo_int = float(rem_int[idx])
                saldo_res = float(dfr.iloc[idx]['Saldo_Fin'])
                d_cap.at[id_c,mes] = round(saldo_cap, 2)
                d_int.at[id_c,mes] = round(saldo_int, 2)
                d_tot.at[id_c,mes] = round(saldo_cap + saldo_int, 2)
                d_res.at[id_c,mes] = round(saldo_res, 2)
            except Exception:
                continue
            
        bar.progress((i+1)/tot,text=f"Procesando {i+1}/{tot}...")
    bar.empty()
    
    df_meta = df.set_index('ID_Contrato')[['Cliente', 'Vehiculo', 'Estatus', 'Fecha_Alta', 'Plazo', 'Fecha_Baja', 'Valor_Sin_IVA', 'Tasa_Calculada']].copy()
    df_meta['Fecha_Alta'] = df_meta['Fecha_Alta'].dt.strftime('%Y-%m-%d')
    df_meta['Fecha_Baja'] = df_meta['Fecha_Baja'].dt.strftime('%Y-%m-%d').fillna('')
    df_meta['Tasa_Anual_%'] = (df_meta['Tasa_Calculada'] * 1200).round(2)
    df_meta.drop(columns=['Tasa_Calculada'], inplace=True)
    df_meta['Valor_Sin_IVA'] = df_meta['Valor_Sin_IVA'].round(2)
    
    out = []
    for d in [d_cap, d_int, d_tot, d_res]: 
        d.columns=MN
        d.dropna(how='all',inplace=True)
        d_merged = df_meta.join(d, how='right')
        
        # Agregar Columna TOTAL AÑO
        d_merged['Total Año'] = d_merged[MN].sum(axis=1)
        
        # Agregar Fila TOTAL CARTERA
        total_row = d_merged[MN + ['Valor_Sin_IVA', 'Total Año']].sum()
        total_row['Cliente'] = 'TOTAL CARTERA'
        d_merged.loc['TOTAL_000'] = total_row  # Use TOTAL_000 so it can sort or be distinct
        
        out.append(d_merged)
        
    return out[0], out[1], out[2], out[3], None


def exportar_saldos_completo(anio, dcap, dint, dtot, dres, avg_cap, avg_int, avg_tot):
    import io
    import pandas as pd
    from reports.excel import excel_con_formato
    import openpyxl
    from openpyxl.chart import AreaChart, Reference
    
    MN = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
    df_resumen = pd.DataFrame({
        'Mes': MN,
        'Saldo Capital': avg_cap.values,
        'Saldo Intereses': avg_int.values,
        'Saldo Total': avg_tot.values
    })
    
    skip_currency = ['ID_Contrato', 'Cliente', 'Vehiculo', 'Estatus', 'Fecha_Alta', 'Plazo', 'Fecha_Baja', 'Tasa_Anual_%']
    def pre(df): return df.reset_index().rename(columns={'index': 'ID_Contrato'})
        
    d1 = pre(dcap); d2 = pre(dint); d3 = pre(dtot); d4 = pre(dres)
    curr_cols_1 = [c for c in d1.columns if c not in skip_currency]
    
    hojas = {
        'Resumen Saldos': df_resumen,
        'Saldo Capital': d1,
        'Saldo Intereses': d2,
        'Saldo Total': d3,
        'Saldo Residual': d4
    }
    
    c_cols = {
        'Resumen Saldos': ['Saldo Capital', 'Saldo Intereses', 'Saldo Total'],
        'Saldo Capital': curr_cols_1,
        'Saldo Intereses': curr_cols_1,
        'Saldo Total': curr_cols_1,
        'Saldo Residual': curr_cols_1
    }
    
    p_cols = {
        'Saldo Capital': ['Tasa_Anual_%'],
        'Saldo Intereses': ['Tasa_Anual_%'],
        'Saldo Total': ['Tasa_Anual_%'],
        'Saldo Residual': ['Tasa_Anual_%']
    }
    
    buf = excel_con_formato(hojas, currency_cols=c_cols, pct_cols=p_cols)
    buf.seek(0)
    wb = openpyxl.load_workbook(buf)
    ws = wb['Resumen Saldos']
    
    from openpyxl.chart import BarChart, Reference
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Saldos Pendientes por Mes"
    chart.y_axis.title = "Monto ($)"
    chart.x_axis.title = "Mes"
    chart.grouping = "clustered"
    chart.gapWidth = 150
    chart.overlap = 0
    
    data = Reference(ws, min_col=2, min_row=3, max_col=4, max_row=15)
    cats = Reference(ws, min_col=1, min_row=4, max_row=15)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    
    # Opcional: Para mostrar las etiquetas de datos encima de las barras de Total (serie 3)
    # En openpyxl es complejo, as que dejaremos la grfica super limpia.
    
    ws.add_chart(chart, "F4")
    chart.width = 25
    chart.height = 14
    
    out_buf = io.BytesIO()
    wb.save(out_buf)
    out_buf.seek(0)
    return out_buf


def reporte_maestro_mensual(anio):
    df=obtener()
    if df.empty: return None,None,None,None,None,None,"Sin contratos registrados"
    MN=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
    ids=df['ID_Contrato'].tolist()
    di=pd.DataFrame(index=ids,columns=range(1,13),dtype=float)
    dr=di.copy(); dc=di.copy(); ds=di.copy(); dcap=di.copy(); drenta=di.copy()
    bar=st.progress(0,"Calculando Reporte Maestro...")
    tot=len(df)
    for i,(_,row) in enumerate(df.iterrows()):
        id_c=row['ID_Contrato']
        try:
            dfa, dfr, pl, _t = _amort_tablas_contrato(row)
        except Exception:
            bar.progress((i+1)/tot,text=f"Procesando {i+1}/{tot}...")
            continue
        try:
            com_mes = round(float(row.get('Comision_Monto') or 0) / pl, 2) if pl else 0.0
        except Exception:
            com_mes = 0.0
        for mes in range(1,13):
            mc = _mes_en_vigencia_contrato(row, anio, mes, inc_primer_mes=True)
            if mc is None:
                continue
            try:
                interes_leasing = round(float(dfa.iloc[mc-1]['Interes']), 2)
                capital_leasing = round(float(dfa.iloc[mc-1]['Capital']), 2)
                di.at[id_c,mes] = interes_leasing
                dcap.at[id_c,mes] = capital_leasing
                drenta.at[id_c,mes] = interes_leasing + capital_leasing
                dr.at[id_c,mes] = round(float(dfr.iloc[mc-1]['Interes']), 2)
                dc.at[id_c,mes] = com_mes
                ds.at[id_c,mes] = round(float(dfr.iloc[mc-1]['Saldo_Fin']), 2)
            except Exception:
                continue
        bar.progress((i+1)/tot,text=f"Procesando {i+1}/{tot}...")
    bar.empty()
    
    # Agregar metadatos descriptivos a cada dataframe
    df_meta = df.set_index('ID_Contrato')[['Cliente', 'Vehiculo', 'Estatus', 'Fecha_Alta', 'Plazo', 'Fecha_Baja', 'Valor_Sin_IVA', 'Tasa_Calculada']].copy()
    # Formatear algunos metadatos para que se vean bien en excel
    df_meta['Fecha_Alta'] = df_meta['Fecha_Alta'].dt.strftime('%Y-%m-%d')
    df_meta['Fecha_Baja'] = df_meta['Fecha_Baja'].dt.strftime('%Y-%m-%d').fillna('')
    df_meta['Tasa_Anual_%'] = (df_meta['Tasa_Calculada'] * 1200).round(2)
    df_meta.drop(columns=['Tasa_Calculada'], inplace=True)
    df_meta['Valor_Sin_IVA'] = df_meta['Valor_Sin_IVA'].round(2)
    

    out = []
    for d in [di,dcap,drenta,dr,dc,ds]: 
        d.columns=MN
        d.dropna(how='all',inplace=True)
        d_merged = df_meta.join(d, how='right')
        
        # Columna Total
        d_merged['Total Año'] = d_merged[MN].sum(axis=1)
        # Fila Total
        total_row = d_merged[MN + ['Valor_Sin_IVA', 'Total Año']].sum()
        total_row['Cliente'] = 'TOTAL CARTERA'
        d_merged.loc['TOTAL_000'] = total_row
        
        out.append(d_merged)
        
    return out[0], out[1], out[2], out[3], out[4], out[5], None


def perdidas_cesion(df_b,cat):
    rows=[]
    for _,c in df_b.iterrows():
        dfp=pol_inicial(c,cat)
        if dfp.empty: continue
        m=dfp['Descripcion'].str.contains('Pérdida cesión|Utilidad cesión',case=False)
        if m.any():
            rows.append({'Mes':c['Fecha_Baja'].strftime('%Y-%m') if pd.notnull(c['Fecha_Baja']) else 'N/A',
                         'Perdida':dfp[m]['Cargo'].sum()-dfp[m]['Abono'].sum()})
    dfd=pd.DataFrame(rows)
    if not dfd.empty: return dfd.groupby('Mes')['Perdida'].sum().reset_index()
    return pd.DataFrame(columns=['Mes','Perdida'])

# Exportaciones
# formatear_hoja_excel / excel_con_formato / _grafica_amort_para_pdf /
# _pie_pdf_marca / pdf_estado_cuenta / pdf_poliza viven ahora en reports/
# (ver docs/PLAN_REFACTOR_CAPAS.md, fase 1) — se importan abajo, junto con
# los helpers de main.py, para no romper ninguna llamada existente.
def backup_db():
    p=get_db_path()
    if os.path.exists(p):
        with open(p,'rb') as f: return f.read()
    return None

def restore_db(uf):
    try:
        with open(get_db_path(),'wb') as f: f.write(uf.getbuffer())
        _tocar_datos()
        return True
    except: return False

def export_excel():
    df=obtener()
    return excel_con_formato(
        {'Contratos': df},
        currency_cols=['Valor_Sin_IVA','Mensualidad_Sin_IVA','Anticipo_Monto','Residual_Monto','Comision_Monto'],
        pct_cols=['Anticipo_Pct','Residual_Pct','Comision_Apertura_Pct'],
        int_cols=['Plazo','Nivel_Morosidad'],
    )

def plantilla():
    df_plantilla = pd.DataFrame({
        'Contrato': ['0554'],
        'Anexo': ['0002'],
        'Cliente': ['Ejemplo S.A. de C.V.'],
        'Fecha De Apertura': [datetime.today().strftime('%Y-%m-%d')],
        'Plazo': [48],
        'Marca': ['RAM'],
        'Versión': ['LIMITED'],
        'Modelo': ['1500'],
        'Valor Cotización': [1600800],
        'Renta': [44483],
        '% Anticipo': [20.0],
        '% Comisión': [2.0],
        'Valor Residual (%)': [10.0],
        'Agencia': ['MOTORMEXA'],
        'Status': ['ACTIVO']
    })
    return excel_con_formato(
        {'Carga': df_plantilla},
        currency_cols=['Valor Cotización','Renta'],
        pct_cols=['% Anticipo','% Comisión','Valor Residual (%)'],
        incluir_titulo=False,
    )

def fmt17(base,id_c=None):
    cfg = get_config_codificacion_cuentas()
    if id_c:
        return cta_sg(base, id_c, cfg=cfg)
    dbase = int(cfg.get('digitos_base', 7))
    total = dbase + int(cfg.get('digitos_contrato', 8)) + len(cfg.get('sufijo',''))
    return base.replace('-','').ljust(total, '0')

def gen_contpaqi():
    df=obtener()
    if df.empty: return None,"Sin contratos"
    cat=cargar_catalogo(); cuentas=set()
    # Cuentas que el usuario dejó en blanco (o desactivó) en "Cuentas (Macro)":
    # se excluyen a propósito y NO deben aparecer en el archivo de Contpaqi.
    omitidas = [k for k,d in cat.items() if not (d.get('cuenta') or '').strip() or not d.get('activa', True)]
    for _,r in df.iterrows():
        id_c=r['ID_Contrato']; cli=r['Cliente']; plz=int(r['Plazo'])
        for k,d in cat.items():
            if k=='CANCELACION_ANTICIPO': continue
            if k in omitidas: continue
            cta=fmt17(d['cuenta'],id_c); nom=f"{id_c} {cli}"[:49]
            ps=d['cuenta'].split('-'); bm=f"{ps[0]}-{ps[1]}-{plz:02d}-00" if len(ps)>=2 else d['cuenta']
            cuentas.add((cta,nom,fmt17(bm)))
    if not cuentas:
        return None, "Todas las cuentas del catálogo están en blanco o desactivadas — no hay nada que generar. Ve a 'Cuentas (Macro)' y captura al menos una cuenta."
    rows=[[''] * 17]*4
    for cta,nom,mayor in cuentas:
        p1=cta[0]; tp,qv=('A','105.01') if p1=='1' else ('D','201.01') if p1=='2' else ('H','401.38') if p1=='4' else ('G','601.84')
        rows.append(['C',"'"+cta,nom[:49],'','\''+mayor,tp,'0','2','0','20250428','11','1','0','0','0','0',qv])
    out=io.StringIO(); pd.DataFrame(rows).to_csv(out,index=False,header=False,encoding='utf-8-sig')
    msg = f"Se generaron {len(cuentas)} cuentas."
    if omitidas:
        msg += f" ({len(omitidas)} claves del catálogo se omitieron por estar en blanco/desactivadas: {', '.join(omitidas)})"
    return out.getvalue().encode('utf-8-sig'), msg



# --- Conciliación de facturas CFDI: funciones de apoyo ---
# El parseo/clasificación puro del XML vive en core/cfdi.py (sin Streamlit,
# probado con pytest en test_cfdi.py). Aquí solo queda lo que sí depende de
# la app: resolver las reglas configurables desde la BD de la empresa activa.
from core.cfdi import (
    parse_cfdi, parse_sat_excel, clasificar_concepto, normalizar_contrato, normalizar_folio,
    REGLAS_CONCEPTO_DEFAULT, MAX_XML_BYTES, NS,
)
from core.alias import resolver_numero_contrato, decidir_accion_alias, sugerir_contrato_similar
# --- Cartera contable: import tolerante (no tumba la app si el core esta viejo) ---
def _fallback_comparar_auxiliar(df_aux, totales_sistema, mapa_columnas=None):
    return pd.DataFrame(
        [{"Concepto_sistema": k, "Monto_sistema": v, "Monto_auxiliar": None, "Diferencia": None}
         for k, v in (totales_sistema or {}).items() if isinstance(v, (int, float))]
    )

def _fallback_cobertura(esperado, facturado):
    esp = float(esperado or 0); fac = float(facturado or 0)
    dif = round(fac - esp, 2)
    pct = round((fac / esp * 100) if esp else 0.0, 2)
    return {"esperado": esp, "facturado": fac, "diferencia": dif, "cobertura_pct": pct}

def _fallback_acumulado_facturacion(df_fact, anio=None, hasta_periodo=None):
    out = {"n_facturas": 0, "total": 0.0, "subtotal": 0.0, "mensual": 0.0, "comision": 0.0, "otros": 0.0}
    if df_fact is None or getattr(df_fact, "empty", True):
        return out
    df = df_fact.copy()
    if "cancelada" in df.columns:
        df = df[df["cancelada"].fillna(0).astype(int) == 0]
    if anio is not None and "periodo" in df.columns:
        df = df[df["periodo"].astype(str).str.startswith(str(anio))]
    if hasta_periodo and "periodo" in df.columns:
        df = df[df["periodo"].astype(str) <= str(hasta_periodo)]
    if df.empty:
        return out
    tipo = df["tipo"].astype(str).str.upper() if "tipo" in df.columns else pd.Series([""] * len(df))
    total = float(df["total"].sum()) if "total" in df.columns else 0.0
    sub = float(df["subtotal"].sum()) if "subtotal" in df.columns else total
    return {
        "n_facturas": int(len(df)), "total": round(total, 2), "subtotal": round(sub, 2),
        "mensual": round(float(df.loc[tipo == "MENSUAL", "total"].sum()) if "total" in df.columns else 0.0, 2),
        "comision": round(float(df.loc[tipo == "COMISION", "total"].sum()) if "total" in df.columns else 0.0, 2),
        "otros": 0.0,
    }

def _fallback_metricas_estilo_tabla_mensual(df_contratos, anio, mes):
    """Delega a totales_intereses_desde_tabla_mensual (misma verdad que Dashboard)."""
    try:
        tot = totales_intereses_desde_tabla_mensual(int(anio), int(mes))
        return {
            "interes_leasing_mes": tot.get("interes_leasing_mes", 0),
            "interes_residual_mes": tot.get("interes_residual_mes", 0),
            "capital_mes": 0.0,
            "renta_mes": 0.0,
            "saldo_residual_mes": tot.get("saldo_residual_mes", 0),
            "saldo_capital_mes": 0.0,
            "interes_leasing_ytd": tot.get("interes_leasing_ytd", 0),
            "interes_residual_ytd": tot.get("interes_residual_ytd", 0),
            "capital_ytd": 0.0,
            "renta_ytd": 0.0,
            "n_contratos_mes": tot.get("n_contratos_mes", 0),
            "n_celdas_ytd": 0,
            "anio": int(anio),
            "mes": int(mes),
        }
    except Exception:
        return {
            "interes_leasing_mes": 0.0, "interes_residual_mes": 0.0, "capital_mes": 0.0, "renta_mes": 0.0,
            "saldo_residual_mes": 0.0, "saldo_capital_mes": 0.0,
            "interes_leasing_ytd": 0.0, "interes_residual_ytd": 0.0, "capital_ytd": 0.0, "renta_ytd": 0.0,
            "n_contratos_mes": 0, "n_celdas_ytd": 0, "anio": int(anio), "mes": int(mes),
        }


def _fallback_consolidar(df_activos, hoy=None):
    hoy = hoy or date.today()
    vacios = {"n_contratos": 0, "inversion_neta": 0.0, "saldo_capital": 0.0, "cxc_cp": 0.0, "cxc_lp": 0.0,
              "cxc_total": 0.0, "renta_mensual": 0.0, "saldo_residual_activo": 0.0}
    if df_activos is None or getattr(df_activos, "empty", True):
        return vacios, pd.DataFrame()
    # minimo: sumar rentas
    renta = float(pd.to_numeric(df_activos.get("Mensualidad_Sin_IVA", 0), errors="coerce").fillna(0).sum())
    vacios["n_contratos"] = len(df_activos)
    vacios["renta_mensual"] = round(renta, 2)
    return vacios, df_activos.copy()

def _fallback_acumulado_amort(df_activos, hoy=None, solo_ejercicio=True):
    hoy = hoy or date.today()
    m = _fallback_metricas_estilo_tabla_mensual(df_activos, hoy.year, hoy.month)
    return {
        "intereses_devengados": m["interes_leasing_ytd"],
        "intereses_residual_devengados": m["interes_residual_ytd"],
        "capital_amortizado": m["capital_ytd"],
        "rentas_esperadas_acum": m["renta_ytd"],
        "intereses_del_mes": m["interes_leasing_mes"],
        "intereses_residual_del_mes": m["interes_residual_mes"],
        "capital_del_mes": m["capital_mes"],
        "n_contratos": m["n_contratos_mes"],
    }

def _fallback_resumen_fact_anio(df_fact):
    return pd.DataFrame()

def _fallback_saldos_contrato(con, hoy=None):
    return {"inversion_neta": 0.0, "saldo_capital": 0.0, "renta_mensual": float(con.get("Mensualidad_Sin_IVA") or 0)}

_cc_err = None
try:
    import core.cartera_contable as _cc
    consolidar_cartera_vigente = getattr(_cc, "consolidar_cartera_vigente", _fallback_consolidar)
    resumen_facturacion_por_anio = getattr(_cc, "resumen_facturacion_por_anio", _fallback_resumen_fact_anio)
    cobertura_facturacion_periodo = getattr(_cc, "cobertura_facturacion_periodo", _fallback_cobertura)
    saldos_contrato_vigente = getattr(_cc, "saldos_contrato_vigente", _fallback_saldos_contrato)
    acumulado_facturacion = getattr(_cc, "acumulado_facturacion", _fallback_acumulado_facturacion)
    acumulado_amortizacion_cartera = getattr(_cc, "acumulado_amortizacion_cartera", _fallback_acumulado_amort)
    metricas_estilo_tabla_mensual = getattr(_cc, "metricas_estilo_tabla_mensual", _fallback_metricas_estilo_tabla_mensual)
    tabla_intereses_mensual = getattr(_cc, "tabla_intereses_mensual", None)
    comparar_auxiliar = getattr(_cc, "comparar_auxiliar", _fallback_comparar_auxiliar)
except Exception as _e_cc:
    _cc_err = _e_cc
    consolidar_cartera_vigente = _fallback_consolidar
    resumen_facturacion_por_anio = _fallback_resumen_fact_anio
    cobertura_facturacion_periodo = _fallback_cobertura
    saldos_contrato_vigente = _fallback_saldos_contrato
    acumulado_facturacion = _fallback_acumulado_facturacion
    acumulado_amortizacion_cartera = _fallback_acumulado_amort
    metricas_estilo_tabla_mensual = _fallback_metricas_estilo_tabla_mensual
    comparar_auxiliar = _fallback_comparar_auxiliar

TOLERANCIA = 0.10

def get_reglas_concepto() -> dict:
    raw = get_cfg('reglas_concepto_json')
    reglas = {k: list(v) for k, v in REGLAS_CONCEPTO_DEFAULT.items()}
    if raw:
        try:
            data = json.loads(raw)
            for k in REGLAS_CONCEPTO_DEFAULT:
                if k in data and isinstance(data[k], list):
                    reglas[k] = [str(p).strip() for p in data[k] if str(p).strip()]
        except Exception:
            pass
    return reglas

def set_reglas_concepto(reglas: dict):
    limpio = {k: [str(p).strip().upper() for p in v if str(p).strip()] for k, v in reglas.items()}
    set_cfg('reglas_concepto_json', json.dumps(limpio, ensure_ascii=False))

# --- Reglas de contrato facturado con otro número (alias) ---
# Estas cuatro funciones son todo lo que necesita el resto de la app para
# aprender y aplicar la equivalencia "el cliente factura como X, en
# realidad es el contrato Y" — quedan juntas para que sea fácil ver de un
# vistazo toda la lógica de esta funcionalidad.

def cargar_alias_contrato() -> dict:
    """numero_facturado -> id_contrato_real. Se resuelve una sola vez antes
    de procesar un lote (igual que `reglas`), no archivo por archivo."""
    conn = get_db()
    filas = conn.execute("SELECT numero_facturado, id_contrato_real FROM contrato_alias").fetchall()
    return {f['numero_facturado']: f['id_contrato_real'] for f in filas}

def registrar_alias_contrato(numero_facturado: str, id_contrato_real: str):
    """Guarda (o corrige) la regla numero_facturado -> id_contrato_real, y
    deja rastro en el historial de auditoría. La decisión de si hay algo
    que guardar (y si es una regla nueva o una corrección) la toma
    `decidir_accion_alias` (core/alias.py, sin tocar la BD) — aquí solo se
    ejecuta esa decisión, para que la regla de 'nunca duplicar' esté en un
    solo lugar y sea la misma que se prueba con pytest."""
    conn = get_db()
    existente = conn.execute(
        "SELECT id_contrato_real FROM contrato_alias WHERE numero_facturado=?", (numero_facturado,)
    ).fetchone()
    id_existente = existente['id_contrato_real'] if existente else None
    accion = decidir_accion_alias(numero_facturado, id_contrato_real, id_existente)
    if accion is None:
        return
    detalle = f"Antes apuntaba a {id_existente}" if accion == 'CORREGIDA' else "Regla nueva, aprendida al resolver una factura a mano"
    conn.execute("""
        INSERT INTO contrato_alias (numero_facturado, id_contrato_real, fecha_creacion)
        VALUES (?,?,CURRENT_TIMESTAMP)
        ON CONFLICT(numero_facturado) DO UPDATE SET id_contrato_real=excluded.id_contrato_real
    """, (numero_facturado, id_contrato_real))
    conn.execute(
        "INSERT INTO contrato_alias_historial (numero_facturado, id_contrato_real, accion, detalle) VALUES (?,?,?,?)",
        (numero_facturado, id_contrato_real, accion, detalle)
    )
    conn.commit()

def marcar_alias_usado_lote(usos: dict):
    """`usos`: numero_facturado -> cuántas veces se aplicó en el lote que se
    acaba de procesar. Se actualiza en un solo lote al terminar, no
    archivo por archivo, para no volver a introducir el mismo tipo de
    cuello de botella que se corrigió en `reglas` y `cache_lote`."""
    if not usos:
        return
    conn = get_db()
    conn.executemany(
        "UPDATE contrato_alias SET veces_aplicado = veces_aplicado + ?, fecha_ultimo_uso = CURRENT_TIMESTAMP "
        "WHERE numero_facturado = ?",
        [(cnt, num) for num, cnt in usos.items()]
    )
    conn.commit()

def obtener_alias_contrato_con_historial():
    """Para la pantalla de auditoría: reglas activas + su historial de
    cambios, para que se pueda ver no solo qué regla existe hoy sino cómo
    llegó a ser lo que es."""
    conn = get_db()
    reglas = conn.execute(
        "SELECT * FROM contrato_alias ORDER BY fecha_ultimo_uso DESC, fecha_creacion DESC"
    ).fetchall()
    historial = conn.execute(
        "SELECT * FROM contrato_alias_historial ORDER BY fecha DESC LIMIT 200"
    ).fetchall()
    return reglas, historial

def eliminar_alias_contrato(numero_facturado: str):
    conn = get_db()
    conn.execute("DELETE FROM contrato_alias WHERE numero_facturado=?", (numero_facturado,))
    conn.execute(
        "INSERT INTO contrato_alias_historial (numero_facturado, id_contrato_real, accion, detalle) "
        "VALUES (?, '', 'ELIMINADA', 'Regla borrada a mano por el usuario')",
        (numero_facturado,)
    )
    conn.commit()

def asignar_contrato_manual(uuid: str, contrato_real: str, numero_detectado: str | None = None):
    """Punto único para asignar a mano el contrato correcto de una factura
    'Sin contrato': actualiza la factura, aprende la regla de alias si el
    número que traía el XML era distinto al contrato real (para que la
    próxima con ese mismo número se resuelva sola), y vuelve a conciliar.
    Las tres pantallas donde se puede resolver un 'Sin contrato' llaman
    aquí, así la regla se aprende sin importar desde cuál se corrigió."""
    conn = get_db()
    conn.execute("UPDATE facturas SET id_contrato=? WHERE uuid=?", (contrato_real, uuid))
    conn.commit()
    if numero_detectado and numero_detectado != contrato_real:
        registrar_alias_contrato(numero_detectado, contrato_real)
    return re_conciliar_factura(uuid)

def obtener_resumen_contratos_no_vigentes(fecha_corte='2025-12'):
    """
    Obtiene el listado agrupado de contratos detectados en facturas que no están
    registrados en la tabla contratos, junto con sus clientes del SAT, fechas y montos.
    """
    conn = get_db()
    query = """
        WITH agrupados AS (
            SELECT 
                COALESCE(NULLIF(f.id_contrato_detectado, ''), NULLIF(f.id_contrato, '')) as num_contrato,
                MAX(f.rfc_receptor) as rfc_receptor,
                COUNT(*) as cant_facturas,
                MIN(f.periodo) as min_periodo,
                MAX(f.periodo) as max_periodo,
                SUM(f.total) as total_monto,
                AVG(CASE WHEN f.tipo = 'MENSUAL' THEN f.subtotal ELSE NULL END) as renta_promedio,
                MAX(CASE WHEN f.periodo >= ? THEN 1 ELSE 0 END) as tiene_recientes
            FROM facturas f
            WHERE (f.id_contrato IS NULL OR TRIM(SUBSTR(f.id_contrato, 1, INSTR(f.id_contrato || ',', ',') - 1)) NOT IN (SELECT ID_Contrato FROM contratos))
              AND f.estatus IN ('SIN_CONTRATO', 'DISCREPANCIA')
              AND (f.cancelada IS NULL OR f.cancelada = 0)
              AND COALESCE(NULLIF(f.id_contrato_detectado, ''), NULLIF(f.id_contrato, '')) IS NOT NULL
            GROUP BY num_contrato
        )
        SELECT 
            a.num_contrato,
            a.rfc_receptor,
            COALESCE(c.nombre, a.rfc_receptor, 'CLIENTE DESCONOCIDO') as cliente_nombre,
            a.cant_facturas,
            a.min_periodo,
            a.max_periodo,
            a.total_monto,
            a.renta_promedio,
            a.tiene_recientes
        FROM agrupados a
        LEFT JOIN clientes_sat c ON a.rfc_receptor = c.rfc
        ORDER BY a.cant_facturas DESC
    """
    return pd.read_sql_query(query, conn, params=[fecha_corte])

def archivar_contratos_historicos_lote(contratos_ids=None, fecha_corte='2025-12', marcar_conciliado=True):
    """
    Registra contratos concluidos en la tabla `contratos` con estatus 'BAJA' (Liquidado)
    y concilia sus facturas históricas para eliminar falsos negativos manteniendo
    la integridad de los reportes por cliente y contrato.
    """
    conn = get_db()
    clausula_contratos = ""
    params = []
    if contratos_ids:
        placeholders = ','.join(['?'] * len(contratos_ids))
        clausula_contratos = f"AND COALESCE(NULLIF(f.id_contrato_detectado, ''), NULLIF(f.id_contrato, '')) IN ({placeholders})"
        params.extend(contratos_ids)
    elif fecha_corte:
        clausula_contratos = "AND f.periodo < ?"
        params.append(fecha_corte)
        
    query = f"""
        WITH agrupados AS (
            SELECT 
                COALESCE(NULLIF(f.id_contrato_detectado, ''), NULLIF(f.id_contrato, '')) as num_contrato,
                MAX(f.rfc_receptor) as rfc_receptor,
                COUNT(*) as cant_facturas,
                MIN(f.periodo) as min_periodo,
                MAX(f.periodo) as max_periodo,
                SUM(f.total) as total_monto,
                AVG(CASE WHEN f.tipo = 'MENSUAL' THEN f.subtotal ELSE 0 END) as renta_promedio
            FROM facturas f
            WHERE (f.id_contrato IS NULL OR TRIM(SUBSTR(f.id_contrato, 1, INSTR(f.id_contrato || ',', ',') - 1)) NOT IN (SELECT ID_Contrato FROM contratos))
              AND f.estatus IN ('SIN_CONTRATO', 'DISCREPANCIA')
              AND COALESCE(NULLIF(f.id_contrato_detectado, ''), NULLIF(f.id_contrato, '')) IS NOT NULL
              {clausula_contratos}
            GROUP BY num_contrato
        )
        SELECT 
            a.*,
            COALESCE(c.nombre, a.rfc_receptor, 'CLIENTE NO IDENTIFICADO') as cliente_nombre
        FROM agrupados a
        LEFT JOIN clientes_sat c ON a.rfc_receptor = c.rfc
    """
    filas = conn.execute(query, params).fetchall()
    if not filas:
        return 0, 0
        
    contratos_creados = 0
    facturas_actualizadas = 0
    c_procesados = []
    
    for r in filas:
        num_c = r['num_contrato']
        c_procesados.append(num_c)
        cliente = r['cliente_nombre']
        min_p = (r['min_periodo'] or '2022-01') + '-01T00:00:00'
        max_p = (r['max_periodo'] or '2025-11') + '-28T00:00:00'
        renta = round(float(r['renta_promedio'] or 0.0), 2)
        plazo = int(r['cant_facturas'] or 1)
        valor_estimado = round(renta * plazo, 2)
        
        existe = conn.execute("SELECT 1 FROM contratos WHERE ID_Contrato=?", (num_c,)).fetchone()
        if not existe:
            conn.execute("""
                INSERT INTO contratos (
                    ID_Contrato, Cliente, Vehiculo, Fecha_Alta, Fecha_Vencimiento, Fecha_Baja,
                    Valor_Sin_IVA, Mensualidad_Sin_IVA, Plazo, Comision_Apertura_Pct, Comision_Monto,
                    Anticipo_Pct, Anticipo_Monto, Residual_Pct, Residual_Monto, Tasa_Calculada,
                    Estatus, Tipo_Baja, Motivo_Baja, Anotaciones
                ) VALUES (
                    ?, ?, 'Arrendamiento Concluido', ?, ?, ?,
                    ?, ?, ?, 0.0, 0.0,
                    0.0, 0.0, 0.0, 0.0, 0.0,
                    'BAJA', 'NATURAL', 'Contrato liquidado previo a Dic 2025 (Histórico)',
                    'Contrato archivado automáticamente a partir del repositorio fiscal histórico.'
                )
            """, (num_c, cliente, min_p, max_p, max_p, valor_estimado, renta, plazo))
            contratos_creados += 1

    # Actualizar facturas en lotes
    if c_procesados:
        for i in range(0, len(c_procesados), 200):
            batch = c_procesados[i:i+200]
            ph = ','.join(['?'] * len(batch))
            if marcar_conciliado:
                q_up = f"""
                    UPDATE facturas 
                    SET id_contrato = COALESCE(NULLIF(id_contrato_detectado, ''), id_contrato),
                        estatus = 'CONCILIADO',
                        observaciones = 'Histórico liquidado (Pre-Dic 2025)',
                        esperado = total,
                        facturado = total,
                        diferencia = 0.0
                    WHERE COALESCE(NULLIF(id_contrato_detectado, ''), NULLIF(id_contrato, '')) IN ({ph})
                      AND estatus = 'SIN_CONTRATO'
                """
                p_up = list(batch)
                if fecha_corte and not contratos_ids:
                    q_up += " AND periodo < ?"
                    p_up.append(fecha_corte)
                res_up = conn.execute(q_up, p_up)
                facturas_actualizadas += res_up.rowcount
            else:
                q_up = f"""
                    UPDATE facturas 
                    SET id_contrato = COALESCE(NULLIF(id_contrato_detectado, ''), id_contrato)
                    WHERE COALESCE(NULLIF(id_contrato_detectado, ''), NULLIF(id_contrato, '')) IN ({ph})
                      AND estatus = 'SIN_CONTRATO'
                """
                p_up = list(batch)
                if fecha_corte and not contratos_ids:
                    q_up += " AND periodo < ?"
                    p_up.append(fecha_corte)
                res_up = conn.execute(q_up, p_up)
                facturas_actualizadas += res_up.rowcount

    conn.commit()
    _tocar_datos()
    return contratos_creados, facturas_actualizadas

def marcar_facturas_historicas_no_aplica(contratos_ids=None, fecha_corte='2025-12'):
    """
    Marca facturas de contratos no vigentes como NO_APLICA para limpiar alertas
    sin crear registros de contratos, conservando la descripción del concepto del CFDI.
    """
    conn = get_db()
    if contratos_ids:
        placeholders = ','.join(['?'] * len(contratos_ids))
        q = f"""
            UPDATE facturas 
            SET estatus = 'NO_APLICA',
                observaciones = COALESCE(
                    (SELECT 'No Aplica Leasing | Concepto: ' || GROUP_CONCAT(fc.descripcion, ' | ') FROM factura_conceptos fc WHERE fc.uuid = facturas.uuid),
                    'No Aplica Leasing | Histórico no vigente (Pre-Dic 2025)'
                )
            WHERE estatus IN ('SIN_CONTRATO', 'DISCREPANCIA')
              AND COALESCE(NULLIF(id_contrato_detectado, ''), NULLIF(id_contrato, '')) IN ({placeholders})
        """
        params = list(contratos_ids)
    else:
        q = """
            UPDATE facturas 
            SET estatus = 'NO_APLICA',
                observaciones = COALESCE(
                    (SELECT 'No Aplica Leasing | Concepto: ' || GROUP_CONCAT(fc.descripcion, ' | ') FROM factura_conceptos fc WHERE fc.uuid = facturas.uuid),
                    'No Aplica Leasing | Histórico no vigente (Pre-Dic 2025)'
                )
            WHERE estatus IN ('SIN_CONTRATO', 'DISCREPANCIA')
              AND periodo < ?
        """
        params = [fecha_corte]
        
    res = conn.execute(q, params)
    n = res.rowcount
    conn.commit()
    _tocar_datos()
    return n

def borrar_facturas_lote(contratos_ids=None, fecha_corte='2025-12'):
    """
    Borra físicamente facturas y conceptos de contratos no vigentes.
    """
    conn = get_db()
    if contratos_ids:
        placeholders = ','.join(['?'] * len(contratos_ids))
        filas = conn.execute(f"""
            SELECT uuid FROM facturas 
            WHERE COALESCE(NULLIF(id_contrato_detectado, ''), NULLIF(id_contrato, '')) IN ({placeholders})
        """, list(contratos_ids)).fetchall()
    else:
        filas = conn.execute("""
            SELECT uuid FROM facturas 
            WHERE estatus IN ('SIN_CONTRATO', 'DISCREPANCIA')
              AND periodo < ?
              AND (id_contrato IS NULL OR TRIM(SUBSTR(id_contrato, 1, INSTR(id_contrato || ',', ',') - 1)) NOT IN (SELECT ID_Contrato FROM contratos))
        """, [fecha_corte]).fetchall()
        
    if not filas:
        return 0
    uuids = [f['uuid'] for f in filas]
    
    for i in range(0, len(uuids), 500):
        lote = uuids[i:i+500]
        ph = ','.join(['?'] * len(lote))
        conn.execute(f"DELETE FROM factura_conceptos WHERE uuid IN ({ph})", lote)
        conn.execute(f"DELETE FROM facturas WHERE uuid IN ({ph})", lote)
    conn.commit()
    _tocar_datos()
    return len(uuids)

def reasignar_facturas_contrato_lote(id_origen: str, id_destino: str, guardar_alias: bool = True):
    """
    Reasigna en lote todas las facturas de un contrato origen a un contrato destino,
    guardando opcionalmente la regla de alias y re-conciliando las facturas.
    """
    conn = get_db()
    filas = conn.execute("""
        SELECT uuid FROM facturas 
        WHERE COALESCE(NULLIF(id_contrato_detectado, ''), NULLIF(id_contrato, '')) = ?
    """, (id_origen,)).fetchall()
    
    if not filas:
        return 0
        
    for f in filas:
        u = f['uuid']
        conn.execute("UPDATE facturas SET id_contrato=? WHERE uuid=?", (id_destino, u))
        re_conciliar_factura(u)
        
    if guardar_alias and id_origen != id_destino:
        registrar_alias_contrato(id_origen, id_destino)
        
    conn.commit()
    _tocar_datos()
    return len(filas)


# ---------------------------------------------------------------------
# Comparación contra analíticas contables (Excel del contador)
# ---------------------------------------------------------------------
# El contador exporta de su sistema contable un Excel con 3 hojas típicas:
# "Ingresos por Intereses", "Ingresos por Valor Residual" y "Arrendamiento
# de Comisión Apertura" — un renglón por contrato, con lo que ya se
# registró en libros mes a mes (ENE-DIC). Aquí se lee ese Excel, se calcula
# lo que el sistema esperaría para cada contrato/mes, y se comparan.

_PAT_ID_ANALITICA_GUION = re.compile(r'(\d{1,4})\s*-\s*(\d{1,4})')
_PAT_ID_ANALITICA_JUNTO = re.compile(r'\b(\d{7,8})\b')

def _extraer_id_contrato_libre(texto: str):
    """El contador escribe el número de contrato de muchas formas distintas
    en su Excel ('CONTRATO 0521-0001 NP300', 'contrato 506-08 (revisar)',
    'Contrato N. 03260002', '0504-0004 Juan Pérez'...). Esto intenta
    reconocer cualquiera de esas formas y devolver el ID normalizado
    (NNNN-NNNN) como lo usa el sistema."""
    if not isinstance(texto, str):
        return None
    m = _PAT_ID_ANALITICA_GUION.search(texto)
    if m:
        return f"{m.group(1).zfill(4)}-{m.group(2).zfill(4)}"
    m2 = _PAT_ID_ANALITICA_JUNTO.search(texto)
    if m2:
        digitos = m2.group(1)
        if len(digitos) == 7:
            digitos = '0' + digitos
        return f"{digitos[:4]}-{digitos[4:]}"
    return None

_MESES_COL_ANALITICA = {'ENE':1,'FEB':2,'MAR':3,'ABR':4,'MAY':5,'JUN':6,
                         'JUL':7,'AGO':8,'SEP':9,'OCT':10,'NOV':11,'DIC':12}
_MAPA_HOJA_TIPO_ANALITICA = [
    (re.compile(r'inter[eé]s', re.I), 'INTERES'),
    (re.compile(r'residual', re.I), 'RESIDUAL'),
    (re.compile(r'comisi[oó]n', re.I), 'COMISION'),
]

def parse_analiticas_excel(file_bytes: bytes):
    """Lee el Excel de analíticas del contador (cualquier subconjunto de las
    3 hojas típicas) y regresa (df_tidy, año_detectado, hojas_no_reconocidas).
    df_tidy: una fila por contrato + mes + tipo con el valor ya registrado
    en la contabilidad."""
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    filas = []
    anio_detectado = None
    hojas_no_reconocidas = []
    for nombre_hoja in xls.sheet_names:
        tipo = None
        for patron, t in _MAPA_HOJA_TIPO_ANALITICA:
            if patron.search(nombre_hoja):
                tipo = t
                break
        if tipo is None:
            hojas_no_reconocidas.append(nombre_hoja)
            continue
        raw = xls.parse(nombre_hoja, header=None)
        if anio_detectado is None:
            for _, r in raw.head(6).iterrows():
                for val in r:
                    if isinstance(val, str):
                        m = re.search(r'\b(20\d{2})\b', val)
                        if m:
                            anio_detectado = int(m.group(1))
                            break
                if anio_detectado:
                    break
        header_row_idx = None
        col_meses = {}
        for i, r in raw.iterrows():
            vals = [str(x).strip().upper() if isinstance(x, str) else '' for x in r]
            hits = {mes: j for j, val in enumerate(vals) for mes in _MESES_COL_ANALITICA if val == mes}
            if len(hits) >= 6:
                header_row_idx = i
                col_meses = hits
                break
        if header_row_idx is None:
            continue
        col_desc = 1
        for i in range(header_row_idx + 1, len(raw)):
            desc = raw.iat[i, col_desc] if col_desc < raw.shape[1] else None
            if not isinstance(desc, str) or not desc.strip():
                continue
            if desc.strip().upper().startswith('TOTAL'):
                continue
            id_contrato = _extraer_id_contrato_libre(desc)
            for mes_nombre, mes_num in _MESES_COL_ANALITICA.items():
                j = col_meses.get(mes_nombre)
                if j is None or j >= raw.shape[1]:
                    continue
                val = raw.iat[i, j]
                if pd.isna(val):
                    continue
                try:
                    val = float(val)
                except Exception:
                    continue
                if val == 0:
                    continue
                filas.append({
                    'Hoja': nombre_hoja, 'Tipo': tipo,
                    'ID_Contrato_Detectado': id_contrato,
                    'Texto_Original': desc.strip(),
                    'Mes': mes_num, 'Valor': val,
                })
    return pd.DataFrame(filas), anio_detectado, hojas_no_reconocidas

def _valor_esperado_contrato_mes(con, tipo: str, mes: int, anio: int):
    """Cuánto debería haber reconocido el sistema para este contrato, este
    mes, para este tipo de ingreso ('INTERES','RESIDUAL','COMISION').
    Devuelve (valor_esperado, motivo_si_es_cero_por_diseño). Usa el mismo
    criterio (mes de contrato, plazo, fecha de baja) que ya usan las
    pólizas contables, para que la comparación sea consistente con el
    resto del sistema."""
    fa = pd.to_datetime(con['Fecha_Alta'])
    fv = pd.to_datetime(con['Fecha_Vencimiento']) if pd.notna(con.get('Fecha_Vencimiento')) \
        else fa + relativedelta(months=int(con['Plazo']))
    fp   = pd.Timestamp(anio, mes, 1)
    fa_m = pd.Timestamp(fa.year, fa.month, 1)
    pl   = int(con['Plazo'])
    mc   = (anio - fa.year) * 12 + (mes - fa.month) + 1

    if str(con.get('Estatus', '')).upper() == 'BAJA' and pd.notna(con.get('Fecha_Baja')) \
       and str(con.get('Fecha_Baja')).strip() not in ('', 'None', 'NaT'):
        fb = pd.to_datetime(con['Fecha_Baja'])
        if fb < fp:
            return 0.0, f"Contrato dado de baja el {fb.strftime('%d/%m/%Y')} — ya no debería generar ingresos en {MN[mes-1]} {anio}."

    if fa_m > fp:
        return 0.0, f"El contrato inicia hasta {fa.strftime('%m/%Y')}; todavía no debería generar ingresos en {MN[mes-1]} {anio}."
    if mc < 1 or mc > pl:
        return 0.0, f"Fuera del plazo del contrato (equivaldría al mes {mc} de {pl})."

    inv = con['Valor_Sin_IVA'] - con['Anticipo_Monto']
    r, t, res = con['Mensualidad_Sin_IVA'], con['Tasa_Calculada'], con['Residual_Monto']
    if tipo == 'INTERES':
        dfa, _, _, _, _ = calc_amort(round(inv, 4), round(r, 4), round(res, 4), pl, round(t, 8))
        return round(float(dfa.iloc[mc-1]['Interes']), 2), None
    if tipo == 'RESIDUAL':
        vpr = vp_res(res, t, pl)
        dfr = calc_res_amort(round(vpr, 4), round(t, 8), pl)
        return round(float(dfr.iloc[mc-1]['Interes']), 2), None
    if tipo == 'COMISION':
        com = con.get('Comision_Monto', 0) or 0
        return (round(com / pl, 2) if pl else 0.0), None
    return 0.0, None

def comparar_analiticas(df_tidy: pd.DataFrame, anio: int, tolerancia: float = 1.0):
    """Compara cada renglón de la analítica contra lo que el sistema
    esperaba, y arma la tabla final con toda la información necesaria para
    decidir qué corregir: valor registrado, esperado, diferencia y una
    observación en texto plano explicando el porqué."""
    df_con = obtener()  # TODOS los contratos, activos y dados de baja
    con_por_id = {r['ID_Contrato']: r for _, r in df_con.iterrows()}
    resultados = []
    for _, fila in df_tidy.iterrows():
        idc = fila['ID_Contrato_Detectado']
        con = con_por_id.get(idc) if idc else None
        if con is None:
            resultados.append({
                'ID_Contrato': idc or '(no identificado)', 'Cliente': '', 'Tipo': fila['Tipo'],
                'Mes': MN[fila['Mes']-1], 'Registrado $': fila['Valor'], 'Esperado $': None,
                'Diferencia $': None, 'Estatus_Contrato': '',
                'Observación': f"No se encontró un contrato con este ID en el sistema. Texto original: \u201c{fila['Texto_Original']}\u201d — revisa si el número está bien capturado en la analítica.",
            })
            continue
        esperado, motivo = _valor_esperado_contrato_mes(con, fila['Tipo'], fila['Mes'], anio)
        dif = round(fila['Valor'] - esperado, 2)
        obs_partes = []
        if motivo:
            obs_partes.append(motivo)
        if abs(con['Mensualidad_Sin_IVA']) > 0 and fila['Tipo'] == 'INTERES' and abs(dif) > tolerancia and not motivo:
            obs_partes.append(
                f"La renta actual en el sistema es ${con['Mensualidad_Sin_IVA']:,.2f} — si el contador usó un "
                f"valor distinto al capturar el contrato, esta diferencia puede venir de ahí."
            )
        if abs(dif) > tolerancia and not obs_partes:
            obs_partes.append("Diferencia sin una causa evidente — revisar a mano.")
        resultados.append({
            'ID_Contrato': idc, 'Cliente': con.get('Cliente', ''), 'Tipo': fila['Tipo'],
            'Mes': MN[fila['Mes']-1], 'Registrado $': fila['Valor'], 'Esperado $': esperado,
            'Diferencia $': dif, 'Estatus_Contrato': con.get('Estatus', ''),
            'Observación': ' '.join(obs_partes) if obs_partes else '',
        })
    df_out = pd.DataFrame(resultados)
    if not df_out.empty:
        df_out['_relevante'] = df_out['Diferencia $'].isna() | (df_out['Diferencia $'].abs() > tolerancia)
    return df_out

def sugerir_ajuste_poliza(df_comparado: pd.DataFrame, cat: dict):
    """A partir de las diferencias encontradas, sugiere una póliza de ajuste
    (una línea de Cargo/Abono por contrato+tipo con diferencia neta) para
    dejar la contabilidad alineada con lo que dice el sistema. Es una
    SUGERENCIA — siempre debe revisarla antes de contabilizarla."""
    _CUENTA_INGRESO = {'INTERES': 'ING_INTERESES', 'RESIDUAL': 'ING_RESIDUAL', 'COMISION': 'COMISION_GASTO'}
    filas = []
    df_v = df_comparado[df_comparado['Diferencia $'].notna() & (df_comparado['Diferencia $'].abs() > 0.01)]
    for (idc, tipo), grp in df_v.groupby(['ID_Contrato', 'Tipo']):
        neto = round(grp['Diferencia $'].sum(), 2)
        if abs(neto) <= 0.01:
            continue
        clave_cta = _CUENTA_INGRESO.get(tipo, 'ING_INTERESES')
        info_cta = cat.get(clave_cta)
        if not info_cta or not info_cta.get('activa', True) or not (info_cta.get('cuenta') or '').strip():
            continue  # cuenta no configurada / desactivada en Cuentas → Macro: no se sugiere ajuste
        cta, nom = info_cta['cuenta'], info_cta['nombre']
        cliente = grp['Cliente'].iloc[0]
        desc = f"Ajuste {tipo.title()} {idc} — registrado de más" if neto > 0 else f"Ajuste {tipo.title()} {idc} — registrado de menos"
        # Si se registró de MÁS de lo que el sistema esperaba (neto positivo),
        # se necesita un cargo a la cuenta de ingreso (para bajarlo);
        # si se registró de MENOS (neto negativo), un abono (para subirlo).
        cargo = abs(neto) if neto > 0 else 0
        abono = abs(neto) if neto < 0 else 0
        filas.append({
            'Cuenta': cta, 'Nombre_Cuenta': nom, 'ID_Contrato': idc, 'Cliente': cliente,
            'Tipo': tipo, 'Cargo': cargo, 'Abono': abono, 'Descripción': desc,
        })
    return pd.DataFrame(filas)

def _res(status, msg, esperado=0.0, facturado=0.0, dif=0.0, detalle=None, **extra):
    r = {
        'status':   status,
        'msg':      msg,
        'esperado': esperado,
        'facturado': facturado,
        'dif':      dif,
        'detalle':  detalle or [],
    }
    r.update(extra)
    return r

def _extraer_concepto_descriptivo(fact: dict) -> str:
    """Extrae las descripciones reales de los conceptos de la factura para
    mostrarlas en la columna de concepto y en observaciones cuando es NO_APLICA."""
    descs = []
    c_raw = fact.get('conceptos_raw') or []
    if isinstance(c_raw, list):
        for c in c_raw:
            if isinstance(c, dict) and c.get('desc'):
                d = re.sub(r'\s+', ' ', str(c['desc'])).strip()
                if d and d not in descs:
                    descs.append(d)
    if not descs and fact.get('uuid'):
        try:
            conn = get_db()
            rows_c = conn.execute(
                "SELECT descripcion FROM factura_conceptos WHERE uuid=? ORDER BY id",
                (fact['uuid'],)
            ).fetchall()
            for r in rows_c:
                d = re.sub(r'\s+', ' ', str(r[0] or '')).strip()
                if d and d not in descs:
                    descs.append(d)
        except Exception:
            pass
    if descs:
        res = " | ".join(descs)
        return res[:217] + "..." if len(res) > 220 else res
    return "Servicios no relacionados con leasing"

def _conciliar_factura_interna(fact: dict, _cache: dict | None = None) -> dict:
    """`_cache` es opcional: cuando se procesa un lote completo (ver
    conciliar_factura), guarda ahí el contrato y su tabla de amortización
    ya calculada por id_contrato. Un mismo contrato suele aparecer en
    varias facturas del mismo lote (una por mes, o renta + comisión del
    mismo anticipo) — sin este caché se repetía la consulta a `contratos`
    y el cálculo completo de amortización una vez por cada factura."""
    if _cache is None:
        _cache = {}
    # Blindaje: ¿la factura es realmente de tu arrendadora? Si configuraste
    # el RFC emisor en la pestaña de Carga, cualquier XML con otro RFC se
    # marca aparte ANTES de intentar conciliarla — así no se cuela por
    # coincidencia de número contra un contrato que no le corresponde.
    rfc_config = _cache.get('rfc_config')
    if rfc_config is None:
        rfc_config = (get_cfg('rfc_arrendadora', '') or '').strip().upper()
        _cache['rfc_config'] = rfc_config
    rfc_fact   = (fact.get('rfc_emisor') or '').strip().upper()
    if rfc_config and rfc_fact and rfc_fact != rfc_config:
        return _res('RFC_INCORRECTO',
                    f"El RFC emisor de este XML ({rfc_fact}) no coincide con el RFC configurado de tu "
                    f"arrendadora ({rfc_config}). Esta factura no se comparó contra ningún contrato.",
                    esperado=0.0, facturado=fact.get('total', 0.0), dif=0.0, detalle=[])

    # Blindaje: ¿es realmente una factura de ingreso (I)? Una nota de
    # crédito (E), un CFDI de nómina (N) o de traslado (T) no se debe
    # comparar como si fuera renta cobrada — antes se trataba igual que
    # cualquier factura y podía compensar o inflar el total sin que se
    # notara.
    tipo_comp = (fact.get('tipo_comprobante') or 'I').upper()
    if tipo_comp != 'I':
        _nombres_comp = {'E': 'nota de crédito / egreso', 'N': 'nómina', 'T': 'traslado', 'P': 'complemento de pago'}
        c_desc = _extraer_concepto_descriptivo(fact)
        return _res('NO_APLICA',
                    f"No Aplica Leasing ({_nombres_comp.get(tipo_comp, tipo_comp)}) | Concepto: {c_desc}",
                    esperado=0.0, facturado=fact.get('total', 0.0), dif=0.0, detalle=[c_desc])

    # Blindaje: si el CFDI no está en pesos, comparar los montos tal cual
    # contra la renta (que sí está en MXN) daría una discrepancia falsa o,
    # peor, un "cuadre" casual sin sentido. Se manda a revisión manual.
    moneda_cfdi = (fact.get('moneda') or 'MXN').upper()
    if moneda_cfdi not in ('MXN', ''):
        return _res('ERROR',
                    f"Este CFDI está emitido en {moneda_cfdi}, no en pesos (MXN) — no se puede comparar "
                    f"directamente contra la renta pactada sin convertir primero. Revísalo a mano.",
                    esperado=0.0, facturado=fact.get('total', 0.0), dif=0.0, detalle=[])

    # Si la factura es de tipo OTRO o INDEMNIZACION, no requiere contrato de leasing
    if fact.get('tipo') in ('OTRO', 'INDEMNIZACION'):
        c_desc = _extraer_concepto_descriptivo(fact)
        return _res('NO_APLICA', f"No Aplica Leasing | Concepto: {c_desc}",
                    esperado=0.0, facturado=fact.get('total', 0.0), dif=0.0, detalle=[c_desc])

    conn = get_db()
    # Detección exhaustiva de contratos (soporta contratos individuales y multi-contrato)
    contratos_detectados = list(fact.get('contratos_detectados') or [])
    for c in (fact.get('conceptos_raw') or []):
        if isinstance(c, dict) and c.get('desc'):
            for m in _PAT_CONTRATO.findall(str(c['desc'])):
                norm_c = normalizar_contrato(m[0], m[1])
                if norm_c not in contratos_detectados:
                    contratos_detectados.append(norm_c)
    if fact.get('uuid'):
        try:
            rows_desc = conn.execute("SELECT descripcion FROM factura_conceptos WHERE uuid=?", (fact['uuid'],)).fetchall()
            for rd in rows_desc:
                for m in _PAT_CONTRATO.findall(str(rd[0] or '')):
                    norm_c = normalizar_contrato(m[0], m[1])
                    if norm_c not in contratos_detectados:
                        contratos_detectados.append(norm_c)
        except Exception:
            pass
    for raw_txt in [fact.get('id_contrato_detectado'), fact.get('id_contrato')]:
        if raw_txt:
            for m in _PAT_CONTRATO.findall(str(raw_txt)):
                norm_c = normalizar_contrato(m[0], m[1])
                if norm_c not in contratos_detectados:
                    contratos_detectados.append(norm_c)

    if not contratos_detectados:
        return _res('SIN_CONTRATO', 'No se pudo extraer el número de contrato del XML')

    contratos_validos = []
    contratos_info = []
    contratos_no_encontrados = []
    for cid in contratos_detectados:
        con_obj = _cache.get('contratos', {}).get(cid)
        if con_obj is None:
            row = conn.execute("SELECT * FROM contratos WHERE ID_Contrato=?", (cid,)).fetchone()
            if row:
                con_obj = dict(row)
                _cache.setdefault('contratos', {})[cid] = con_obj
        if con_obj:
            contratos_validos.append(cid)
            contratos_info.append(con_obj)
        else:
            contratos_no_encontrados.append(cid)

    if not contratos_validos:
        cids_faltantes = ', '.join(contratos_detectados)
        return _res('SIN_CONTRATO', f'Contrato(s) {cids_faltantes} no encontrado(s) en la base de datos')

    es_multi = len(contratos_validos) > 1
    contratos_str = ', '.join(contratos_validos)
    n_contratos = len(contratos_validos)
    id_c = contratos_validos[0]
    con = contratos_info[0]

    # Fecha de referencia de esta factura (para todas las validaciones de
    # vigencia de aquí en adelante): el período detectado del XML si se
    # pudo armar, si no la fecha de emisión del CFDI.
    fecha_factura = None
    if fact.get('periodo'):
        try:
            _y, _m = fact['periodo'].split('-')
            fecha_factura = pd.Timestamp(int(_y), int(_m), 1)
        except Exception:
            fecha_factura = None
    if fecha_factura is None and fact.get('fecha'):
        try:
            fecha_factura = pd.to_datetime(fact['fecha']).replace(day=1)
        except Exception:
            fecha_factura = None

    conceptos = fact.get('conceptos') or {}

    # =========================================================================
    # LÓGICA ESPECIAL PARA FACTURAS MULTI-CONTRATO (CONSOLIDACIÓN DE RENTAS/APERTURAS)
    # =========================================================================
    if es_multi:
        if fact.get('tipo') == 'MENSUAL':
            renta_pura_fac = round(float(conceptos.get('RENTA') or 0.0), 2)
            admin_fac      = round(float(conceptos.get('ADMIN') or 0.0), 2)
            geoloc_fac     = round(float(conceptos.get('GEOLOC') or 0.0), 2)
            renta_fac      = round(renta_pura_fac + admin_fac + geoloc_fac, 2)
            
            renta_total_esp = round(sum(float(c.get('Mensualidad_Sin_IVA') or 0.0) for c in contratos_info), 2)
            dif = round(renta_fac - renta_total_esp, 2)
            
            errores = []
            obs_lista = []
            obs_multi = f"Factura multi-contrato ({n_contratos} contratos: {contratos_str})"
            
            if abs(dif) <= TOLERANCIA:
                status = 'CONCILIADO'
                obs_lista.append(f"{obs_multi} conciliada. Renta combinada esperada ${renta_total_esp:,.2f} coincide con lo facturado ${renta_fac:,.2f}.")
            else:
                status = 'DISCREPANCIA'
                errores.append(f"{obs_multi}: esperado combinado ${renta_total_esp:,.2f}, facturado ${renta_fac:,.2f} (dif ${dif:+,.2f})")
                
            if contratos_no_encontrados:
                obs_lista.append(f"Aviso: contrato(s) {', '.join(contratos_no_encontrados)} en el texto no existen en el sistema")
                
            todas_obs = errores + obs_lista
            return _res(status, '; '.join(todas_obs) if todas_obs else 'OK',
                        esperado=renta_total_esp, facturado=renta_fac, dif=dif, detalle=todas_obs,
                        es_multi_contrato=True,
                        contratos_considerados=contratos_validos,
                        contratos_str=contratos_str,
                        renta_esp=renta_total_esp, renta_fac=renta_fac, dif_renta=dif,
                        renta_pura_fac=renta_pura_fac, admin_fac=admin_fac, geoloc_fac=geoloc_fac)

        elif fact.get('tipo') == 'ANTICIPO':
            tot_ant_esp = round(sum(float(c.get('Anticipo_Monto') or 0.0) for c in contratos_info), 2)
            tot_com_esp = round(sum(float(c.get('Comision_Monto') or 0.0) for c in contratos_info), 2)
            tot_apertura_esp = round(tot_ant_esp + tot_com_esp, 2)

            ant_fac = round(float(conceptos.get('ANTICIPO') or 0.0), 2)
            com_fac = round(float(conceptos.get('COMISION') or 0.0), 2)
            tot_apertura_fac = round(ant_fac + com_fac, 2)
            dif_apertura = round(tot_apertura_fac - tot_apertura_esp, 2)

            renta_mes1_fac = round(float(conceptos.get('RENTA') or 0.0) + float(conceptos.get('ADMIN') or 0.0) + float(conceptos.get('GEOLOC') or 0.0), 2)
            renta_mes1_esp = round(sum(float(c.get('Mensualidad_Sin_IVA') or 0.0) for c in contratos_info), 2) if renta_mes1_fac > 0 else 0.0
            dif_r1 = round(renta_mes1_fac - renta_mes1_esp, 2) if renta_mes1_fac > 0 else 0.0

            total_esp = round(tot_apertura_esp + renta_mes1_esp, 2)
            total_fac = round(tot_apertura_fac + renta_mes1_fac, 2)
            dif_total = round(total_fac - total_esp, 2)

            errores = []
            obs_lista = []
            obs_multi = f"Factura multi-contrato de apertura ({n_contratos} contratos: {contratos_str})"

            apertura_ok = (abs(dif_apertura) <= TOLERANCIA)
            renta_ok = (renta_mes1_fac == 0 or abs(dif_r1) <= TOLERANCIA)

            if apertura_ok and renta_ok:
                status = 'CONCILIADO'
                obs_lista.append(f"{obs_multi} conciliada. Apertura combinada esperada ${tot_apertura_esp:,.2f} coincide con lo facturado ${tot_apertura_fac:,.2f}.")
            else:
                status = 'DISCREPANCIA'
                if not apertura_ok:
                    errores.append(f"{obs_multi}: apertura esperada ${tot_apertura_esp:,.2f}, facturada ${tot_apertura_fac:,.2f} (dif ${dif_apertura:+,.2f})")
                if not renta_ok:
                    errores.append(f"{obs_multi}: renta mes 1 esperada ${renta_mes1_esp:,.2f}, facturada ${renta_mes1_fac:,.2f} (dif ${dif_r1:+,.2f})")

            todas_obs = errores + obs_lista
            return _res(status, '; '.join(todas_obs) if todas_obs else 'OK',
                        esperado=total_esp, facturado=total_fac, dif=dif_total, detalle=todas_obs,
                        es_multi_contrato=True,
                        contratos_considerados=contratos_validos,
                        contratos_str=contratos_str,
                        ant_esp=tot_ant_esp, ant_fac=ant_fac,
                        com_esp=tot_com_esp, com_fac=com_fac,
                        tot_apertura_esp=tot_apertura_esp, tot_apertura_fac=tot_apertura_fac, dif_apertura=dif_apertura,
                        renta_esp=renta_mes1_esp, renta_fac=renta_mes1_fac, dif_renta=dif_r1)

    # Blindaje: ¿el contrato ya estaba dado de baja cuando se emitió esta
    # factura, o la factura es de ANTES de que el contrato siquiera
    # empezara? Antes esto no se revisaba — si por casualidad el monto
    # facturado coincidía con lo esperado, una factura de mayo se marcaba
    # CONCILIADO aunque el contrato se hubiera dado de baja en febrero. El
    # monto puede cuadrar por coincidencia; la vigencia del contrato no.
    if fecha_factura is not None:
        if str(con.get('Estatus', '')).upper() == 'BAJA' and con.get('Fecha_Baja') and \
           str(con.get('Fecha_Baja')).strip() not in ('', 'None', 'NaT', '0'):
            try:
                fb = pd.to_datetime(con['Fecha_Baja'])
                if pd.Timestamp(fb.year, fb.month, 1) < fecha_factura:
                    return _res('FUERA_DE_VIGENCIA',
                        f"Este contrato se dio de baja el {fb.strftime('%d/%m/%Y')} — esta factura es de "
                        f"{fecha_factura.strftime('%m/%Y')}, un período posterior a la baja. No debería haberse "
                        f"generado renta para un contrato ya dado de baja.",
                        esperado=0.0, facturado=fact.get('total', 0.0), dif=fact.get('total', 0.0), detalle=[])
            except Exception:
                pass  # si la fecha de baja está mal capturada, no se bloquea la conciliación por esto
        if fact.get('tipo') == 'MENSUAL':
            try:
                fa_chk = pd.to_datetime(con['Fecha_Alta'])
                if pd.Timestamp(fa_chk.year, fa_chk.month, 1) > fecha_factura:
                    return _res('FUERA_DE_VIGENCIA',
                        f"Este contrato no inicia hasta {fa_chk.strftime('%m/%Y')} — esta factura es de "
                        f"{fecha_factura.strftime('%m/%Y')}, anterior al alta. No debería existir renta para "
                        f"un período previo al inicio del contrato.",
                        esperado=0.0, facturado=fact.get('total', 0.0), dif=fact.get('total', 0.0), detalle=[])
            except Exception:
                pass

    inv  = round(con['Valor_Sin_IVA'] - con['Anticipo_Monto'], 4)
    pl   = int(con['Plazo'])
    tasa = round(con['Tasa_Calculada'], 8)
    renta = round(con['Mensualidad_Sin_IVA'], 4)
    res  = round(con['Residual_Monto'], 4)

    dfa = _cache.get('amortizaciones', {}).get(id_c)
    if dfa is None:
        try:
            dfa, _, _, _, _ = calc_amort(inv, renta, res, pl, tasa)
        except Exception as e:
            return _res('ERROR', f'Error al calcular amortización: {e}')
        _cache.setdefault('amortizaciones', {})[id_c] = dfa

    # Deteccion exhaustiva de meses del contrato (soporta facturas con pago acumulado de multiples meses)
    meses_detectados = list(fact.get('meses_detectados') or [])
    for c in (fact.get('conceptos_raw') or []):
        if isinstance(c, dict):
            desc_c = str(c.get('desc') or '')
            clv_c = str(c.get('clave') or '')
            if clv_c in ('OTRO', 'SEGURO', 'GESTORIA') and 'RENTA' not in desc_c.upper():
                continue
            for m_str, pl_str in _PAT_MES.findall(desc_c):
                try:
                    m_val = int(m_str)
                    if 1 <= m_val <= pl and m_val not in meses_detectados:
                        meses_detectados.append(m_val)
                except Exception:
                    pass
    if fact.get('uuid'):
        try:
            rows_c = conn.execute("SELECT descripcion, clave FROM factura_conceptos WHERE uuid=?", (fact['uuid'],)).fetchall()
            for rc in rows_c:
                desc_c = str(rc[0] or '')
                clv_c = str(rc[1] or '')
                if clv_c in ('OTRO', 'SEGURO', 'GESTORIA') and 'RENTA' not in desc_c.upper():
                    continue
                for m_str, pl_str in _PAT_MES.findall(desc_c):
                    try:
                        m_val = int(m_str)
                        if 1 <= m_val <= pl and m_val not in meses_detectados:
                            meses_detectados.append(m_val)
                    except Exception:
                        pass
        except Exception:
            pass
    if fact.get('mes_contrato'):
        for part in str(fact['mes_contrato']).split(','):
            part = part.strip().replace('.0', '')
            if part.isdigit():
                m_val = int(part)
                if 1 <= m_val <= pl and m_val not in meses_detectados:
                    meses_detectados.append(m_val)

    meses_detectados = sorted(meses_detectados)
    es_multi_mes = len(meses_detectados) > 1
    cant_meses = len(meses_detectados) if es_multi_mes else 1
    mes_str = ', '.join(str(m) for m in meses_detectados) if es_multi_mes else (str(meses_detectados[0]) if meses_detectados else (str(fact.get('mes_contrato') or 1)))

    mes = meses_detectados[0] if meses_detectados else None
    if mes is None:
        try:
            mes = int(str(fact.get('mes_contrato', '')).split(',')[0].strip().replace('.0', ''))
        except Exception:
            mes = None

    if not mes or mes < 1 or mes > pl:
        if fact.get('tipo') == 'ANTICIPO':
            mes = 1
        else:
            return _res('ERROR', f'Mes del contrato ({fact.get("mes_contrato")}) fuera de rango 1-{pl}')

    # Correspondencia operativa entre mes de concepto y fecha calendario de la factura
    aviso_mes_info = None
    if fecha_factura is not None and fact.get('tipo') == 'MENSUAL':
        try:
            fa_ref = pd.to_datetime(con['Fecha_Alta'])
            mes_por_fecha = (fecha_factura.year - fa_ref.year) * 12 + (fecha_factura.month - fa_ref.month) + 1
            dif_meses = mes_por_fecha - mes
            if dif_meses == 1:
                # Desfase operativo estandar: Firma + Anticipo + Renta 1 en Mes 1, Mes 2 sin factura, Mes 3 factura Renta 2
                aviso_mes_info = (
                    f"Desfase operativo de 1 mes (renta inicial anticipada al inicio del contrato; "
                    f"concepto mes {mes} de {pl} vs calendario mes {mes_por_fecha})."
                )
            elif dif_meses < 0:
                # Pago adelantado: el cliente facturo mensualidades por adelantado
                adelanto = abs(dif_meses)
                aviso_mes_info = (
                    f"Pago adelantado de {adelanto} mensualidad(es) respecto al calendario normal del contrato "
                    f"(concepto indica mes {mes} de {pl} en periodo de mes {mes_por_fecha})."
                )
            elif dif_meses > 1:
                # Desfase mayor a 1 mes
                aviso_mes_info = (
                    f"Desfase de {dif_meses} meses respecto al calendario "
                    f"(concepto indica mes {mes} de {pl} en periodo de mes {mes_por_fecha})."
                )
        except Exception:
            pass

    conceptos = fact['conceptos']
    fila = dfa.iloc[mes - 1]
    renta_pura_esp = round(fila['Interes'] + fila['Capital'], 2)
    renta_total_esp = round(con['Mensualidad_Sin_IVA'], 2)

    if fact['tipo'] == 'VENTA_VEHICULO':
        res_esp = round(con.get('Residual_Monto', 0.0), 2)
        res_fac = round(fact.get('subtotal', fact.get('total', 0.0)), 2)
        dif_res = round(res_fac - res_esp, 2)
        errores = []
        if abs(dif_res) > TOLERANCIA:
            errores.append(f"Residual/Venta: esperado ${res_esp:,.2f}, facturado ${res_fac:,.2f} (dif ${dif_res:+,.2f})")
        
        estatus_con = str(con.get('Estatus', '')).upper()
        if estatus_con != 'BAJA':
            errores.append("ATENCION: Factura por Venta de Vehiculo (cobro de residual), pero el contrato figura como ACTIVO en el sistema. Se sugiere procesar la Baja en Gestor de Bajas.")
        
        status = 'CONCILIADO' if not errores else 'DISCREPANCIA'
        return _res(status, '; '.join(errores) if errores else 'OK (Venta de Vehiculo / Residual)',
                    esperado=res_esp, facturado=res_fac, dif=dif_res, detalle=errores)

    elif fact['tipo'] == 'INDEMNIZACION':
        total_fac = round(fact.get('total', 0.0), 2)
        errores = []
        
        estatus_con = str(con.get('Estatus', '')).upper()
        has_evento = False
        try:
            ev_count = conn.execute("SELECT COUNT(*) FROM eventos_especiales WHERE id_contrato=?", (id_c,)).fetchone()[0]
            if ev_count > 0:
                has_evento = True
        except Exception:
            pass

        if estatus_con != 'BAJA' and not has_evento:
            errores.append("ATENCION: Factura por Indemnizacion de Seguro recibida, pero el contrato figura ACTIVO y no existe evento especial de Siniestro/Baja registrado.")

        status = 'CONCILIADO' if not errores else 'DISCREPANCIA'
        return _res(status, '; '.join(errores) if errores else 'OK (Indemnizacion por Siniestro)',
                    esperado=total_fac, facturado=total_fac, dif=0.0, detalle=errores)

    elif fact['tipo'] == 'MENSUAL':
        renta_pura_fac = round(float(conceptos.get('RENTA') or 0.0), 2)
        admin_fac      = round(float(conceptos.get('ADMIN') or 0.0), 2)
        geoloc_fac     = round(float(conceptos.get('GEOLOC') or 0.0), 2)
        renta_fac      = round(renta_pura_fac + admin_fac + geoloc_fac, 2)
        
        renta_mensual_esp = round(con['Mensualidad_Sin_IVA'], 2)
        renta_total_esp   = round(renta_mensual_esp * cant_meses, 2)
        
        subtotal_fac = round(float(fact.get('subtotal') or 0.0), 2)
        if es_multi_mes:
            if abs(renta_fac - renta_total_esp) > TOLERANCIA and abs(subtotal_fac - renta_total_esp) <= TOLERANCIA:
                renta_fac = subtotal_fac
        
        dif            = round(renta_fac - renta_total_esp, 2)
        errores        = []
        observaciones_lista = []
        
        # Validar equivalencia de IVA (16%): renta capturada con IVA o doble deduccion de IVA
        if abs(dif) > TOLERANCIA:
            # Caso A: Renta en contrato se capturo con IVA incluido (dif negativa)
            renta_fac_civa = round(renta_fac * 1.16, 2)
            subtot_fac_civa = round(subtotal_fac * 1.16, 2) if subtotal_fac > 0 else 0.0
            if abs(renta_fac_civa - renta_total_esp) <= 10.0 or abs(subtot_fac_civa - renta_total_esp) <= 10.0:
                dif = 0.0
                renta_total_esp = renta_fac
                observaciones_lista.append(
                    f"OK (Equivalencia IVA 16%: renta capturada con IVA ${con['Mensualidad_Sin_IVA']:,.2f} vs subtotal CFDI ${renta_fac:,.2f})"
                )
            else:
                # Caso B: Renta en contrato con doble deduccion de IVA (dif positiva)
                renta_esp_civa = round(renta_total_esp * 1.16, 2)
                if abs(renta_esp_civa - renta_fac) <= 10.0 or abs(renta_esp_civa - subtotal_fac) <= 10.0:
                    dif = 0.0
                    renta_total_esp = renta_fac
                    observaciones_lista.append(
                        f"OK (Equivalencia IVA 16%: deduccion doble de IVA en contrato ${con['Mensualidad_Sin_IVA']:,.2f} vs subtotal CFDI ${renta_fac:,.2f})"
                    )
        
        if es_multi_mes:
            obs_multi_mes = f"Factura con pago acumulado de {cant_meses} mensualidades (meses {mes_str} de {pl})"
            if abs(dif) <= TOLERANCIA:
                observaciones_lista.append(f"{obs_multi_mes} conciliada. Renta combinada esperada ${renta_total_esp:,.2f} coincide con lo facturado ${renta_fac:,.2f}.")
            else:
                errores.append(f"{obs_multi_mes}: esperado combinado ${renta_total_esp:,.2f}, facturado ${renta_fac:,.2f} (dif ${dif:+,.2f})")
        else:
            if abs(dif) > TOLERANCIA:
                errores.append(
                    f"Renta (Renta+Admin+GPS): esperado ${renta_total_esp:,.2f}, "
                    f"facturado ${renta_fac:,.2f} (dif ${dif:+,.2f})"
                )
            if aviso_mes_info:
                observaciones_lista.append(aviso_mes_info)

        status = 'CONCILIADO' if not errores else 'DISCREPANCIA'
        todas_obs = errores + observaciones_lista
        return _res(status, '; '.join(todas_obs) if todas_obs else 'OK',
                    esperado=renta_total_esp, facturado=renta_fac, dif=dif, detalle=todas_obs,
                    es_multi_mes=es_multi_mes, cant_meses=cant_meses,
                    meses_detectados=meses_detectados, mes_contrato_str=mes_str,
                    renta_esp=renta_total_esp, renta_fac=renta_fac, dif_renta=dif,
                    renta_pura_fac=renta_pura_fac,
                    admin_fac=admin_fac,
                    geoloc_fac=geoloc_fac)
    else:  # ANTICIPO
        errores = []
        obs_informativas = []

        # 1. Canal Apertura: Anticipo y Comision por Apertura
        ant_esp = round(float(con.get('Anticipo_Monto') or 0.0), 2)
        com_esp = round(float(con.get('Comision_Monto') or 0.0), 2)
        tot_apertura_esp = round(ant_esp + com_esp, 2)

        ant_fac = round(float(conceptos.get('ANTICIPO') or 0.0), 2)
        com_fac = round(float(conceptos.get('COMISION') or 0.0), 2)
        tot_apertura_fac = round(ant_fac + com_fac, 2)
        dif_apertura = round(tot_apertura_fac - tot_apertura_esp, 2)

        dif_ant = round(ant_fac - ant_esp, 2)
        dif_com = round(com_fac - com_esp, 2)

        # Diagnostico de agrupacion entre anticipo y comision
        es_agrupado_comision = (ant_fac == 0 and com_fac > 0 and ant_esp > 0)
        es_agrupado_anticipo = (com_fac == 0 and ant_fac > 0 and com_esp > 0)

        if abs(dif_apertura) <= TOLERANCIA:
            if es_agrupado_comision:
                obs_informativas.append("Anticipo facturado agrupado bajo concepto de comision por apertura")
            elif es_agrupado_anticipo:
                obs_informativas.append("Comision por apertura facturada agrupada bajo concepto de anticipo")
        else:
            nota_agrup = ""
            if es_agrupado_comision:
                nota_agrup = " [Anticipo facturado en comision]"
            elif es_agrupado_anticipo:
                nota_agrup = " [Comision facturada en anticipo]"
            errores.append(
                f"Apertura (Anticipo+Comision): esperado ${tot_apertura_esp:,.2f}, "
                f"facturado ${tot_apertura_fac:,.2f} (dif ${dif_apertura:+,.2f}){nota_agrup}"
            )

        # 2. Canal Renta Mes 1: Renta + Admin + GPS (solo si viene en la factura inicial)
        renta_mes1_fac = round(float(conceptos.get('RENTA') or 0.0) + float(conceptos.get('ADMIN') or 0.0) + float(conceptos.get('GEOLOC') or 0.0), 2)
        renta_mes1_esp = renta_total_esp if renta_mes1_fac > 0 else 0.0
        dif_r1 = round(renta_mes1_fac - renta_mes1_esp, 2) if renta_mes1_fac > 0 else 0.0

        if renta_mes1_fac > 0 and abs(dif_r1) > TOLERANCIA:
            errores.append(
                f"Renta mes 1 (Renta+Admin+GPS): esperado ${renta_mes1_esp:,.2f}, "
                f"facturado ${renta_mes1_fac:,.2f} (dif ${dif_r1:+,.2f})"
            )

        # 3. Totales consolidados de la factura
        total_esp = round(tot_apertura_esp + renta_mes1_esp, 2)
        total_fac = round(tot_apertura_fac + renta_mes1_fac, 2)
        dif_total = round(total_fac - total_esp, 2)

        status = 'CONCILIADO' if not errores else 'DISCREPANCIA'
        todas_obs = errores + obs_informativas
        obs_texto = '; '.join(todas_obs) if todas_obs else 'OK'

        return _res(status, obs_texto,
                    esperado=total_esp, facturado=total_fac, dif=dif_total, detalle=todas_obs,
                    ant_esp=ant_esp, ant_fac=ant_fac, dif_ant=dif_ant,
                    com_esp=com_esp, com_fac=com_fac, dif_com=dif_com,
                    tot_apertura_esp=tot_apertura_esp, tot_apertura_fac=tot_apertura_fac, dif_apertura=dif_apertura,
                    renta_esp=renta_mes1_esp, renta_fac=renta_mes1_fac, dif_renta=dif_r1)

def conciliar_factura(fact: dict, _cache: dict | None = None) -> dict:
    # Si la factura viene cancelada del repositorio o SAT, no se concilia y no cuenta para ningún cálculo
    if fact.get('cancelada') == 1 or str(fact.get('status')).upper() == 'CANCELADA':
        fec_c = str(fact.get('fecha_cancelacion') or '')[:10]
        msg_c = f"Comprobante fiscal cancelado en repositorio/SAT ({fec_c})" if fec_c else "Comprobante fiscal cancelado en repositorio/SAT"
        return {
            'status': 'CANCELADA',
            'msg': msg_c,
            'esperado': 0.0,
            'facturado': 0.0,
            'dif': 0.0,
            'id_contrato': fact.get('id_contrato'),
            'id_contrato_detectado': fact.get('id_contrato_detectado'),
            'mes_contrato': fact.get('mes_contrato'),
            'tipo': fact.get('tipo', 'MENSUAL'),
            'periodo': fact.get('periodo', ''),
            'cancelada': 1,
            'fecha_cancelacion': fact.get('fecha_cancelacion'),
        }

    res = _conciliar_factura_interna(fact, _cache)

    def _agregar_aviso(texto):
        res['msg'] = (res['msg'] + ' | ' + texto) if res['msg'] and res['msg'] != 'OK' else texto

    # Avisos de calidad de datos — no cambian el estatus, pero valen la
    # pena tenerlos a la vista.
    if res['status'] in ('CONCILIADO', 'DISCREPANCIA'):
        for _av in (fact.get('avisos_xml') or []):
            _agregar_aviso(_av)
        if fact.get('fecha'):
            try:
                if pd.to_datetime(fact['fecha']) > pd.Timestamp(hoy_ref()) + pd.Timedelta(days=1):
                    _agregar_aviso(f"Ojo: la fecha de emisión ({fact['fecha'][:10]}) es posterior a hoy — revisa si el XML está bien, o si el reloj de quien la timbró estaba mal.")
            except Exception:
                pass
        if not (fact.get('folio') or '').strip():
            _agregar_aviso("Ojo: este CFDI no trae folio — será más difícil rastrearlo después; considera pedirle al emisor que siempre lo capture.")

    # Aviso adicional (no cambia el estatus): ¿ya existe OTRA factura — con
    # distinto UUID — para este mismo contrato, mismo período y mismo tipo?
    # El cuadre de montos no detecta una doble facturación si por error se
    # subió/generó dos veces la renta del mismo mes; esto sí.
    if res['status'] in ('CONCILIADO', 'DISCREPANCIA') and fact.get('id_contrato') and fact.get('periodo'):
        try:
            conn = get_db()
            otras = conn.execute(
                """SELECT folio, uuid FROM facturas
                   WHERE id_contrato=? AND periodo=? AND tipo=? AND uuid != ?
                     AND (cancelada IS NULL OR cancelada=0)""",
                (fact['id_contrato'], fact['periodo'], fact.get('tipo', ''), fact.get('uuid', ''))
            ).fetchall()
            if otras:
                fol_actual_norm = normalizar_folio(fact.get('folio'))
                otras_distintas = [
                    o for o in otras
                    if not (fol_actual_norm and normalizar_folio(o['folio']) == fol_actual_norm)
                ]
                if otras_distintas:
                    folios = ', '.join(o['folio'] or o['uuid'][:8] for o in otras_distintas)
                    _agregar_aviso(f"Ojo: ya existe otra factura para este contrato en {fact['periodo']} (folio(s): {folios}) — revisa que no sea una doble facturación.")
        except Exception:
            pass
    if res['status'] in ('CONCILIADO', 'DISCREPANCIA') and fact.get('folio') and fact.get('uuid'):
        try:
            conn = get_db()
            fol_actual_norm = normalizar_folio(fact.get('folio'))
            if fol_actual_norm:
                todos_folios = conn.execute(
                    """SELECT uuid, folio, id_contrato FROM facturas
                       WHERE uuid != ? AND (cancelada IS NULL OR cancelada=0)""",
                    (fact['uuid'],)
                ).fetchall()
                mismo_folio = [
                    m for m in todos_folios
                    if normalizar_folio(m['folio']) == fol_actual_norm
                       and (m['id_contrato'] or '') != (fact.get('id_contrato') or '')
                ]
                if mismo_folio:
                    otros_c = ', '.join(f"{m['id_contrato'] or '—'} ({m['uuid'][:8]}…)" for m in mismo_folio)
                    _agregar_aviso(f"Ojo: el folio {fact['folio']} ya existe en otra factura de otro contrato ({otros_c}) — revisa si es un folio reutilizado.")
        except Exception:
            pass
    return res

def calcular_avance_pago(con: dict, facturas_dict: dict = None) -> dict:
    """Con base en las facturas ya conciliadas (no en el nivel de morosidad
    manual), calcula si el cliente va adelantado, atrasado o al corriente
    con sus mensualidades. 'Adelantado' significa que ya pagó meses futuros
    a los que le tocarían hoy por calendario; 'atrasado' significa que
    faltan meses ya vencidos por facturar/pagar (con el detalle de cuáles)."""
    fa = pd.to_datetime(con['Fecha_Alta'])
    pl = int(con['Plazo'])
    hoy = hoy_ref()
    mes_esperado_hoy = max(0, (hoy.year - fa.year) * 12 + (hoy.month - fa.month) + 1)
    mes_esperado_hoy = min(mes_esperado_hoy, pl)

    if facturas_dict is not None:
        meses_facturados = set()
        for m_val in facturas_dict.get(con['ID_Contrato'], set()):
            for m_chunk in str(m_val).split(','):
                m_chunk = m_chunk.strip().replace('.0', '')
                if m_chunk.isdigit():
                    meses_facturados.add(int(m_chunk))
    else:
        conn = get_db()
        rows = conn.execute(
            """SELECT DISTINCT mes_contrato FROM facturas
               WHERE (id_contrato=? OR instr(id_contrato, ?) > 0 OR instr(coalesce(id_contrato_detectado,''), ?) > 0)
                 AND tipo='MENSUAL' AND estatus IN ('CONCILIADO','DISCREPANCIA')
                 AND (cancelada IS NULL OR cancelada=0)""",
            (con['ID_Contrato'], con['ID_Contrato'], con['ID_Contrato'])
        ).fetchall()
        meses_facturados = set()
        for r in rows:
            if r['mes_contrato']:
                for m_chunk in str(r['mes_contrato']).split(','):
                    m_chunk = m_chunk.strip().replace('.0', '')
                    if m_chunk.isdigit():
                        meses_facturados.add(int(m_chunk))
    mes_max_facturado = max(meses_facturados) if meses_facturados else 0

    tope = min(mes_esperado_hoy, pl)
    meses_faltantes = sorted(m for m in range(1, tope + 1) if m not in meses_facturados)

    if meses_faltantes:
        estado = 'ATRASADO'
    elif mes_max_facturado > mes_esperado_hoy:
        estado = 'ADELANTADO'
    else:
        estado = 'AL_CORRIENTE'

    return {
        'estado': estado,
        'mes_esperado_hoy': mes_esperado_hoy,
        'mes_max_facturado': mes_max_facturado,
        'meses_adelanto': max(0, mes_max_facturado - mes_esperado_hoy),
        'meses_atraso': len(meses_faltantes),
        'meses_faltantes': meses_faltantes,
    }

def detectar_duplicados(lista: list) -> set:
    """Identificadores (UUIDs) de la lista que YA existen en la base de datos
    (por UUID exacto o por folio numérico idéntico en comprobantes activos)."""
    if not lista:
        return set()
    conn = get_db()
    uuids = [r['uuid'] for r in lista if r.get('uuid')]
    duplicados_uuids = set()

    # 1. Coincidencia por UUID exacto en comprobantes vigentes
    if uuids:
        placeholders = ','.join('?' * len(uuids))
        rows = conn.execute(
            f"SELECT uuid FROM facturas WHERE uuid IN ({placeholders}) AND (cancelada IS NULL OR cancelada=0)",
            uuids
        ).fetchall()
        for r in rows:
            duplicados_uuids.add(r[0])

    # 2. Coincidencia por folio numérico (la F- o serie no cuenta)
    folios_in = {}
    for r in lista:
        f_norm = normalizar_folio(r.get('folio'))
        if f_norm and r.get('uuid'):
            folios_in.setdefault(f_norm, []).append(r['uuid'])

    if folios_in:
        rows_fol = conn.execute(
            "SELECT uuid, folio FROM facturas WHERE folio IS NOT NULL AND (cancelada IS NULL OR cancelada=0)"
        ).fetchall()
        for r_u, r_fol in rows_fol:
            rf_norm = normalizar_folio(r_fol)
            if rf_norm in folios_in:
                for in_u in folios_in[rf_norm]:
                    duplicados_uuids.add(in_u)

    return duplicados_uuids

def guardar_facturas_batch(lista: list, sobrescribir_existentes: bool = True):
    conn = get_db()
    cur  = conn.cursor()
    existentes = detectar_duplicados(lista) if not sobrescribir_existentes else set()
    rows_fact = []
    uuids_guardadas = []
    conceptos_por_uuid = {}
    omitidas = []
    folios_en_batch = set()

    # Si se permite sobrescribir, desactivamos comprobantes previos activos que tengan
    # el mismo folio numérico bajo distinto UUID para evitar duplicar registros en BD
    if sobrescribir_existentes and lista:
        try:
            db_activos = conn.execute(
                "SELECT uuid, folio FROM facturas WHERE folio IS NOT NULL AND (cancelada IS NULL OR cancelada=0)"
            ).fetchall()
            db_fol_map = {}
            for db_u, db_fol in db_activos:
                fn = normalizar_folio(db_fol)
                if fn:
                    db_fol_map.setdefault(fn, set()).add(db_u)

            for r in lista:
                fn = normalizar_folio(r.get('folio'))
                u_in = r.get('uuid')
                if fn and fn in db_fol_map:
                    for old_u in db_fol_map[fn]:
                        if old_u != u_in:
                            cur.execute(
                                "UPDATE facturas SET cancelada=1, fecha_cancelacion=CURRENT_TIMESTAMP WHERE uuid=?",
                                (old_u,)
                            )
        except Exception:
            pass

    for r in lista:
        fn = normalizar_folio(r.get('folio'))
        if r['uuid'] in existentes or (fn and fn in folios_en_batch):
            omitidas.append(r.get('folio', r['uuid']))
            continue

        if fn:
            folios_en_batch.add(fn)

        folio_guardar = fn if fn else (r.get('folio') or '')
        idc_guardar = r.get('contratos_str') or r.get('id_contrato_detectado') or r.get('id_contrato')
        es_canc = 1 if (r.get('cancelada') == 1 or str(r.get('status')).upper() == 'CANCELADA') else 0
        fec_canc = str(r.get('fecha_cancelacion') or '')[:10] if es_canc else None
        st_guardar = 'CANCELADA' if es_canc else r['status']
        esp_guardar = 0.0 if es_canc else (r.get('esperado', 0) or 0)
        fac_guardar = 0.0 if es_canc else (r.get('facturado', 0) or 0)
        dif_guardar = 0.0 if es_canc else (r.get('dif', 0) or 0)

        rows_fact.append((
            r['uuid'],
            idc_guardar,
            r['fecha'],
            folio_guardar,
            r['subtotal'],
            r['total'],
            r['tipo'],
            r['periodo'],
            r.get('mes_contrato_str') or r.get('mes_contrato'),
            st_guardar,
            r.get('msg', ''),
            r.get('rfc_emisor'),
            r.get('rfc_receptor'),
            esp_guardar,
            fac_guardar,
            dif_guardar,
            idc_guardar,
            1 if r.get('alias_aplicado') else 0,
            es_canc,
            fec_canc,
        ))
        uuids_guardadas.append(r['uuid'])
        conceptos_por_uuid[r['uuid']] = r.get('conceptos_raw', [])
    if rows_fact:
        cur.executemany("""
            INSERT OR REPLACE INTO facturas
            (uuid, id_contrato, fecha_emision, folio, subtotal, total,
             tipo, periodo, mes_contrato, estatus, observaciones,
             rfc_emisor, rfc_receptor, esperado, facturado, diferencia,
             id_contrato_detectado, alias_aplicado, cancelada, fecha_cancelacion)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, rows_fact)
        # Se borran los renglones de concepto anteriores de estas facturas
        # antes de volver a insertarlos — si no, al reprocesar/sobrescribir
        # el mismo XML se irían acumulando líneas duplicadas.
        cur.executemany("DELETE FROM factura_conceptos WHERE uuid=?", [(u,) for u in uuids_guardadas])
        rows_conc = []
        for uuid, conceptos in conceptos_por_uuid.items():
            for cpto in conceptos:
                if cpto.get('importe'):
                    rows_conc.append((uuid, cpto['clave'], cpto.get('desc', ''), cpto['importe']))
        if rows_conc:
            cur.executemany("""
                INSERT INTO factura_conceptos (uuid, clave, descripcion, importe, manual)
                VALUES (?,?,?,?,0)
            """, rows_conc)
        conn.commit()
    return omitidas

def registrar_corrida_conciliacion(periodo: str, resultados: list, omitidas: int = 0):
    """Deja un renglón en el historial de corridas cada vez que se procesa un
    lote — a diferencia del historial de facturas (que se puede sobreescribir
    si reprocesas el mismo período), esto siempre suma un registro nuevo, para
    tener a la mano cuándo se concilió cada lote y con qué resultado."""
    conteos = {}
    for r in resultados:
        conteos[r['status']] = conteos.get(r['status'], 0) + 1
    conn = get_db()
    conn.execute("""
        INSERT INTO conciliacion_corridas
        (fecha_corrida, periodo, total_facturas, conciliadas, discrepancias,
         sin_contrato, errores, no_aplica, rfc_incorrecto, omitidas, rfc_configurado)
        VALUES (?,?,?,?,?,?,?,?,?,?,?)
    """, (
        datetime.now().isoformat(timespec='seconds'), periodo, len(resultados),
        conteos.get('CONCILIADO', 0), conteos.get('DISCREPANCIA', 0),
        conteos.get('SIN_CONTRATO', 0), conteos.get('ERROR', 0),
        conteos.get('NO_APLICA', 0), conteos.get('RFC_INCORRECTO', 0),
        omitidas, (get_cfg('rfc_arrendadora', '') or '').strip().upper(),
    ))
    conn.commit()

def reclasificar_concepto(concepto_id: int, nueva_clave: str):
    """Reclasifica a mano un concepto que el sistema no reconoció (o que
    reconoció distinto). Queda marcado como 'manual' para distinguirlo de
    los que se clasificaron solos por el texto del XML."""
    conn = get_db()
    conn.execute("UPDATE factura_conceptos SET clave=?, manual=1 WHERE id=?", (nueva_clave, concepto_id))
    conn.commit()

def desmarcar_concepto(concepto_id: int):
    """Deshace una reclasificación manual: el concepto vuelve a 'OTRO'
    (fuera de la comparación contra la póliza) y deja de contar como manual."""
    conn = get_db()
    conn.execute("UPDATE factura_conceptos SET clave='OTRO', manual=0 WHERE id=?", (concepto_id,))
    conn.commit()

def _totales_desde_conceptos(uuid: str) -> dict:
    """Reconstruye los totales por clave (Renta/Admin/Geoloc/Anticipo/
    Comisión/Otro) a partir de lo que quedó guardado en factura_conceptos —
    respetando cualquier reclasificación manual que se haya hecho."""
    conn = get_db()
    rows = conn.execute("SELECT clave, importe FROM factura_conceptos WHERE uuid=?", (uuid,)).fetchall()
    totales = {'RENTA': 0.0, 'ADMIN': 0.0, 'GEOLOC': 0.0, 'ANTICIPO': 0.0, 'COMISION': 0.0, 'OTRO': 0.0}
    for r in rows:
        totales[r['clave']] = totales.get(r['clave'], 0.0) + (r['importe'] or 0.0)
    return totales


def marcar_factura_cancelada(uuid: str, cancelada: bool = True):
    """Marca (o desmarca) una factura como cancelada ante el SAT. Una factura
    cancelada deja de contar en los totales de control y en 'contratos sin
    factura' — pero se conserva en el historial para trazabilidad, no se
    borra."""
    conn = get_db()
    conn.execute(
        "UPDATE facturas SET cancelada=?, fecha_cancelacion=? WHERE uuid=?",
        (1 if cancelada else 0, hoy_ref().isoformat() if cancelada else None, uuid)
    )
    conn.commit()

def re_conciliar_factura(uuid: str):
    """Re-conciliar una factura existente (por ejemplo, después de asignar
    contrato, actualizar comisión, o reclasificar un concepto a mano). Los
    totales se recalculan a partir de lo que quedó en factura_conceptos —no
    volviendo a leer el XML— para que cualquier reclasificación manual se
    respete en vez de perderse cada vez que se vuelve a conciliar."""
    conn = get_db()
    row = conn.execute("SELECT * FROM facturas WHERE uuid=?", (uuid,)).fetchone()
    if not row:
        return None
    if row['cancelada']:
        return {'status': 'CANCELADA', 'msg': 'Esta factura está marcada como cancelada — no se reconcilia.'}
    try:
        fact = dict(row)
        fact['conceptos'] = _totales_desde_conceptos(uuid)
        rows_c = conn.execute("SELECT descripcion, clave, importe FROM factura_conceptos WHERE uuid=?", (uuid,)).fetchall()
        fact['conceptos_raw'] = [{'desc': rc['descripcion'], 'clave': rc['clave'], 'importe': rc['importe']} for rc in rows_c]
        # Si al reclasificar a mano un concepto la factura pasó de "ningún
        # concepto de leasing reconocido" a "sí tiene alguno", se actualiza
        # el tipo (igual que se hace al leer el XML por primera vez) — si
        # no, se quedaría marcada para siempre como NO_APLICA aunque ya se
        # haya corregido la clasificación.
        conceptos_leasing = ('RENTA', 'ADMIN', 'GEOLOC', 'ANTICIPO', 'COMISION')
        tiene_concepto_lease = any(fact['conceptos'].get(k, 0) for k in conceptos_leasing)
        if tiene_concepto_lease:
            es_anticipo = fact['conceptos'].get('ANTICIPO', 0) or fact['conceptos'].get('COMISION', 0)
            fact['tipo'] = 'ANTICIPO' if es_anticipo else 'MENSUAL'
        res = conciliar_factura(fact)
        nuevo_idc = res.get('contratos_str') or fact.get('id_contrato')
        nuevo_idd = res.get('contratos_str') or fact.get('id_contrato_detectado') or fact.get('id_contrato')
        nuevo_mes = res.get('mes_contrato_str') or res.get('mes_contrato') or fact.get('mes_contrato')
        conn.execute(
            """UPDATE facturas
               SET estatus=?, observaciones=?, tipo=?, esperado=?, facturado=?, diferencia=?,
                   id_contrato=?, id_contrato_detectado=?, mes_contrato=?
               WHERE uuid=?""",
            (res['status'], res['msg'], fact['tipo'],
             res.get('esperado', 0) or 0, res.get('facturado', 0) or 0, res.get('dif', 0) or 0,
             nuevo_idc, nuevo_idd, nuevo_mes, uuid)
        )
        conn.commit()
        return res
    except Exception as e:
        return {'status':'ERROR', 'msg':str(e)}


def es_contrato_nuevo(con_or_id) -> bool:
    """Determina si un contrato es nuevo (iniciado a partir de septiembre 2026),
    por lo que no pertenece al histórico cerrado de enero a agosto 2026."""
    if not con_or_id:
        return False
    fa_val = None
    if isinstance(con_or_id, str):
        conn = get_db()
        row = conn.execute("SELECT Fecha_Alta FROM contratos WHERE ID_Contrato=?", (con_or_id.strip(),)).fetchone()
        if not row:
            return False
        fa_val = row['Fecha_Alta']
    elif isinstance(con_or_id, dict) or hasattr(con_or_id, 'get'):
        fa_val = con_or_id.get('Fecha_Alta')
    else:
        try:
            fa_val = con_or_id['Fecha_Alta']
        except Exception:
            return False
    if not fa_val or pd.isna(fa_val):
        return False
    try:
        fa = pd.to_datetime(fa_val)
        return fa >= pd.Timestamp('2026-09-01')
    except Exception:
        return False


def detectar_descuadre_iva_contrato(renta_contrato: float, renta_facturada: float) -> dict | None:
    """Analiza si la diferencia entre la renta del contrato y la facturada
    se debe a captura con IVA (+16%) o a deducción errónea de IVA (-13.8%).
    Retorna un diccionario con el diagnóstico o None si no es por IVA."""
    try:
        rc = float(renta_contrato or 0.0)
        rf = float(renta_facturada or 0.0)
    except Exception:
        return None
    if rc <= 0 or rf <= 0:
        return None
    if abs(rc - rf) <= 1.0:
        return None
    
    # Caso 1: Renta en contrato capturada con IVA (renta contrato es ~16% mayor que facturado sin IVA)
    # rf * 1.16 ~= rc
    rf_civa = round(rf * 1.16, 2)
    if abs(rf_civa - rc) <= 10.0:
        pct = round(((rc - rf) / rf) * 100, 1)
        return {
            'tipo': 'CON_IVA',
            'mensaje': f"Renta capturada con IVA en contrato (${rc:,.2f}) vs subtotal CFDI sin IVA (${rf:,.2f})",
            'detalle': f"Diferencia de +{pct}% (equivale exactamente al 16% de IVA trasladado).",
            'renta_correcta': rf,
            'renta_actual': rc
        }
    
    # Caso 2: Renta en contrato con deducción errónea de IVA (se restó IVA dos veces, renta contrato es ~13.8% menor)
    # rc * 1.16 ~= rf
    rc_civa = round(rc * 1.16, 2)
    if abs(rc_civa - rf) <= 10.0:
        pct = round(((rf - rc) / rf) * 100, 1)
        return {
            'tipo': 'DOBLE_DEDUCCION',
            'mensaje': f"Deducción errónea de IVA en contrato (${rc:,.2f}) vs subtotal CFDI real sin IVA (${rf:,.2f})",
            'detalle': f"Diferencia de -{pct}% (se restó IVA a un valor que ya venía sin IVA).",
            'renta_correcta': rf,
            'renta_actual': rc
        }
    return None


def corregir_renta_contrato_nuevo(id_contrato: str, nueva_renta: float) -> tuple[bool, str]:
    """Corrige la mensualidad sin IVA y recalcula la corrida para un contrato NUEVO.
    Protege inmutablemente cualquier contrato del periodo cerrado (< 2026-09-01)."""
    conn = get_db()
    c = conn.execute("SELECT * FROM contratos WHERE ID_Contrato=?", (id_contrato.strip(),)).fetchone()
    if not c:
        return False, f"Contrato {id_contrato} no encontrado."
    if not es_contrato_nuevo(dict(c)):
        fa_str = str(c['Fecha_Alta'])[:10]
        return False, f"El contrato {id_contrato} inició el {fa_str} (periodo cerrado). No se puede alterar directamente para proteger los saldos auditados."
    
    pl = int(c['Plazo'])
    inv = float(c['Valor_Sin_IVA']) - float(c['Anticipo_Monto'])
    res = float(c['Residual_Monto'])
    t_nueva = calc_tasa(pl, float(nueva_renta), inv, res)
    if t_nueva is None:
        return False, "No se pudo calcular una tasa financiera válida con la nueva renta."
    
    vpr = float(vp_res(res, t_nueva, pl))
    conn.execute(
        "UPDATE contratos SET Mensualidad_Sin_IVA=?, Tasa_Calculada=?, VP_Residual=? WHERE ID_Contrato=?",
        (round(float(nueva_renta), 2), t_nueva, vpr, id_contrato.strip())
    )
    conn.commit()
    _tocar_datos()
    
    # Re-conciliar facturas del contrato
    facs = conn.execute(
        "SELECT uuid FROM facturas WHERE (id_contrato=? OR id_contrato_detectado=?) AND (cancelada IS NULL OR cancelada=0)",
        (id_contrato.strip(), id_contrato.strip())
    ).fetchall()
    reconciliadas = 0
    for f in facs:
        re_conciliar_factura(f['uuid'])
        reconciliadas += 1
        
    return True, f"Contrato {id_contrato} actualizado: Renta sin IVA ajustada a ${nueva_renta:,.2f}, tasa recalculada a {t_nueva*100:.2f}% mensual. Se re-conciliaron {reconciliadas} factura(s)."


# --- Pantalla de conciliación de facturas ---
def _render_conciliacion():
    from reports.excel import excel_con_formato
    st.title("Conciliación de Facturas CFDI")
    st.caption("Revisa y cuadra las facturas emitidas contra lo programado en tus contratos.")

    TABS_CONCIL = [
        "Carga y Conciliación",
        "Historial",
        "Comparativo de Cifras",
        "Resolver Pendientes",
        "Conceptos sin Reconocer",
        "Avance de Pago",
        "Auditoría por Contrato",
        "Reporte Consolidado"
    ]

    st.markdown("""
        <style>
        div[data-testid="stSegmentedControl"] {
            display: flex;
            flex-wrap: wrap;
            gap: 6px;
            background: #F8F9FA;
            padding: 6px;
            border-radius: 8px;
            border: 1px solid #E5E7EB;
            margin-bottom: 1.2rem;
        }
        div[data-testid="stSegmentedControl"] button {
            border: none !important;
            border-radius: 6px !important;
            font-weight: 600 !important;
            font-size: 0.88rem !important;
            padding: 7px 15px !important;
            transition: all 0.15s ease-in-out !important;
        }
        div[data-testid="stSegmentedControl"] button[aria-selected="true"] {
            background-color: #0E7090 !important;
            color: #FFFFFF !important;
            box-shadow: 0 1px 3px rgba(0,0,0,0.12) !important;
        }
        </style>
    """, unsafe_allow_html=True)

    tab_activa = st.segmented_control(
        "Navegación de Conciliación",
        options=TABS_CONCIL,
        default=st.session_state.get("concil_tab_activa", TABS_CONCIL[0]),
        key="concil_tab_activa",
        label_visibility="collapsed"
    ) or TABS_CONCIL[0]

    # ===================================================================
    # TAB 1 – CARGA Y CONCILIACIÓN
    # ===================================================================
    if tab_activa == "Carga y Conciliación":
        with st.expander("RFC emisor de tu empresa", expanded=not (get_cfg('rfc_arrendadora', '') or '').strip()):
            st.caption("Solo se procesarán facturas emitidas por este RFC.")
            _rfc_actual = get_cfg('rfc_arrendadora', '') or ''
            _rfc_nuevo = st.text_input("RFC de tu arrendadora", value=_rfc_actual, max_chars=13,
                                        placeholder="Ej. ABC010101AB1").strip().upper()
            if st.button("Guardar RFC"):
                set_cfg('rfc_arrendadora', _rfc_nuevo)
                st.success("RFC guardado. Se aplicará a partir de la próxima carga.")

        # ===================================================================
        # PIPELINE DE CONCILIACIÓN REUTILIZABLE
        # ===================================================================
        def _ejecutar_conciliacion_lote(facts_lote, sobrescribir=True, barra=None, estado=None, label_lote="Procesando"):
            total = len(facts_lote)
            if total == 0:
                return [], [], []
            reglas = get_reglas_concepto()
            cache_lote = {}
            alias_map = cargar_alias_contrato()
            usos_alias = {}
            resultados = []
            errores = []
            uuids_en_lote = set()
            folios_en_lote = set()
            step_update = max(1, total // 40)

            for i, fact in enumerate(facts_lote):
                if barra and (i % step_update == 0 or i == total - 1):
                    pct = (i + 1) / total
                    fol = str(fact.get("folio") or fact.get("uuid", "")[:8])
                    barra.progress(pct, text=f"{label_lote} ({i+1}/{total}) - Folio {fol}")
                    if estado:
                        c_id = fact.get("id_contrato") or "Sin contrato"
                        estado.caption(f"Conciliando {i+1} de {total}: Folio `{fol}` | Contrato `{c_id}` | Periodo `{fact.get('periodo', '-')}`")

                u = fact.get("uuid")
                fn = normalizar_folio(fact.get("folio"))
                if (u and u in uuids_en_lote) or (fn and fn in folios_en_lote):
                    continue
                if u:
                    uuids_en_lote.add(u)
                if fn:
                    folios_en_lote.add(fn)

                try:
                    detectado = fact.get("id_contrato")
                    id_final, se_aplico = resolver_numero_contrato(detectado, alias_map)
                    fact["id_contrato"] = id_final
                    fact["alias_aplicado"] = 1 if se_aplico else 0
                    if se_aplico and detectado:
                        usos_alias[detectado] = usos_alias.get(detectado, 0) + 1

                    res = conciliar_factura(fact, cache_lote)
                    merged = {**fact, **res, "status": res["status"]}
                    resultados.append(merged)
                except Exception as e:
                    errores.append(f"Factura {fact.get('folio', i)}: {e}")

            if barra:
                barra.progress(1.0, text=f"Guardando {len(resultados)} facturas en base de datos...")
            if estado:
                estado.caption("Guardando registros en SQLite...")

            marcar_alias_usado_lote(usos_alias)
            omitidas = guardar_facturas_batch(resultados, sobrescribir_existentes=sobrescribir)

            pers = sorted(list({r.get("periodo") for r in resultados if r.get("periodo")}))
            per_corrida = f"{pers[0]} a {pers[-1]}" if len(pers) > 1 else (pers[0] if pers else "Varios")
            registrar_corrida_conciliacion(per_corrida, resultados, omitidas=len(omitidas))

            return resultados, omitidas, errores

        def _mostrar_resumen_lote(resultados, omitidas, errores, titulo="Lote Procesado"):
            if not resultados:
                st.error("No se procesó ninguna factura correctamente.")
                st.session_state["conciliacion_resultados"] = None
                return

            st.session_state["conciliacion_resultados"] = resultados
            st.session_state["conciliacion_omitidas"] = omitidas
            n_conc = sum(1 for r in resultados if r.get('status') == 'CONCILIADO')
            n_disc = sum(1 for r in resultados if r.get('status') == 'DISCREPANCIA')
            n_sin  = sum(1 for r in resultados if r.get('status') == 'SIN_CONTRATO')
            n_canc = sum(1 for r in resultados if (r.get('status') == 'CANCELADA' or r.get('cancelada') == 1))
            n_otro = sum(1 for r in resultados if r.get('status') in ('NO_APLICA', 'RFC_INCORRECTO', 'ERROR'))
            monto_proc = sum(float(r.get('total') or 0) for r in resultados if not (r.get('status') == 'CANCELADA' or r.get('cancelada') == 1))
            pers = sorted(list({r.get('periodo') for r in resultados if r.get('periodo') and not (r.get('status') == 'CANCELADA' or r.get('cancelada') == 1)}))
            rango_per = f"{pers[0]} a {pers[-1]}" if len(pers) > 1 else (pers[0] if pers else "-")

            st.success(
                f"**¡{titulo} exitosamente!** Se procesaron **{len(resultados):,}** comprobante(s) "
                f"cubriendo el período **{rango_per}** por un total vigente de **${monto_proc:,.2f}**."
            )
            c_k1, c_k2, c_k3, c_k4, c_k5, c_k6 = st.columns(6)
            c_k1.metric("Conciliadas", f"{n_conc:,}")
            c_k2.metric("Discrepancia", f"{n_disc:,}")
            c_k3.metric("Sin Contrato", f"{n_sin:,}")
            c_k4.metric("Canceladas (SAT)", f"{n_canc:,}")
            c_k5.metric("Otros / No Aplica", f"{n_otro:,}")
            c_k6.metric("Omitidas (ya en BD)", f"{len(omitidas):,}")

            if n_sin > 0:
                st.info(f"Nota: Hay **{n_sin:,}** facturas sin contrato asignado. Ve a la pestaña **'Resolver Pendientes'** para vincularlas.")

            if errores:
                with st.expander(f"{len(errores)} observación(es) durante el proceso", expanded=False):
                    for em in errores[:30]:
                        st.warning(em)
                    if len(errores) > 30:
                        st.caption(f"... y {len(errores)-30} más.")
            st.session_state["_refresh"] = True

        metodo_carga = st.radio(
            "Modalidad de carga:",
            [
                "Carga desde carpeta (Repositorio Y:)",
                "Subida manual (Archivos o ZIP)",
            ],
            horizontal=True,
            key="radio_metodo_carga_cfdi",
        )

        # ===================================================================
        # MODALIDAD 1: CARGA DIRECTA DESDE REPOSITORIO EN DISCO
        # ===================================================================
        if "Carga desde carpeta" in metodo_carga:
            c_btn1, c_btn2 = st.columns(2)
            with c_btn1:
                if st.button("Usar Repositorio Excel (Normas Internacionales)", use_container_width=True, key="btn_path_excel"):
                    st.session_state["repo_dir_val"] = r"Y:\Leasy Arrendamientos\NORMAS INTERNACIONALES\REPOSITORIOS"
                    st.rerun()
            with c_btn2:
                if st.button("Usar Repositorio XMLs (REPOSITORIOS SAT)", use_container_width=True, key="btn_path_xml"):
                    st.session_state["repo_dir_val"] = r"Y:\Leasy Arrendamientos\REPOSITORIOS SAT"
                    st.rerun()

            default_repo = st.session_state.get("repo_dir_val", r"Y:\Leasy Arrendamientos\NORMAS INTERNACIONALES\REPOSITORIOS")
            repo_dir = st.text_input("Ruta de la carpeta del repositorio:", value=default_repo, key="txt_repo_dir")

            if not os.path.exists(repo_dir):
                st.warning(f"La ruta `{repo_dir}` no fue encontrada o no está accesible. Verifica que la unidad de red esté conectada.")
            else:
                carpetas_anios = sorted([d for d in os.listdir(repo_dir) if os.path.isdir(os.path.join(repo_dir, d)) and d.isdigit()])
                
                c_sel_a, c_tipo_f = st.columns([2, 1])
                with c_sel_a:
                    anios_sel = st.multiselect(
                        "Años a procesar:",
                        options=carpetas_anios,
                        default=carpetas_anios,
                        help="Selecciona los años que deseas cargar a la base de datos",
                        key="ms_anios_disco"
                    )
                with c_tipo_f:
                    es_repo_excel = "NORMAS INTERNACIONALES" in repo_dir.upper() or any("emitid" in f.lower() for _, _, files in os.walk(repo_dir) for f in files[:2])
                    tipo_rep = st.radio(
                        "Tipo de archivo a buscar:",
                        ["Archivos Excel Emitidos (*.xlsx)", "Archivos XML (*.xml)"],
                        index=0 if es_repo_excel else 1,
                        key="radio_tipo_disco"
                    )

                c_chk1, c_chk2 = st.columns(2)
                with c_chk1:
                    sobrescribir_disco = st.checkbox(
                        "Sobrescribir facturas existentes",
                        value=True,
                        help="Actualiza los datos si la factura ya estaba cargada.",
                        key="chk_sobre_disco"
                    )
                with c_chk2:
                    omitir_canceladas_disco = st.checkbox(
                        "Omitir canceladas en el SAT",
                        value=False,
                        help="No sube comprobantes con estatus cancelado.",
                        key="chk_omitir_canceladas_disco"
                    )

                if st.button("Cargar y Conciliar Repositorio en Disco", type="primary", use_container_width=True, key="btn_ejecutar_disco"):
                    if not anios_sel:
                        st.warning("Selecciona al menos un año para procesar.")
                    else:
                        reglas = get_reglas_concepto()
                        barra_d = st.progress(0, text="Buscando archivos en el repositorio...")
                        estado_d = st.empty()
                        facts_totales = []
                        archivos_encontrados = []

                        # Escanear archivos
                        if "Excel" in tipo_rep:
                            estado_d.caption("Localizando archivos Excel de comprobantes emitidos...")
                            for yr in anios_sel:
                                y_path = os.path.join(repo_dir, yr)
                                if os.path.exists(y_path):
                                    for fname in os.listdir(y_path):
                                        fl = fname.lower()
                                        if fl.endswith((".xlsx", ".csv")) and "emitid" in fl and "egreso" not in fl and "pago" not in fl:
                                            archivos_encontrados.append(os.path.join(y_path, fname))
                            
                            # Si no se encontraron por "emitid", agregar todos los xlsx del año
                            if not archivos_encontrados:
                                for yr in anios_sel:
                                    y_path = os.path.join(repo_dir, yr)
                                    if os.path.exists(y_path):
                                        for fname in os.listdir(y_path):
                                            if fname.lower().endswith((".xlsx", ".csv")):
                                                archivos_encontrados.append(os.path.join(y_path, fname))

                            total_arch = len(archivos_encontrados)
                            if total_arch == 0:
                                st.warning("No se encontraron archivos Excel para los años seleccionados.")
                            else:
                                for idx_a, f_path in enumerate(archivos_encontrados):
                                    f_name = os.path.basename(f_path)
                                    estado_d.caption(f"Leyendo Excel {idx_a+1} de {total_arch}: `{f_name}`...")
                                    try:
                                        facts_arch = parse_sat_excel(f_path, reglas)
                                        facts_totales.extend(facts_arch)
                                    except Exception as e_p:
                                        st.warning(f"Error al leer `{f_name}`: {e_p}")

                        else:
                            estado_d.caption("Buscando archivos XML en las carpetas de los años seleccionados...")
                            for yr in anios_sel:
                                y_path = os.path.join(repo_dir, yr)
                                if os.path.exists(y_path):
                                    for root, _, files in os.walk(y_path):
                                        for f in files:
                                            if f.lower().endswith(".xml"):
                                                archivos_encontrados.append(os.path.join(root, f))

                            total_arch = len(archivos_encontrados)
                            if total_arch == 0:
                                st.warning("No se encontraron archivos XML para los años seleccionados.")
                            else:
                                step_xml = max(1, total_arch // 50)
                                for idx_x, x_path in enumerate(archivos_encontrados):
                                    if idx_x % step_xml == 0 or idx_x == total_arch - 1:
                                        pct = (idx_x + 1) / total_arch
                                        barra_d.progress(pct, text=f"Leyendo XMLs: {idx_x+1}/{total_arch}")
                                    try:
                                        with open(x_path, "rb") as xf:
                                            raw_x = xf.read()
                                        fact_x = parse_cfdi(raw_x, reglas)
                                        facts_totales.append(fact_x)
                                    except Exception as e_x:
                                        pass

                        if facts_totales:
                            if omitir_canceladas_disco:
                                facts_totales = [f for f in facts_totales if f.get("cancelada") != 1 and f.get("status") != "CANCELADA"]
                            res_d, omit_d, err_d = _ejecutar_conciliacion_lote(
                                facts_totales,
                                sobrescribir=sobrescribir_disco,
                                barra=barra_d,
                                estado=estado_d,
                                label_lote="Conciliando comprobantes"
                            )
                            barra_d.empty()
                            estado_d.empty()
                            _mostrar_resumen_lote(res_d, omit_d, err_d, titulo="Repositorio Procesado")
                        else:
                            barra_d.empty()
                            estado_d.empty()
                            st.warning("No se extrajeron comprobantes de las rutas seleccionadas.")

        # ===================================================================
        # MODALIDAD 2: SUBIDA MANUAL POR NAVEGADOR
        # ===================================================================
        else:
            col_per, col_over = st.columns(2)
            with col_per:
                periodo_ref = st.date_input(
                    "Período de respaldo",
                    value=hoy_ref().replace(day=1),
                    help="Solo se usará si un comprobante no incluye fecha legible.",
                    key="carga_periodo_ref",
                )
            with col_over:
                st.write("")
                sobrescribir = st.checkbox(
                    "Sobrescribir facturas existentes",
                    value=True,
                    help="Actualiza los datos si el UUID ya existe en la base de datos.",
                    key="carga_sobrescribir",
                )
                omitir_canceladas_manual = st.checkbox(
                    "Omitir canceladas en el SAT",
                    value=False,
                    help="No sube comprobantes con estatus cancelado.",
                    key="carga_omitir_canceladas_manual",
                )

            archivos = st.file_uploader(
                "Arrastra tus archivos XML, ZIP o Excel aquí",
                type=["xml", "zip", "xlsx", "csv"],
                accept_multiple_files=True,
                key="carga_cfdi_uploader_unico",
            )

            def _expandir_archivos_subidos(lista):
                import zipfile
                out = []
                errores = []
                if not lista:
                    return out, errores
                for arch in lista:
                    name = arch.name or "sin_nombre"
                    name_l = name.lower()
                    try:
                        raw = arch.getvalue() if hasattr(arch, "getvalue") else arch.read()
                        if not raw:
                            errores.append(f"{name}: archivo vacio (0 bytes).")
                            continue
                        if name_l.endswith(".zip"):
                            try:
                                with zipfile.ZipFile(io.BytesIO(raw)) as zf:
                                    xml_names = [n for n in zf.namelist()
                                                 if n.lower().endswith(".xml") and not n.startswith("__MACOSX")]
                                    if not xml_names:
                                        errores.append(f"{name}: el ZIP no contiene ningun archivo .xml")
                                        continue
                                    for n in xml_names:
                                        out.append((os.path.basename(n.replace("\\", "/")) or n, zf.read(n)))
                            except zipfile.BadZipFile:
                                errores.append(f"{name}: archivo ZIP dañado o inválido.")
                        elif name_l.endswith(".xml"):
                            out.append((name, raw))
                        elif name_l.endswith((".xlsx", ".csv")):
                            out.append((name, raw))
                        else:
                            errores.append(f"{name}: extensión no soportada.")
                    except Exception as e:
                        errores.append(f"{name}: error al leer ({e})")
                return out, errores

            col_btn_up, col_inf_up = st.columns([1, 3])
            with col_btn_up:
                procesar_clic = st.button(
                    "Procesar Archivos",
                    type="primary",
                    disabled=not bool(archivos),
                    key="btn_procesar_cfdi",
                    use_container_width=True,
                )
            with col_inf_up:
                if archivos:
                    st.caption(f"{len(archivos)} archivo(s) seleccionado(s). Haz clic en Procesar Archivos.")
                else:
                    st.caption("Arrastra o selecciona archivos XML, ZIP, XLSX o CSV arriba.")

            if procesar_clic and archivos:
                items, err_prev = _expandir_archivos_subidos(archivos)
                n_xml = sum(1 for n, _ in items if n.lower().endswith(".xml"))
                n_xls = sum(1 for n, _ in items if n.lower().endswith((".xlsx", ".csv")))
                peso_kb = sum(len(b) for _, b in items) / 1024.0

                st.info(f"{len(items)} archivo(s) expandidos ({n_xml} XML, {n_xls} Excel/CSV) - {peso_kb:,.0f} KB total")

                if err_prev:
                    with st.expander(f"Observaciones ({len(err_prev)})", expanded=True):
                        for msg in err_prev:
                            st.warning(msg)

                if items:
                    facts_navegador = []
                    errores_parse = list(err_prev)
                    reglas = get_reglas_concepto()

                    for nombre, raw in items:
                        name_lower = nombre.lower()
                        try:
                            if name_lower.endswith(".xml"):
                                if len(raw) > MAX_XML_BYTES:
                                    raise ValueError(f"pesa {len(raw)/1024:.0f} KB (límite {MAX_XML_BYTES//1024//1024} MB).")
                                facts_navegador.append(parse_cfdi(raw, reglas))
                            elif name_lower.endswith(".xlsx"):
                                try:
                                    facts_sat = parse_sat_excel(raw, reglas)
                                    facts_navegador.extend(facts_sat)
                                except Exception:
                                    df_arch = pd.read_excel(io.BytesIO(raw))
                                    df_arch.columns = [str(col).strip().upper() for col in df_arch.columns]
                                    for _, r in df_arch.iterrows():
                                        fec = str(r.get("FECHA", "") or "")
                                        facts_navegador.append({
                                            "uuid": str(r.get("UUID", "") or ""),
                                            "fecha": fec,
                                            "folio": str(r.get("FOLIO", "") or ""),
                                            "subtotal": float(r.get("SUBTOTAL", 0) or 0),
                                            "total": float(r.get("TOTAL", 0) or 0),
                                            "id_contrato": str(r.get("CONTRATO", "") or "") or None,
                                            "mes_contrato": None,
                                            "tipo": "MENSUAL",
                                            "periodo": fec[:7] if fec else "",
                                            "conceptos": {"RENTA": float(r.get("SUBTOTAL", 0) or 0)},
                                            "conceptos_raw": [], "rfc_emisor": None, "rfc_receptor": None,
                                            "tipo_comprobante": "I", "moneda": "MXN", "avisos_xml": [], "alias_aplicado": 0,
                                        })
                            elif name_lower.endswith(".csv"):
                                df_arch = pd.read_csv(io.BytesIO(raw))
                                df_arch.columns = [str(col).strip().upper() for col in df_arch.columns]
                                for _, r in df_arch.iterrows():
                                    fec = str(r.get("FECHA", "") or "")
                                    facts_navegador.append({
                                        "uuid": str(r.get("UUID", "") or ""),
                                        "fecha": fec,
                                        "folio": str(r.get("FOLIO", "") or ""),
                                        "subtotal": float(r.get("SUBTOTAL", 0) or 0),
                                        "total": float(r.get("TOTAL", 0) or 0),
                                        "id_contrato": str(r.get("CONTRATO", "") or "") or None,
                                        "mes_contrato": None,
                                        "tipo": "MENSUAL",
                                        "periodo": fec[:7] if fec else "",
                                        "conceptos": {"RENTA": float(r.get("SUBTOTAL", 0) or 0)},
                                        "conceptos_raw": [], "rfc_emisor": None, "rfc_receptor": None,
                                        "tipo_comprobante": "I", "moneda": "MXN", "avisos_xml": [], "alias_aplicado": 0,
                                    })
                        except Exception as e_arch:
                            errores_parse.append(f"**{nombre}**: {e_arch}")

                    # Asegurar período de respaldo solo si falta
                    for fact in facts_navegador:
                        if not fact.get("periodo"):
                            fact["periodo"] = periodo_ref.strftime("%Y-%m")

                    if omitir_canceladas_manual:
                        facts_navegador = [f for f in facts_navegador if f.get("cancelada") != 1 and f.get("status") != "CANCELADA"]

                    barra_n = st.progress(0, text="Iniciando conciliación...")
                    estado_n = st.empty()

                    res_n, omit_n, err_n = _ejecutar_conciliacion_lote(
                        facts_navegador,
                        sobrescribir=sobrescribir,
                        barra=barra_n,
                        estado=estado_n,
                        label_lote="Procesando archivos"
                    )
                    barra_n.empty()
                    estado_n.empty()
                    _mostrar_resumen_lote(res_n, omit_n, errores_parse + err_n, titulo="Archivos Subidos")
                else:
                    st.warning("No se encontraron archivos válidos en la selección.")



        st.markdown("---")
        # Resultados guardados permanentemente

        _omitidas_ultima = st.session_state.get('conciliacion_omitidas') or []
        if _omitidas_ultima:
            st.warning(
                f"**{len(_omitidas_ultima)} factura(s) ya existían** y se omitieron para no perder resoluciones "
                f"previas: {', '.join(str(x) for x in _omitidas_ultima[:15])}"
                + (f" y {len(_omitidas_ultima)-15} más…" if len(_omitidas_ultima) > 15 else "")
                + " — marca 'Sobrescribir facturas ya cargadas' arriba si de verdad quieres reemplazarlas."
            )

        st.divider()
        st.markdown("**Historial de facturas procesadas**")
        st.caption("Consulta el detalle de los períodos previamente procesados.")
        conn_res = get_db()
        _periodos_todos = [r[0] for r in conn_res.execute(
            "SELECT DISTINCT periodo FROM facturas WHERE periodo IS NOT NULL ORDER BY periodo DESC"
        ).fetchall()]

        if not _periodos_todos:
            st.info("Aún no has procesado ningún XML.")
        else:
            _opciones_periodos = ["TODOS LOS MESES (VER ACUMULADO ANUAL)"] + _periodos_todos
            _periodo_default = st.session_state.get('conciliacion_periodo_lote')
            _idx_default = _opciones_periodos.index(_periodo_default) if _periodo_default in _opciones_periodos else 0
            periodo_ver = st.selectbox("Mes / Período a revisar", _opciones_periodos, index=_idx_default, key="periodo_ver_resultados")

            if periodo_ver == "TODOS LOS MESES (VER ACUMULADO ANUAL)":
                filas = conn_res.execute(
                    """SELECT uuid, folio, id_contrato, tipo, mes_contrato, subtotal, total,
                              estatus, observaciones, esperado, facturado, diferencia, cancelada, periodo
                       FROM facturas ORDER BY periodo DESC, estatus, id_contrato"""
                ).fetchall()
            else:
                filas = conn_res.execute(
                    """SELECT uuid, folio, id_contrato, tipo, mes_contrato, subtotal, total,
                              estatus, observaciones, esperado, facturado, diferencia, cancelada, periodo
                       FROM facturas WHERE periodo=? ORDER BY estatus, id_contrato""",
                    (periodo_ver,)
                ).fetchall()

            if not filas:
                st.info(f"No hay facturas guardadas para {periodo_ver}.")
            else:
                n_ok   = sum(1 for f in filas if f['estatus'] == 'CONCILIADO')
                n_disc = sum(1 for f in filas if f['estatus'] == 'DISCREPANCIA')
                n_sin  = sum(1 for f in filas if f['estatus'] == 'SIN_CONTRATO')
                n_canc = sum(1 for f in filas if f['estatus'] == 'CANCELADA' or f['cancelada'] == 1)
                n_err  = sum(1 for f in filas if f['estatus'] == 'ERROR')
                n_noa  = sum(1 for f in filas if f['estatus'] == 'NO_APLICA')
                n_rfc  = sum(1 for f in filas if f['estatus'] == 'RFC_INCORRECTO')

                k1, k2, k3, k4, k5, k6, k7 = st.columns(7)
                k1.metric("Conciliados",   n_ok)
                k2.metric("Discrepancia",  n_disc)
                k3.metric("Sin contrato",  n_sin)
                k4.metric("Canceladas",    n_canc)
                k5.metric("Error",         n_err)
                k6.metric("No aplica",     n_noa)
                k7.metric("RFC incorrecto", n_rfc)

                # Tabla de resultados
                df_res = pd.DataFrame([{
                    'Folio':        f['folio'] or '-',
                    'UUID':         (f['uuid'] or '')[:8] + '…',
                    'Contrato':     f['id_contrato'] or '—',
                    'Tipo':         f['tipo'] or '-',
                    'Mes':          str(f['mes_contrato']).replace('.0', '') if f['mes_contrato'] is not None else '-',
                    'Facturado $':  f['facturado'] or 0,
                    'Esperado $':   f['esperado'] or 0,
                    'Diferencia $': f['diferencia'] or 0,
                    'Estatus':      f['estatus'] or '-',
                    'Observación':  (f['observaciones'] or '')[:80],
                } for f in filas])

                def colorear(val):
                    if val == 'CONCILIADO':  return 'background-color:#E1F2E7;color:#155C3B'
                    if val == 'DISCREPANCIA': return 'background-color:#FBF0DA;color:#7A5209'
                    if val == 'SIN_CONTRATO': return 'background-color:#F3EDD3;color:#6B5A1F'
                    if val == 'ERROR':        return 'background-color:#FAE3E1;color:#8A2019'
                    if val == 'NO_APLICA':    return 'background-color:#E9EBEE;color:#565E68'
                    if val == 'RFC_INCORRECTO': return 'background-color:#f5d0d0;color:#7A1015;font-weight:700'
                    if val == 'CANCELADA':    return 'background-color:#e2e2e2;color:#555;text-decoration:line-through'
                    if val == 'FUERA_DE_VIGENCIA': return 'background-color:#e2c2c2;color:#6b1515;font-weight:700;text-decoration:line-through'
                    return ''

                fmt = {'Facturado $': '${:,.2f}', 'Esperado $': '${:,.2f}', 'Diferencia $': '${:+,.2f}'}
                styled = df_res.style.format(fmt).map(colorear, subset=['Estatus'])
                st.dataframe(styled, width='stretch', height=400, key="df_002")

                if n_disc > 0 or n_sin > 0:
                    st.info(
                        f"**Acciones pendientes:** Hay **{n_disc}** factura(s) con discrepancia y **{n_sin}** sin contrato en este período. "
                        f"Puedes resolverlas, asignar contratos o actualizar comisiones en la pestaña **'Resolver Pendientes'**."
                    )

                csv_bytes = df_res.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    "Descargar resultados CSV",
                    csv_bytes,
                    f"conciliacion_{periodo_ver}.csv",
                    mime='text/csv'
                )

    # ===================================================================
    # TAB 2 – HISTORIAL
    # ===================================================================
    elif tab_activa == "Historial":
        st.subheader("Historial de Facturas Procesadas")
        conn = get_db()
        periodos = conn.execute(
            "SELECT DISTINCT periodo FROM facturas WHERE periodo IS NOT NULL ORDER BY periodo DESC"
        ).fetchall()
        lista_per = [r[0] for r in periodos]

        if not lista_per:
            st.info("Aún no hay facturas procesadas.")
        else:
            per_sel = st.selectbox("Período", lista_per)
            estatus_fil = st.multiselect(
                "Filtrar por estatus",
                ['CONCILIADO', 'DISCREPANCIA', 'SIN_CONTRATO', 'PENDIENTE', 'ERROR', 'NO_APLICA', 'RFC_INCORRECTO', 'CANCELADA'],
                default=['CONCILIADO', 'DISCREPANCIA', 'SIN_CONTRATO', 'ERROR', 'NO_APLICA', 'RFC_INCORRECTO']
            )
            placeholders = ','.join('?' * len(estatus_fil))
            rows = conn.execute(
                f"""SELECT f.uuid, f.folio, f.id_contrato, f.tipo, f.mes_contrato,
                           f.subtotal, f.total, f.estatus,
                           COALESCE((SELECT GROUP_CONCAT(fc.descripcion, ' | ') FROM factura_conceptos fc WHERE fc.uuid = f.uuid), f.observaciones) as concepto,
                           f.observaciones, f.fecha_emision
                    FROM facturas f
                    WHERE f.periodo=? AND f.estatus IN ({placeholders})
                    ORDER BY f.estatus, f.id_contrato""",
                [per_sel] + estatus_fil
            ).fetchall()

            if rows:
                df_hist = pd.DataFrame(rows, columns=[
                    'UUID', 'Folio', 'Contrato', 'Tipo', 'Mes',
                    'Subtotal', 'Total', 'Estatus', 'Concepto', 'Observación', 'Fecha'
                ])
                df_hist['UUID'] = df_hist['UUID'].str[:8] + '…'
                st.dataframe(df_hist, width='stretch', height=500, key="df_003")

                all_rows = conn.execute(
                    "SELECT estatus, COUNT(*) as n FROM facturas WHERE periodo=? GROUP BY estatus",
                    (per_sel,)
                ).fetchall()
                kk = {r[0]: r[1] for r in all_rows}
                c1, c2, c3, c4, c5, c6 = st.columns(6)
                c1.metric("Conciliados",  kk.get('CONCILIADO', 0))
                c2.metric("Discrepancia", kk.get('DISCREPANCIA', 0))
                c3.metric("Sin contrato", kk.get('SIN_CONTRATO', 0))
                c4.metric("Canceladas",   kk.get('CANCELADA', 0))
                c5.metric("Error",        kk.get('ERROR', 0))
                c6.metric("No aplica",    kk.get('NO_APLICA', 0))
            else:
                st.info("No hay facturas con esos filtros para ese período.")

        st.divider()
        st.subheader("Historial de Corridas")
        st.caption("Registro de lotes de facturación procesados.")
        corridas = conn.execute(
            """SELECT fecha_corrida, periodo, total_facturas, conciliadas, discrepancias,
                      sin_contrato, errores, no_aplica, rfc_incorrecto, omitidas, rfc_configurado
               FROM conciliacion_corridas ORDER BY fecha_corrida DESC LIMIT 200"""
        ).fetchall()
        if not corridas:
            st.info("Aún no hay corridas registradas.")
        else:
            df_corridas = pd.DataFrame(corridas, columns=[
                'Fecha', 'Período', 'Total', 'Conciliadas', 'Discrepancias',
                'Sin contrato', 'Errores', 'No aplica', 'RFC incorrecto', 'Omitidas', 'RFC configurado'
            ])
            st.dataframe(df_corridas, width='stretch', height=320, key="df_corridas")

    # ===================================================================
    # TAB COMPARATIVO DE CIFRAS (DISCREPANCIAS NUMÉRICAS)
    # ===================================================================
    elif tab_activa == "Comparativo de Cifras":
        st.subheader("Comparativo de Cifras y Análisis de Discrepancias")
        st.caption("Facturas donde el monto facturado no coincide con lo programado en el contrato.")
        conn = get_db()

        query_comp = """
            SELECT 
                f.uuid, 
                f.folio, 
                f.id_contrato, 
                c.Cliente, 
                f.periodo, 
                f.mes_contrato, 
                f.tipo, 
                f.esperado, 
                f.facturado, 
                f.diferencia,
                (SELECT GROUP_CONCAT(fc.descripcion, ' | ') FROM factura_conceptos fc WHERE fc.uuid = f.uuid) as concepto,
                f.observaciones,
                f.fecha_emision,
                f.estatus
            FROM facturas f
            LEFT JOIN contratos c ON c.ID_Contrato = TRIM(SUBSTR(f.id_contrato, 1, INSTR(f.id_contrato || ',', ',') - 1))
            WHERE f.estatus IN ('CONCILIADO', 'DISCREPANCIA') AND (f.cancelada IS NULL OR f.cancelada = 0)
            ORDER BY f.periodo DESC, ABS(f.diferencia) DESC
        """
        rows_comp = conn.execute(query_comp).fetchall()

        if not rows_comp:
            st.success("No hay facturas procesadas registradas en el sistema.")
        else:
            df_comp = pd.DataFrame(rows_comp, columns=[
                'UUID', 'Folio', 'Contrato', 'Cliente', 'Período', 'Mes', 'Tipo',
                'Esperado', 'Facturado', 'Diferencia', 'Concepto', 'Observación', 'Fecha Emisión', 'Estatus'
            ])

            df_comp['Cliente'] = df_comp['Cliente'].fillna('Sin Contrato / Desconocido')
            df_comp['Contrato'] = df_comp['Contrato'].fillna('—')
            df_comp['Folio'] = df_comp['Folio'].fillna('—')
            df_comp['Esperado'] = pd.to_numeric(df_comp['Esperado'], errors='coerce').fillna(0.0)
            df_comp['Facturado'] = pd.to_numeric(df_comp['Facturado'], errors='coerce').fillna(0.0)
            df_comp['Diferencia'] = pd.to_numeric(df_comp['Diferencia'], errors='coerce').fillna(0.0)
            df_comp['Dif_Abs'] = df_comp['Diferencia'].abs()
            
            df_comp['Pct_Var'] = np.where(
                df_comp['Esperado'] > 0,
                (df_comp['Diferencia'] / df_comp['Esperado']) * 100,
                0.0
            )

            df_disc = df_comp[df_comp['Estatus'] == 'DISCREPANCIA'].copy()
            df_conc = df_comp[df_comp['Estatus'] == 'CONCILIADO'].copy()

            tot_facturas = len(df_comp)
            n_conciliadas = len(df_conc)
            n_discrepancias = len(df_disc)
            pct_efectividad = (n_conciliadas / tot_facturas * 100) if tot_facturas else 0.0
            tot_facturado = float(df_comp['Facturado'].sum())
            tot_esperado = float(df_comp['Esperado'].sum())
            tot_dif = float(df_disc['Diferencia'].sum())

            cnt_redondeo = int((df_disc['Dif_Abs'] <= 5.0).sum())
            mto_redondeo = float(df_disc.loc[df_disc['Dif_Abs'] <= 5.0, 'Diferencia'].sum())

            cnt_menor = int(((df_disc['Dif_Abs'] > 5.0) & (df_disc['Dif_Abs'] <= 500.0)).sum())
            mto_menor = float(df_disc.loc[(df_disc['Dif_Abs'] > 5.0) & (df_disc['Dif_Abs'] <= 500.0), 'Diferencia'].sum())

            cnt_media = int(((df_disc['Dif_Abs'] > 500.0) & (df_disc['Dif_Abs'] <= 5000.0)).sum())
            mto_media = float(df_disc.loc[(df_disc['Dif_Abs'] > 500.0) & (df_disc['Dif_Abs'] <= 5000.0), 'Diferencia'].sum())

            cnt_mayor = int((df_disc['Dif_Abs'] > 5000.0).sum())
            mto_mayor = float(df_disc.loc[df_disc['Dif_Abs'] > 5000.0, 'Diferencia'].sum())

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Facturas Analizadas", f"{tot_facturas:,}")
            k2.metric("Conciliadas (Correctas)", f"{n_conciliadas:,}", f"{pct_efectividad:.1f}% efectividad")
            k3.metric("Con Discrepancia", f"{n_discrepancias:,}", delta=f"-{100-pct_efectividad:.1f}%", delta_color="inverse")
            k4.metric("Diferencia Neta Total", f"${tot_dif:+,.2f}")

            if n_discrepancias > 0:
                st.markdown("##### Desglose de Discrepancias por Magnitud")
                rg1, rg2, rg3, rg4 = st.columns(4)
                rg1.metric("Redondeo (<= $5.00)", f"{cnt_redondeo} facturas", f"${mto_redondeo:+,.2f}")
                rg2.metric("Menores ($5.01 - $500)", f"{cnt_menor} facturas", f"${mto_menor:+,.2f}")
                rg3.metric("Medias ($500.01 - $5,000)", f"{cnt_media} facturas", f"${mto_media:+,.2f}")
                rg4.metric("Mayores (> $5,000)", f"{cnt_mayor} facturas", f"${mto_mayor:+,.2f}")

                if cnt_redondeo > 0:
                    with st.expander(f"Conciliar {cnt_redondeo} facturas por redondeo (<= $5.00)", expanded=True):
                        st.caption("Diferencias mínimas de centavos o IVA que puedes aprobar en bloque.")
                        if st.button("Aceptar diferencias de redondeo (<= $5.00) como Conciliadas", type="primary", key="btn_conciliar_redondeo"):
                            conn.execute("""
                                UPDATE facturas
                                SET estatus = 'CONCILIADO',
                                    observaciones = 'Conciliado con tolerancia de redondeo (' || printf('$+%.2f', diferencia) || ') | ' || observaciones,
                                    diferencia = 0.0
                                WHERE estatus = 'DISCREPANCIA'
                                  AND (cancelada IS NULL OR cancelada = 0)
                                  AND ABS(diferencia) <= 5.0
                            """)
                            conn.commit()
                            _tocar_datos()
                            st.success(f"Se conciliaron exitosamente {cnt_redondeo} facturas con diferencias de centavos/redondeo.")
                            st.rerun()

            st.divider()

            st.markdown("##### Modalidad de Comparativo")
            sub_vista = st.radio(
                "Selecciona la vista a consultar:",
                [
                    "Vista Consolidada (General)",
                    "Contratos Faltantes por Facturar (Control de Cartera / Sin CFDI)",
                    "Comparativa de Anticipos y Apertura",
                    "Comparativa de Rentas (Mensualidades)"
                ],
                horizontal=True,
                key="radio_subvista_comparativo"
            )

            def _color_diferencia(val):
                if val < 0:
                    return 'background-color:#FAE3E1;color:#8A2019;font-weight:600'
                elif val > 0:
                    return 'background-color:#FBF2D5;color:#785A00;font-weight:600'
                return ''

            def _color_estatus(val):
                if val == 'CONCILIADO':
                    return 'background-color:#EBFBF3;color:#0F5132;font-weight:600'
                elif val == 'DISCREPANCIA':
                    return 'background-color:#FBF0DA;color:#7A5209;font-weight:600'
                return ''

            if sub_vista == "Vista Consolidada (General)":
                f_est, fc1, fc2, fc3 = st.columns([1.5, 1.2, 1.3, 1.8])
                with f_est:
                    filtro_estatus_c = st.selectbox(
                        "Filtrar por estatus:",
                        [
                            "Todos los Contratos (Conciliados y Discrepancias)",
                            "Solo Contratos Correctos (Conciliados)",
                            "Solo Discrepancias"
                        ],
                        key="filtro_comp_estatus"
                    )
                with fc1:
                    periodos_disc = ['TODOS'] + sorted([p for p in df_comp['Período'].dropna().unique().tolist() if p], reverse=True)
                    filtro_per_c = st.selectbox("Filtrar por período:", periodos_disc, key="filtro_disc_periodo")
                with fc2:
                    filtro_mag = st.selectbox(
                        "Filtrar por magnitud:",
                        [
                            "Todas las magnitudes",
                            "Solo redondeo (<= $5.00)",
                            "Diferencias menores ($5.01 - $500)",
                            "Diferencias medias ($500.01 - $5,000)",
                            "Diferencias mayores (> $5,000)"
                        ],
                        key="filtro_disc_magnitud"
                    )
                with fc3:
                    txt_filtro_disc = st.text_input(
                        "Buscar por Folio, Contrato o Cliente:",
                        "", placeholder="Ej. 0718 o HARVEST", key="txt_busq_disc"
                    )

                df_vista = df_comp.copy()
                if filtro_estatus_c == "Solo Contratos Correctos (Conciliados)":
                    df_vista = df_vista[df_vista['Estatus'] == 'CONCILIADO']
                elif filtro_estatus_c == "Solo Discrepancias":
                    df_vista = df_vista[df_vista['Estatus'] == 'DISCREPANCIA']

                if filtro_per_c != 'TODOS':
                    df_vista = df_vista[df_vista['Período'] == filtro_per_c]
                
                if filtro_mag == "Solo redondeo (<= $5.00)":
                    df_vista = df_vista[df_vista['Dif_Abs'] <= 5.0]
                elif filtro_mag == "Diferencias menores ($5.01 - $500)":
                    df_vista = df_vista[(df_vista['Dif_Abs'] > 5.0) & (df_vista['Dif_Abs'] <= 500.0)]
                elif filtro_mag == "Diferencias medias ($500.01 - $5,000)":
                    df_vista = df_vista[(df_vista['Dif_Abs'] > 500.0) & (df_vista['Dif_Abs'] <= 5000.0)]
                elif filtro_mag == "Diferencias mayores (> $5,000)":
                    df_vista = df_vista[df_vista['Dif_Abs'] > 5000.0]

                if txt_filtro_disc.strip():
                    patt = txt_filtro_disc.strip()
                    df_vista = df_vista[
                        df_vista['Folio'].astype(str).str.contains(patt, case=False, na=False) |
                        df_vista['Contrato'].astype(str).str.contains(patt, case=False, na=False) |
                        df_vista['Cliente'].astype(str).str.contains(patt, case=False, na=False)
                    ]

                st.write(f"Mostrando **{len(df_vista):,}** facturas ({filtro_estatus_c}):")

                cols_mostrar = ['Folio', 'Contrato', 'Cliente', 'Período', 'Mes', 'Tipo', 'Estatus', 'Esperado', 'Facturado', 'Diferencia', 'Pct_Var', 'Concepto', 'Observación']
                df_render = df_vista[cols_mostrar].copy()

                st.dataframe(
                    df_render.style
                    .format({
                        'Esperado': '${:,.2f}',
                        'Facturado': '${:,.2f}',
                        'Diferencia': '${:+,.2f}',
                        'Pct_Var': '{:+.1f}%'
                    })
                    .map(_color_diferencia, subset=['Diferencia'])
                    .map(_color_estatus, subset=['Estatus']),
                    width='stretch',
                    height=450,
                    key="df_comparativo_cifras_grid"
                )

                try:
                    buf_excel = excel_con_formato(
                        {
                            'Comparativo_Filtrado': df_render,
                            'Conciliados_Correctos': df_comp[df_comp['Estatus'] == 'CONCILIADO'][cols_mostrar],
                            'Discrepancias': df_comp[df_comp['Estatus'] == 'DISCREPANCIA'][cols_mostrar]
                        },
                        currency_cols=['Esperado', 'Facturado', 'Diferencia']
                    )
                    st.download_button(
                        "Descargar Comparativo de Cifras (Excel con Conciliados y Discrepancias)",
                        data=buf_excel,
                        file_name="comparativo_cifras_completo.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_comp_cifras_excel"
                    )
                except Exception:
                    csv_data = df_render.to_csv(index=False).encode('utf-8-sig')
                    st.download_button(
                        "Descargar Comparativo de Cifras (CSV)",
                        data=csv_data,
                        file_name="comparativo_cifras.csv",
                        mime="text/csv",
                        key="dl_comp_cifras_csv"
                    )

            elif sub_vista == "Contratos Faltantes por Facturar (Control de Cartera / Sin CFDI)":
                st.caption("Auditoría de cartera: contratos activos en vigencia que no cuentan con factura CFDI mensual emitida en el período seleccionado.")
                
                pers_disponibles = sorted([p for p in conn.execute("SELECT DISTINCT periodo FROM facturas WHERE periodo IS NOT NULL").fetchall() if p[0]], reverse=True)
                pers_lista = [p[0] for p in pers_disponibles if str(p[0]).strip()]
                if not pers_lista:
                    pers_lista = [hoy_ref().strftime('%Y-%m')]
                
                c_p1, c_p2 = st.columns([1.5, 2.5])
                with c_p1:
                    per_sel_falt = st.selectbox("Período a revisar:", pers_lista, index=0, key="sel_per_faltante")
                with c_p2:
                    st.caption(f"Auditando qué contratos activos debieron emitir renta en **{per_sel_falt}** pero no tienen CFDI.")
                
                try:
                    p_partes = per_sel_falt.split('-')
                    anio_f, mes_f = int(p_partes[0]), int(p_partes[1])
                except Exception:
                    anio_f, mes_f = hoy_ref().year, hoy_ref().month

                facs_per = conn.execute(
                    "SELECT id_contrato, id_contrato_detectado FROM facturas WHERE periodo=? AND tipo='MENSUAL' AND (cancelada IS NULL OR cancelada=0)",
                    (per_sel_falt,)
                ).fetchall()
                c_facturados_per = set()
                for fp in facs_per:
                    for campo in [fp['id_contrato'], fp['id_contrato_detectado']]:
                        if campo:
                            for c_item in str(campo).split(','):
                                if c_item.strip():
                                    c_facturados_per.add(c_item.strip())

                contratos_act = conn.execute("SELECT * FROM contratos WHERE Estatus='ACTIVO'").fetchall()
                total_en_vigencia = 0
                faltantes_per = []
                
                ML_MAP = {0: "0-Al corriente", 1: "1-Atraso", 2: "2-Convenio", 3: "3-Devuelve no paga", 4: "4-Judicial"}

                for c in contratos_act:
                    c_dict = dict(c)
                    mc = _mes_en_vigencia_contrato(c_dict, anio_f, mes_f, inc_primer_mes=True)
                    if mc is not None:
                        total_en_vigencia += 1
                        idc_c = c_dict['ID_Contrato']
                        if idc_c not in c_facturados_per:
                            nm_val = int(c_dict.get('Nivel_Morosidad') or 0)
                            faltantes_per.append({
                                'Contrato': idc_c,
                                'Cliente': c_dict.get('Cliente', ''),
                                'Vehículo': c_dict.get('Vehiculo', ''),
                                'Mes': f"Mes {mc} de {c_dict.get('Plazo', '')}",
                                'Renta_Esperada': round(float(c_dict.get('Mensualidad_Sin_IVA') or 0.0), 2),
                                'Morosidad': ML_MAP.get(nm_val, f"Nivel {nm_val}"),
                                'Nivel_Num': nm_val,
                                'Excluido_Poliza': 'SI' if c_dict.get('Fecha_Excl_Poliza') else 'NO',
                                'Motivo_Exclusion': c_dict.get('Motivo_Excl_Poliza') or '—',
                            })

                cnt_facturados = len(c_facturados_per)
                cnt_faltantes = len(faltantes_per)
                mto_faltante = sum(f['Renta_Esperada'] for f in faltantes_per)

                kf1, kf2, kf3, kf4 = st.columns(4)
                kf1.metric("Activos en Vigencia", f"{total_en_vigencia:,}")
                kf2.metric("Facturados en Período", f"{cnt_facturados:,}", f"{(cnt_facturados/total_en_vigencia*100) if total_en_vigencia else 0:.1f}%")
                kf3.metric("Faltantes por Facturar", f"{cnt_faltantes:,}")
                kf4.metric("Renta Total Faltante", f"${mto_faltante:,.2f}")

                if not faltantes_per:
                    st.success(f"Excelente: Todos los contratos activos en vigencia ({total_en_vigencia}) tienen factura CFDI emitida en {per_sel_falt}.")
                else:
                    df_falt = pd.DataFrame(faltantes_per)

                    cm1, cm2 = st.columns([1.5, 2.5])
                    with cm1:
                        filtro_mora = st.selectbox(
                            "Filtrar por estatus de cobranza:",
                            ["Todos los faltantes", "Al corriente (Nivel 0)", "Con atraso o convenio (Nivel 1-2)", "Deterioro / Judicial (Nivel 3-4)"],
                            key="filtro_mora_faltante"
                        )
                    with cm2:
                        txt_busq_falt = st.text_input("Buscar faltante por contrato o cliente:", "", key="txt_busq_falt")

                    df_falt_v = df_falt.copy()
                    if filtro_mora == "Al corriente (Nivel 0)":
                        df_falt_v = df_falt_v[df_falt_v['Nivel_Num'] == 0]
                    elif filtro_mora == "Con atraso o convenio (Nivel 1-2)":
                        df_falt_v = df_falt_v[df_falt_v['Nivel_Num'].isin([1, 2])]
                    elif filtro_mora == "Deterioro / Judicial (Nivel 3-4)":
                        df_falt_v = df_falt_v[df_falt_v['Nivel_Num'].isin([3, 4])]

                    if txt_busq_falt.strip():
                        patt_f = txt_busq_falt.strip()
                        df_falt_v = df_falt_v[
                            df_falt_v['Contrato'].str.contains(patt_f, case=False, na=False) |
                            df_falt_v['Cliente'].str.contains(patt_f, case=False, na=False)
                        ]

                    st.write(f"Mostrando **{len(df_falt_v):,}** contratos pendientes de facturar:")
                    cols_falt_ver = ['Contrato', 'Cliente', 'Vehículo', 'Mes', 'Renta_Esperada', 'Morosidad', 'Excluido_Poliza', 'Motivo_Exclusion']
                    st.dataframe(
                        df_falt_v[cols_falt_ver].style.format({'Renta_Esperada': '${:,.2f}'}),
                        width='stretch',
                        height=350,
                        key="grid_contratos_faltantes"
                    )

                    with st.expander("Mitigar diferencia: Reclasificar Morosidad o Enviar a Deterioro", expanded=True):
                        st.caption("Asigna el estatus de cobranza correspondiente para justificar contablemente la falta de facturación y mitigar el riesgo.")
                        lista_opc_falt = [f"{f['Contrato']} — {f['Cliente']} (${f['Renta_Esperada']:,.2f})" for f in faltantes_per]
                        mapa_opc_falt = {f"{f['Contrato']} — {f['Cliente']} (${f['Renta_Esperada']:,.2f})": f['Contrato'] for f in faltantes_per}
                        
                        sel_etiq_falt = st.selectbox("Selecciona contrato a clasificar:", lista_opc_falt, key="sel_falt_mitigar")
                        c_target_falt = mapa_opc_falt.get(sel_etiq_falt)
                        
                        col_acc1, col_acc2 = st.columns(2)
                        with col_acc1:
                            st.markdown("**1. Reclasificar Nivel de Morosidad**")
                            nm_nuevo = st.selectbox(
                                "Nuevo nivel:",
                                [1, 2, 3, 4],
                                format_func=lambda x: {
                                    1: "1 - Atraso (Sin pago primeros días)",
                                    2: "2 - Convenio (En negociación / Reestructura)",
                                    3: "3 - Devuelve no paga (Pase a Deterioro / Excluir)",
                                    4: "4 - Judicial (Demanda legal / Excluir)"
                                }[x],
                                key=f"nm_sel_{c_target_falt}"
                            )
                            motivo_mora = st.text_input("Nota o justificación:", value=f"Falta de facturación en {per_sel_falt}", key=f"mot_mora_{c_target_falt}")
                            if st.button("Actualizar Morosidad", key=f"btn_mora_{c_target_falt}"):
                                hoy_str = hoy_ref().strftime('%Y-%m-%d')
                                if nm_nuevo >= 3:
                                    conn.execute(
                                        "UPDATE contratos SET Nivel_Morosidad=?, Fecha_Excl_Poliza=?, Motivo_Excl_Poliza=? WHERE ID_Contrato=?",
                                        (nm_nuevo, hoy_str, motivo_mora, c_target_falt)
                                    )
                                else:
                                    conn.execute(
                                        "UPDATE contratos SET Nivel_Morosidad=? WHERE ID_Contrato=?",
                                        (nm_nuevo, c_target_falt)
                                    )
                                conn.commit()
                                _tocar_datos()
                                st.success(f"Contrato {c_target_falt} actualizado a Nivel {nm_nuevo}.")
                                st.session_state['_refresh'] = True
                                st.rerun()

                        with col_acc2:
                            st.markdown("**2. Enviar a Deterioro de Cartera**")
                            st.caption("Marca el contrato con Nivel 3, lo excluye de pólizas y registra el evento especial de Deterioro.")
                            nota_det = st.text_input("Observación de deterioro:", value=f"Deterioro por falta de emisión de CFDI en {per_sel_falt}", key=f"nota_det_{c_target_falt}")
                            if st.button("Enviar a Deterioro de Cartera", key=f"btn_det_{c_target_falt}", type="primary"):
                                hoy_str = hoy_ref().strftime('%Y-%m-%d')
                                conn.execute(
                                    "UPDATE contratos SET Nivel_Morosidad=3, Fecha_Excl_Poliza=?, Motivo_Excl_Poliza=? WHERE ID_Contrato=?",
                                    (hoy_str, nota_det, c_target_falt)
                                )
                                conn.execute("""
                                    INSERT INTO eventos_especiales (
                                        ID_Contrato, Tipo_Evento, Fecha_Evento, Fecha_Registro, Observaciones
                                    ) VALUES (?, 'DETERIORO', ?, ?, ?)
                                """, (c_target_falt, f"{per_sel_falt}-01", datetime.now().isoformat(), nota_det))
                                conn.execute("""
                                    INSERT INTO anotaciones (ID_Contrato, Fecha, Tipo, Texto)
                                    VALUES (?, ?, 'DETERIORO', ?)
                                """, (c_target_falt, hoy_str, f"Enviado a deterioro de cartera desde conciliación: {nota_det}"))
                                conn.commit()
                                _tocar_datos()
                                st.success(f"Contrato {c_target_falt} enviado a Deterioro de Cartera y excluido de pólizas.")
                                st.session_state['_refresh'] = True
                                st.rerun()

                    buf_falt = excel_con_formato(
                        {'Faltantes_Por_Facturar': df_falt_v[cols_falt_ver]},
                        currency_cols=['Renta_Esperada']
                    )
                    st.download_button(
                        f"Descargar Faltantes por Facturar ({per_sel_falt}) en Excel",
                        data=buf_falt,
                        file_name=f"faltantes_facturar_{per_sel_falt}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_dl_faltantes_excel"
                    )

            elif sub_vista == "Comparativa de Anticipos y Apertura":
                st.caption("Revisión de pagos iniciales: anticipo y comisión.")
                q_ant = """
                    SELECT 
                        f.uuid, 
                        f.folio, 
                        f.id_contrato, 
                        c.Cliente, 
                        f.periodo, 
                        f.fecha_emision,
                        CASE WHEN INSTR(f.id_contrato, ',') > 0 THEN f.esperado ELSE COALESCE(c.Anticipo_Monto, 0.0) END as ant_esp,
                        CASE WHEN INSTR(f.id_contrato, ',') > 0 THEN 0.0 ELSE COALESCE(c.Comision_Monto, 0.0) END as com_esp,
                        COALESCE((SELECT SUM(fc.importe) FROM factura_conceptos fc WHERE fc.uuid = f.uuid AND fc.clave = 'ANTICIPO'), 0.0) as ant_fac,
                        COALESCE((SELECT SUM(fc.importe) FROM factura_conceptos fc WHERE fc.uuid = f.uuid AND fc.clave = 'COMISION'), 0.0) as com_fac,
                        f.esperado, 
                        f.facturado, 
                        f.diferencia, 
                        f.estatus, 
                        f.observaciones
                    FROM facturas f
                    LEFT JOIN contratos c ON c.ID_Contrato = TRIM(SUBSTR(f.id_contrato, 1, INSTR(f.id_contrato || ',', ',') - 1))
                    WHERE (f.tipo = 'ANTICIPO' OR EXISTS (SELECT 1 FROM factura_conceptos fc WHERE fc.uuid = f.uuid AND fc.clave IN ('ANTICIPO', 'COMISION')))
                      AND (f.cancelada IS NULL OR f.cancelada = 0)
                    ORDER BY f.periodo DESC, ABS(f.diferencia) DESC
                """
                rows_ant = conn.execute(q_ant).fetchall()
                if not rows_ant:
                    st.info("No se encontraron facturas con conceptos de anticipo o comisión.")
                else:
                    df_ant = pd.DataFrame(rows_ant, columns=[
                        'UUID', 'Folio', 'Contrato', 'Cliente', 'Período', 'Fecha',
                        'Anticipo Esp', 'Comisión Esp', 'Anticipo Fac', 'Comisión Fac',
                        'Esperado Total', 'Facturado Total', 'Diferencia Total', 'Estatus', 'Observación'
                    ])
                    df_ant['Cliente'] = df_ant['Cliente'].fillna('Sin Contrato / Desconocido')
                    df_ant['Contrato'] = df_ant['Contrato'].fillna('—')
                    df_ant['Folio'] = df_ant['Folio'].fillna('—')
                    for col_m in ['Anticipo Esp', 'Comisión Esp', 'Anticipo Fac', 'Comisión Fac', 'Esperado Total', 'Facturado Total', 'Diferencia Total']:
                        df_ant[col_m] = pd.to_numeric(df_ant[col_m], errors='coerce').fillna(0.0)

                    df_ant['Apertura Esp'] = df_ant['Anticipo Esp'] + df_ant['Comisión Esp']
                    df_ant['Apertura Fac'] = df_ant['Anticipo Fac'] + df_ant['Comisión Fac']
                    df_ant['Dif Apertura'] = df_ant['Apertura Fac'] - df_ant['Apertura Esp']
                    df_ant['Dif_Abs'] = df_ant['Dif Apertura'].abs()

                    def _calc_diag_ant(r):
                        ae = r['Anticipo Esp']
                        ce = r['Comisión Esp']
                        af = r['Anticipo Fac']
                        cf = r['Comisión Fac']
                        da = r['Dif Apertura']
                        if abs(da) <= 1.0:
                            if af == 0 and cf > 0 and ae > 0:
                                return "Anticipo agrupado en comisión (OK)"
                            elif cf == 0 and af > 0 and ce > 0:
                                return "Comisión agrupada en anticipo (OK)"
                            return "Conciliado OK"
                        else:
                            if af == 0 and cf > 0 and ae > 0:
                                return "Anticipo en comisión (Diferencia)"
                            elif cf == 0 and af > 0 and ce > 0:
                                return "Comisión en anticipo (Diferencia)"
                            return "Diferencia Apertura"

                    df_ant['Diagnóstico'] = df_ant.apply(_calc_diag_ant, axis=1)

                    fa1, fa2, fa3 = st.columns([1.5, 1.5, 2])
                    with fa1:
                        pers_ant = ['TODOS'] + sorted([p for p in df_ant['Período'].dropna().unique().tolist() if p], reverse=True)
                        fil_per_ant = st.selectbox("Filtrar por período:", pers_ant, key="fil_per_ant")
                    with fa2:
                        fil_est_ant = st.selectbox(
                            "Filtrar por resultado:",
                            [
                                "Todas las facturas de apertura",
                                "Solo con Discrepancia en Apertura (Dif != $0)",
                                "Solo Conciliadas / Sin Diferencia",
                                "Anticipo agrupado en comisión"
                            ],
                            key="fil_est_ant"
                        )
                    with fa3:
                        txt_busq_ant = st.text_input(
                            "Buscar por Folio, Contrato o Cliente:",
                            "", placeholder="Ej. 0755 o HARVEST", key="txt_busq_ant"
                        )

                    df_v_ant = df_ant.copy()
                    if fil_per_ant != 'TODOS':
                        df_v_ant = df_v_ant[df_v_ant['Período'] == fil_per_ant]
                    if fil_est_ant == "Solo con Discrepancia en Apertura (Dif != $0)":
                        df_v_ant = df_v_ant[df_v_ant['Dif_Abs'] > 1.0]
                    elif fil_est_ant == "Solo Conciliadas / Sin Diferencia":
                        df_v_ant = df_v_ant[df_v_ant['Dif_Abs'] <= 1.0]
                    elif fil_est_ant == "Anticipo agrupado en comisión":
                        df_v_ant = df_v_ant[df_v_ant['Diagnóstico'].str.contains('Anticipo agrupado|Anticipo en comisión', case=False, na=False)]

                    if txt_busq_ant.strip():
                        patt = txt_busq_ant.strip()
                        df_v_ant = df_v_ant[
                            df_v_ant['Folio'].astype(str).str.contains(patt, case=False, na=False) |
                            df_v_ant['Contrato'].astype(str).str.contains(patt, case=False, na=False) |
                            df_v_ant['Cliente'].astype(str).str.contains(patt, case=False, na=False)
                        ]

                    st.write(f"Mostrando **{len(df_v_ant):,}** facturas de apertura:")

                    cols_ant_render = [
                        'Folio', 'Contrato', 'Cliente', 'Período',
                        'Anticipo Esp', 'Anticipo Fac',
                        'Comisión Esp', 'Comisión Fac',
                        'Apertura Esp', 'Apertura Fac',
                        'Dif Apertura', 'Diagnóstico', 'Observación'
                    ]
                    df_ant_render = df_v_ant[cols_ant_render].copy()

                    st.dataframe(
                        df_ant_render.style
                        .format({
                            'Anticipo Esp': '${:,.2f}',
                            'Anticipo Fac': '${:,.2f}',
                            'Comisión Esp': '${:,.2f}',
                            'Comisión Fac': '${:,.2f}',
                            'Apertura Esp': '${:,.2f}',
                            'Apertura Fac': '${:,.2f}',
                            'Dif Apertura': '${:+,.2f}'
                        })
                        .map(_color_diferencia, subset=['Dif Apertura']),
                        width='stretch',
                        height=450,
                        key="df_comparativo_anticipos_grid"
                    )

                    try:
                        buf_excel_ant = excel_con_formato(
                            {'Apertura': df_ant_render},
                            currency_cols=['Anticipo Esp', 'Anticipo Fac', 'Comisión Esp', 'Comisión Fac', 'Apertura Esp', 'Apertura Fac', 'Dif Apertura']
                        )
                        st.download_button(
                            "Descargar Comparativa de Anticipos (Excel)",
                            data=buf_excel_ant,
                            file_name="comparativo_anticipos_apertura.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_comp_ant_excel"
                        )
                    except Exception:
                        csv_data_ant = df_ant_render.to_csv(index=False).encode('utf-8-sig')
                        st.download_button(
                            "Descargar Comparativa de Anticipos (CSV)",
                            data=csv_data_ant,
                            file_name="comparativo_anticipos_apertura.csv",
                            mime="text/csv",
                            key="dl_comp_ant_csv"
                        )

            elif sub_vista == "Comparativa de Rentas (Mensualidades)":
                st.caption("Revisión de rentas mensuales (renta, administración y GPS).")
                q_rent = """
                    SELECT 
                        f.uuid, 
                        f.folio, 
                        f.id_contrato, 
                        c.Cliente, 
                        f.periodo, 
                        f.mes_contrato, 
                        f.fecha_emision,
                        CASE WHEN INSTR(f.id_contrato, ',') > 0 OR INSTR(f.mes_contrato, ',') > 0 THEN f.esperado ELSE COALESCE(c.Mensualidad_Sin_IVA, 0.0) END as renta_esp,
                        COALESCE((SELECT SUM(fc.importe) FROM factura_conceptos fc WHERE fc.uuid = f.uuid AND fc.clave = 'RENTA'), 0.0) as renta_fac,
                        COALESCE((SELECT SUM(fc.importe) FROM factura_conceptos fc WHERE fc.uuid = f.uuid AND fc.clave = 'ADMIN'), 0.0) as admin_fac,
                        COALESCE((SELECT SUM(fc.importe) FROM factura_conceptos fc WHERE fc.uuid = f.uuid AND fc.clave = 'GEOLOC'), 0.0) as geoloc_fac,
                        f.esperado, 
                        f.facturado, 
                        f.diferencia, 
                        f.estatus, 
                        f.observaciones
                    FROM facturas f
                    LEFT JOIN contratos c ON c.ID_Contrato = TRIM(SUBSTR(f.id_contrato, 1, INSTR(f.id_contrato || ',', ',') - 1))
                    WHERE (f.tipo = 'MENSUAL' OR EXISTS (SELECT 1 FROM factura_conceptos fc WHERE fc.uuid = f.uuid AND fc.clave IN ('RENTA', 'ADMIN', 'GEOLOC')))
                      AND (f.cancelada IS NULL OR f.cancelada = 0)
                    ORDER BY f.periodo DESC, ABS(f.diferencia) DESC
                """
                rows_rent = conn.execute(q_rent).fetchall()
                if not rows_rent:
                    st.info("No se encontraron facturas con conceptos de renta.")
                else:
                    df_rent = pd.DataFrame(rows_rent, columns=[
                        'UUID', 'Folio', 'Contrato', 'Cliente', 'Período', 'Mes', 'Fecha',
                        'Renta Esp', 'Renta Fac', 'Admin Fac', 'GPS Fac',
                        'Esperado Total', 'Facturado Total', 'Diferencia Total', 'Estatus', 'Observación'
                    ])
                    df_rent['Cliente'] = df_rent['Cliente'].fillna('Sin Contrato / Desconocido')
                    df_rent['Contrato'] = df_rent['Contrato'].fillna('—')
                    df_rent['Folio'] = df_rent['Folio'].fillna('—')
                    df_rent['Mes'] = df_rent['Mes'].fillna('—')

                    for col_m in ['Renta Esp', 'Renta Fac', 'Admin Fac', 'GPS Fac', 'Esperado Total', 'Facturado Total', 'Diferencia Total']:
                        df_rent[col_m] = pd.to_numeric(df_rent[col_m], errors='coerce').fillna(0.0)

                    df_rent['Total Renta Fac'] = df_rent['Renta Fac'] + df_rent['Admin Fac'] + df_rent['GPS Fac']
                    df_rent['Total Renta Fac'] = df_rent.apply(
                        lambda r: r['Facturado Total'] if (r['Estatus'] == 'CONCILIADO' and abs(r['Facturado Total'] - r['Renta Esp']) <= 1.0) else r['Total Renta Fac'],
                        axis=1
                    )
                    df_rent['Dif Renta'] = df_rent['Total Renta Fac'] - df_rent['Renta Esp']
                    df_rent['Dif_Abs'] = df_rent['Dif Renta'].abs()

                    def _calc_diag_rent(r):
                        dif = r['Dif Renta']
                        obs = str(r.get('Observación', '') or '')
                        if abs(dif) <= 1.0:
                            if 'pago acumulado' in obs.lower() or 'acumulado de' in obs.lower():
                                return "Pago acumulado multimes (OK)"
                            elif 'Desfase operativo de 1 mes' in obs:
                                return "Desfase operativo 1 mes (OK)"
                            elif 'Pago adelantado' in obs:
                                return "Pago adelantado (OK)"
                            return "Conciliado OK"
                        else:
                            if 'pago acumulado' in obs.lower() or 'acumulado de' in obs.lower():
                                return "Pago acumulado multimes (Dif monto)"
                            elif 'Desfase' in obs:
                                return "Desfase calendario + Dif monto"
                            elif 'Pago adelantado' in obs:
                                return "Pago adelantado + Dif monto"
                            return "Diferencia de monto"

                    df_rent['Diagnóstico'] = df_rent.apply(_calc_diag_rent, axis=1)

                    fr1, fr2, fr3 = st.columns([1.5, 1.5, 2])
                    with fr1:
                        pers_rent = ['TODOS'] + sorted([p for p in df_rent['Período'].dropna().unique().tolist() if p], reverse=True)
                        fil_per_rent = st.selectbox("Filtrar por período:", pers_rent, key="fil_per_rent")
                    with fr2:
                        fil_est_rent = st.selectbox(
                            "Filtrar por resultado:",
                            [
                                "Todas las rentas",
                                "Solo con Discrepancia en Renta (Dif != $0)",
                                "Solo Conciliadas / Sin Diferencia",
                                "Con desfase operativo o pago adelantado"
                            ],
                            key="fil_est_rent"
                        )
                    with fr3:
                        txt_busq_rent = st.text_input(
                            "Buscar por Folio, Contrato o Cliente:",
                            "", placeholder="Ej. 0648 o HARVEST", key="txt_busq_rent"
                        )

                    df_v_rent = df_rent.copy()
                    if fil_per_rent != 'TODOS':
                        df_v_rent = df_v_rent[df_v_rent['Período'] == fil_per_rent]
                    if fil_est_rent == "Solo con Discrepancia en Renta (Dif != $0)":
                        df_v_rent = df_v_rent[df_v_rent['Dif_Abs'] > 1.0]
                    elif fil_est_rent == "Solo Conciliadas / Sin Diferencia":
                        df_v_rent = df_v_rent[df_v_rent['Dif_Abs'] <= 1.0]
                    elif fil_est_rent == "Con desfase operativo o pago adelantado":
                        df_v_rent = df_v_rent[df_v_rent['Diagnóstico'].str.contains('Desfase|adelantado', case=False, na=False)]

                    if txt_busq_rent.strip():
                        patt = txt_busq_rent.strip()
                        df_v_rent = df_v_rent[
                            df_v_rent['Folio'].astype(str).str.contains(patt, case=False, na=False) |
                            df_v_rent['Contrato'].astype(str).str.contains(patt, case=False, na=False) |
                            df_v_rent['Cliente'].astype(str).str.contains(patt, case=False, na=False)
                        ]

                    st.write(f"Mostrando **{len(df_v_rent):,}** facturas de renta:")

                    cols_rent_render = [
                        'Folio', 'Contrato', 'Cliente', 'Período', 'Mes',
                        'Renta Esp', 'Renta Fac', 'Admin Fac', 'GPS Fac',
                        'Total Renta Fac', 'Dif Renta', 'Diagnóstico', 'Observación'
                    ]
                    df_rent_render = df_v_rent[cols_rent_render].copy()

                    st.dataframe(
                        df_rent_render.style
                        .format({
                            'Renta Esp': '${:,.2f}',
                            'Renta Fac': '${:,.2f}',
                            'Admin Fac': '${:,.2f}',
                            'GPS Fac': '${:,.2f}',
                            'Total Renta Fac': '${:,.2f}',
                            'Dif Renta': '${:+,.2f}'
                        })
                        .map(_color_diferencia, subset=['Dif Renta']),
                        width='stretch',
                        height=450,
                        key="df_comparativo_rentas_grid"
                    )

                    try:
                        buf_excel_rent = excel_con_formato(
                            {'Rentas': df_rent_render},
                            currency_cols=['Renta Esp', 'Renta Fac', 'Admin Fac', 'GPS Fac', 'Total Renta Fac', 'Dif Renta']
                        )
                        st.download_button(
                            "Descargar Comparativa de Rentas (Excel)",
                            data=buf_excel_rent,
                            file_name="comparativo_rentas.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_comp_rent_excel"
                        )
                    except Exception:
                        csv_data_rent = df_rent_render.to_csv(index=False).encode('utf-8-sig')
                        st.download_button(
                            "Descargar Comparativa de Rentas (CSV)",
                            data=csv_data_rent,
                            file_name="comparativo_rentas.csv",
                            mime="text/csv",
                            key="dl_comp_rent_csv"
                        )

    # ===================================================================
    # TAB 3 – CONCEPTOS SIN RECONOCER
    # ===================================================================
    elif tab_activa == "Conceptos sin Reconocer":
        st.subheader("Conceptos sin Reconocer")
        st.caption("Conceptos no identificados automáticamente. Puedes asignarlos a Renta, Anticipo o dejarlos fuera.")
        conn = get_db()
        periodos_c = conn.execute(
            "SELECT DISTINCT periodo FROM facturas WHERE periodo IS NOT NULL ORDER BY periodo DESC"
        ).fetchall()
        lista_per_c = [r[0] for r in periodos_c]
        if not lista_per_c:
            st.info("Primero procesa un lote de XMLs en la pestaña de Carga.")
        else:
            per_c = st.selectbox("Período", lista_per_c, key='per_conceptos')
            filas_sin_reconocer = conn.execute(
                """SELECT fcn.id, fcn.uuid, fcn.descripcion, fcn.importe, fcn.manual,
                          f.folio, f.id_contrato, f.estatus
                   FROM factura_conceptos fcn
                   JOIN facturas f ON f.uuid = fcn.uuid
                   WHERE f.periodo=? AND fcn.clave='OTRO'
                     AND f.tipo IN ('MENSUAL','ANTICIPO')
                     AND (f.cancelada IS NULL OR f.cancelada=0)
                   ORDER BY f.id_contrato, fcn.id""",
                (per_c,)
            ).fetchall()
            if not filas_sin_reconocer:
                st.success(f"No hay conceptos sin reconocer en {per_c} — todos los conceptos de ese período coinciden con las palabras configuradas.")
            else:
                st.warning(f"**{len(filas_sin_reconocer)} concepto(s)** sin reconocer en {per_c}. Ninguno de estos importes se está tomando en cuenta en la conciliación.")
                for fila in filas_sin_reconocer:
                    with st.container(key=f"concepto_card_{fila['id']}"):
                        cc1, cc2 = st.columns([3, 1])
                        with cc1:
                            st.markdown(
                                f"**Contrato {fila['id_contrato'] or '—'}** · Folio {fila['folio'] or '—'} "
                                f"· Estatus factura: {fila['estatus']}"
                                + (" · <span style='color:#96660C;'>reclasificado a mano</span>" if fila['manual'] else "")
                            , unsafe_allow_html=True)
                            st.markdown(f"Texto del concepto: *\u201c{fila['descripcion'] or '(sin descripción)'}\u201d*")
                            st.markdown(f"Importe: **${fila['importe']:,.2f}**")
                        with cc2:
                            nueva_clave = st.selectbox(
                                "Reclasificar como", ["RENTA", "ADMIN", "GEOLOC", "ANTICIPO", "COMISION"],
                                key=f"reclas_{fila['id']}"
                            )
                            if st.button("Aplicar y reconciliar", key=f"aplicar_{fila['id']}"):
                                reclasificar_concepto(fila['id'], nueva_clave)
                                res2 = re_conciliar_factura(fila['uuid'])
                                if res2:
                                    st.success(f"Reclasificado como {nueva_clave}. Nuevo estatus de la factura: {res2['status']}")
                                else:
                                    st.warning("Reclasificado, pero no se encontró la factura para re-conciliarla.")
                                st.session_state['_refresh'] = True
                            if fila['manual']:
                                if st.button("Desmarcar (volver a excluir)", key=f"desmarcar_{fila['id']}"):
                                    desmarcar_concepto(fila['id'])
                                    res2 = re_conciliar_factura(fila['uuid'])
                                    st.info("Se deshizo la reclasificación manual; el concepto vuelve a quedar excluido.")
                                    st.session_state['_refresh'] = True
                        st.divider()

    # ===================================================================
    # TAB 4 – AVANCE DE PAGO POR CONTRATO
    # ===================================================================
    elif tab_activa == "Avance de Pago":
        st.subheader("Avance de Pago por Contrato")
        st.caption("Estatus de mensualidades pagadas y pendientes por cada contrato activo.")
        df_avance_base = obtener('ACTIVO')
        if df_avance_base.empty:
            st.info("No hay contratos activos.")
        else:
            with st.spinner("Calculando avance de pago de la cartera..."):
                conn = get_db()
                todas_fact = conn.execute(
                    """SELECT id_contrato, mes_contrato FROM facturas
                       WHERE tipo='MENSUAL' AND estatus IN ('CONCILIADO','DISCREPANCIA')
                         AND (cancelada IS NULL OR cancelada=0) AND mes_contrato IS NOT NULL"""
                ).fetchall()
                facturas_dict = {}
                for r in todas_fact:
                    id_raw = r['id_contrato'] or ''
                    m_raw = str(r['mes_contrato'] or '')
                    meses_in_row = []
                    for m_chunk in m_raw.split(','):
                        m_chunk = m_chunk.strip().replace('.0', '')
                        if m_chunk.isdigit():
                            meses_in_row.append(int(m_chunk))
                    for c_part in [p.strip() for p in id_raw.split(',') if p.strip()]:
                        for m_val in meses_in_row:
                            facturas_dict.setdefault(c_part, set()).add(m_val)
                
                filas_avance = []
                for _con_av in df_avance_base.to_dict('records'):
                    _avp = calcular_avance_pago(_con_av, facturas_dict)
                    _renta = float(_con_av.get('Mensualidad_Sin_IVA') or 0)
                    _monto_atraso = round(_renta * _avp['meses_atraso'], 2)
                    filas_avance.append({
                        'ID_Contrato': _con_av['ID_Contrato'], 'Cliente': _con_av['Cliente'],
                        'Estado': _avp['estado'],
                        'Renta mensual': _renta,
                        'Monto atraso': _monto_atraso,
                        'Mes esperado hoy': _avp['mes_esperado_hoy'],
                        'Mes máx. facturado': _avp['mes_max_facturado'],
                        'Meses adelanto': _avp['meses_adelanto'],
                        'Meses atraso': _avp['meses_atraso'],
                        'Meses faltantes': ', '.join(str(m) for m in _avp['meses_faltantes']) if _avp['meses_faltantes'] else '',
                    })
            df_avance = pd.DataFrame(filas_avance)
            n_atr = int((df_avance['Estado']=='ATRASADO').sum())
            n_ade = int((df_avance['Estado']=='ADELANTADO').sum())
            n_cor = int((df_avance['Estado']=='AL_CORRIENTE').sum())
            ka1,ka2,ka3 = st.columns(3)
            _monto_atr_tot = float(df_avance.loc[df_avance['Estado']=='ATRASADO', 'Monto atraso'].sum()) if 'Monto atraso' in df_avance.columns else 0.0
            ka1.metric("Atrasados", n_atr)
            ka2.metric("Al corriente", n_cor)
            ka3.metric("Adelantados", n_ade)
            ka1b, ka2b = st.columns(2)
            ka1b.metric("Monto estimado en atraso", f"${_monto_atr_tot:,.2f}")
            ka2b.metric("Contratos en atraso", f"{n_atr}")

            filtro_estado = st.multiselect(
                "Filtrar por estado", ['ATRASADO','AL_CORRIENTE','ADELANTADO'],
                default=['ATRASADO','ADELANTADO'], key="filtro_avance_pago"
            )
            df_avance_mostrar = df_avance[df_avance['Estado'].isin(filtro_estado)] if filtro_estado else df_avance
            df_avance_mostrar = df_avance_mostrar.sort_values(['Estado','Meses atraso'], ascending=[True, False])

            def _color_estado_avance(val):
                if val == 'ATRASADO': return 'background-color:#FAE3E1;color:#8A2019;font-weight:700'
                if val == 'ADELANTADO': return 'background-color:#E1F2E7;color:#155C3B;font-weight:700'
                return ''
            st.dataframe(
                df_avance_mostrar.style.map(_color_estado_avance, subset=['Estado']),
                width='stretch', height=380, key="df_avance_pago"
            )

        st.divider()
        st.subheader("Contratos Activos sin Factura (por período)")
        conn = get_db()
        periodos_f = conn.execute(
            "SELECT DISTINCT periodo FROM facturas WHERE periodo IS NOT NULL ORDER BY periodo DESC"
        ).fetchall()
        lista_per_f = [r[0] for r in periodos_f]

        if not lista_per_f:
            st.info("Primero procesa un lote de XMLs en la pestaña de Carga.")
        else:
            per_f = st.selectbox("Período a verificar", lista_per_f, key='per_falt')
            df_act = obtener('ACTIVO')

            if df_act.empty:
                st.info("No hay contratos activos.")
            else:
                rows_fact = conn.execute(
                    """SELECT DISTINCT id_contrato FROM facturas
                       WHERE periodo=? AND estatus IN ('CONCILIADO','DISCREPANCIA','PENDIENTE')
                       AND (cancelada IS NULL OR cancelada=0)""",
                    (per_f,)
                ).fetchall()
                ids_facturados = {r[0] for r in rows_fact}
                df_falt = df_act[~df_act['ID_Contrato'].isin(ids_facturados)].copy()

                if df_falt.empty:
                    st.success(f"Todos los contratos activos tienen factura registrada para {per_f}.")
                else:
                    st.warning(
                        f"**{len(df_falt)} contrato(s)** activos NO tienen factura "
                        f"registrada para **{per_f}**"
                    )
                    cols_show = [c for c in ['ID_Contrato', 'Cliente', 'Vehiculo',
                                             'Mensualidad_Sin_IVA', 'Fecha_Alta'] if c in df_falt.columns]
                    fmt_falt = {'Mensualidad_Sin_IVA': '${:,.2f}'} if 'Mensualidad_Sin_IVA' in df_falt.columns else {}
                    st.dataframe(
                        df_falt[cols_show].style.format(fmt_falt),
                        width='stretch',
    key="df_004")
                    monto_pendiente = df_falt['Mensualidad_Sin_IVA'].sum() if 'Mensualidad_Sin_IVA' in df_falt.columns else 0
                    st.metric("Renta mensual no facturada", f"${monto_pendiente:,.2f}")

                    # Obtener facturas SIN_CONTRATO del mismo período para emparejar
                    facturas_sin = conn.execute(
                        "SELECT uuid, folio, total, subtotal, mes_contrato, id_contrato, id_contrato_detectado "
                        "FROM facturas WHERE periodo=? AND estatus='SIN_CONTRATO' AND (cancelada IS NULL OR cancelada = 0)",
                        (per_f,)
                    ).fetchall()
                    if facturas_sin:
                        st.subheader("Emparejamiento automático sugerido")
                        st.caption("Facturas sin contrato que podrían corresponder a contratos sin factura (basado en mes y monto). "
                                    "Cada contrato solo se sugiere una vez por lote; si no es el correcto, puedes elegir otro manualmente.")

                        # Todos los contratos (no solo ACTIVO/sin-factura) para el selector manual de respaldo
                        df_todos = obtener()
                        df_todos_sorted = df_todos.sort_values('ID_Contrato') if not df_todos.empty else df_todos
                        etiquetas_todas = (
                            (df_todos_sorted['ID_Contrato'] + ' — ' + df_todos_sorted['Cliente'].fillna('') +
                             ' (' + df_todos_sorted['Estatus'].fillna('') + ')').tolist()
                            if not df_todos_sorted.empty else []
                        )
                        mapa_etiqueta_id_todas = dict(zip(etiquetas_todas, df_todos_sorted['ID_Contrato'])) if not df_todos_sorted.empty else {}

                        contratos_ya_sugeridos = set()  # evita sugerir el mismo contrato dos veces en este lote
                        for fact in facturas_sin:
                            mes_f = fact['mes_contrato']
                            candidato = None
                            if mes_f:
                                for idx, c in df_falt.iterrows():
                                    if c['ID_Contrato'] in contratos_ya_sugeridos:
                                        continue  # ya fue sugerido/emparejado con otra factura en este lote
                                    if abs(c['Mensualidad_Sin_IVA'] - fact['subtotal']) < TOLERANCIA:
                                        candidato = c
                                        break

                            if candidato is not None:
                                contratos_ya_sugeridos.add(candidato['ID_Contrato'])
                                st.markdown(f"**Factura {fact['folio']}** (${fact['total']:,.2f}) → **Contrato sugerido {candidato['ID_Contrato']}** (${candidato['Mensualidad_Sin_IVA']:,.2f})")
                                col_auto, col_manual = st.columns(2)
                                with col_auto:
                                    if st.button(f"Asignar sugerido a {candidato['ID_Contrato']}", key=f"auto_assign_{fact['uuid']}"):
                                        _detectado = fact['id_contrato_detectado'] or fact['id_contrato']
                                        res2 = asignar_contrato_manual(fact['uuid'], candidato['ID_Contrato'], _detectado)
                                        if res2:
                                            if _detectado and _detectado != candidato['ID_Contrato']:
                                                st.success(f"Asignado y reconciliado. Estatus: {res2['status']}. "
                                                           f"Se guardó como regla: '{_detectado}' se asignará solo a "
                                                           f"{candidato['ID_Contrato']} de ahora en adelante.")
                                            else:
                                                st.success(f"Asignado y reconciliado. Estatus: {res2['status']}")
                                        else:
                                            st.warning("Asignado pero no se pudo re-conciliar.")
                                        st.session_state['_refresh'] = True
                            else:
                                st.markdown(f"**Factura {fact['folio']}** (${fact['total']:,.2f}) — sin coincidencia automática")
                                col_auto, col_manual = st.columns(2)

                            with col_manual:
                                if etiquetas_todas:
                                    etiqueta_manual = st.selectbox(
                                        "O elige otro contrato (busca por número o cliente)",
                                        etiquetas_todas, key=f"manual_sel_{fact['uuid']}"
                                    )
                                    if st.button("Asignar este contrato", key=f"manual_btn_{fact['uuid']}"):
                                        id_manual = mapa_etiqueta_id_todas[etiqueta_manual]
                                        _detectado = fact['id_contrato_detectado'] or fact['id_contrato']
                                        res2 = asignar_contrato_manual(fact['uuid'], id_manual, _detectado)
                                        if res2:
                                            if _detectado and _detectado != id_manual:
                                                st.success(f"Asignado y reconciliado. Estatus: {res2['status']}. "
                                                           f"Se guardó como regla: '{_detectado}' se asignará solo a "
                                                           f"{id_manual} de ahora en adelante.")
                                            else:
                                                st.success(f"Asignado y reconciliado. Estatus: {res2['status']}")
                                        else:
                                            st.warning("Asignado pero no se pudo re-conciliar.")
                                        st.session_state['_refresh'] = True
                            st.divider()

                    csv_falt = df_falt[cols_show].to_csv(index=False).encode('utf-8-sig')
                    st.download_button(
                        "Exportar contratos sin factura",
                        csv_falt,
                        f"sin_factura_{per_f}.csv",
                        mime='text/csv'
                    )

    # ===================================================================
    # TAB 5 – RESOLVER PENDIENTES (vista general)
    # ===================================================================
    elif tab_activa == "Resolver Pendientes":
        st.subheader("Resolución de Pendientes y Depuración de Contratos")
        st.caption("Gestiona facturas pendientes, contratos anteriores y reglas de asignación.")
        conn = get_db()

        sub_tab_lote, sub_tab_indiv, sub_tab_alias = st.tabs([
            "Gestor Rápido de Contratos No Vigentes (Por Lote)",
            "Facturas Pendientes (Paginadas)",
            "Reglas Aprendidas (Alias)"
        ])

        # ---------------------------------------------------------------
        # SUBTAB 1: GESTOR RÁPIDO DE CONTRATOS NO VIGENTES
        # ---------------------------------------------------------------
        with sub_tab_lote:
            st.markdown("### Depuración y Archivo de Contratos Previos a Dic 2025")
            st.info("Pertenecen a contratos concluidos antes de diciembre 2025. Puedes archivarlos para conciliar sus facturas y limpiar la lista de pendientes.")

            df_no_vig = obtener_resumen_contratos_no_vigentes(fecha_corte='2025-12')
            
            # Métricas resumen
            m_col1, m_col2, m_col3, m_col4 = st.columns(4)
            c_hist = df_no_vig[df_no_vig['tiene_recientes'] == 0] if not df_no_vig.empty else pd.DataFrame()
            c_rec  = df_no_vig[df_no_vig['tiene_recientes'] == 1] if not df_no_vig.empty else pd.DataFrame()
            
            tot_fac_hist = int(c_hist['cant_facturas'].sum()) if not c_hist.empty else 0
            tot_mto_hist = float(c_hist['total_monto'].sum()) if not c_hist.empty else 0.0
            tot_fac_rec  = int(c_rec['cant_facturas'].sum()) if not c_rec.empty else 0
            
            m_col1.metric("Contratos no registrados", len(df_no_vig))
            m_col2.metric("Facturas < Dic 2025", f"{tot_fac_hist:,}")
            m_col3.metric("Monto histórico", f"${tot_mto_hist:,.2f}")
            m_col4.metric("Facturas en 2026", f"{tot_fac_rec:,}")

            st.divider()

            # Botones de Acción Global
            st.markdown("#### Acciones Rápidas Globales")
            b_col1, b_col2 = st.columns([1, 1])

            with b_col1:
                st.markdown("**1. Archivar y Conciliar (Recomendado)**")
                st.caption("Los registra como concluidos y concilia sus facturas.")
                if st.button("Archivar Contratos Históricos (< Dic 2025) y Conciliar", type="primary", key="btn_archivar_global"):
                    with st.spinner("Procesando contratos históricos..."):
                        c_creados, f_act = archivar_contratos_historicos_lote(fecha_corte='2025-12', marcar_conciliado=True)
                        st.success(f"¡Éxito! Se registraron {c_creados} contratos históricos concluidos y se reconciliaron {f_act} facturas.")
                        st.session_state['_refresh'] = True
                        st.rerun()

            with b_col2:
                st.markdown("**2. Marcar como NO APLICA (Silenciar Alertas)**")
                st.caption("Las marca como no aplicables para quitarlas de pendientes.")
                if st.button("Marcar facturas (< Dic 2025) como NO APLICA", key="btn_noaplica_global"):
                    with st.spinner("Marcando facturas..."):
                        n_marcadas = marcar_facturas_historicas_no_aplica(fecha_corte='2025-12')
                        st.success(f"Se marcaron {n_marcadas} facturas como NO APLICA.")
                        st.session_state['_refresh'] = True
                        st.rerun()

            with st.expander("Zona de Peligro: Purgar / Borrar Facturas de la Base de Datos"):
                st.warning("Atención: Esta acción borra definitivamente las facturas de la base de datos.")
                chk_confirma_borrar = st.checkbox("Confirmo que deseo borrar físicamente estas facturas de la base de datos", key="chk_conf_borrar")
                if st.button("Borrar definitivamente facturas históricas (< Dic 2025)", type="secondary", disabled=not chk_confirma_borrar, key="btn_borrar_global"):
                    with st.spinner("Borrando facturas..."):
                        n_borradas = borrar_facturas_lote(fecha_corte='2025-12')
                        st.success(f"Se eliminaron {n_borradas} facturas históricas de la base de datos.")
                        st.session_state['_refresh'] = True
                        st.rerun()

            st.divider()

            # Tabla Interactiva y Gestión Contrato por Contrato
            st.markdown("#### Detalle de Contratos Detectados No Registrados")
            if df_no_vig.empty:
                st.success("¡Excelente! No hay contratos pendientes de registrar.")
            else:
                filtro_vista = st.radio(
                    "Filtrar lista:",
                    ["Solo anteriores a Dic 2025 (Históricos)", "Con facturas recientes (2026)", "Ver todos"],
                    horizontal=True, key="rad_filtro_vista_contratos"
                )
                
                df_mostrar = df_no_vig.copy()
                if filtro_vista == "Solo anteriores a Dic 2025 (Históricos)":
                    df_mostrar = df_mostrar[df_mostrar['tiene_recientes'] == 0]
                elif filtro_vista == "Con facturas recientes (2026)":
                    df_mostrar = df_mostrar[df_mostrar['tiene_recientes'] == 1]
                    
                txt_busq = st.text_input("Filtrar por contrato o cliente:", "", placeholder="Ej. 0528 o HARVEST", key="txt_busq_c_novig")
                if txt_busq:
                    mask = (
                        df_mostrar['num_contrato'].str.contains(txt_busq, case=False, na=False) |
                        df_mostrar['cliente_nombre'].str.contains(txt_busq, case=False, na=False) |
                        df_mostrar['rfc_receptor'].str.contains(txt_busq, case=False, na=False)
                    )
                    df_mostrar = df_mostrar[mask]

                st.write(f"Mostrando **{len(df_mostrar)}** contratos ({int(df_mostrar['cant_facturas'].sum()):,} facturas)")

                df_tabla_c = df_mostrar[['num_contrato', 'cliente_nombre', 'rfc_receptor', 'cant_facturas', 'min_periodo', 'max_periodo', 'total_monto', 'renta_promedio']].copy()
                df_tabla_c.columns = ['Contrato', 'Cliente SAT', 'RFC', 'Facturas', 'Desde', 'Hasta', 'Total Facturado', 'Renta Promedio']
                st.dataframe(
                    df_tabla_c.style.format({
                        'Total Facturado': '${:,.2f}',
                        'Renta Promedio': '${:,.2f}',
                        'Facturas': '{:,}'
                    }),
                    use_container_width=True,
                    height=280
                )

                st.markdown("##### Acciones sobre un contrato específico")
                sel_num_c = st.selectbox(
                    "Selecciona un contrato para operar:",
                    df_mostrar['num_contrato'].tolist(),
                    format_func=lambda x: f"{x} — {df_mostrar.loc[df_mostrar['num_contrato']==x, 'cliente_nombre'].values[0]} ({df_mostrar.loc[df_mostrar['num_contrato']==x, 'cant_facturas'].values[0]} facturas)",
                    key="sel_c_operar"
                )

                if sel_num_c:
                    row_c = df_mostrar[df_mostrar['num_contrato'] == sel_num_c].iloc[0]
                    act_c1, act_c2, act_c3, act_c4 = st.columns([1.2, 1.8, 1, 1])

                    with act_c1:
                        if st.button(f"Archivar {sel_num_c}", key=f"btn_arch_single_{sel_num_c}"):
                            c_c, f_a = archivar_contratos_historicos_lote(contratos_ids=[sel_num_c], fecha_corte=None, marcar_conciliado=True)
                            st.success(f"Contrato {sel_num_c} archivado ({f_a} facturas conciliadas).")
                            st.session_state['_refresh'] = True
                            st.rerun()

                    with act_c2:
                        df_todos_vig = obtener()
                        if not df_todos_vig.empty:
                            opts_reasis = df_todos_vig.sort_values('ID_Contrato')['ID_Contrato'].tolist()
                            c_target = st.selectbox("Reasignar a existente:", opts_reasis, key=f"sel_reasis_{sel_num_c}")
                            if st.button(f"Reasignar a {c_target}", key=f"btn_reasis_{sel_num_c}"):
                                n_re = reasignar_facturas_contrato_lote(sel_num_c, c_target, guardar_alias=True)
                                st.success(f"Se reasignaron {n_re} facturas a {c_target} y se guardó la regla.")
                                st.session_state['_refresh'] = True
                                st.rerun()

                    with act_c3:
                        if st.button(f"NO APLICA", key=f"btn_noap_single_{sel_num_c}"):
                            n_m = marcar_facturas_historicas_no_aplica(contratos_ids=[sel_num_c])
                            st.success(f"{n_m} facturas marcadas como NO APLICA.")
                            st.session_state['_refresh'] = True
                            st.rerun()

                    with act_c4:
                        if st.button(f"Borrar", key=f"btn_del_single_{sel_num_c}"):
                            n_b = borrar_facturas_lote(contratos_ids=[sel_num_c])
                            st.success(f"{n_b} facturas eliminadas de {sel_num_c}.")
                            st.session_state['_refresh'] = True
                            st.rerun()

        # ---------------------------------------------------------------
        # SUBTAB 2: FACTURAS PENDIENTES (PAGINADAS - RÁPIDO)
        # ---------------------------------------------------------------
        with sub_tab_indiv:
            st.markdown("### Facturas Pendientes de Conciliación")
            
            f_col1, f_col2, f_col3 = st.columns([1.5, 2, 1])
            with f_col1:
                filtro_est = st.selectbox(
                    "Filtrar por estatus:",
                    ['TODOS', 'SIN_CONTRATO', 'DISCREPANCIA', 'FUERA_DE_VIGENCIA', 'RFC_INCORRECTO', 'POSIBLES_DUPLICADAS'],
                    key="indiv_filtro_est"
                )
            with f_col2:
                filtro_txt = st.text_input(
                    "Buscar por Folio, Contrato o UUID:",
                    "", placeholder="Ej. 13657 o 0528", key="indiv_filtro_txt"
                )
            with f_col3:
                tam_pag = st.selectbox("Por página:", [20, 50, 100], index=0, key="indiv_tam_pag")

            where_clauses = ["(cancelada IS NULL OR cancelada=0)"]
            params_count = []
            
            if filtro_est == 'POSIBLES_DUPLICADAS':
                where_clauses.append("(observaciones LIKE '%ya existe otra factura%' OR observaciones LIKE '%revisa que no sea una doble%')")
            elif filtro_est != 'TODOS':
                where_clauses.append("estatus = ?")
                params_count.append(filtro_est)
            else:
                where_clauses.append("estatus IN ('DISCREPANCIA','SIN_CONTRATO','RFC_INCORRECTO','FUERA_DE_VIGENCIA')")
                
            if filtro_txt.strip():
                t_pat = f"%{filtro_txt.strip()}%"
                where_clauses.append("(folio LIKE ? OR id_contrato LIKE ? OR id_contrato_detectado LIKE ? OR uuid LIKE ?)")
                params_count.extend([t_pat, t_pat, t_pat, t_pat])
                
            where_sql = " AND ".join(where_clauses)
            
            total_pendientes = conn.execute(f"SELECT COUNT(*) FROM facturas WHERE {where_sql}", params_count).fetchone()[0]
            
            if total_pendientes == 0:
                st.success("¡No hay facturas pendientes con los filtros seleccionados!")
            else:
                total_paginas = max(1, (total_pendientes + tam_pag - 1) // tam_pag)
                
                p_col1, p_col2, p_col3 = st.columns([1, 2, 1])
                with p_col2:
                    pagina = st.number_input(f"Página (de {total_paginas:,} | Total: {total_pendientes:,} facturas)", min_value=1, max_value=total_paginas, value=1, step=1, key="indiv_pagina_num")
                
                offset = (pagina - 1) * tam_pag
                params_query = list(params_count) + [tam_pag, offset]
                
                filas_pend = conn.execute(f"""
                    SELECT uuid, folio, id_contrato, id_contrato_detectado, tipo, estatus, observaciones, periodo, mes_contrato, 
                           fecha_emision, fecha_registro, rfc_emisor, total, rfc_receptor
                    FROM facturas 
                    WHERE {where_sql}
                    ORDER BY periodo DESC, estatus
                    LIMIT ? OFFSET ?
                """, params_query).fetchall()

                df_todos = obtener()
                mapa_etiqueta_id = {}
                etiquetas_contratos = []
                if not df_todos.empty:
                    df_todos_sorted = df_todos.sort_values('ID_Contrato')
                    etiquetas_contratos = (
                        df_todos_sorted['ID_Contrato'] + ' — ' + df_todos_sorted['Cliente'].fillna('') +
                        ' (' + df_todos_sorted['Estatus'].fillna('') + ')'
                    ).tolist()
                    mapa_etiqueta_id = dict(zip(etiquetas_contratos, df_todos_sorted['ID_Contrato']))

                for p in filas_pend:
                    c_num = p['id_contrato'] or p['id_contrato_detectado'] or '—'
                    with st.expander(f"{p['folio'] or 'Sin Folio'} | {p['periodo']} | {p['tipo']} | {p['estatus']} | Contrato: {c_num} | ${p['total']:,.2f}"):
                        st.write(f"**UUID:** `{p['uuid']}` | **Emisión:** {p['fecha_emision']}")
                        st.write(f"**Observaciones:** {p['observaciones']}")

                        if p['estatus'] == 'RFC_INCORRECTO':
                            st.error(f"RFC emisor detectado: **{p['rfc_emisor'] or '—'}**. No coincide con tu arrendadora.")

                        elif p['estatus'] == 'FUERA_DE_VIGENCIA':
                            st.warning("El contrato ya estaba dado de baja o aún no iniciaba cuando se emitió esta factura.")

                        elif p['estatus'] == 'DISCREPANCIA' and p['tipo'] == 'ANTICIPO' and 'Comisión' in (p['observaciones'] or ''):
                            if st.button(f"Actualizar comisión para contrato {p['id_contrato']}", key=f"res_com_{p['uuid']}"):
                                f_com = conn.execute("SELECT COALESCE(SUM(importe),0) FROM factura_conceptos WHERE uuid=? AND clave='COMISION'", (p['uuid'],)).fetchone()[0]
                                conn.execute("UPDATE contratos SET Comision_Monto=? WHERE ID_Contrato=?", (f_com, p['id_contrato']))
                                conn.commit()
                                re_conciliar_factura(p['uuid'])
                                st.success("Comisión actualizada y reconciliada.")
                                st.session_state['_refresh'] = True
                                st.rerun()

                        elif p['estatus'] == 'DISCREPANCIA' and p['tipo'] == 'MENSUAL' and p['id_contrato']:
                            idc_act = str(p['id_contrato']).split(',')[0].strip()
                            c_info = conn.execute("SELECT * FROM contratos WHERE ID_Contrato=?", (idc_act,)).fetchone()
                            if c_info:
                                c_dict = dict(c_info)
                                es_nuevo = es_contrato_nuevo(c_dict)
                                subtotal_fac = float(conn.execute("SELECT subtotal FROM facturas WHERE uuid=?", (p['uuid'],)).fetchone()[0] or 0.0)
                                renta_con = float(c_dict.get('Mensualidad_Sin_IVA') or 0.0)
                                diag_iva = detectar_descuadre_iva_contrato(renta_con, subtotal_fac)
                                if diag_iva:
                                    if es_nuevo:
                                        if diag_iva['tipo'] == 'CON_IVA':
                                            st.warning(f"Contrato nuevo detectado con captura con IVA: {diag_iva['mensaje']}. {diag_iva['detalle']}")
                                            if st.button(f"Corregir renta a ${diag_iva['renta_correcta']:,.2f} sin IVA y recalcular corrida", key=f"btn_fix_iva_{p['uuid']}"):
                                                ok, msg = corregir_renta_contrato_nuevo(idc_act, diag_iva['renta_correcta'])
                                                if ok:
                                                    st.success(msg)
                                                    st.session_state['_refresh'] = True
                                                    st.rerun()
                                                else:
                                                    st.error(msg)
                                        elif diag_iva['tipo'] == 'DOBLE_DEDUCCION':
                                            st.warning(f"Contrato nuevo detectado con deducción errónea de IVA: {diag_iva['mensaje']}. {diag_iva['detalle']}")
                                            if st.button(f"Corregir renta a ${diag_iva['renta_correcta']:,.2f} sin IVA real y recalcular corrida", key=f"btn_fix_diva_{p['uuid']}"):
                                                ok, msg = corregir_renta_contrato_nuevo(idc_act, diag_iva['renta_correcta'])
                                                if ok:
                                                    st.success(msg)
                                                    st.session_state['_refresh'] = True
                                                    st.rerun()
                                                else:
                                                    st.error(msg)
                                    else:
                                        st.info(f"Contrato histórico {idc_act} (cerrado a agosto 2026): {diag_iva['mensaje']}. Conciliado bajo regla de equivalencia para proteger la contabilidad cerrada.")

                        elif p['estatus'] in ('SIN_CONTRATO', 'DISCREPANCIA'):
                            if etiquetas_contratos:
                                _det = p['id_contrato_detectado'] or p['id_contrato']
                                col_asig, col_btn = st.columns([3, 1])
                                with col_asig:
                                    etiq_sel = st.selectbox(
                                        "Asignar contrato:",
                                        etiquetas_contratos,
                                        key=f"res_asig_{p['uuid']}"
                                    )
                                with col_btn:
                                    st.write("")
                                    st.write("")
                                    c_target_sel = mapa_etiqueta_id.get(etiq_sel)
                                    if st.button(f"Asignar", key=f"btn_asig_{p['uuid']}"):
                                        res_asig = asignar_contrato_manual(p['uuid'], c_target_sel, _det)
                                        st.success(f"Asignado a {c_target_sel}. Estatus: {res_asig['status']}")
                                        st.session_state['_refresh'] = True
                                        st.rerun()

                        st.caption("Opciones adicionales:")
                        c_opt1, c_opt2 = st.columns(2)
                        with c_opt1:
                            if st.button("No considerar (Omitir / Cancelada)", key=f"btn_canc_{p['uuid']}"):
                                marcar_factura_cancelada(p['uuid'], True)
                                st.success("Factura excluida. Ya no se considerará en los cálculos.")
                                st.session_state['_refresh'] = True
                                st.rerun()
                        with c_opt2:
                            if st.button("Marcar como NO APLICA", key=f"btn_noap_{p['uuid']}"):
                                c_m = conn.execute("SELECT GROUP_CONCAT(descripcion, ' | ') FROM factura_conceptos WHERE uuid=?", (p['uuid'],)).fetchone()[0] or 'Marcada manualmente como no aplicable'
                                if len(c_m) > 200: c_m = c_m[:197] + '...'
                                conn.execute("UPDATE facturas SET estatus='NO_APLICA', observaciones=? WHERE uuid=?", (f"No Aplica Leasing | Concepto: {c_m}", p['uuid']))
                                conn.commit()
                                st.success("Marcada como NO APLICA con su concepto.")
                                st.session_state['_refresh'] = True
                                st.rerun()

        # ---------------------------------------------------------------
        # SUBTAB 3: REGLAS APRENDIDAS (ALIAS)
        # ---------------------------------------------------------------
        with sub_tab_alias:
            st.markdown("### Equivalencias de Contratos")
            st.caption("Reglas para asociar números de factura no estándar a sus contratos correspondientes.")
            reglas_alias, historial_alias = obtener_alias_contrato_con_historial()
            if not reglas_alias:
                st.info("No hay reglas activas todavía.")
            else:
                st.markdown(f"**{len(reglas_alias)} regla(s) activa(s)**")
                for regla in reglas_alias:
                    col_r, col_del = st.columns([5, 1])
                    with col_r:
                        usado = f"— usada {regla['veces_aplicado']} vez(es), la última el {regla['fecha_ultimo_uso']}" \
                                if regla['veces_aplicado'] else "— todavía no se ha vuelto a usar"
                        st.write(f"**'{regla['numero_facturado']}'** → **{regla['id_contrato_real']}** "
                                 f"{usado} (creada el {regla['fecha_creacion']})")
                    with col_del:
                        if st.button("Borrar", key=f"del_alias_{regla['numero_facturado']}"):
                            eliminar_alias_contrato(regla['numero_facturado'])
                            st.session_state['_refresh'] = True
                            st.rerun()
                if historial_alias:
                    st.markdown("**Historial de cambios** (últimos 200)")
                    df_hist = pd.DataFrame([dict(h) for h in historial_alias])
                    st.dataframe(
                        df_hist[['fecha', 'numero_facturado', 'id_contrato_real', 'accion', 'detalle']],
                        use_container_width=True, key="df_hist_alias"
                    )

    # ===================================================================
    # TAB 6 – AUDITORÍA Y CERTIFICACIÓN DE FACTURACIÓN POR CONTRATO
    # ===================================================================
    elif tab_activa == "Auditoría por Contrato":
        st.subheader("Auditoría por Contrato")
        st.caption("Expediente de facturación mensual por contrato vs lo programado.")
        
        df_contratos_aud = obtener()
        if df_contratos_aud.empty:
            st.info("No hay contratos registrados.")
        else:
            opts_aud = df_contratos_aud['ID_Contrato'].tolist()
            sel_aud = st.selectbox(
                "Selecciona un contrato para auditar",
                opts_aud,
                format_func=lambda x: f"{x} — {df_contratos_aud.loc[df_contratos_aud['ID_Contrato']==x, 'Cliente'].values[0]} ({df_contratos_aud.loc[df_contratos_aud['ID_Contrato']==x, 'Vehiculo'].values[0]})",
                key="sel_contrato_audit"
            )
            
            row_aud = df_contratos_aud[df_contratos_aud['ID_Contrato'] == sel_aud].iloc[0]
            conn_aud = get_db()
            
            def _get_contrato_match_data_audit(cid, con_row, fact_r, conn):
                uuid = fact_r.get('uuid')
                rows_c = conn.execute("SELECT clave, importe, descripcion FROM factura_conceptos WHERE uuid=?", (uuid,)).fetchall()
                
                cid_parts = str(cid).split('-')
                p_num = cid_parts[0] if len(cid_parts) > 0 else '0'
                p_suf = cid_parts[1] if len(cid_parts) > 1 else '0'
                try:
                    p_num_int, p_suf_int = int(p_num), int(p_suf)
                except Exception:
                    p_num_int, p_suf_int = 0, 0
                es_multi_c = (',' in str(fact_r.get('id_contrato') or ''))
                
                conceptos_este_con = []
                meses_en_conceptos = []
                
                for rc in rows_c:
                    desc = str(rc['descripcion'] or '')
                    clv = str(rc['clave'] or '')
                    
                    if es_multi_c:
                        pertenece = False
                        for m in _PAT_CONTRATO.findall(desc):
                            if normalizar_contrato(m[0], m[1]) == cid:
                                pertenece = True
                                break
                        if not pertenece and p_num_int > 0:
                            patterns_var = [f"{p_num_int:04d}-{p_suf_int}", f"{p_num_int}-{p_suf_int:02d}", f"{p_num_int}-{p_suf_int}"]
                            for pv in patterns_var:
                                if pv in desc:
                                    pertenece = True
                                    break
                        if not pertenece:
                            continue
                    
                    if clv in ('RENTA', 'ADMIN', 'GEOLOC'):
                        conceptos_este_con.append(rc)
                        for m_str, _ in _PAT_MES.findall(desc):
                            try:
                                m_val = int(m_str)
                                if m_val not in meses_en_conceptos:
                                    meses_en_conceptos.append(m_val)
                            except Exception:
                                pass
                                
                meses_en_conceptos = sorted(meses_en_conceptos)
                if meses_en_conceptos:
                    meses_contrato = meses_en_conceptos
                else:
                    if not es_multi_c and fact_r.get('mes_contrato'):
                        meses_contrato = [int(p.strip().replace('.0','')) for p in str(fact_r['mes_contrato']).split(',') if p.strip().replace('.0','').isdigit()]
                    else:
                        meses_contrato = []
                        
                cant_m = len(meses_contrato) or 1
                
                if conceptos_este_con:
                    renta_sin_iva = sum(float(c['importe'] or 0) for c in conceptos_este_con)
                    if fact_r.get('estatus') == 'CONCILIADO':
                        monto_con_iva = round(float(con_row.get('Mensualidad_Sin_IVA') or 0.0) * cant_m * 1.16, 2)
                    else:
                        monto_con_iva = round(renta_sin_iva * 1.16, 2)
                else:
                    if fact_r.get('estatus') == 'CONCILIADO':
                        monto_con_iva = round(float(con_row.get('Mensualidad_Sin_IVA') or 0.0) * cant_m * 1.16, 2)
                    else:
                        if not es_multi_c:
                            monto_sin_iva = float(fact_r.get('facturado') or fact_r.get('subtotal') or 0.0)
                            monto_con_iva = round(monto_sin_iva * 1.16, 2)
                        else:
                            monto_con_iva = 0.0

                return meses_contrato, monto_con_iva, cant_m

            # Consultar facturas del contrato
            facts_con = conn_aud.execute(
                """SELECT uuid, folio, fecha_emision, periodo, mes_contrato, tipo, subtotal, total, estatus, observaciones, cancelada, id_contrato
                   FROM facturas
                   WHERE (id_contrato=? OR instr(id_contrato, ?) > 0 OR instr(coalesce(id_contrato_detectado,''), ?) > 0)
                     AND (cancelada IS NULL OR cancelada=0)
                   ORDER BY CASE WHEN estatus='CONCILIADO' THEN 0 ELSE 1 END, periodo ASC, fecha_emision ASC""",
                (sel_aud, sel_aud, sel_aud)
            ).fetchall()
            
            # Header del Contrato
            estatus_badge_color = "#1C7A4D" if str(row_aud['Estatus']).upper() == 'ACTIVO' else "#B3261E"
            st.markdown(f"""
            <div style="background:#FFFFFF;border:1px solid #DCE0E5;border-left:5px solid {estatus_badge_color};border-radius:4px;padding:12px 16px;margin-bottom:12px;">
                <div style="display:flex;justify-content:space-between;align-items:center;">
                    <div>
                        <span style="font-size:1.1rem;font-weight:700;color:#20242B;">Contrato: {row_aud['ID_Contrato']}</span>
                        <span style="background:{estatus_badge_color};color:#fff;border-radius:4px;padding:2px 8px;font-size:0.75rem;font-weight:700;margin-left:8px;">{row_aud['Estatus']}</span>
                    </div>
                    <div style="font-size:0.85rem;color:#565E68;">Alta: <b>{str(row_aud['Fecha_Alta'])[:10]}</b> | Plazo: <b>{row_aud['Plazo']} meses</b></div>
                </div>
                <div style="font-size:0.9rem;color:#565E68;margin-top:4px;">
                    Cliente: <b>{row_aud['Cliente']}</b> | Vehículo: <b>{row_aud['Vehiculo']}</b>
                </div>
                <div style="font-size:0.85rem;color:#20242B;margin-top:6px;">
                    Renta Mensual: <b>${float(row_aud['Mensualidad_Sin_IVA']):,.2f}</b> | Residual Pactado: <b>${float(row_aud['Residual_Monto']):,.2f}</b>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            # Auditoría Mes a Mes (Plazo)
            plazo_num = int(row_aud['Plazo'])
            renta_esperada = round(float(row_aud['Mensualidad_Sin_IVA']), 2)
            fa_dt = pd.to_datetime(row_aud['Fecha_Alta'])
            
            # Pre-procesar facturas con asignación precisa de conceptos por contrato
            facts_parsed = []
            for f in facts_con:
                f_dict = dict(f)
                meses_c, monto_c_iva, cant_m = _get_contrato_match_data_audit(sel_aud, row_aud, f_dict, conn_aud)
                f_dict['meses_contrato_especificos'] = meses_c
                f_dict['monto_renta_con_iva'] = monto_c_iva
                f_dict['cant_meses_especificos'] = cant_m
                facts_parsed.append(f_dict)

            # Deduplicar en facts_parsed comprobantes que correspondan a la misma factura
            # (mismo folio numérico o UUID, ej. '25236' y 'F-25236'), priorizando CONCILIADO
            facts_dedup = []
            seen_p_fols = {}
            for fp in facts_parsed:
                fn = normalizar_folio(fp.get('folio'))
                u = (fp.get('uuid') or '').strip().upper()
                key = fn if fn else u
                if key in seen_p_fols:
                    prev_idx = seen_p_fols[key]
                    if fp.get('estatus') == 'CONCILIADO' and facts_dedup[prev_idx].get('estatus') != 'CONCILIADO':
                        facts_dedup[prev_idx] = fp
                    continue
                seen_p_fols[key] = len(facts_dedup)
                facts_dedup.append(fp)
            facts_parsed = facts_dedup
            
            df_facts_con = pd.DataFrame(facts_parsed) if facts_parsed else pd.DataFrame()
            
            audit_rows = []
            total_esperado_rentas = 0.0
            total_facturado_rentas = 0.0
            meses_con_multiples = []
            
            for m_num in range(1, plazo_num + 1):
                f_mes_dt = fa_dt + relativedelta(months=m_num-1)
                f_mes_str = f_mes_dt.strftime('%Y-%m')
                
                # Buscar todas las facturas correspondientes a este mes (se suman automáticamente si son facturas distintas)
                facts_match_list = []
                # 1. Coincidencia por mes específico de este contrato
                for f in facts_parsed:
                    if f['tipo'] == 'MENSUAL' and m_num in f['meses_contrato_especificos']:
                        facts_match_list.append(f)
                        
                # 2. Coincidencia por período si no hubo ninguna por mes_contrato específico
                if not facts_match_list:
                    for f in facts_parsed:
                        if f['tipo'] == 'MENSUAL' and f.get('periodo') == f_mes_str:
                            if ',' in str(f.get('id_contrato') or '') and f['monto_renta_con_iva'] == 0:
                                continue
                            facts_match_list.append(f)

                # Deduplicar comprobantes que sean el mismo (mismo folio numérico o UUID).
                # Facturas con folio numérico DISTINTO SÍ se suman (pagos complementarios).
                if facts_match_list:
                    seen_m_fols = {}
                    facts_unicas = []
                    for f_m in facts_match_list:
                        fn = normalizar_folio(f_m.get('folio'))
                        u = (f_m.get('uuid') or '').strip().upper()
                        key = fn if fn else u
                        if key in seen_m_fols:
                            prev_idx = seen_m_fols[key]
                            if f_m.get('estatus') == 'CONCILIADO' and facts_unicas[prev_idx].get('estatus') != 'CONCILIADO':
                                facts_unicas[prev_idx] = f_m
                            continue
                        seen_m_fols[key] = len(facts_unicas)
                        facts_unicas.append(f_m)
                    facts_match_list = facts_unicas
                
                if facts_match_list:
                    if len(facts_match_list) > 1:
                        meses_con_multiples.append((m_num, f_mes_str, facts_match_list))
                    
                    monto_fac = sum(round(float(f['monto_renta_con_iva']) / (f['cant_meses_especificos'] or 1), 2) for f in facts_match_list)
                    monto_fac = round(monto_fac, 2)
                    dif_mes = round(monto_fac - (renta_esperada * 1.16), 2)
                    total_facturado_rentas += monto_fac
                    
                    if abs(dif_mes) <= TOLERANCIA * 2 or all(str(f.get('estatus')) == 'CONCILIADO' for f in facts_match_list):
                        est_audit = "CONCILIADO"
                    else:
                        est_audit = f"DISCREPANCIA (${dif_mes:+,.2f})"
                        
                    folio_str = ', '.join(f['folio'] or f['uuid'][:8] for f in facts_match_list)
                    
                    if len(facts_match_list) > 1:
                        obs_str = f"Suma de {len(facts_match_list)} facturas ({folio_str}): ${monto_fac:,.2f}"
                    elif ',' in str(facts_match_list[0].get('id_contrato') or ''):
                        obs_str = f"Factura multi-contrato ({facts_match_list[0]['id_contrato']})"
                    elif facts_match_list[0]['cant_meses_especificos'] > 1:
                        obs_str = f"Factura acumulada de {facts_match_list[0]['cant_meses_especificos']} meses ({facts_match_list[0]['mes_contrato']})"
                    else:
                        obs_str = facts_match_list[0]['observaciones'] or "OK"
                else:
                    monto_fac = 0.0
                    dif_mes = -round(renta_esperada * 1.16, 2)
                    est_audit = "SIN FACTURA"
                    folio_str = "—"
                    obs_str = "Mes transcurrido sin CFDI registrado" if f_mes_dt <= pd.Timestamp(hoy_ref()) else "Mes futuro"
                
                total_esperado_rentas += (renta_esperada * 1.16)
                audit_rows.append({
                    'Mes': m_num,
                    'Período': f_mes_str,
                    'Renta Esperada (c/IVA)': round(renta_esperada * 1.16, 2),
                    'Facturado (c/IVA)': monto_fac,
                    'Diferencia': dif_mes,
                    'Folio CFDI': folio_str,
                    'Estatus Audit': est_audit,
                    'Observaciones': obs_str
                })
            
            df_audit_table = pd.DataFrame(audit_rows)
            
            kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
            kpi_col1.metric("Facturación Esperada Total", f"${total_esperado_rentas:,.2f}")
            kpi_col2.metric("Facturación Registrada Total", f"${total_facturado_rentas:,.2f}")
            kpi_col3.metric("Diferencia Acumulada", f"${total_facturado_rentas - total_esperado_rentas:+,.2f}")
            pct_cob = (total_facturado_rentas / total_esperado_rentas * 100) if total_esperado_rentas > 0 else 0.0
            kpi_col4.metric("Cumplimiento de Facturación", f"{pct_cob:.1f}%")
            
            st.markdown("---")
            st.markdown("#### Auditoría Mes a Mes del Plazo del Contrato")
            st.dataframe(
                df_audit_table.style.format({
                    'Renta Esperada (c/IVA)': '${:,.2f}',
                    'Facturado (c/IVA)': '${:,.2f}',
                    'Diferencia': '${:+,.2f}'
                }),
                width='stretch',
                height=320,
                key="df_audit_contract_table"
            )
            
            # Consultar facturas excluidas / canceladas para este contrato
            facts_canceladas = conn_aud.execute(
                """SELECT uuid, folio, fecha_emision, periodo, mes_contrato, tipo, subtotal, total, cancelada, fecha_cancelacion
                   FROM facturas
                   WHERE (id_contrato=? OR instr(id_contrato, ?) > 0 OR instr(coalesce(id_contrato_detectado,''), ?) > 0)
                     AND cancelada = 1
                   ORDER BY fecha_emision DESC""",
                (sel_aud, sel_aud, sel_aud)
            ).fetchall()
            
            if meses_con_multiples or facts_canceladas:
                st.markdown("---")
                st.markdown("#### Facturas Múltiples en el Mismo Mes")
                st.caption("Si hay varias facturas en un mes, sus importes se suman. Puedes excluir duplicados si es necesario.")
                
                for m_num_m, p_str_m, f_list_m in meses_con_multiples:
                    m_renta_tot = sum(round(float(f_item['monto_renta_con_iva']) / (f_item['cant_meses_especificos'] or 1), 2) for f_item in f_list_m)
                    with st.expander(f"Mes {m_num_m} ({p_str_m}) — {len(f_list_m)} facturas sumadas (${m_renta_tot:,.2f} c/IVA)", expanded=True):
                        cols_m = st.columns(len(f_list_m))
                        for idx_f, f_item in enumerate(f_list_m):
                            with cols_m[idx_f]:
                                f_folio = f_item['folio'] or f_item['uuid'][:8]
                                f_monto_renta = round(float(f_item['monto_renta_con_iva']) / (f_item['cant_meses_especificos'] or 1), 2)
                                st.markdown(f"""
                                <div style="background:#F9FAFB;border:1px solid #DCE0E5;border-radius:6px;padding:12px;margin-bottom:8px;">
                                    <div style="font-weight:700;font-size:1rem;color:#20242B;">Folio: {f_folio}</div>
                                    <div style="font-size:0.75rem;color:#717882;word-break:break-all;">UUID: {f_item['uuid']}</div>
                                    <div style="font-size:0.85rem;color:#565E68;margin-top:4px;">
                                        Fecha: <b>{str(f_item['fecha_emision'])[:10]}</b> | Total CFDI: <b>${float(f_item['total']):,.2f}</b>
                                    </div>
                                    <div style="font-size:0.9rem;color:#1C7A4D;margin-top:4px;">
                                        Renta asignada (c/IVA): <b>${f_monto_renta:,.2f}</b>
                                    </div>
                                    <div style="font-size:0.8rem;color:#565E68;margin-top:2px;">
                                        Estatus actual: <b>{f_item.get('estatus', '—')}</b>
                                    </div>
                                    <div style="font-size:0.75rem;color:#858D96;margin-top:4px;">
                                        {f_item.get('observaciones') or '—'}
                                    </div>
                                </div>
                                """, unsafe_allow_html=True)
                                
                                if st.button(f"No considerar {f_folio}", key=f"btn_excl_{sel_aud}_{m_num_m}_{f_item['uuid']}", width='stretch'):
                                    conn_aud.execute(
                                        "UPDATE facturas SET cancelada=1, fecha_cancelacion=? WHERE uuid=?",
                                        (hoy_ref().isoformat(), f_item['uuid'])
                                    )
                                    conn_aud.commit()
                                    st.success(f"Factura {f_folio} excluida del cálculo.")
                                    st.session_state['_refresh'] = True
                                    st.rerun()

                if facts_canceladas:
                    st.markdown("##### Facturas Excluidas / Omitidas de este Contrato")
                    for fc in facts_canceladas:
                        fc_folio = fc['folio'] or fc['uuid'][:8]
                        c_f1, c_f2 = st.columns([4, 1])
                        with c_f1:
                            st.write(f"• **Folio {fc_folio}** (Período: {fc['periodo']}) — Total CFDI: ${float(fc['total']):,.2f} — Excluida el {str(fc['fecha_cancelacion'] or '')[:10]}")
                        with c_f2:
                            if st.button("Volver a considerar", key=f"btn_reactivar_{fc['uuid']}"):
                                conn_aud.execute("UPDATE facturas SET cancelada=0, fecha_cancelacion=NULL, estatus='PENDIENTE' WHERE uuid=?", (fc['uuid'],))
                                conn_aud.commit()
                                st.success(f"Factura {fc_folio} reactivada e incluida nuevamente en los cálculos.")
                                st.session_state['_refresh'] = True
                                st.rerun()
            
            # Auditoría Especial (Venta de Vehículo / Residual e Indemnización)
            st.markdown("---")
            st.markdown("#### Eventos Especiales y Cierre de Contrato (Residual / Indemnización)")
            
            fact_venta = df_facts_con[df_facts_con['tipo'] == 'VENTA_VEHICULO'] if not df_facts_con.empty else pd.DataFrame()
            fact_indem = df_facts_con[df_facts_con['tipo'] == 'INDEMNIZACION'] if not df_facts_con.empty else pd.DataFrame()
            
            ac1, ac2 = st.columns(2)
            with ac1:
                st.markdown("**1. Cobro de Residual / Venta de Vehículo**")
                if not fact_venta.empty:
                    fv_row = fact_venta.iloc[0]
                    st.success(f"Factura de Venta de Vehículo detectada: Folio **{fv_row['folio']}** | Monto: **${float(fv_row['total']):,.2f}**")
                    if str(row_aud['Estatus']).upper() != 'BAJA':
                        st.warning("**ALERTA:** Existe factura por Venta de Vehículo pero el contrato figura como **ACTIVO**. Procesa la baja en Gestor de Bajas.")
                    else:
                        st.info(f"Contrato dado de BAJA el {str(row_aud.get('Fecha_Baja'))[:10]}. Residual liquidado.")
                else:
                    st.caption(f"Residual pactado: **${float(row_aud['Residual_Monto']):,.2f}** (Sin IVA). No se ha recibido CFDI de Venta de Vehículo.")
                    
            with ac2:
                st.markdown("**2. Reclamaciones / Indemnización de Seguro**")
                if not fact_indem.empty:
                    fi_row = fact_indem.iloc[0]
                    st.success(f"Factura por Indemnización recibida: Folio **{fi_row['folio']}** | Monto: **${float(fi_row['total']):,.2f}**")
                else:
                    st.caption("Sin facturas por Indemnización de seguro registradas para este contrato.")
            
            # Resumen Ejecutivo y Certificación de Auditoría para Dirección / Ventas / Administración
            st.markdown("---")
            st.markdown("#### Certificación de Estado de Facturación y Cumplimiento")
            
            # Evaluar estatus global
            hay_alertas_venta = (not fact_venta.empty and str(row_aud['Estatus']).upper() != 'BAJA')
            hay_alertas_indem = (not fact_indem.empty and str(row_aud['Estatus']).upper() != 'BAJA')
            meses_atrasados_cont = len(df_audit_table[df_audit_table['Estatus Audit'].str.contains('SIN FACTURA') & (pd.to_datetime(df_audit_table['Período'] + '-01') <= pd.Timestamp(hoy_ref()))])
            
            c_cert1, c_cert2, c_cert3 = st.columns(3)
            with c_cert1:
                st.markdown("**Para Administración / Contabilidad**")
                st.write(f"• Meses transcurridos al corte: **{min(plazo_num, (pd.Timestamp(hoy_ref()).year - fa_dt.year)*12 + pd.Timestamp(hoy_ref()).month - fa_dt.month + 1)} de {plazo_num}**")
                st.write(f"• Meses con facturación omitida: **{meses_atrasados_cont}**")
                st.write(f"• Acumulado Facturado (con IVA): **${total_facturado_rentas:,.2f}**")
            
            with c_cert2:
                st.markdown("**Para Ventas / Atención a Clientes**")
                if pct_cob >= 99.0:
                    st.success("Cliente al corriente en facturación de rentas.")
                elif pct_cob >= 80.0:
                    st.warning(f"Facturación al {pct_cob:.1f}% (Revisar facturas pendientes).")
                else:
                    st.error(f"Facturación retrasada ({pct_cob:.1f}% de avance).")
                st.write(f"• Próxima fecha de corte: **01/{(pd.Timestamp(hoy_ref()) + relativedelta(months=1)).strftime('%m/%Y')}**")

            with c_cert3:
                st.markdown("**Para Dirección General / Auditoría**")
                if hay_alertas_venta:
                    st.error("ALERTA: ALERTA DIRECCIÓN: Cobro de Residual realizado pero contrato figura ACTIVO.")
                elif hay_alertas_indem:
                    st.error("ALERTA: ALERTA DIRECCIÓN: Indemnización de seguro recibida sin evento de baja/siniestro.")
                elif meses_atrasados_cont > 0:
                    st.warning(f"ATENCIÓN: {meses_atrasados_cont} mes(es) pendientes de facturar.")
                else:
                    st.success("EXPEDIENTE 100% CERTIFICADO Y SIN ANOMALÍAS")

            # Descargar reporte de auditoría en Excel
            try:
                buf_aud = excel_con_formato(
                    {
                        'Auditoría_Contrato': df_audit_table,
                        'Resumen_Ejecutivo': pd.DataFrame([
                            {'Indicador': 'ID Contrato', 'Valor': sel_aud},
                            {'Indicador': 'Cliente', 'Valor': row_aud['Cliente']},
                            {'Indicador': 'Vehículo', 'Valor': row_aud['Vehiculo']},
                            {'Indicador': 'Estatus Contrato', 'Valor': row_aud['Estatus']},
                            {'Indicador': 'Plazo', 'Valor': f"{plazo_num} meses"},
                            {'Indicador': 'Renta Mensual (Sin IVA)', 'Valor': f"${float(row_aud['Mensualidad_Sin_IVA']):,.2f}"},
                            {'Indicador': 'Facturación Esperada Total', 'Valor': f"${total_esperado_rentas:,.2f}"},
                            {'Indicador': 'Facturación Registrada Total', 'Valor': f"${total_facturado_rentas:,.2f}"},
                            {'Indicador': 'Diferencia Acumulada', 'Valor': f"${total_facturado_rentas - total_esperado_rentas:+,.2f}"},
                            {'Indicador': 'Porcentaje de Cumplimiento', 'Valor': f"{pct_cob:.1f}%"},
                            {'Indicador': 'Factura Venta Vehículo', 'Valor': f"${float(fact_venta.iloc[0]['total']):,.2f}" if not fact_venta.empty else "No"},
                            {'Indicador': 'Factura Indemnización', 'Valor': f"${float(fact_indem.iloc[0]['total']):,.2f}" if not fact_indem.empty else "No"}
                        ])
                    },
                    currency_cols=['Renta Esperada (c/IVA)', 'Facturado (c/IVA)', 'Diferencia']
                )
                st.download_button(
                    f"Descargar Expediente Auditado Completo de {sel_aud} (Excel)",
                    buf_aud,
                    file_name=f"auditoria_expediente_{sel_aud}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_aud_excel_{sel_aud}"
                )
            except Exception:
                csv_aud = df_audit_table.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    f"Descargar Expediente Auditado Completo de {sel_aud} (CSV)",
                    csv_aud,
                    file_name=f"auditoria_expediente_{sel_aud}.csv",
                    mime="text/csv",
                    key=f"dl_aud_csv_{sel_aud}"
                )

    # ===================================================================
    # TAB 7 – REPORTE CONSOLIDADO
    # ===================================================================
    elif tab_activa == "Reporte Consolidado":
        st.subheader("Reporte Consolidado de Conciliación")
        conn = get_db()
        periodos = conn.execute(
            "SELECT DISTINCT periodo FROM facturas WHERE periodo IS NOT NULL ORDER BY periodo DESC"
        ).fetchall()
        lista_per = [r[0] for r in periodos]

        # Analisis multi-anio (todos los XML cargados, varios anios a la vez)
        with st.expander("Analisis multi-anio de facturacion (todos los XML cargados)", expanded=False):
            st.caption("Resumen anual y por tipo de factura para cuadrar con contabilidad.")
            try:
                _df_all_f = pd.read_sql_query(
                    """SELECT periodo, tipo, total, subtotal, cancelada, estatus
                       FROM facturas WHERE periodo IS NOT NULL""",
                    conn,
                )
            except Exception:
                _df_all_f = pd.DataFrame()
            if _df_all_f.empty:
                st.info("Sin facturas registradas.")
            else:
                _res_anio = resumen_facturacion_por_anio(_df_all_f, solo_vigentes=True)
                if not _res_anio.empty:
                    _anios = sorted(_res_anio["anio"].unique())
                    st.markdown(f"**Años con facturas:** {', '.join(_anios)}  |  **Total vigentes:** ${_res_anio['total'].sum():,.2f}")
                    c_chart, c_tab = st.columns([1, 1])
                    with c_chart:
                        try:
                            fig_an = go.Figure()
                            for col in ["mensual", "comision", "otros"]:
                                if col in _res_anio.columns:
                                    fig_an.add_trace(go.Bar(x=_res_anio["anio"], y=_res_anio[col], name=col.capitalize()))
                            fig_an.update_layout(barmode="stack", title="Facturado vigente por año y tipo", height=320)
                            st.plotly_chart(sfig(fig_an, h=320), width="stretch", key="pc_multi_anio")
                        except Exception:
                            pass
                    with c_tab:
                        st.dataframe(
                            _res_anio.style.format({"total": "${:,.2f}", "subtotal": "${:,.2f}", "mensual": "${:,.2f}", "comision": "${:,.2f}", "otros": "${:,.2f}"}),
                            width="stretch", height=300, key="df_multi_anio",
                        )
                    try:
                        from reports.excel import excel_con_formato as _exc_ma
                        _buf_ma = _exc_ma(
                            {"Por año": _res_anio},
                            currency_cols=["total", "subtotal", "mensual", "comision", "otros"],
                        )
                        st.download_button(
                            "Excel multi-año",
                            data=_buf_ma,
                            file_name="facturacion_multi_anio.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_multi_anio",
                        )
                    except Exception:
                        pass

        if not lista_per:
            st.info("Aún no hay facturas procesadas. Carga algunos XMLs primero.")
        else:
            per_rep = st.selectbox("Selecciona un período", lista_per, key="per_reporte")
            rows = conn.execute(
                """SELECT f.uuid, f.folio, f.id_contrato, f.tipo, f.mes_contrato,
                          f.subtotal, f.total, f.estatus,
                          COALESCE((SELECT GROUP_CONCAT(fc.descripcion, ' | ') FROM factura_conceptos fc WHERE fc.uuid = f.uuid), f.observaciones) as concepto,
                          f.observaciones, f.fecha_emision, f.fecha_registro, f.cancelada
                   FROM facturas f WHERE f.periodo=? ORDER BY f.fecha_emision""",
                (per_rep,)
            ).fetchall()
            if not rows:
                st.warning("No hay facturas para este período.")
            else:
                        df_rep = pd.DataFrame(rows, columns=[
                            'UUID', 'Folio', 'Contrato', 'Tipo', 'Mes',
                            'Subtotal', 'Total', 'Estatus', 'Concepto', 'Observación', 'Fecha Emisión', 'Fecha Registro', 'Cancelada'
                        ])
                        df_rep['UUID'] = df_rep['UUID'].str[:8] + '…'
                        df_rep['Mes'] = df_rep['Mes'].astype(str).replace(['nan', 'None', '<NA>', 'NaN'], '-').fillna('-')
                        df_rep['Cancelada'] = df_rep['Cancelada'].fillna(0).astype(int).map({1:'Sí', 0:'No'})

                        total_fact = len(df_rep)
                        conciliados = len(df_rep[df_rep['Estatus'] == 'CONCILIADO'])
                        discrepancia = len(df_rep[df_rep['Estatus'] == 'DISCREPANCIA'])
                        sin_contrato = len(df_rep[df_rep['Estatus'] == 'SIN_CONTRATO'])
                        no_aplica = len(df_rep[df_rep['Estatus'] == 'NO_APLICA'])
                        monto_total = df_rep['Total'].sum()

                        col1, col2, col3, col4, col5, col6 = st.columns(6)
                        col1.metric("Total Facturas", f"{total_fact:,}")
                        col2.metric("Conciliados", f"{conciliados:,}", delta=f"{conciliados/total_fact*100:.1f}%" if total_fact else "")
                        col3.metric("Discrepancia", f"{discrepancia:,}", delta=f"{discrepancia/total_fact*100:.1f}%" if total_fact else "", delta_color="inverse")
                        col4.metric("Sin Contrato", f"{sin_contrato:,}", delta=f"{sin_contrato/total_fact*100:.1f}%" if total_fact else "", delta_color="inverse")
                        col5.metric("No aplica", f"{no_aplica:,}", delta=f"{no_aplica/total_fact*100:.1f}%" if total_fact else "")
                        col6.metric("Monto Total", f"${monto_total:,.2f}")

                        st.markdown("---")
                        st.markdown('<span class="section-label">Control global del período</span>', unsafe_allow_html=True)
                        st.caption("Compara la renta pactada de contratos activos contra lo facturado en el mes.")
                        df_act_rep = obtener('ACTIVO')
                        esperado_periodo = float(df_act_rep['Mensualidad_Sin_IVA'].sum()) if not df_act_rep.empty else 0.0
                        facturado_vigente = float(df_rep.loc[df_rep['Cancelada']=='No', 'Total'].sum())
                        dif_control = facturado_vigente - esperado_periodo
                        pct_cobertura = (facturado_vigente/esperado_periodo*100) if esperado_periodo else 0
                        cg1, cg2, cg3, cg4 = st.columns(4)
                        cg1.metric("Renta esperada (cartera activa)", f"${esperado_periodo:,.2f}")
                        cg2.metric("Facturado vigente (sin canceladas)", f"${facturado_vigente:,.2f}")
                        cg3.metric("Diferencia", f"${dif_control:+,.2f}", delta_color="inverse" if abs(dif_control) > esperado_periodo*0.02 else "normal")
                        cg4.metric("% de cobertura", f"{pct_cobertura:.1f}%")
                        if esperado_periodo and abs(dif_control) > esperado_periodo * 0.02:
                            st.warning(
                                f"La diferencia es de ${abs(dif_control):,.2f} ({abs(pct_cobertura-100):.1f} puntos), "
                                f"más del 2% de la renta esperada — vale la pena revisar 'Contratos sin Factura' y "
                                f"'Resolver Pendientes' para este período antes de dar por cerrado el mes."
                            )
                        else:
                            st.success("Lo facturado vigente cuadra razonablemente contra la renta esperada de la cartera activa.")

                        st.markdown("---")

                        r1c1, r1c2 = st.columns(2)
                        with r1c1:
                            status_counts = df_rep['Estatus'].value_counts().reset_index()
                            status_counts.columns = ['Estatus', 'Cantidad']
                            colores_estatus = {
                                'CONCILIADO': C_PASTEL['success'],
                                'DISCREPANCIA': C_PASTEL['warning'],
                                'SIN_CONTRATO': C_PASTEL['accent'],
                                'PENDIENTE': C_PASTEL['info'],
                                'ERROR': '#E39490',
                                'NO_APLICA': '#CBC6B8',
                                'RFC_INCORRECTO': '#D97C7C',
                                'FUERA_DE_VIGENCIA': '#8A2019',
                                'CANCELADA': '#9AA3AC'
                            }
                            fig_pie = px.pie(status_counts, values='Cantidad', names='Estatus',
                                             title='Distribución por Estatus',
                                             color='Estatus', color_discrete_map=colores_estatus,
                                             hole=0.4)
                            fig_pie = sfig(fig_pie, h=280)
                            fig_pie.update_traces(textposition='inside', textinfo='percent+label')
                            st.plotly_chart(fig_pie, width='stretch', key="pc_002")
                            explain("Composición de las facturas del período",
                                    "Muestra el porcentaje de facturas conciliadas, con discrepancias, sin contrato o no aplica.")

                        with r1c2:
                            df_rep['Mes_Emision'] = pd.to_datetime(df_rep['Fecha Emisión']).dt.to_period('M').astype(str)
                            evo = df_rep.groupby(['Mes_Emision', 'Estatus']).size().reset_index(name='Cantidad')
                            if not evo.empty:
                                fig_line = px.line(evo, x='Mes_Emision', y='Cantidad', color='Estatus',
                                                   title='Evolución Mensual de Facturas',
                                                   markers=True, color_discrete_map=colores_estatus)
                                fig_line = sfig(fig_line, h=280)
                                fig_line.update_traces(line_width=2.5)
                                st.plotly_chart(fig_line, width='stretch', key="pc_003")
                                explain("Tendencia de facturación por mes",
                                        "Cuántas facturas se registraron cada mes, desglosadas por estatus.")
                            else:
                                st.info("No hay datos para evolución mensual.")

                        st.markdown("---")
                        st.markdown("#### Detalle de Facturas")
                        fmt_rep = {'Subtotal': '${:,.2f}', 'Total': '${:,.2f}'}
                        styled_rep = df_rep.style.format(fmt_rep)
                        def color_status(val):
                            if val == 'CONCILIADO': return 'background-color:#E1F2E7;color:#155C3B'
                            elif val == 'DISCREPANCIA': return 'background-color:#FBF0DA;color:#7A5209'
                            elif val == 'SIN_CONTRATO': return 'background-color:#F3EDD3;color:#6B5A1F'
                            elif val == 'ERROR': return 'background-color:#FAE3E1;color:#8A2019'
                            elif val == 'RFC_INCORRECTO': return 'background-color:#f5d0d0;color:#7A1015;font-weight:700'
                            elif val == 'CANCELADA': return 'background-color:#e2e2e2;color:#555;text-decoration:line-through'
                            elif val == 'FUERA_DE_VIGENCIA': return 'background-color:#e2c2c2;color:#6b1515;font-weight:700;text-decoration:line-through'
                            elif val == 'NO_APLICA': return 'background-color:#E9EBEE;color:#565E68'
                            return ''
                        styled_rep = styled_rep.map(color_status, subset=['Estatus'])
                        st.dataframe(styled_rep, width='stretch', height=400, key="df_005")

                        excel_data = exportar_reporte_conciliacion(per_rep, df_rep)
                        st.download_button(
                            "Descargar Reporte Excel",
                            excel_data,
                            f"reporte_conciliacion_{per_rep}.xlsx",
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )

def _render_analiticas():
    st.title("Comparar Analíticas Contables")
    st.caption("Compara el Excel de analíticas contables contra los devengamientos del sistema para detectar diferencias.")
    arch_analitica = st.file_uploader("Excel de analíticas", type=["xlsx", "xls"], key="analiticas_uploader")

    if arch_analitica is not None:
        try:
            df_tidy, anio_detectado, hojas_no_rec = parse_analiticas_excel(arch_analitica.read())
        except Exception as e:
            st.error(f"No se pudo leer el archivo: {e}")
            df_tidy, anio_detectado, hojas_no_rec = pd.DataFrame(), None, []

        if hojas_no_rec:
            st.info(f"Hojas no reconocidas (omitidas): {', '.join(hojas_no_rec)}.")

        if df_tidy.empty:
            st.warning("No se encontraron registros en las hojas reconocidas.")
        else:
            anio_cmp = st.number_input(
                "Año que cubre esta analítica", min_value=2015, max_value=2100,
                value=anio_detectado or hoy_ref().year, step=1
            )
            n_sin_id = df_tidy['ID_Contrato_Detectado'].isna().sum()
            if n_sin_id:
                st.caption(f"{n_sin_id} renglones sin número de contrato identificable.")

            if st.button("Comparar contra el sistema", type="primary"):
                with st.spinner("Calculando lo esperado y comparando…"):
                    df_cmp = comparar_analiticas(df_tidy, int(anio_cmp))
                st.session_state['analiticas_comparado'] = df_cmp
                st.session_state['analiticas_anio'] = int(anio_cmp)

    df_cmp = st.session_state.get('analiticas_comparado')
    if df_cmp is not None and not df_cmp.empty:
        st.divider()
        total_renglones = len(df_cmp)
        con_diferencia = int(df_cmp['_relevante'].sum())
        no_encontrados = int((df_cmp['Esperado $'].isna()).sum())
        monto_dif_neto = df_cmp['Diferencia $'].sum(skipna=True)

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Renglones comparados", f"{total_renglones:,}")
        k2.metric("Con diferencia relevante", f"{con_diferencia:,}")
        k3.metric("Contratos no encontrados", f"{no_encontrados:,}")
        k4.metric("Diferencia neta total", f"${monto_dif_neto:,.2f}")

        solo_diferencias = st.checkbox("Mostrar solo renglones con diferencia", value=True)
        df_mostrar = df_cmp[df_cmp['_relevante']] if solo_diferencias else df_cmp
        df_mostrar = df_mostrar.drop(columns=['_relevante'])

        if df_mostrar.empty:
            st.success("Todo lo que trae el Excel coincide con lo que el sistema esperaba — sin diferencias relevantes.")
        else:
            def _color_dif(val):
                if pd.isna(val):
                    return 'background-color:#FAE3E1;color:#8A2019'
                if abs(val) > 1:
                    return 'background-color:#FBF0DA;color:#7A5209;font-weight:700'
                return ''
            fmt_cmp = {'Registrado $': '${:,.2f}', 'Esperado $': '${:,.2f}', 'Diferencia $': '${:+,.2f}'}
            styled_cmp = df_mostrar.style.format(fmt_cmp, na_rep='—').map(_color_dif, subset=['Diferencia $'])
            st.dataframe(styled_cmp, width='stretch', height=440, key="df_analiticas_cmp")

            buf_cmp = excel_con_formato(
                {'Comparativo': df_cmp.drop(columns=['_relevante'])},
                currency_cols=['Registrado $', 'Esperado $', 'Diferencia $']
            )
            st.download_button(
                "Descargar comparativo Excel (con todas las observaciones)",
                buf_cmp,
                f"comparativo_analiticas_{st.session_state.get('analiticas_anio','')}.xlsx",
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

            st.divider()
            st.markdown("**Sugerencia de póliza de ajuste**")
            st.caption("Propuesta de cargos y abonos para cuadrar las diferencias encontradas.")
            df_ajuste = sugerir_ajuste_poliza(df_cmp, CAT)
            if df_ajuste.empty:
                st.info("No hay diferencias netas que ameriten un ajuste.")
            else:
                tc_aj, ta_aj = df_ajuste['Cargo'].sum(), df_ajuste['Abono'].sum()
                ca1, ca2 = st.columns(2)
                ca1.metric("Total Cargos sugeridos", f"${tc_aj:,.2f}")
                ca2.metric("Total Abonos sugeridos", f"${ta_aj:,.2f}")
                fmt_aj = {'Cargo': '${:,.2f}', 'Abono': '${:,.2f}'}
                st.dataframe(df_ajuste.style.format(fmt_aj), width='stretch', height=300, key="df_ajuste_analiticas")
                buf_aj = excel_con_formato({'Ajuste_Sugerido': df_ajuste}, currency_cols=['Cargo', 'Abono'])
                st.download_button(
                    "Descargar póliza de ajuste sugerida (Excel)",
                    buf_aj,
                    f"poliza_ajuste_sugerida_{st.session_state.get('analiticas_anio','')}.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

def exportar_reporte_conciliacion(periodo: str, df: pd.DataFrame) -> bytes:
    """Genera un archivo Excel con el reporte de conciliación para el período dado."""
    resumen = pd.DataFrame({
        'Métrica': ['Período', 'Total Facturas', 'Conciliados', 'Discrepancia', 'Sin Contrato', 'No Aplica', 'Monto Total'],
        'Valor': [
            periodo,
            len(df),
            len(df[df['Estatus'] == 'CONCILIADO']),
            len(df[df['Estatus'] == 'DISCREPANCIA']),
            len(df[df['Estatus'] == 'SIN_CONTRATO']),
            len(df[df['Estatus'] == 'NO_APLICA']),
            df['Total'].sum()
        ]
    })
    df_detalle = df.copy()

    if 'Mes_Emision' in df.columns:
        stats = df.groupby(['Mes_Emision', 'Estatus']).size().unstack(fill_value=0).reset_index()
    else:
        df['Mes_Emision'] = pd.to_datetime(df['Fecha Emisión']).dt.to_period('M').astype(str)
        stats = df.groupby(['Mes_Emision', 'Estatus']).size().unstack(fill_value=0).reset_index()

    buf = excel_con_formato(
        {'Resumen': resumen, 'Detalle': df_detalle, 'Estadísticas': stats},
        currency_cols={'Detalle': ['Total', 'Subtotal']},
    )
    return buf.getvalue()

# Arranque
def _pantalla_bd_danada(db_path: str, error_tecnico: str):
    st.title("La base de datos no se pudo abrir")
    st.error(
        "El archivo de base de datos se dañó y no se puede abrir. "
        "Puedes restaurar uno de los respaldos automáticos a continuación."
    )
    respaldos = listar_respaldos_automaticos(db_path)
    if respaldos:
        st.success(f"Hay {len(respaldos)} respaldo(s) disponible(s) para restaurar.")
        opciones = {
            f"{fecha.strftime('%Y-%m-%d %H:%M:%S')}" + (" (cifrado)" if cifrado else ""): archivo
            for archivo, fecha, cifrado in respaldos
        }
        elegido = st.selectbox("Elige el respaldo a restaurar (el más reciente arriba):", list(opciones.keys()))
        archivo_elegido = opciones[elegido]
        clave_restaurar = None
        if archivo_elegido.endswith(".zip"):
            clave_restaurar = st.text_input("Contraseña del respaldo", type="password")
        st.caption("El archivo dañado se conserva por seguridad con otro nombre.")
        if st.button("Restaurar este respaldo y continuar", type="primary"):
            try:
                restaurar_respaldo_automatico(db_path, archivo_elegido, password=clave_restaurar)
                st.session_state.clear()
                st.success("Restaurado. Vuelve a abrir la aplicación.")
                st.stop()
            except RuntimeError as e:
                st.error(str(e))
    else:
        st.warning("No se encontraron respaldos automáticos disponibles para este archivo.")
    st.markdown("---")
    st.markdown("**Recomendación:** Guarda la carpeta del sistema en un disco local y fuera de carpetas sincronizadas en la nube.")
    with st.expander("Detalle técnico"):
        st.code(error_tecnico)
    st.stop()

_db_path_actual = get_db_path()
# Bug de rendimiento real: esta línea no tenía ningún candado, así que
# corría en CADA click de todo el sistema (Streamlit vuelve a ejecutar el
# script completo en cada interacción) — y "crear un respaldo" no es
# barato: revisa la base de datos entera con PRAGMA integrity_check y
# después copia el archivo completo a disco. Eso, repetido en cada botón
# que se toca durante toda la sesión, era buena parte de la lentitud
# general que se sentía en todo el sistema, no solo al abrirlo. La propia
# función ya decía en su comentario "se llama una vez por sesión" — nunca
# se hizo cumplir. Ahora sí: mismo candado que ya se usa para init_db().
_bkey = f"bkp_{_db_path_actual}"
if not st.session_state.get(_bkey):
    try:
        _resp_pass = get_cfg('respaldo_password') or None
    except Exception:
        _resp_pass = None  # primer arranque de la vida del sistema: la tabla de configuración todavía no existe
    crear_respaldo_automatico(_db_path_actual, password=_resp_pass)
    st.session_state[_bkey] = True
ikey=f"dbi_{_db_path_actual}"
if not st.session_state.get(ikey):
    try:
        init_db()
        st.session_state[ikey]=True
    except sqlite3.DatabaseError as _e:
        import time
        time.sleep(1.0) # Wait 1s for Antivirus to release the lock
        try:
            init_db() # Retry
            st.session_state[ikey]=True
        except sqlite3.DatabaseError as _e2:
            _pantalla_bd_danada(_db_path_actual, str(_e2))
CAT=cargar_catalogo()

def _pantalla_error_amigable(e: Exception, contexto: str = ""):
    """Se dispara cuando algo truena dentro de una pantalla. Antes esto se
    veía en el navegador como el error de Python crudo (traceback completo,
    nombres de función, números de línea) — confuso y con pinta de que el
    programa está roto. Ahora se ve una tarjeta clara, en el mismo estilo
    del resto del sistema, y el detalle técnico completo se sigue
    imprimiendo en la consola/terminal (para quien de soporte lo necesite)
    pero ya NO en la pantalla de quien está usando la app."""
    print(f"\n[O-Leasing] Error atrapado en '{contexto}':", file=sys.stderr)
    traceback.print_exc()
    st.markdown("""<div class="empty-state" style="border-color:var(--bad);">
      <div class="es-icon"></div>
      <span class="es-title">Esta pantalla no pudo cargar</span>
      <span class="es-msg">
        Ocurrió un problema temporal al cargar esta pantalla. Intenta recargar o volver al menú principal.
      </span>
    </div>""", unsafe_allow_html=True)
    c_err1, c_err2 = st.columns(2)
    with c_err1:
        if st.button("Reintentar", width='stretch'):
            st.session_state['_refresh'] = True
    with c_err2:
        if st.button("Ir al Dashboard", width='stretch'):
            st.session_state['menu_item'] = "Dashboard & Cartera"
            st.session_state['menu_grupo'] = "Cartera"
            st.session_state['_refresh'] = True
    with st.expander("Detalle técnico (para soporte)"):
        st.code(f"{type(e).__name__}: {e}")

MN=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
DN=['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo']

def fecha_larga(d):
    """Fecha en español ('miércoles 8 de julio de 2026') sin depender del
    idioma/locale configurado en el sistema operativo. strftime('%A'/'%B')
    puede salir en inglés en un servidor Windows si no tiene el locale en
    español instalado — esto evita ese problema por completo."""
    return f"{DN[d.weekday()]} {d.day} de {MN[d.month-1].lower()} de {d.year}"

vkey=f"venc_{get_db_path()}_{date.today().isoformat()}"
if not st.session_state.get(vkey):
    n_venc=procesar_vencimientos_naturales()
    st.session_state[vkey]=True
    if n_venc>0: st.session_state['_aviso_venc_auto']=n_venc

emp=get_empresa_actual()
nombre_empresa=get_cfg('nombre_empresa',emp.get('nombre','Orange Leasing'))
logo_data=get_cfg('logo')
_logo_sidebar_ok=False
if logo_data:
    try:
        # Se envuelve en una tarjeta clara: un logo subido por el usuario
        # puede traer trazos oscuros pensados para un fondo blanco, y sobre
        # el sidebar oscuro se perdería igual que le pasaba al wordmark por
        # defecto. Con esta tarjeta cualquier logo que suban se ve bien,
        # sin importar de qué color sea.
        with st.sidebar.container(key="logo_personalizado_wrap"):
            st.image(Image.open(io.BytesIO(base64.b64decode(logo_data))),width=176)
        _logo_sidebar_ok=True
    except Exception:
        pass  # logo guardado corrupto: cae al logo por defecto de abajo, no se queda en blanco
if not _logo_sidebar_ok:
    # Una sola marca al inicio, no dos: antes se mostraba el ícono solo aquí
    # arriba Y el nombre completo "O-Leasing" más abajo, a la mitad de la
    # lista de controles — dos logos separados sin relación visual clara.
    # Ahora el nombre completo (que ya incluye el ícono) es la marca
    # principal desde el primer momento.
    if _LOGO_WORDMARK_CLARO_B64:
        st.sidebar.markdown(f"""
        <div class="sidebar-wordmark">
          <img src="data:image/svg+xml;base64,{_LOGO_WORDMARK_CLARO_B64}" alt="O-Leasing">
        </div>
        """, unsafe_allow_html=True)
    elif _LOGO_SOLO_B64:
        st.sidebar.markdown(
            f'<div class="logo-seal"><img src="data:image/svg+xml;base64,{_LOGO_SOLO_B64}" alt="O-Leasing"></div>',
            unsafe_allow_html=True)
    else:
        st.sidebar.markdown('<div class="logo-seal"><span>O</span></div>', unsafe_allow_html=True)

_hora = datetime.now().hour
_saludo = "Buenos días" if _hora < 12 else ("Buenas tardes" if _hora < 19 else "Buenas noches")
st.sidebar.markdown(f'<div class="saludo-sidebar">{_saludo}</div>', unsafe_allow_html=True)

# "Empresa activa" y "Fecha de análisis" son controles relacionados (ambos
# cambian DESDE dónde estás viendo el sistema) — antes flotaban como dos
# widgets sueltos, sin nada que los agrupara visualmente como una sola
# sección de contexto. Ahora comparten una sola tarjeta.
with st.sidebar.container(key="contexto_wrap", border=True):
    st.markdown('<div class="emp-label">Empresa activa</div>', unsafe_allow_html=True)
    _emp_lista = get_empresas_permitidas()
    _emp_nombres = [e['nombre'] for e in _emp_lista]
    _emp_idx_actual = next((i for i, e in enumerate(_emp_lista) if e['id'] == emp['id']), 0)
    _emp_sel = st.selectbox(
        "Cambiar de empresa", _emp_nombres, index=_emp_idx_actual,
        key="sidebar_empresa_switch", label_visibility="collapsed",
        help="Cambia de empresa sin salir de la pantalla en la que estás."
    )
    if _emp_sel != _emp_lista[_emp_idx_actual]['nombre']:
        _emp_destino = next(e for e in _emp_lista if e['nombre'] == _emp_sel)
        st.session_state['empresa_id'] = _emp_destino['id']
        limpiar_seleccion_contrato()
        st.session_state['_refresh'] = True

    with st.expander(f"Fecha de análisis — {hoy_ref().strftime('%d/%b/%Y')}" + (" (cierre pasado)" if viendo_fecha_pasada() else " (hoy)")):
        _fa = st.date_input(
            "Ver el sistema como si fuera:", value=hoy_ref(), key="fecha_analisis_input",
            help="Cambia esta fecha para revisar cómo se veían los contratos, el punto de equilibrio y las "
                 "alertas en el cierre de un mes anterior. Afecta Dashboard, Estado de Cuenta y Punto de "
                 "Equilibrio. Vuelve a 'Hoy' cuando termines."
        )
        fcol1, fcol2 = st.columns(2)
        if fcol1.button("Aplicar", key="aplicar_fecha_analisis", width='stretch'):
            st.session_state['fecha_analisis'] = _fa
            st.rerun()
        if fcol2.button("Volver a hoy", key="reset_fecha_analisis", width='stretch'):
            st.session_state.pop('fecha_analisis', None)
            st.rerun()
if viendo_fecha_pasada():
    st.sidebar.warning(f"Viendo el sistema al **{hoy_ref().strftime('%d/%m/%Y')}**, no a hoy.")

st.sidebar.markdown("---")

# GRUPOS: estructura del menú (grupo -> páginas que contiene).
# No cambiar las claves de texto: se usan como identificador de página
# en todo el sistema (session_state y navegación cruzada entre pantallas).
GRUPOS = {
    "Cartera": [
        "Dashboard & Cartera",
        "Estado de Cuenta",
        "Carga Masiva y Altas",
        "Editar / Eliminar",
        "Gestor de Bajas",
        "Tabla Mensual por Contrato",
        "Reporte Maestro",
        "Reporte Maestro Saldos",
    ],
    "Gestión de Riesgo": [
        "Gestión de Morosidad",
        "Eventos Especiales",
        "Anotaciones",
    ],
    "Finanzas & Contabilidad": [
        "Pólizas Contables",
        "Intereses del Mes",
        "Facturación de Intereses",
        "Conciliación de Facturas",
        "Cierre y Conciliación Mensual",
        "Comparar Analíticas",
        "Tablas de Amortización",
    ],
    "Análisis": [
        "Proyección Financiera",
        "Cotizador Comercial",
        "Análisis de Rentabilidad",
        "Punto de Equilibrio",
        "Reportes por Cliente",
    ],
    "Configuración": [
        "Cuentas (Macro)",
        "Contpaqi (Cuentas)",
        "Respaldo y Restauración",
        "Multiempresa",
        "Usuarios y Roles",
    ],
}

# Descripción corta de cada grupo y cada página: se usa como tooltip al
# pasar el mouse por el menú y como encabezado de orientación arriba de
# cada pantalla, para que siempre quede claro qué se puede hacer ahí.
DESCRIPCIONES_GRUPO = {
    "Cartera": "Alta, edición, consulta y bajas de tus contratos.",
    "Gestión de Riesgo": "Morosidad, eventos especiales y seguimiento.",
    "Finanzas & Contabilidad": "Pólizas, intereses, facturación y amortización.",
    "Análisis": "Proyecciones, rentabilidad, punto de equilibrio y reportes.",
    "Configuración": "Ajustes generales, cuentas contables y respaldos.",
}
DESCRIPCIONES = {
    "Dashboard & Cartera": "Cartera vigente comparable con contabilidad: inversion neta, saldo capital, CxC, residual, cobertura del periodo y export Excel.",
    "Estado de Cuenta": "Consulta el estado de cuenta detallado de un contrato o cliente específico.",
    "Carga Masiva y Altas": "Sube archivos para dar de alta contratos nuevos de forma masiva.",
    "Editar / Eliminar": "Modifica o elimina la información de un contrato ya existente.",
    "Gestor de Bajas": "Administra los contratos que han sido dados de baja o terminados.",
    "Tabla Mensual por Contrato": "Consulta las tablas base contables (Interés, Residual, Comisión).",
    "Reporte Maestro": "Reporte gerencial completo de AMORTIZACIÓN (flujo) por contrato.",
    "Reporte Maestro Saldos": "Reporte gerencial de SALDOS INSOLUTOS (lo que deben) por contrato.",
    "Gestión de Morosidad": "Da seguimiento y actualiza el nivel de morosidad de los contratos.",
    "Eventos Especiales": "Registra y resuelve eventos especiales como siniestros o incidencias.",
    "Anotaciones": "Guarda notas y comentarios de seguimiento sobre tus contratos y clientes.",
    "Pólizas Contables": "Consulta las pólizas contables generadas por el sistema.",
    "Intereses del Mes": "Calcula y revisa los intereses correspondientes al mes en curso.",
    "Facturación de Intereses": "Genera y administra la facturación de los intereses cobrados.",
    "Conciliación de Facturas": "Concilia las facturas CFDI recibidas contra tus contratos registrados.",
    "Cierre y Conciliación Mensual": "Resumen ejecutivo mensual de movimientos, facturación, cobranza y pólizas para cierre de mes contable con exportación PDF.",
    "Comparar Analíticas": "Sube las analíticas contables del contador y compáralas contra lo que el sistema esperaba, con sugerencia de póliza de ajuste.",
    "Tablas de Amortización": "Consulta la tabla de amortización completa de cualquier contrato.",
    "Proyección Financiera": "Proyecta el comportamiento financiero futuro de tu cartera.",
    "Cotizador Comercial": "Simulador interactivo de nuevos arrendamientos puros con desglose inicial, amortización y exportación de cotización en PDF.",
    "Análisis de Rentabilidad": "Analiza la rentabilidad de tus contratos y de la cartera en general.",
    "Punto de Equilibrio": "Calcula el punto de equilibrio de tu operación.",
    "Reportes por Cliente": "Genera reportes personalizados agrupados por cliente.",
    "Cuentas (Macro)": "Configura las cuentas contables macro utilizadas por el sistema.",
    "Contpaqi (Cuentas)": "Configura la relación de cuentas para la integración con Contpaqi.",
    "Respaldo y Restauración": "Crea respaldos de tu información o restaura una copia anterior.",
    "Multiempresa": "Administra y cambia entre las distintas empresas registradas en el sistema.",
    "Usuarios y Roles": "Crea cuentas de acceso, asigna roles y desactiva usuarios que ya no deben entrar.",
}

# Los ítems de Configuración implican cambios de fondo (cuentas contables,
# respaldos, altas/bajas de empresas) — solo el rol admin los ve en el menú.
AUTH_USER = st.session_state.get("auth_user", {"username": "—", "nombre_completo": "—", "rol": "admin"})
ROL_ACTUAL = AUTH_USER.get("rol", "admin")
if ROL_ACTUAL not in ("admin", "super_usuario"):
    GRUPOS = {g: its for g, its in GRUPOS.items() if g != "Configuración"}

# Permisos por pantalla (admin marca con palomitas qué ve cada usuario)
_TODAS_PANTALLAS = [item for _g, _its in GRUPOS.items() for item in _its]
_uid_act = AUTH_USER.get("id")
_pantallas_usr = None
if _uid_act and ROL_ACTUAL not in ("admin", "super_usuario"):
    try:
        _pantallas_usr = obtener_pantallas_usuario(AUTH_DB, int(_uid_act))
    except Exception:
        _pantallas_usr = None
    if _pantallas_usr is not None:
        _permitidas = set(_pantallas_usr)
        GRUPOS = {
            g: [i for i in its if i in _permitidas]
            for g, its in GRUPOS.items()
        }
        GRUPOS = {g: its for g, its in GRUPOS.items() if its}

# El rol lectura puede consultar todo pero no capturar/editar/borrar — se
# bloquean aquí las pantallas cuyo propósito central es escribir datos.
PANTALLAS_SOLO_ESCRITURA = {"Carga Masiva y Altas", "Editar / Eliminar", "Gestor de Bajas"}

if 'menu_grupo' not in st.session_state:
    st.session_state['menu_grupo'] = "Cartera"
if 'menu_item'  not in st.session_state:
    st.session_state['menu_item']  = "Dashboard & Cartera"
if st.session_state['menu_grupo'] not in GRUPOS:
    st.session_state['menu_grupo'] = "Cartera"
    st.session_state['menu_item']  = "Dashboard & Cartera"

_QUICKNAV = [
    ("Dashboard", "Dashboard & Cartera", "Cartera", "Ir al Dashboard"),
    ("Estado cta.", "Estado de Cuenta", "Cartera", "Ir a Estado de Cuenta"),
    ("Conciliación", "Conciliación de Facturas", "Finanzas & Contabilidad", "Ir a Conciliación de Facturas"),
    ("Pto. equilibrio", "Punto de Equilibrio", "Análisis", "Ir a Punto de Equilibrio"),
]
st.sidebar.markdown('<div class="nav-hint" style="padding-bottom:4px;">ACCESOS RÁPIDOS</div>', unsafe_allow_html=True)
_qn_cols_1 = st.sidebar.columns(2)
_qn_cols_2 = st.sidebar.columns(2)
for _qi, (_label, _target, _grupo_t, _tip) in enumerate(_QUICKNAV):
    col = _qn_cols_1[_qi] if _qi < 2 else _qn_cols_2[_qi - 2]
    if col.button(_label, key=f"qn_{_qi}", help=_tip, use_container_width=True):
        st.session_state['menu_item'] = _target
        st.session_state['menu_grupo'] = _grupo_t
        st.rerun()

st.sidebar.markdown('<div class="nav-hint">MENÚ · toca una sección para ver sus páginas</div>', unsafe_allow_html=True)

for grupo, items in GRUPOS.items():
    abierto = (st.session_state['menu_grupo'] == grupo)
    if st.sidebar.button(
        grupo,
        key=f"grp_{grupo}",
        width='stretch',
        type="primary" if abierto else "secondary",
        help=DESCRIPCIONES_GRUPO.get(grupo, "")
    ):
        st.session_state['menu_grupo'] = grupo
        st.session_state['menu_item']  = items[0]
        st.session_state['_refresh'] = True
    if abierto:
        for item in items:
            activo = (st.session_state['menu_item'] == item)
            if st.sidebar.button(item, key=f"item_{item}", width='stretch',
                                 type="primary" if activo else "secondary",
                                 help=DESCRIPCIONES.get(item, "")):
                st.session_state['menu_item'] = item
                st.session_state['_refresh'] = True

menu = st.session_state['menu_item']

# Si el usuario no tiene permiso a la pantalla actual, redirigir a la primera permitida
if ROL_ACTUAL not in ("admin", "super_usuario") and _pantallas_usr is not None:
    if menu not in _pantallas_usr:
        _first = next((it for its in GRUPOS.values() for it in its), None)
        if _first:
            st.session_state['menu_item'] = _first
            for _g, _its in GRUPOS.items():
                if _first in _its:
                    st.session_state['menu_grupo'] = _g
                    break
            st.warning(f"No tienes permiso para ver **{menu}**. Se muestra **{_first}**.")
            menu = _first
        else:
            st.error("Tu usuario no tiene pantallas asignadas. Pide al administrador que marque al menos una.")
            st.stop()

st.sidebar.markdown(f"""<div class="empresa-badge">
  <span class="emp-label">Sesión activa</span>
  <span class="emp-name">{AUTH_USER.get('nombre_completo','—')}</span>
  <span class="emp-label" style="margin-top:2px;display:block;">{ROL_LABELS.get(ROL_ACTUAL, ROL_ACTUAL)}</span>
</div>""", unsafe_allow_html=True)

if st.sidebar.button("Cambiar mi contraseña", key="btn_cambiar_pwd"):
    st.session_state['mostrar_perfil'] = True

if st.session_state.get('mostrar_perfil'):
    @st.dialog("Mi Perfil - Cambiar Contraseña")
    def dialog_perfil():
        st.write("Cambia la contraseña de tu cuenta:")
        _pwd_act = st.text_input("Contraseña actual", type="password")
        _pwd_new = st.text_input("Nueva contraseña", type="password")
        _pwd_new2 = st.text_input("Confirmar nueva contraseña", type="password")
        if st.button("Guardar contraseña"):
            if _pwd_new != _pwd_new2:
                st.error("Las contraseñas nuevas no coinciden.")
            else:
                ok, msg = resetear_password(AUTH_DB, AUTH_USER['id'], _pwd_new, password_actual=_pwd_act, require_actual=True)
                if ok:
                    st.success(msg)
                    st.session_state['mostrar_perfil'] = False
                    st.rerun()
                else:
                    st.error(msg)
    dialog_perfil()

if st.sidebar.button("Cerrar sesión", key="btn_logout", width='stretch'):
    _tok_logout = st.query_params.get("st_auth")
    if _tok_logout:
        try:
            cerrar_sesion(AUTH_DB, _tok_logout)
        except Exception:
            pass
        st.query_params.pop("st_auth", None)
    st.session_state.pop("auth_user", None)
    st.rerun()

st.sidebar.markdown('<div class="sidebar-footer">O-Leasing v6 · © 2026 · by Javier Illán<br>Cifras en Pesos Mexicanos (MXN)</div>',unsafe_allow_html=True)
if _LOGO_ORANGE_B64:
    st.sidebar.markdown(
        f'<div class="sidebar-submarca">'
        f'<img src="data:image/svg+xml;base64,{_LOGO_ORANGE_B64}" alt="Orange">'
        f'<span>Una app de Orange</span></div>',
        unsafe_allow_html=True)

# --- REFRESH CONTROL ---
if st.session_state.get('_refresh', False):
    st.session_state['_refresh'] = False
    st.rerun()

if st.session_state.get('_aviso_venc_auto'):
    n_aviso=st.session_state.pop('_aviso_venc_auto')
    st.info(f"{n_aviso} contrato(s) llegaron al final de su plazo y se pasaron automáticamente a **BAJA por terminación natural**. Puedes verlos en Gestor de Bajas.")

if st.session_state.get('flash_msg'):
    _tipo_flash, _texto_flash = st.session_state.pop('flash_msg')
    getattr(st, _tipo_flash)(_texto_flash)

# Encabezado con la orientación de la pantalla: sección, qué se puede hacer
# ahí, y la fecha con la que están calculados los saldos/gráficas (por
# default la fecha real, salvo que uses el selector de "Fecha de análisis").
_grupo_de_item = {it: g for g, its in GRUPOS.items() for it in its}
_grupo_actual  = _grupo_de_item.get(menu, "")
_desc_actual   = DESCRIPCIONES.get(menu, "")
_grupo_nombre  = _grupo_actual.split(" ", 1)[-1] if _grupo_actual else ""
_hoy_ref = hoy_ref()
_hoy_ref_txt = f"{_hoy_ref.day} de {MN[_hoy_ref.month-1].lower()} de {_hoy_ref.year}"

# Barra superior fija con el logo: se pinta en TODAS las pantallas, aquí en
# el área principal (no solo en el sidebar), para que nunca desaparezca
# aunque el usuario colapse el menú lateral.
if logo_data:
    try:
        _thb_fmt = (Image.open(io.BytesIO(base64.b64decode(logo_data))).format or "PNG").lower()
        _thb_logo_html = f'<img src="data:image/{_thb_fmt};base64,{logo_data}" alt="{nombre_empresa}">'
    except Exception:
        _thb_logo_html = '<div class="thb-seal">O</div>'
elif _LOGO_WORDMARK_B64:
    _thb_logo_html = f'<img src="data:image/svg+xml;base64,{_LOGO_WORDMARK_B64}" alt="O-Leasing">'
else:
    _thb_logo_html = '<div class="thb-seal">O</div>'
st.markdown(f"""
<div class="top-header-bar">
  <div class="thb-brand">
    {_thb_logo_html}
    <span class="thb-empresa"><span class="thb-tag">Empresa activa</span>{nombre_empresa}</span>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown(f"""
<div style="background:#FFFFFF;border:1px solid #DCE0E5;border-radius:6px;
            padding:.65rem 1.1rem;margin-bottom:1.1rem;
            display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:8px;">
  <div>
    <div style="font-size:.68rem;font-weight:700;color:#8A929C;text-transform:uppercase;letter-spacing:.7px;">
      {_grupo_nombre} {'›' if _grupo_nombre else ''} <span style="color:#1E5C4F;">{menu}</span>
    </div>
    {f'<div style="font-size:.85rem;color:#565E68;margin-top:3px;">{_desc_actual}</div>' if _desc_actual else ''}
  </div>
  <div title="Todos los cálculos, gráficas y proyecciones de esta pantalla toman esta fecha como 'hoy'.{' Estás viendo un cierre pasado, no el día de hoy.' if viendo_fecha_pasada() else ''}"
       style="background:{'#96660C' if viendo_fecha_pasada() else '#1E5C4F'};
              color:#fff;font-size:.72rem;font-weight:700;
              padding:.3rem .8rem;border-radius:4px;white-space:nowrap;">
    {'Viendo cierre al' if viendo_fecha_pasada() else 'Cálculos al'} {_hoy_ref_txt}
  </div>
</div>
""", unsafe_allow_html=True)

# --- BÚSQUEDA GLOBAL / COMMAND CENTER ---
with st.expander("Búsqueda Global — Buscar Contratos, Clientes, Series o Facturas", expanded=st.session_state.get("_open_search", False)):
    _q_glob = st.text_input("Ingresa cualquier término (ID Contrato, Cliente, Vehículo, Serie/VIN o Anotaciones):", key="global_search_q_input", placeholder="Ej: 0773, Besthelg, Ford, 3FA6P...").strip()
    if _q_glob:
        _df_con_g = obtener()
        if not _df_con_g.empty:
            _mask_c = (
                _df_con_g['ID_Contrato'].astype(str).str.contains(_q_glob, case=False, na=False) |
                _df_con_g['Cliente'].astype(str).str.contains(_q_glob, case=False, na=False) |
                _df_con_g['Vehiculo'].astype(str).str.contains(_q_glob, case=False, na=False) |
                _df_con_g.get('Serie', pd.Series(dtype=str)).astype(str).str.contains(_q_glob, case=False, na=False) |
                _df_con_g.get('Anotaciones', pd.Series(dtype=str)).astype(str).str.contains(_q_glob, case=False, na=False)
            )
            _res_con = _df_con_g[_mask_c]
            st.markdown(f"**Contratos encontrados ({len(_res_con)}):**")
            if not _res_con.empty:
                for _, _r_g in _res_con.head(10).iterrows():
                    _gcol1, _gcol2, _gcol3, _gcol4 = st.columns([2, 3, 2, 2])
                    _gcol1.markdown(f"**{_r_g['ID_Contrato']}**")
                    _gcol2.markdown(f"**{_r_g['Cliente']}**<br><small style='color:#666;'>{_r_g['Vehiculo']}</small>", unsafe_allow_html=True)
                    _gcol3.markdown(f"Renta: **${float(_r_g.get('Mensualidad_Sin_IVA',0)):,.2f}**")
                    if _gcol4.button("Estado Cta", key=f"btn_gs_ec_{_r_g['ID_Contrato']}"):
                        st.session_state['ec_contrato'] = _r_g['ID_Contrato']
                        st.session_state['menu_item'] = "Estado de Cuenta"
                        st.session_state['menu_grupo'] = "Cartera"
                        st.session_state['_open_search'] = False
                        st.rerun()
            else:
                st.caption("No se encontraron coincidencias en la cartera.")

if ROL_ACTUAL == "lectura" and menu in PANTALLAS_SOLO_ESCRITURA:
    st.warning(
        f"Tu usuario tiene rol de **solo lectura** y esta pantalla ({menu}) es para capturar o modificar "
        "información. Pide a un usuario con rol de captura o administrador que haga este cambio."
    )
    st.stop()

# Dashboard
try:
    if menu=="Dashboard & Cartera":
        # Solo cifras de cartera vigente + mes seleccionado + acumulado a ese mes.
        hoy_sistema = hoy_ref()
        df_all = obtener()
        df = obtener('ACTIVO')
        _tot_c = {}
        _det_c = pd.DataFrame()
        _acu_prev = {}
        _acu_am = {}
        _acu_anio = {}

        if _cc_err is not None:
            st.warning(
                f"core/cartera_contable.py no cargo bien ({type(_cc_err).__name__}: {_cc_err}). "
                "Se usan formulas internas de respaldo (igual Tabla Mensual). "
                "Copia el archivo del ZIP a core\\cartera_contable.py y borra core\\__pycache__."
            )
        st.markdown(
            '<div style="background:#1E5C4F;color:#fff;padding:10px 14px;border-radius:6px;margin-bottom:12px;font-weight:600;">'
            'DASHBOARD CARTERA / CONTABILIDAD — elige mes y año para ver el mes y el acumulado a ese corte'
            '</div>',
            unsafe_allow_html=True,
        )

        col_m, col_a, col_info = st.columns([1, 1, 2])
        with col_m:
            mes_sel = st.selectbox(
                "Mes de análisis",
                list(range(1, 13)),
                index=max(0, hoy_sistema.month - 1),
                format_func=lambda m: MN[m - 1],
                key="dash_mes_analisis",
            )
        with col_a:
            anio_sel = st.number_input(
                "Año de análisis",
                min_value=2000,
                max_value=2100,
                value=int(hoy_sistema.year),
                step=1,
                key="dash_anio_analisis",
            )
        # Corte = último día del mes seleccionado (para saldos y acumulado)
        try:
            _corte = (date(int(anio_sel), int(mes_sel), 1) + relativedelta(months=1) - relativedelta(days=1))
        except Exception:
            _corte = hoy_sistema
        if _corte > hoy_sistema:
            _corte = hoy_sistema
        _periodo = f"{int(anio_sel):04d}-{int(mes_sel):02d}"
        with col_info:
            st.markdown(
                f"**Corte:** {_corte.isoformat()}  \n"
                f"**Periodo del mes:** `{_periodo}`  \n"
                f"Primer mes en firma = mes 1 de amortización."
            )

        busq_top = st.text_input(
            "Buscar contrato o cliente → Estado de cuenta",
            placeholder="ID o nombre…",
            key="dash_busq_top",
        )
        if busq_top.strip() and not df_all.empty:
            _match_top = df_all[
                df_all['ID_Contrato'].str.contains(busq_top, case=False, na=False)
                | df_all['Cliente'].str.contains(busq_top, case=False, na=False)
            ]
            if not _match_top.empty:
                for _, _mr in _match_top.head(5).iterrows():
                    if st.button(
                        f"{_mr['ID_Contrato']} — {_mr['Cliente']}",
                        key=f"jump_{_mr['ID_Contrato']}",
                        width='stretch',
                    ):
                        st.session_state['ec_contrato'] = _mr['ID_Contrato']
                        st.session_state['menu_item'] = "Estado de Cuenta"
                        for g, its in GRUPOS.items():
                            if "Estado de Cuenta" in its:
                                st.session_state['menu_grupo'] = g
                        st.rerun()
            else:
                st.caption("Sin coincidencias.")

        # --- VERDAD = Tabla Mensual por Contrato (misma funcion / mismos numeros del Excel) ---
        st.markdown("---")
        st.subheader(f"Comparativos contables — {_periodo}")
        st.caption("Cálculo de intereses y amortización para el período seleccionado.")

        with st.spinner("Calculando con logica de Tabla Mensual…"):
            _tot_tm = totales_intereses_desde_tabla_mensual(int(anio_sel), int(mes_sel))

        if _tot_tm.get("error"):
            st.error(_tot_tm["error"])

        _m_calc = metricas_estilo_tabla_mensual(df_all, int(anio_sel), int(mes_sel))
        _m_tm = {
            "interes_leasing_mes": _tot_tm.get("interes_leasing_mes", 0),
            "interes_residual_mes": _tot_tm.get("interes_residual_mes", 0),
            "comision_mes": _tot_tm.get("comision_mes", 0),
            "interes_leasing_ytd": _tot_tm.get("interes_leasing_ytd", 0),
            "interes_residual_ytd": _tot_tm.get("interes_residual_ytd", 0),
            "comision_ytd": _tot_tm.get("comision_ytd", 0),
            "saldo_residual_mes": _tot_tm.get("saldo_residual_mes", 0),
            "n_contratos_mes": _tot_tm.get("n_contratos_mes", 0),
            "capital_mes": _m_calc.get("capital_mes", 0.0),
            "capital_ytd": _m_calc.get("capital_ytd", 0.0),
            "renta_mes": _m_calc.get("renta_mes", 0.0),
            "renta_ytd": _m_calc.get("renta_ytd", 0.0),
            "saldo_capital_mes": _m_calc.get("saldo_capital_mes", 0.0),
        }

        # Control: totales por mes = suma de columnas de los Excel de Tabla Mensual
        if _tot_tm.get("por_mes_leasing") or _tot_tm.get("por_mes_comision"):
            _df_ctrl = pd.DataFrame([
                {
                    "Mes": MN[m - 1],
                    "Intereses leasing": _tot_tm["por_mes_leasing"].get(m, 0),
                    "Intereses residual": _tot_tm["por_mes_residual"].get(m, 0),
                }
                for m in range(1, 13)
            ])
            st.dataframe(_df_ctrl, width="stretch", key="df_ctrl_int_mes_v4")
            st.success(
                f"Del mes ({MN[int(mes_sel)-1]}): "
                f"leasing **${_m_tm['interes_leasing_mes']:,.2f}** · "
                f"residual **${_m_tm['interes_residual_mes']:,.2f}**  |  "
                f"Acum ene–{MN[int(mes_sel)-1][:3].lower()}: "
                f"leasing **${_m_tm['interes_leasing_ytd']:,.2f}** · "
                f"residual **${_m_tm['interes_residual_ytd']:,.2f}**"
            )

        try:
            _tot_c, _det_c = consolidar_cartera_vigente(df, _corte)
        except Exception as _e_c:
            _tot_c, _det_c = {}, pd.DataFrame()
            st.warning(f"Saldos cartera: {_e_c}")

        # Facturacion del mes (si hay CFDI cargados)
        try:
            _conn_d = get_db()
            _row_f = _conn_d.execute(
                """SELECT COALESCE(SUM(total),0) AS t,
                          COALESCE(SUM(CASE WHEN tipo='MENSUAL' THEN total ELSE 0 END),0) AS tm,
                          COUNT(*) AS n
                   FROM facturas
                   WHERE periodo=? AND (cancelada IS NULL OR cancelada=0)
                     AND estatus IN ('CONCILIADO','DISCREPANCIA','PENDIENTE')""",
                (_periodo,),
            ).fetchone()
            _fact_mes = float(_row_f['t'] if _row_f else 0)
            _fact_mes_renta = float(_row_f['tm'] if _row_f else 0)
            _n_fact_mes = int(_row_f['n'] if _row_f else 0)
        except Exception:
            _fact_mes = _fact_mes_renta = 0.0
            _n_fact_mes = 0

        # ---- 1) DEL MES ----
        st.markdown(f"### 1. Del mes: {MN[int(mes_sel)-1]} {int(anio_sel)}")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Intereses leasing (del mes)", f"${_m_tm.get('interes_leasing_mes', 0):,.2f}")
        c2.metric("Intereses residual (del mes)", f"${_m_tm.get('interes_residual_mes', 0):,.2f}")
        c3.metric("Saldo residual activo (del mes)", f"${_m_tm.get('saldo_residual_mes', 0):,.2f}")
        c4.metric("Contratos en el mes", f"{_m_tm.get('n_contratos_mes', 0):,}")
        c5, = st.columns(1)
        c5.metric("Facturado rentas (CFDI)", f"${_fact_mes_renta:,.2f}")

        # ---- 2) ACUMULADO EJERCICIO ----
        st.markdown(f"### 2. Acumulado ejercicio {int(anio_sel)} (enero → {MN[int(mes_sel)-1].lower()})")
        a1, a2, a3 = st.columns(3)
        a1.metric("Intereses leasing (acum.)", f"${_m_tm.get('interes_leasing_ytd', 0):,.2f}")
        a2.metric("Intereses residual (acum.)", f"${_m_tm.get('interes_residual_ytd', 0):,.2f}")
        a3.metric("Contratos (ref. mes)", f"{_m_tm.get('n_contratos_mes', 0):,}")

        try:
            _df_fact_all = pd.read_sql_query(
                """SELECT periodo, tipo, total, subtotal, cancelada FROM facturas""",
                get_db(),
            )
        except Exception:
            _df_fact_all = pd.DataFrame()
        _acu_anio = acumulado_facturacion(_df_fact_all, anio=int(anio_sel), hasta_periodo=_periodo)
        b1, b2 = st.columns(2)
        b1.metric(f"Facturado total {int(anio_sel)} YTD", f"${_acu_anio.get('total', 0):,.2f}")
        b2.metric(f"Rentas facturadas {int(anio_sel)} YTD", f"${_acu_anio.get('mensual', 0):,.2f}")

        _dif_mes = round(_fact_mes_renta - _m_tm.get('renta_mes', 0), 2)
        _dif_ytd = round(_acu_anio.get('mensual', 0) - _m_tm.get('renta_ytd', 0), 2)
        st.info(
            f"Cuadre rentas del mes: CFDI ${_fact_mes_renta:,.2f} vs esperado ${_m_tm.get('renta_mes', 0):,.2f} "
            f"→ dif. ${_dif_mes:+,.2f}.  |  "
            f"YTD: CFDI ${_acu_anio.get('mensual', 0):,.2f} vs esperado ${_m_tm.get('renta_ytd', 0):,.2f} "
            f"→ dif. ${_dif_ytd:+,.2f}."
        )

        # Guardar para auxiliar y export
        _acu_am = {
            "intereses_devengados": _m_tm.get("interes_leasing_ytd", 0),
            "intereses_residual_devengados": _m_tm.get("interes_residual_ytd", 0),
            "comision_apertura_acum": _m_tm.get("comision_ytd", 0),
            "comision_apertura_mes": _m_tm.get("comision_mes", 0),
            "capital_amortizado": _m_tm.get("capital_ytd", 0),
            "rentas_esperadas_acum": _m_tm.get("renta_ytd", 0),
            "intereses_del_mes": _m_tm.get("interes_leasing_mes", 0),
            "intereses_residual_del_mes": _m_tm.get("interes_residual_mes", 0),
            "capital_del_mes": _m_tm.get("capital_mes", 0),
            "n_contratos": _m_tm.get("n_contratos_mes", 0),
        }
        _acu_prev = _acu_am

        if _m_tm.get("n_contratos_mes", 0) == 0:
            st.warning("No hay contratos con amortización registrados en este mes.")

        # --- Auxiliar contable ---
        st.markdown("---")
        st.subheader(f"4. Comparar con auxiliar contable (ejercicio {int(anio_sel)})")
        st.caption("Compara tu balanza o auxiliar contable contra las cifras del sistema.")
        _aux_file = st.file_uploader(
            "Auxiliar contable (.xlsx o .csv)",
            type=["xlsx", "csv"],
            key="dash_aux_uploader",
        )
        if _aux_file is not None:
            try:
                if _aux_file.name.lower().endswith(".csv"):
                    _df_aux = pd.read_csv(_aux_file)
                else:
                    _df_aux = pd.read_excel(_aux_file)
                st.write("Vista previa auxiliar:", _df_aux.head(8))
                _tot_sis = {
                    "intereses_devengados": _acu_am.get("intereses_devengados", 0),
                    "capital_amortizado": _acu_am.get("capital_amortizado", 0),
                    "rentas_esperadas_acum": _acu_am.get("rentas_esperadas_acum", 0),
                    "intereses_residual_devengados": _acu_am.get("intereses_residual_devengados", 0),
                    "facturado_ytd": _acu_anio.get("total", 0),
                }
                _cmp = comparar_auxiliar(_df_aux, _tot_sis)
                st.dataframe(_cmp, width="stretch", key="df_cmp_aux")
                try:
                    from reports.excel import excel_con_formato as _exc_aux
                    _buf_aux = _exc_aux(
                        {"Comparacion": _cmp, "Auxiliar": _df_aux, "Totales sistema": pd.DataFrame([_tot_sis])},
                    )
                    st.download_button(
                        "Excel comparacion vs auxiliar",
                        data=_buf_aux,
                        file_name=f"comparacion_auxiliar_{_periodo}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_cmp_aux",
                    )
                except Exception:
                    pass
            except Exception as _e_aux:
                st.error(f"No se pudo leer el auxiliar: {_e_aux}")

        tab_det, tab_cli, tab_anio, tab_exp = st.tabs([
            "Detalle por contrato al corte",
            "Por cliente",
            "Facturación por año",
            "Exportar Excel",
        ])
        with tab_det:
            if _det_c is not None and not _det_c.empty:
                st.dataframe(_det_c, width='stretch', height=360, key="df_cartera_mes")
            else:
                st.info("Sin detalle.")
        with tab_cli:
            if _det_c is not None and not _det_c.empty and "Cliente" in _det_c.columns:
                _por_cli = (
                    _det_c.groupby("Cliente", dropna=False)
                    .agg(
                        Contratos=("ID_Contrato", "count"),
                        Inversion_Neta=("Inversion_Neta", "sum"),
                        Saldo_Capital=("Saldo_Capital", "sum"),
                        Renta_Mensual=("Renta_Mensual", "sum"),
                        CxC_CP=("CxC_CP", "sum"),
                        CxC_LP=("CxC_LP", "sum"),
                    )
                    .reset_index()
                )
                _por_cli["CxC_Total"] = _por_cli["CxC_CP"] + _por_cli["CxC_LP"]
                st.dataframe(_por_cli.sort_values("Saldo_Capital", ascending=False), width='stretch', height=320, key="df_cli_mes")
            else:
                st.info("Sin datos.")
        with tab_anio:
            if not _df_fact_all.empty:
                _res_a = resumen_facturacion_por_anio(_df_fact_all, solo_vigentes=True)
                if not _res_a.empty:
                    st.dataframe(_res_a, width='stretch', height=280, key="df_fact_anio_dash")
                    try:
                        _pivot = _res_a.pivot_table(index="Anio", columns="Tipo", values="Total", aggfunc="sum", fill_value=0).reset_index()
                        fig_an = go.Figure()
                        for col in [c for c in _pivot.columns if c != "Anio"]:
                            fig_an.add_trace(go.Bar(x=_pivot["Anio"], y=_pivot[col], name=str(col)))
                        fig_an.update_layout(barmode="stack", title="Facturado por año", height=300)
                        st.plotly_chart(sfig(fig_an, h=300), width="stretch", key="pc_dash_anio")
                    except Exception:
                        pass
                else:
                    st.info("Sin facturas agrupables.")
            else:
                st.info("Aún no hay facturas cargadas. Sube XML en Conciliación → Carga.")
        with tab_exp:
            try:
                from reports.excel import excel_con_formato
                _hojas = {
                    "Saldos corte": _det_c if _det_c is not None else pd.DataFrame(),
                    "Totales corte": pd.DataFrame([_tot_c]) if _tot_c else pd.DataFrame(),
                    "Del mes": pd.DataFrame([{
                        "periodo": _periodo,
                        "renta_esperada": _cob["esperado"],
                        "facturado": _cob["facturado"],
                        "diferencia": _cob["diferencia"],
                        "cobertura_pct": _cob["cobertura_pct"],
                        "capital_mes": round(_cap_mes, 2),
                        "intereses_mes": round(_int_mes, 2),
                        "intereses_residual_mes": round(_int_res_mes, 2),
                    }]),
                    "Acumulado amort": pd.DataFrame([_acu_am]),
                    "Facturado hasta mes": pd.DataFrame([_acu_hasta]),
                    "Facturado YTD anio": pd.DataFrame([_acu_anio]),
                }
                _buf = excel_con_formato(_hojas)
                st.download_button(
                    f"Descargar Excel — {_periodo} (mes + acumulado)",
                    data=_buf,
                    file_name=f"cartera_{_periodo}_mes_y_acumulado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_mes_acu",
                )
            except Exception as _e_exp:
                st.error(f"No se pudo armar el Excel: {_e_exp}")
                if _det_c is not None and not _det_c.empty:
                    st.download_button(
                        "CSV detalle cartera",
                        data=_det_c.to_csv(index=False).encode("utf-8-sig"),
                        file_name=f"cartera_{_periodo}.csv",
                        mime="text/csv",
                        key="dl_csv_mes",
                    )

    elif menu=="Estado de Cuenta":
        st.title("Estado de Cuenta por Contrato")
        st.caption(f"Empresa activa: **{get_empresa_actual().get('nombre','—')}** — los contratos que ves abajo pertenecen solo a esta empresa.")
        df_todos = obtener()
        if df_todos.empty:
            estado_vacio("Todavía no hay contratos registrados",
                         "En cuanto des de alta el primer contrato desde 'Carga Masiva y Altas', aparecerá aquí su estado de cuenta completo.")
        else:
            # Los IDs se numeran por empresa (ej. "0635-0003"), así que el mismo
            # ID puede repetirse en otra empresa apuntando a OTRO cliente. Por
            # eso la selección de contrato se limpia cada vez que se cambia de
            # empresa (ver limpiar_seleccion_contrato) — nunca debe sobrevivir
            # un ID que no exista en la empresa activa.
            opts_ec = df_todos['ID_Contrato'].tolist()
            # Salto rápido desde otra pantalla ("Ver estado de cuenta — X"):
            # se fuerza la preselección ANTES de crear el widget, escribiendo
            # directamente en su estado — nunca junto con "index=", que es la
            # combinación que causaba que la casilla mostrara un contrato
            # mientras los datos de abajo correspondían a otro distinto.
            pre = st.session_state.pop('ec_contrato', None)
            if pre and pre in opts_ec:
                st.session_state['ec_sel'] = pre
            elif 'ec_sel' not in st.session_state or st.session_state['ec_sel'] not in opts_ec:
                st.session_state['ec_sel'] = opts_ec[0]
            sel_ec  = st.selectbox(
                "Contrato",
                opts_ec,
                format_func=lambda x: f"{x}  —  {df_todos.loc[df_todos['ID_Contrato']==x,'Cliente'].values[0]}",
                key="ec_sel"
            )
            _match = df_todos[df_todos['ID_Contrato']==sel_ec]
            if _match.empty:
                # Salvaguarda: si por cualquier motivo el contrato elegido ya no
                # existe en esta empresa, nunca mostramos datos de otro contrato
                # "por accidente" — se avisa y se fuerza a re-elegir.
                st.error("El contrato seleccionado no pertenece a la empresa activa. Vuelve a elegirlo.")
                st.session_state.pop('ec_sel', None)
                st.stop()
            row = _match.iloc[0]
            assert str(row['ID_Contrato']) == str(sel_ec), "Inconsistencia de selección de contrato detectada."

            with st.container(key=f"tabbox_header_{sel_ec}"):
                nm = int(row.get('Nivel_Morosidad',0) or 0)
                ML_ec = {0:"Al corriente",1:"Atraso",2:"Convenio",3:"Devuelve no paga",4:"Judicial"}
                MC_ec = {0:C['success'],1:C['gold'],2:C['warning'],3:C['accent'],4:"#7A1015"}
                excl_poliza = str(row.get('Fecha_Excl_Poliza','') or '').strip()
                excl_poliza = '' if excl_poliza in ('0','None','nan') else excl_poliza
                motivo_excl = str(row.get('Motivo_Excl_Poliza','') or '').strip()
                motivo_excl = '' if motivo_excl in ('0','None','nan') else motivo_excl
                _fb_hdr = str(row.get('Fecha_Baja','') or '').strip()
                _fb_hdr = '' if _fb_hdr in ('0','None','nan','NaT') else _fb_hdr
                es_baja_hdr = str(row.get('Estatus','')).upper() == 'BAJA'

                badge_mora  = f'<span style="background:{MC_ec[nm]};color:#fff;border-radius:6px;padding:2px 10px;font-size:.8rem;font-weight:700;">Nivel {nm} — {ML_ec[nm]}</span>'
                badge_baja  = (f'<span style="background:#B3261E;color:#fff;border-radius:4px;padding:2px 10px;font-size:.8rem;font-weight:700;">Dado de baja el {_fb_hdr[:10]}</span>'
                               if es_baja_hdr and _fb_hdr else
                               ('<span style="background:#96660C;color:#fff;border-radius:4px;padding:2px 10px;font-size:.8rem;font-weight:700;">Dado de baja (sin fecha capturada)</span>' if es_baja_hdr else ''))
                badge_excl  = (f'<span style="background:#B3261E;color:#fff;border-radius:4px;padding:2px 10px;font-size:.8rem;font-weight:700;">Excluido de pólizas desde {excl_poliza}</span>'
                               if excl_poliza else '')
                badge_avance = ''
                if not es_baja_hdr:
                    _avp = calcular_avance_pago(row)
                    if _avp['estado'] == 'ATRASADO':
                        _txt_falt = ', '.join(str(m) for m in _avp['meses_faltantes'][:8]) + \
                                    (f" y {len(_avp['meses_faltantes'])-8} más" if len(_avp['meses_faltantes']) > 8 else '')
                        badge_avance = (f'<span style="background:#B3261E;color:#fff;border-radius:4px;padding:2px 10px;font-size:.8rem;font-weight:700;" '
                                        f'title="Meses de plazo sin factura conciliada: {_txt_falt}">'
                                        f"Va atrasado {_avp['meses_atraso']} mes(es) — falta{'n' if _avp['meses_atraso']>1 else ''} el/los mes(es) {_txt_falt}</span>")
                    elif _avp['estado'] == 'ADELANTADO':
                        badge_avance = (f'<span style="background:#1E5C4F;color:#fff;border-radius:4px;padding:2px 10px;font-size:.8rem;font-weight:700;">'
                                        f"Va adelantado {_avp['meses_adelanto']} mes(es) — ya tiene facturado hasta el mes {_avp['mes_max_facturado']} de {int(row.get('Plazo',0))}</span>")
                _campo_vencimiento = (
                    f"<div><b>Fecha Vencimiento (original)</b><br>{str(row.get('Fecha_Vencimiento','—'))[:10]}</div>"
                    if es_baja_hdr else
                    f"<div><b>Fecha Vencimiento</b><br>{str(row.get('Fecha_Vencimiento','—'))[:10]}</div>"
                )
                st.markdown(f"""
                <div style="background:#FFFFFF;border:1px solid #DCE0E5;border-radius:6px;padding:18px 22px;
                            margin-bottom:16px;">
                  <div style="font-size:1.3rem;font-weight:700;color:#1E5C4F;margin-bottom:8px;">
                    {sel_ec} &nbsp;·&nbsp; {row['Cliente']}
                  </div>
                  <div style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:10px;">
                    {badge_mora} {badge_baja} {badge_excl} {badge_avance}
                  </div>
                  <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:8px;font-size:.85rem;color:#565E68;">
                    <div><b>Vehículo</b><br>{row.get('Vehiculo','—')}</div>
                    <div><b>Fecha Alta</b><br>{str(row.get('Fecha_Alta','—'))[:10]}</div>
                    {_campo_vencimiento}
                    <div><b>Estatus</b><br>{row.get('Estatus','—')}</div>
                    <div><b>Plazo</b><br>{int(row.get('Plazo',0))} meses</div>
                    <div><b>Valor (s/IVA)</b><br>${float(row.get('Valor_Sin_IVA',0)):,.2f}</div>
                    <div><b>Anticipo</b><br>${float(row.get('Anticipo_Monto',0)):,.2f} ({float(row.get('Anticipo_Pct',0)):.1f}%)</div>
                    <div><b>Comisión</b><br>${float(row.get('Comision_Monto',0)):,.2f}</div>
                    <div><b>Renta (s/IVA)</b><br>${float(row.get('Mensualidad_Sin_IVA',0)):,.2f}</div>
                    <div><b>Residual</b><br>${float(row.get('Residual_Monto',0)):,.2f}</div>
                    <div><b>Tasa anual impl.</b><br>{float(row.get('Tasa_Calculada',0))*1200:.2f}%</div>
                  </div>
                  {('<div style="margin-top:10px;font-size:.8rem;color:#B3261E;"><b>Motivo exclusi\u00f3n:</b> ' + motivo_excl + '</div>') if motivo_excl else ""}
                </div>
                """, unsafe_allow_html=True)

                try:
                    g, m, _pe_lineal = rentabilidad(row); t_ec = tir(row)
                    inv_neta = max(float(row['Valor_Sin_IVA']) - float(row['Anticipo_Monto']), 0.01)
                except: g=m=t_ec=inv_neta=0

                # Punto de equilibrio: mismo criterio que la página de
                # Reportes (renta acumulada vs. inversión neta). No se usa
                # solo la columna Capital de la amortización porque casi
                # nunca cuadra si el contrato tiene residual, ya que ese se
                # recupera aparte al final.
                try:
                    mes_pe = mes_pe_rentas(row)
                except Exception:
                    mes_pe = None

                km1,km2,km3,km4 = st.columns(4)
                km1.metric("Inversión neta", f"${inv_neta:,.2f}")
                km2.metric("Ganancia proy.", f"${g:,.2f}")
                km3.metric("Margen total",   f"{m:.1f}%")
                km4.metric("Punto Equil.",   f"Mes {mes_pe}" if mes_pe else "N/A", help="Mes en que se recupera la inversión")

                # Segunda fila: lectura "de contador" — cobranza y saldo
                fa_hdr = pd.to_datetime(row.get('Fecha_Alta'))
                hoy_ts = pd.Timestamp(hoy_ref())
                estatus_hdr = str(row.get('Estatus','')).upper()
                plazo_hdr = int(row.get('Plazo',0) or 0)
                meses_transcurridos = 0
                saldo_insoluto = None
                utilidad_devengada = None
                if estatus_hdr == 'ACTIVO' and pd.notnull(fa_hdr):
                    meses_transcurridos = max((hoy_ts.year-fa_hdr.year)*12 + (hoy_ts.month-fa_hdr.month), 0)
                    dia_corte = fa_hdr.day
                    prox_pago = (fa_hdr + relativedelta(months=meses_transcurridos)).replace(day=min(dia_corte,28))
                    if prox_pago < hoy_ts:
                        prox_pago = prox_pago + relativedelta(months=1)
                    dias_prox = (prox_pago.normalize() - hoy_ts.normalize()).days
                    txt_prox = prox_pago.strftime('%Y-%m-%d')
                    txt_dias = f"En {dias_prox} días" if dias_prox >= 0 else f"{abs(dias_prox)} días de atraso"
                    try:
                        _inv_h = float(row['Valor_Sin_IVA']) - float(row['Anticipo_Monto'])
                        if _inv_h > 0 and plazo_hdr > 0:
                            _dfa_h, _, _, _, _ = calc_amort(round(_inv_h,4), float(row['Mensualidad_Sin_IVA']),
                                                              float(row['Residual_Monto']), plazo_hdr, float(row['Tasa_Calculada']))
                            _idx = min(meses_transcurridos, plazo_hdr) - 1
                            saldo_insoluto = float(_dfa_h.iloc[_idx]['Saldo']) if _idx >= 0 else _inv_h
                            utilidad_devengada = float(_dfa_h.iloc[:max(_idx+1,0)]['Interes'].sum()) if _idx >= 0 else 0.0
                    except Exception:
                        pass
                else:
                    txt_prox, txt_dias = "—", "Contrato no activo"
                pagos_restantes = max(plazo_hdr - meses_transcurridos, 0) if estatus_hdr=='ACTIVO' else 0

                kf1,kf2,kf3,kf4 = st.columns(4)
                kf1.metric("Próximo pago",          txt_prox)
                kf2.metric("Vence en / atraso",     txt_dias)
                kf3.metric("Pagos realizados / restantes", f"{meses_transcurridos} / {pagos_restantes}")
                kf4.metric("Saldo insoluto (capital)", f"${saldo_insoluto:,.2f}" if saldo_insoluto is not None else "—",
                           help="Capital que aún se debe hoy, según la tabla de amortización — no es lo mismo que la inversión neta inicial.")

                kg1,kg2 = st.columns(2)
                kg1.metric("Utilidad (interés) devengada a la fecha", f"${utilidad_devengada:,.2f}" if utilidad_devengada is not None else "—")
                _avance = (meses_transcurridos/plazo_hdr*100) if (estatus_hdr=='ACTIVO' and plazo_hdr>0) else 0
                kg2.metric("Avance del plazo", f"{_avance:.1f}%")
                if estatus_hdr == 'ACTIVO' and plazo_hdr > 0:
                    st.progress(min(max(_avance/100,0),1.0), text=f"Mes {meses_transcurridos} de {plazo_hdr}")

                # Insignia de Punto de Equilibrio: ¿ya se alcanzó o no?
                if estatus_hdr == 'ACTIVO' and mes_pe:
                    _avance_pe = min(meses_transcurridos/mes_pe, 1.0) if mes_pe else 0
                    if meses_transcurridos >= mes_pe:
                        st.markdown(f"""
                        <div class="milestone-banner">
                          <div class="mb-text">
                            <b>¡Punto de equilibrio alcanzado!</b> — se recuperó el capital invertido en el
                            <b>mes {mes_pe}</b> de {plazo_hdr} (llevas {meses_transcurridos} meses transcurridos).
                            Todo lo que se cobre de aquí en adelante es utilidad sobre el capital. 
                          </div>
                        </div>
                        """, unsafe_allow_html=True)
                    else:
                        _falta = mes_pe - meses_transcurridos
                        st.warning(
                            f"**Aún no llega a su punto de equilibrio** — se alcanza en el **mes {mes_pe}** de "
                            f"{plazo_hdr} (van {meses_transcurridos}, faltan {_falta} pago(s) más)."
                        )
                        st.progress(_avance_pe, text=f"Camino al punto de equilibrio: {_avance_pe*100:.0f}%")
                elif estatus_hdr == 'ACTIVO' and not mes_pe:
                    st.info("Con la renta y plazo pactados, este contrato no recupera el capital dentro del plazo por sí solo (depende del valor residual al final).")

                if estatus_hdr == 'ACTIVO' and pd.notnull(fa_hdr) and plazo_hdr > 0:
                    _dia_corte = fa_hdr.day
                    _prox_pagos = []
                    _m0 = meses_transcurridos
                    for _j in range(3):
                        _mm = min(_m0 + _j, plazo_hdr)
                        _f = (fa_hdr + relativedelta(months=_mm)).replace(day=min(_dia_corte,28))
                        if _f < hoy_ts and _j == 0:
                            _f = _f + relativedelta(months=1)
                        _prox_pagos.append({'Fecha': _f.strftime('%Y-%m-%d'), 'Monto programado': float(row.get('Mensualidad_Sin_IVA',0) or 0)})
                    with st.expander("Próximos 3 pagos programados (para planear cobranza)"):
                        st.dataframe(pd.DataFrame(_prox_pagos), width='stretch', height=140, key=f"df_prox3_{sel_ec}")

            st.markdown("---")

            tab_amort, tab_res, tab_evts, tab_anot, tab_cont = st.tabs([
                "Amortización Leasing",
                "Acumulación Residual",
                "Eventos Especiales",
                "Anotaciones",
                "Contabilidad"
            ])

            inv = float(row['Valor_Sin_IVA']) - float(row['Anticipo_Monto'])
            pl  = int(row['Plazo'])
            t   = round(float(row['Tasa_Calculada']), 8)
            r_m = round(float(row['Mensualidad_Sin_IVA']), 4)
            res = round(float(row['Residual_Monto']), 4)
            fa  = pd.to_datetime(row['Fecha_Alta'])

            # Si el contrato se dio de baja ANTES de terminar su plazo original,
            # las tablas de abajo deben cortarse en el mes real de la baja — de lo
            # contrario mostrarían meses de renta que nunca se llegaron a cobrar.
            fecha_baja_ec = None
            _fb_raw = row.get('Fecha_Baja')
            if str(row.get('Estatus','')).upper()=='BAJA' and pd.notnull(_fb_raw) and str(_fb_raw).strip() not in ('','None','NaT','0'):
                try: fecha_baja_ec = pd.to_datetime(_fb_raw)
                except Exception: fecha_baja_ec = None
            mes_corte_ec = None
            if fecha_baja_ec is not None:
                mes_corte_ec = max((fecha_baja_ec.year-fa.year)*12 + (fecha_baja_ec.month-fa.month), 0)
                if mes_corte_ec >= pl:
                    mes_corte_ec = None  # terminó justo en su plazo natural: no hay nada que truncar

            with tab_amort:
                with st.container(key=f"tabbox_amort_{sel_ec}"):
                    if inv > 0:
                        dfa, _, _, _, _ = calc_amort(round(inv,4), r_m, res, pl, t)
                        dfa = dfa.copy()
                        dfa['Fecha']     = dfa['Mes'].apply(lambda m: (fa + relativedelta(months=m)).strftime('%Y-%m'))
                        saldos_ini = [round(inv,4)] + list(dfa['Saldo'].iloc[:-1].round(4))
                        dfa.insert(dfa.columns.get_loc('Interes'), 'Saldo_Ini', saldos_ini)
                        dfa.rename(columns={'Saldo':'Saldo_Fin'}, inplace=True)

                        if mes_corte_ec is not None:
                            dfa = dfa[dfa['Mes'] <= mes_corte_ec].copy()
                            st.warning(
                                f"Este contrato se dio de **baja anticipada** el **{fecha_baja_ec.strftime('%Y-%m-%d')}** "
                                f"(mes {mes_corte_ec} de {pl} originalmente pactados). La tabla y la gráfica solo muestran "
                                "hasta ese punto — los meses posteriores nunca se cobraron."
                            )

                        mes_hoy = date.today().strftime('%Y-%m')
                        mes_act = dfa[dfa['Fecha'] <= mes_hoy]
                        mes_prox = dfa[dfa['Fecha'] > mes_hoy]
                        capital_amort = mes_act['Capital'].sum() if not mes_act.empty else 0
                        capital_pend  = mes_prox['Capital'].sum() if not mes_prox.empty else 0
                        saldo_actual  = mes_act['Saldo_Fin'].iloc[-1] if not mes_act.empty else inv

                        pa1,pa2,pa3 = st.columns(3)
                        pa1.metric("Capital amortizado", f"${capital_amort:,.2f}")
                        pa2.metric("Saldo vigente",       f"${saldo_actual:,.4f}")
                        pa3.metric("Capital pendiente",   f"${capital_pend:,.2f}")

                        fig_a = go.Figure()
                        fig_a.add_trace(go.Bar(x=dfa['Fecha'], y=dfa['Capital'], name='Capital', marker_color=C_PASTEL['primary']))
                        fig_a.add_trace(go.Bar(x=dfa['Fecha'], y=dfa['Interes'], name='Interés', marker_color=C_PASTEL['accent']))
                        fig_a.add_trace(go.Scatter(x=dfa['Fecha'], y=dfa['Saldo_Fin'], name='Saldo', mode='lines+markers',
                            line=dict(color=C_PASTEL['success'], width=2.5), yaxis='y2'))
                        
                        if mes_hoy in list(dfa['Fecha'].values):
                            idx_hoy = list(dfa['Fecha'].values).index(mes_hoy)
                            fig_a.add_shape(type='line', x0=idx_hoy-0.5, x1=idx_hoy-0.5, y0=0, y1=1, xref='x', yref='paper',
                                line=dict(color=C['gold'], width=2, dash='dash'))
                            fig_a.add_annotation(x=idx_hoy, y=1.04, xref='x', yref='paper', text='Hoy', showarrow=False, font=dict(color=C['gold'], size=11))
                        
                        if mes_corte_ec is not None and not dfa.empty:
                            idx_baja = len(dfa) - 1
                            fig_a.add_shape(type='line', x0=idx_baja+0.5, x1=idx_baja+0.5, y0=0, y1=1, xref='x', yref='paper',
                                line=dict(color=C['accent'], width=2, dash='dash'))
                            fig_a.add_annotation(x=idx_baja, y=1.04, xref='x', yref='paper', text='Baja', showarrow=False, font=dict(color=C['accent'], size=11))
                        
                        # Fix X-axis to display months cleanly without omitting ticks if possible, and adjust bar width
                        fig_a.update_layout(
                            barmode='stack', 
                            yaxis2=dict(overlaying='y', side='right', showgrid=False),
                            legend=dict(orientation='h', y=1.06, x=0),
                            xaxis=dict(
                                type='category', 
                                tickangle=-45,
                                dtick=1 if len(dfa) <= 36 else 2 # Adapt to plazo
                            ),
                            margin=dict(l=10, r=10, t=30, b=10)
                        )
                        fig_a = sfig(fig_a, h=350) # Make slightly taller to fit labels
                        st.plotly_chart(fig_a, width='stretch', key=f"pc_007_{sel_ec}", use_container_width=True)

                        cols_a = ['Fecha','Mes','Saldo_Ini','Interes','Capital','Saldo_Fin']
                        fmt_a  = {c:'${:,.4f}' for c in ['Saldo_Ini','Interes','Capital','Saldo_Fin']}
                        _styler_a = dfa[cols_a].style.format(fmt_a)
                        if mes_corte_ec is not None:
                            _styler_a = _styler_a.set_properties(
                                subset=pd.IndexSlice[[dfa.index[-1]], :],
                                **{'background-color': '#FAE3E1', 'color': '#8A2019', 'font-weight': '700'}
                            )
                        st.dataframe(_styler_a, width='stretch', height=300, key=f"df_008_{sel_ec}")

                        buf_ea = excel_con_formato({'Amort_Leasing': dfa[cols_a]},
                            currency_cols=['Saldo_Ini','Interes','Capital','Saldo_Fin'])
                        st.download_button("Descargar Excel", buf_ea, f"amort_{sel_ec}.xlsx", key=f"dl_amort_{sel_ec}")
                    else:
                        st.info("Inversión neta cero — sin tabla de amortización.")

            with tab_res:
                with st.container(key=f"tabbox_res_{sel_ec}"):
                    if inv > 0 and res > 0:
                        vpr = vp_res(res, t, pl)
                        dfr = calc_res_amort(round(vpr,4), t, pl).copy()
                        dfr['Fecha'] = dfr['Mes'].apply(lambda m: (fa + relativedelta(months=m)).strftime('%Y-%m'))
                        dfr['Residual_Pactado'] = res
                        dfr['VP_Residual']      = round(vpr, 4)

                        if mes_corte_ec is not None:
                            dfr = dfr[dfr['Mes'] <= mes_corte_ec].copy()
                            st.warning(
                                f"Este contrato se dio de **baja anticipada** el **{fecha_baja_ec.strftime('%Y-%m-%d')}** "
                                f"(mes {mes_corte_ec} de {pl} originalmente pactados). La tabla y la gráfica solo muestran "
                                "hasta ese punto."
                            )

                        pr1,pr2,pr3 = st.columns(3)
                        pr1.metric("VP del Residual",    f"${vpr:,.4f}")
                        pr2.metric("Residual pactado",   f"${res:,.2f}")
                        pr3.metric("Intereses residual", f"${dfr['Interes'].sum():,.4f}")

                        fig_r = go.Figure()
                        fig_r.add_trace(go.Scatter(
                            x=dfr['Fecha'], y=dfr['Saldo_Fin'],
                            name='Saldo acumulado', fill='tozeroy', mode='lines',
                            line=dict(color=C_PASTEL['gold'], width=2.5)
                        ))
                        fig_r.add_hline(y=res, line_dash='dash', line_color=C['accent'],
                            annotation_text=f"Residual pactado ${res:,.2f}")
                        fig_r.add_hline(y=vpr, line_dash='dot', line_color=C['info'],
                            annotation_text=f"VP ${vpr:,.4f}")
                        if mes_corte_ec is not None and not dfr.empty:
                            idx_baja_r = len(dfr) - 1
                            fig_r.add_shape(type='line', x0=idx_baja_r, x1=idx_baja_r,
                                y0=0, y1=1, xref='x', yref='paper',
                                line=dict(color=C['accent'], width=2, dash='dash'))
                            fig_r.add_annotation(x=idx_baja_r, y=1.04, xref='x', yref='paper',
                                text='Baja', showarrow=False, font=dict(color=C['accent'], size=11))
                        fig_r = sfig(fig_r, h=280)
                        st.plotly_chart(fig_r, width='stretch', key=f"pc_008_{sel_ec}")

                        cols_r = ['Fecha','Mes','VP_Residual','Saldo_Ini','Interes','Saldo_Fin','Residual_Pactado']
                        fmt_r  = {c:'${:,.4f}' for c in ['VP_Residual','Saldo_Ini','Interes','Saldo_Fin','Residual_Pactado']}
                        _styler_r = dfr[cols_r].style.format(fmt_r)
                        if mes_corte_ec is not None:
                            _styler_r = _styler_r.set_properties(
                                subset=pd.IndexSlice[[dfr.index[-1]], :],
                                **{'background-color': '#FAE3E1', 'color': '#8A2019', 'font-weight': '700'}
                            )
                        st.dataframe(_styler_r, width='stretch', height=300, key=f"df_009_{sel_ec}")
                    else:
                        st.info("Sin residual o inversión neta cero.")

            with tab_evts:
                with st.container(key=f"tabbox_evts_{sel_ec}"):
                    df_ev_c = obtener_eventos_especiales(id_c=sel_ec)
                    if df_ev_c.empty:
                        st.info("Sin eventos especiales registrados para este contrato.")
                    else:
                        for _, ev in df_ev_c.iterrows():
                            est  = str(ev.get('Estatus_Seguro','') or '')
                            col_e = C['success'] if est=='CONFIRMADO' else (C['accent'] if est=='PENDIENTE' else C['warning'])
                            with st.container(key=f"evtcard_{sel_ec}_{int(ev['id'])}"):
                                st.markdown(f"""
                                <div style="border:1px solid #DCE0E5;border-left:4px solid {col_e};background:#FFFFFF;border-radius:0 4px 4px 0;
                                            padding:10px 14px;margin-bottom:8px;">
                                  <b>{ev['Tipo_Evento']}</b>
                                  <span style="float:right;background:{col_e};color:#fff;border-radius:4px;
                                               padding:1px 8px;font-size:.75rem;">{est}</span><br>
                                  <span style="font-size:.8rem;color:#565E68;">Fecha evento: {ev['Fecha_Evento']}
                                  &nbsp;·&nbsp; Registrado: {ev['Fecha_Registro']}</span><br>
                                  <span style="font-size:.82rem;color:#20242B;">
                                    Valor recuperable: <b>${float(ev.get('Valor_Recuperable',0)):,.2f}</b>
                                    &nbsp;·&nbsp; Reclamado: <b>${float(ev.get('Monto_Reclamado',0)):,.2f}</b>
                                    &nbsp;·&nbsp; Confirmado: <b>${float(ev.get('Monto_Confirmado',0)):,.2f}</b>
                                  </span>
                                  {"<br><span style='font-size:.78rem;color:#8A929C;'>"+str(ev.get('Observaciones',''))+"</span>" if ev.get('Observaciones') else ""}
                                </div>
                                """, unsafe_allow_html=True)

            with tab_anot:
                with st.container(key=f"tabbox_anot_{sel_ec}"):
                    df_an_c = obtener_anotaciones(id_c=sel_ec)
                    TIPO_COLORS_EC = {"General":"#3E6FA6","Ajuste contable":"#B3261E","Nota legal":"#6E5A9C",
                                      "Seguimiento":"#1E5C4F","Alerta":"#96660C","Acuerdo con cliente":"#1C7A4D","Otro":"#8A6D2F"}

                    with st.expander("Agregar anotación", expanded=df_an_c.empty):
                        TIPOS_AN = ["General","Ajuste contable","Nota legal","Seguimiento","Alerta","Acuerdo con cliente","Otro"]
                        ec_an1, ec_an2 = st.columns([3,1])
                        ec_txt  = ec_an1.text_area("Texto", height=70, key="ec_anot_txt", placeholder="Escribe aquí…")
                        ec_tipo = ec_an2.selectbox("Tipo", TIPOS_AN, key="ec_anot_tipo")
                        if st.button("Guardar", key="ec_anot_save"):
                            if ec_txt.strip():
                                agregar_anotacion(sel_ec, ec_txt.strip(), ec_tipo)
                                st.success("Anotación guardada.")
                                st.session_state['_refresh'] = True

                    if df_an_c.empty:
                        st.info("Sin anotaciones para este contrato.")
                    else:
                        df_an_c_s = df_an_c.sort_values('Fecha', ascending=False)
                        for _, an in df_an_c_s.iterrows():
                            tipo_an = str(an.get('Tipo','General') or 'General')
                            col_an  = TIPO_COLORS_EC.get(tipo_an,'#0369A1')
                            confirm_del_ec = (st.session_state.get('anot_confirm_del') == int(an['id']))
                            with st.container(key=f"anotcard_{sel_ec}_{int(an['id'])}"):
                                st.markdown(f"""
                                <div style="border:1px solid #DCE0E5;border-left:4px solid {col_an};background:#FFFFFF;border-radius:0 4px 4px 0;
                                            padding:10px 14px;margin-bottom:6px;">
                                  <span style="font-weight:700;color:{col_an};font-size:.82rem;">{tipo_an}</span>
                                  <span style="float:right;font-size:.75rem;color:#8A929C;">{an['Fecha']}</span><br>
                                  <span style="font-size:.88rem;color:#20242B;white-space:pre-wrap;">{an['Texto']}</span>
                                </div>
                                """, unsafe_allow_html=True)
                                bc1_ec, bc2_ec, _ = st.columns([1,1,5])
                                if confirm_del_ec:
                                    bc1_ec.warning("¿Borrar?")
                                    if bc2_ec.button("Sí", key=f"ec_del_ok_{an['id']}", width='stretch'):
                                        eliminar_anotacion(int(an['id']))
                                        st.session_state['anot_confirm_del'] = None
                                        st.session_state['_refresh'] = True
                                    if _.button("No", key=f"ec_del_no_{an['id']}", width='stretch'):
                                        st.session_state['anot_confirm_del'] = None
                                        st.session_state['_refresh'] = True
                                else:
                                    if bc2_ec.button("Eliminar", key=f"ec_del_{an['id']}", width='stretch'):
                                        st.session_state['anot_confirm_del'] = int(an['id'])
                                        st.session_state['_refresh'] = True

            with tab_cont:
                with st.container(key=f"tabbox_cont_{sel_ec}"):
                    st.markdown('<span class="section-label">Composición financiera del contrato</span>', unsafe_allow_html=True)
                    if inv > 0:
                        _dfa_c, _icp_c, _ilp_c, _r12_c, _rresto_c = calc_amort(round(inv,4), r_m, res, pl, t)
                        _cap_total = float(_dfa_c['Capital'].sum())
                        _int_total = _icp_c + _ilp_c
                        fc1,fc2,fc3,fc4 = st.columns(4)
                        fc1.metric("Capital a recuperar (plazo)", f"${_cap_total:,.2f}")
                        fc2.metric("Interés total del plazo",     f"${_int_total:,.2f}")
                        fc3.metric("Interés primeros 12 meses",   f"${_icp_c:,.2f}")
                        fc4.metric("Interés resto del plazo",     f"${_ilp_c:,.2f}")
                        st.caption("Desglose del valor del contrato entre capital, interés y residual.")
                    else:
                        st.info("No se puede calcular la composición: inversión neta no positiva.")

                    st.markdown("---")
                    st.markdown('<span class="section-label">Cuentas contables asignadas a este contrato</span>', unsafe_allow_html=True)
                    st.caption("Cuentas contables asignadas a este contrato.")
                    cat_ec = cargar_catalogo()
                    filas_cat = []
                    for k_ec, d_ec in cat_ec.items():
                        if k_ec == 'CANCELACION_ANTICIPO':
                            continue
                        cuenta_base = str(d_ec.get('cuenta') or '').strip()
                        if not cuenta_base:
                            continue
                        filas_cat.append({
                            'Clave': k_ec,
                            'Concepto': d_ec['nombre'],
                            'Cuenta Contpaqi (detalle)': fmt17(cuenta_base, sel_ec),
                        })
                    if filas_cat:
                        st.dataframe(pd.DataFrame(filas_cat), width='stretch', height=280, key=f"df_cont_{sel_ec}")
                    else:
                        st.warning("No hay cuentas activas configuradas en el catálogo maestro.")

                    st.markdown("---")
                    st.markdown('<span class="section-label">Otros contratos del mismo cliente</span>', unsafe_allow_html=True)
                    _otros = df_todos[(df_todos['Cliente']==row['Cliente']) & (df_todos['ID_Contrato']!=sel_ec)]
                    if not _otros.empty:
                        st.dataframe(
                            _otros[['ID_Contrato','Vehiculo','Estatus','Plazo','Mensualidad_Sin_IVA','Nivel_Morosidad']]
                                .rename(columns={'Mensualidad_Sin_IVA':'Renta mensual'}),
                            width='stretch', height=180, key=f"df_otros_{sel_ec}"
                        )
                    else:
                        st.caption("Sin otros contratos registrados para este cliente.")

            st.markdown("---")
            if st.button("Exportar estado de cuenta completo (Excel)", width='stretch'):
                _hojas_ec = {}
                _hojas_ec['Datos Generales'] = pd.DataFrame([{
                    'Campo': k, 'Valor': str(row.get(k,''))
                } for k in ['ID_Contrato','Cliente','Vehiculo','Estatus','Fecha_Alta',
                             'Fecha_Vencimiento','Plazo','Valor_Sin_IVA','Anticipo_Monto',
                             'Mensualidad_Sin_IVA','Residual_Monto','Tasa_Calculada',
                             'Nivel_Morosidad','Fecha_Excl_Poliza','Motivo_Excl_Poliza']
                ])
                if inv > 0:
                    dfa_exp, _, _, _, _ = calc_amort(round(inv,4), r_m, res, pl, t)
                    dfa_exp['Fecha'] = dfa_exp['Mes'].apply(lambda m: (fa+relativedelta(months=m)).strftime('%Y-%m'))
                    _hojas_ec['Amort Leasing'] = dfa_exp
                if inv > 0 and res > 0:
                    vpr_e = vp_res(res, t, pl)
                    dfr_e = calc_res_amort(round(vpr_e,4), t, pl)
                    dfr_e['Fecha'] = dfr_e['Mes'].apply(lambda m: (fa+relativedelta(months=m)).strftime('%Y-%m'))
                    _hojas_ec['Acum Residual'] = dfr_e
                df_an_exp = obtener_anotaciones(id_c=sel_ec)
                if not df_an_exp.empty:
                    _hojas_ec['Anotaciones'] = df_an_exp
                df_ev_exp = obtener_eventos_especiales(id_c=sel_ec)
                if not df_ev_exp.empty:
                    _hojas_ec['Eventos Especiales'] = df_ev_exp
                _hojas_ec['Resumen Financiero'] = pd.DataFrame([
                    {'Concepto':'Próximo pago','Monto':txt_prox},
                    {'Concepto':'Pagos realizados','Monto':meses_transcurridos},
                    {'Concepto':'Pagos restantes','Monto':pagos_restantes},
                    {'Concepto':'Saldo insoluto (capital)','Monto':saldo_insoluto if saldo_insoluto is not None else ''},
                    {'Concepto':'Utilidad (interés) devengada a la fecha','Monto':utilidad_devengada if utilidad_devengada is not None else ''},
                ])
                if filas_cat:
                    _hojas_ec['Cuentas Contpaqi'] = pd.DataFrame(filas_cat)

                buf_ec = excel_con_formato(
                    _hojas_ec,
                    currency_cols={
                        'Amort Leasing': ['Saldo_Ini','Interes','Capital','Saldo'],
                        'Acum Residual': ['VP_Residual','Saldo_Ini','Interes','Saldo_Fin','Residual_Pactado'],
                        'Eventos Especiales': ['Valor_Recuperable','Monto_Reclamado','Monto_Confirmado'],
                    },
                )
                st.download_button(
                    "Descargar Excel completo",
                    buf_ec,
                    f"estado_cuenta_{sel_ec}.xlsx",
                    key="ec_download"
                )

            if inv > 0:
                with st.expander("Exportar a PDF", expanded=False):
                    pdf_key = f"pdf_buf_{sel_ec}"
                    if st.button("Generar Documento PDF", width='stretch', key="ec_pdf_gen_btn"):
                        with st.spinner("Creando PDF... (puede tomar un par de segundos)"):
                            dfa_pdf, _, _, _, _ = calc_amort(round(inv,4), r_m, res, pl, t)
                            dfa_pdf = dfa_pdf.copy()
                            dfa_pdf['Fecha'] = dfa_pdf['Mes'].apply(lambda m: (fa + relativedelta(months=m)).strftime('%Y-%m'))
                            saldos_ini_pdf = [round(inv,4)] + list(dfa_pdf['Saldo'].iloc[:-1].round(4))
                            dfa_pdf.insert(dfa_pdf.columns.get_loc('Interes'), 'Saldo_Ini', saldos_ini_pdf)
                            dfa_pdf.rename(columns={'Saldo':'Saldo_Fin'}, inplace=True)
                            # Se genera el PDF completo siempre a peticion del usuario
                            dfr_pdf = None
                            if res > 0:
                                vpr_pdf = vp_res(res, t, pl)
                                dfr_pdf = calc_res_amort(round(vpr_pdf,4), t, pl).copy()
                                dfr_pdf['Fecha'] = dfr_pdf['Mes'].apply(lambda m: (fa + relativedelta(months=m)).strftime('%Y-%m'))
                                dfr_pdf['Residual_Pactado'] = res
                                dfr_pdf['VP_Residual']      = round(vpr_pdf, 4)
                                # Se genera el PDF completo siempre a peticion del usuario
                            _avp_pdf = calcular_avance_pago(row) if str(row.get('Estatus','')).upper() != 'BAJA' else None
                            pdf_buf = pdf_estado_cuenta(
                                row, dfa_pdf, dfr=dfr_pdf, avance=_avp_pdf, mes_corte=mes_corte_ec,
                                fecha_hoy_txt=fecha_larga(hoy_ref())
                            )
                            st.session_state[pdf_key] = pdf_buf

                    if pdf_key in st.session_state:
                        st.download_button(
                            "Descargar PDF Ahora",
                            st.session_state[pdf_key],
                            f"estado_cuenta_{sel_ec}.pdf",
                            mime="application/pdf",
                            key=f"ec_pdf_download_{sel_ec}"
                        )


    elif menu=="Carga Masiva y Altas":
        st.title("Carga Masiva y Altas de Contratos")
        tab_m,tab_man=st.tabs(["Carga Masiva","Alta Manual"])
        with tab_m:
            st.subheader("Importar desde Excel")
            if st.button("Descargar Plantilla"):
                st.download_button("Guardar plantilla",plantilla(),"plantilla.xlsx")
            arch = st.file_uploader("Sube .xlsx o .csv",type=['xlsx','csv'])
            
            hojas_sel = []
            xls = None
            if arch and arch.name.endswith('.xlsx'):
                xls = pd.ExcelFile(arch)
                opciones = xls.sheet_names
                # Preseleccionar julio y agosto 2026 si existen, por requerimiento directo
                defs = [h for h in opciones if "JULIO 2026" in h.upper() or "AGOSTO 2026" in h.upper() or h.upper().strip() == "CARTERA"]
                hojas_sel = st.multiselect("Selecciona las hojas a procesar (dejalo vacio para procesar la primera):", opciones, default=defs)
            
            if arch and st.button("Procesar y Cargar"):
                try:
                    if arch.name.endswith('.xlsx'):
                        if not hojas_sel: hojas_sel = [xls.sheet_names[0]]
                        dfs = []
                        for h in hojas_sel:
                            # Buscar dinamicamente la fila de headers (hasta fila 10)
                            df_raw = pd.read_excel(xls, sheet_name=h, header=None)
                            header_idx = 0
                            for r_idx in range(min(10, len(df_raw))):
                                row_vals = [str(x).upper() for x in df_raw.iloc[r_idx].values]
                                if 'CONTRATO' in row_vals or 'CLIENTE' in row_vals:
                                    header_idx = r_idx
                                    break
                            df_temp = pd.read_excel(xls, sheet_name=h, header=header_idx)
                            
                            def clean_col(c):
                                import unicodedata
                                c = str(c).strip().upper()
                                c = unicodedata.normalize('NFKD', c).encode('ASCII', 'ignore').decode('utf-8')
                                return c
                            df_temp.columns = [clean_col(c) for c in df_temp.columns]
                            
                            col_contrato = next((c for c in df_temp.columns if str(c).strip() == 'CONTRATO'), None)
                            if col_contrato:
                                df_temp = df_temp.dropna(subset=[col_contrato])
                            dfs.append(df_temp)
                        dfu = pd.concat(dfs, ignore_index=True)
                    else:
                        dfu = pd.read_csv(arch)
                        def clean_col(c):
                            import unicodedata
                            c = str(c).strip().upper()
                            c = unicodedata.normalize('NFKD', c).encode('ASCII', 'ignore').decode('utf-8')
                            return c
                        dfu.columns = [clean_col(c) for c in dfu.columns]
                        col_c_dfu = next((c for c in dfu.columns if str(c).strip() == 'CONTRATO'), None)
                        if col_c_dfu:
                            dfu = dfu.dropna(subset=[col_c_dfu])
                        
                    nuevos=act=0; bar=st.progress(0); tot_f=len(dfu)
                    for i,r in dfu.iterrows():
                        try:
                            r = r.fillna('')
                            
                            # Contrato y Anexo
                            c_val = str(r.get('CONTRATO','')).replace('.0','').strip()
                            if not c_val: continue
                            a_val = str(r.get('ANEXO','')).replace('.0','').strip()
                            
                            if '-' in c_val:
                                c_parts = c_val.split('-')
                                id_c = f"{c_parts[0].zfill(4)}-{c_parts[1].zfill(4)}"
                            else:
                                if a_val:
                                    id_c = f"{c_val.zfill(4)}-{a_val.zfill(4)}"
                                else:
                                    id_c = c_val.zfill(4) + "-0001"
                                
                            cliente = str(r.get('CLIENTE','')).strip()
                            
                            # Vehiculo
                            if 'UNIDAD' in r and str(r['UNIDAD']).strip() != '':
                                vehiculo = str(r['UNIDAD']).strip()
                            else:
                                vehiculo = f"{str(r.get('MARCA','')).strip()} {str(r.get('VERSION','')).strip()} {str(r.get('MODELO','')).strip()}".strip()
                            if not vehiculo: vehiculo = "VEHICULO NO ESPECIFICADO"
                            
                            # Fecha Apertura
                            if 'FECHA DE APERTURA' in r and str(r['FECHA DE APERTURA']).strip() != '':
                                fa = pd.to_datetime(r['FECHA DE APERTURA'])
                            else:
                                fa = datetime.today()
                                
                            pl = int(float(r.get('PLAZO', 1)))
                            if pl < 1: pl = 1
                            
                            # Valores financieros (quitar IVA segun regla de negocio)
                            val_raw = r.get('VALOR COTIZACION', r.get('MOI', 0))
                            v_con_iva = float(str(val_raw).replace('$','').replace(',','').strip() or 0)
                            v_sin_iva = v_con_iva / 1.16
                            if v_sin_iva <= 0: raise ValueError("Valor Cotizacion invalido o 0.")
                            
                            renta_raw = r.get('RENTA', r.get('RENTAS', 0))
                            rn_con_iva = float(str(renta_raw).replace('$','').replace(',','').strip() or 0)
                            # Si la columna se llama RENTAS, asumimos que es el total
                            if 'RENTAS' in r and 'RENTA' not in r:
                                rn_con_iva = rn_con_iva / pl if pl > 0 else rn_con_iva
                            rn_sin_iva = rn_con_iva / 1.16
                            
                            def safe_pct(cols_posibles):
                                for col in cols_posibles:
                                    if col in r and str(r[col]).strip() != '':
                                        val = str(r[col]).replace('%','').strip()
                                        try:
                                            v_f = float(val)
                                            return v_f*100 if v_f<=1 and v_f>0 else v_f
                                        except: pass
                                return 0.0
                            
                            pc = safe_pct(['% COMISION', 'COMISION %', 'COMISION'])
                            pa = safe_pct(['% ANTICIPO', 'ANTICIPO %', 'ANTICIPO'])
                            
                            # Residual
                            pr = safe_pct(['VALOR RESIDUAL (%)', 'RESIDUAL %', 'RESIDUAL', '%'])
                            if pr == 0.0 and 'RESIDUAL SIN IVA' in r and str(r['RESIDUAL SIN IVA']).strip() != '':
                                val_res = float(str(r['RESIDUAL SIN IVA']).replace('$','').replace(',','').strip())
                                pr = (val_res / v_sin_iva) * 100 if v_sin_iva > 0 else 0.0
                                
                            est = str(r.get('STATUS','ACTIVO')).strip().upper()
                            fb = pd.to_datetime(r['FECHA DE BAJA']) if 'FECHA DE BAJA' in r and str(r['FECHA DE BAJA']).strip() != '' else None
                            if est == 'BAJA' and fb is None: fb = datetime.today()
                            
                            existe = get_db().execute("SELECT 1 FROM contratos WHERE ID_Contrato=?",(id_c,)).fetchone()
                            
                            anotaciones = ""
                            for extra_col in ['SERIE', 'AGENCIA', 'EMISOR', 'ORIGEN']:
                                if extra_col in r and str(r[extra_col]).strip() != '':
                                    anotaciones += f"{extra_col.capitalize()}: {r[extra_col]}\n"
                            
                            ct = dict(
                                ID_Contrato=id_c, Cliente=cliente, Vehiculo=vehiculo,
                                Fecha_Alta=fa, Fecha_Vencimiento=fa+relativedelta(months=pl),
                                Valor_Sin_IVA=v_sin_iva, Mensualidad_Sin_IVA=rn_sin_iva, Plazo=pl,
                                Comision_Apertura_Pct=pc, Comision_Monto=v_sin_iva*pc/100,
                                Anticipo_Pct=pa, Anticipo_Monto=v_sin_iva*pa/100,
                                Residual_Pct=pr, Residual_Monto=v_sin_iva*pr/100,
                                Tasa_Calculada=0.0, Estatus=est, Fecha_Baja=fb,
                                residual_transferred=0, Nivel_Morosidad=0,
                                Anotaciones=anotaciones.strip()
                            )
                            guardar(ct)
                            if existe: act+=1
                            else: nuevos+=1
                        except Exception as e: st.error(f"Fila {i+1} (Contrato {r.get('CONTRATO','')}): {e}")
                    bar.progress(1.0)
                    st.success(f"Procesamiento listo: {nuevos} contratos nuevos, {act} actualizados.")
                except Exception as e:
                    st.error(f"Error procesando el archivo: {e}")
                    
        with tab_man:
            with st.form("alta"):
                c1,c2=st.columns(2)
                id_c=c1.text_input("ID Contrato (ej. 0472-0003)"); cliente=c2.text_input("Cliente")
                vehiculo=st.text_input("Vehículo"); fa=st.date_input("Fecha de Alta")
                v=st.number_input("Valor sin IVA",min_value=0.0,step=1000.0); renta=st.number_input("Renta mensual sin IVA",min_value=0.0,step=100.0)
                plazo=st.number_input("Plazo (meses)",min_value=1,value=36); pa=st.number_input("% Anticipo",value=30.0,step=1.0)
                pr=st.number_input("% Residual",value=10.0,step=1.0); pc=st.number_input("% Comisión",value=2.0,step=.5)
                est=st.selectbox("Estatus",["ACTIVO","BAJA"])
                nm=st.selectbox("Nivel Morosidad",[0,1,2,3,4],format_func=lambda x:{0:"0-Al corriente",1:"1-Atraso",2:"2-Convenio",3:"3-Devuelve no paga",4:"4-Judicial"}[x])
                fb=st.date_input("Fecha de Baja",value=None) if est=="BAJA" else None
                if st.form_submit_button("Guardar"):
                    if id_c and cliente and v>0:
                        pan=norm_pct(pa); prn=norm_pct(pr); pcn=norm_pct(pc)
                        ct=dict(ID_Contrato=id_c,Cliente=cliente,Vehiculo=vehiculo,
                                Fecha_Alta=pd.to_datetime(fa),Fecha_Vencimiento=pd.to_datetime(fa)+relativedelta(months=int(plazo)),
                                Valor_Sin_IVA=v,Mensualidad_Sin_IVA=renta,Plazo=int(plazo),
                                Comision_Apertura_Pct=pcn,Comision_Monto=v*pcn/100,Anticipo_Pct=pan,Anticipo_Monto=v*pan/100,
                                Residual_Pct=prn,Residual_Monto=v*prn/100,Tasa_Calculada=0.0,Estatus=est,Fecha_Baja=fb,
                                residual_transferred=0,Nivel_Morosidad=nm)
                        guardar(ct); st.success("Guardado.")
                        st.session_state['_refresh'] = True

                        # Revisar facturas para alertar discrepancia de IVA en contrato nuevo
                        f_ex = get_db().execute(
                            "SELECT subtotal FROM facturas WHERE (id_contrato=? OR id_contrato_detectado=?) AND tipo='MENSUAL' AND (cancelada IS NULL OR cancelada=0) ORDER BY fecha_emision DESC LIMIT 1",
                            (id_c.strip(), id_c.strip())
                        ).fetchone()
                        if f_ex and f_ex['subtotal']:
                            subt = float(f_ex['subtotal'])
                            diag = detectar_descuadre_iva_contrato(renta, subt)
                            if diag:
                                st.session_state['alerta_iva_nuevo_contrato'] = {
                                    'id_contrato': id_c.strip(),
                                    'diag': diag,
                                    'subtotal': subt
                                }
                    else: st.error("Completa los campos obligatorios.")

            alerta_c = st.session_state.get('alerta_iva_nuevo_contrato')
            if alerta_c and alerta_c.get('id_contrato'):
                idc_a = alerta_c['id_contrato']
                diag_a = alerta_c['diag']
                subt_a = alerta_c['subtotal']
                st.warning(f"Alerta en Contrato Nuevo {idc_a}: {diag_a['mensaje']}. {diag_a['detalle']}")
                if st.button(f"Corregir renta a ${subt_a:,.2f} y recalcular corrida", key=f"btn_corr_alta_{idc_a}"):
                    ok, msg = corregir_renta_contrato_nuevo(idc_a, subt_a)
                    if ok:
                        st.success(msg)
                        st.session_state.pop('alerta_iva_nuevo_contrato', None)
                        st.session_state['_refresh'] = True
                        st.rerun()
                    else:
                        st.error(msg)

    elif menu=="Editar / Eliminar":
        st.title("Edición y Eliminación de Contratos")
        dfa=obtener()
        if dfa.empty: st.info("Sin contratos.")
        else:
            ids=st.selectbox("Contrato",dfa['ID_Contrato'],key="edit_sel_contrato"); row=dfa[dfa['ID_Contrato']==ids].iloc[0]

            with st.expander(f"Cambiar número de contrato ({ids})"):
                st.caption("Actualiza el número de contrato y sus registros asociados.")
                nuevo_id = st.text_input("Nuevo número (formato NNNN-NNNN)", value=ids, key=f"nuevo_id_{ids}")
                if st.button("Renumerar contrato"):
                    ok, msg = renumerar_contrato(ids, nuevo_id)
                    if ok:
                        st.success(msg)
                    else:
                        st.error(msg)
                    st.session_state['_refresh'] = True

            with st.form("editar"):
                c1,c2=st.columns(2)
                cli=c1.text_input("Cliente",row['Cliente']); veh=c2.text_input("Vehículo",row['Vehiculo'])
                fa=st.date_input("Fecha Alta",row['Fecha_Alta'])
                v=st.number_input("Valor sin IVA",value=float(row['Valor_Sin_IVA']),min_value=0.01); r=st.number_input("Renta",value=float(row['Mensualidad_Sin_IVA']))
                pl=st.number_input("Plazo",value=int(row['Plazo']),min_value=1); pa=st.number_input("% Anticipo",value=float(row['Anticipo_Pct']))
                pr_=st.number_input("% Residual",value=float(row['Residual_Pct']))
                nm=st.selectbox("Nivel Morosidad",[0,1,2,3,4],index=int(row.get('Nivel_Morosidad',0)),
                                 format_func=lambda x:{0:"0-Al corriente",1:"1-Atraso",2:"2-Convenio",3:"3-Devuelve no paga",4:"4-Judicial"}[x])
                fb_a=row['Fecha_Baja'] if pd.notnull(row['Fecha_Baja']) else None
                fb=st.date_input("Fecha de Baja",value=fb_a)
                if st.form_submit_button("Actualizar"):
                    ant=v*pa/100; res_m=v*pr_/100; inv_=v-ant; est_='BAJA' if fb else row['Estatus']
                    upd=row.to_dict(); upd.update(Cliente=cli,Vehiculo=veh,Fecha_Alta=pd.to_datetime(fa),
                        Fecha_Vencimiento=pd.to_datetime(fa)+relativedelta(months=int(pl)),
                        Valor_Sin_IVA=v,Mensualidad_Sin_IVA=r,Plazo=int(pl),Anticipo_Pct=pa,Anticipo_Monto=ant,
                        Residual_Pct=pr_,Residual_Monto=res_m,Tasa_Calculada=calc_tasa(int(pl),r,inv_,res_m),
                        Nivel_Morosidad=nm,Estatus=est_,Fecha_Baja=pd.to_datetime(fb) if fb else None)
                    guardar(upd); st.success("Actualizado.")
                    st.session_state['_refresh'] = True

            st.divider()
            n_fact_rel = get_db().execute("SELECT COUNT(*) FROM facturas WHERE (id_contrato=? OR instr(id_contrato, ?) > 0)", (ids, ids)).fetchone()[0]
            n_anot_rel = get_db().execute("SELECT COUNT(*) FROM anotaciones WHERE ID_Contrato=?", (ids,)).fetchone()[0]
            n_evt_rel  = get_db().execute("SELECT COUNT(*) FROM eventos_especiales WHERE ID_Contrato=?", (ids,)).fetchone()[0]
            if n_fact_rel or n_anot_rel or n_evt_rel:
                st.caption(
                    f"Este contrato tiene {n_fact_rel} factura(s), {n_anot_rel} anotación(es) y "
                    f"{n_evt_rel} evento(s) especial(es) asociados — eliminarlo también los borra a ellos."
                )
            if st.session_state.get('confirmar_borrado_contrato') == ids:
                st.error(f"¿Seguro que quieres eliminar **{ids}** y todo lo que le pertenece? Esto no se puede deshacer.")
                cb1, cb2 = st.columns(2)
                if cb1.button("Sí, eliminar definitivamente", type="primary", key=f"del_ok_{ids}"):
                    borrado = eliminar_contrato_completo(ids)
                    st.cache_data.clear()  # por si acaso: limpia TODO caché de Streamlit, no solo el de contratos
                    if borrado:
                        st.session_state['flash_msg'] = ('success', f"Se eliminó **{ids}** de forma permanente, junto con sus facturas, anotaciones y eventos especiales. Ya no debería aparecer en ninguna otra pantalla.")
                    else:
                        st.session_state['flash_msg'] = ('error', f"No se encontró **{ids}** en la base de datos — es posible que ya se hubiera eliminado antes.")
                    st.session_state['confirmar_borrado_contrato'] = None
                    st.rerun()
                if cb2.button("Cancelar", key=f"del_no_{ids}"):
                    st.session_state['confirmar_borrado_contrato'] = None
                    st.session_state['_refresh'] = True
            else:
                if st.button("Eliminar permanentemente",type="primary"):
                    st.session_state['confirmar_borrado_contrato'] = ids
                    st.session_state['_refresh'] = True

            st.divider()
            st.subheader("Anotaciones de este contrato")
            notas=obtener_anotaciones(ids)
            if notas.empty:
                st.info("Sin anotaciones registradas para este contrato.")
            else:
                for _,n in notas.iterrows():
                    nc1,nc2=st.columns([6,1])
                    nc1.markdown(f"**{n['Fecha']}** · _{n['Tipo']}_  \n{n['Texto']}")
                    if nc2.button("Eliminar",key=f"del_anot_{n['id']}"):
                        eliminar_anotacion(n['id'])
                        st.session_state['_refresh'] = True
            with st.form("nueva_anotacion",clear_on_submit=True):
                ca1,ca2=st.columns([1,3])
                tipo_a=ca1.selectbox("Tipo",["Ajuste contable","Aclaración","Nota general","Otro"])
                texto_a=ca2.text_area("Anotación",height=80)
                if st.form_submit_button("Agregar anotación"):
                    if texto_a.strip():
                        agregar_anotacion(ids,texto_a,tipo_a); st.success("Anotación agregada.")
                        st.session_state['_refresh'] = True
                    else:
                        st.warning("Escribe una anotación antes de guardar.")

    elif menu=="Gestor de Bajas":
        st.title("Gestor de Bajas y Análisis de Cartera")
        st.caption("Los contratos vencidos se dan de baja automáticamente por terminación natural.")
        if st.button("Revisar vencimientos ahora"):
            n_chk=procesar_vencimientos_naturales()
            if n_chk>0: st.success(f"{n_chk} contrato(s) pasaron a BAJA por terminación natural.")
            else: st.info("No hay contratos pendientes de vencimiento. Todo está al día.")
            st.session_state['_refresh'] = True

        act = obtener('ACTIVO')
        bj  = obtener('BAJA')

        tab_reg, tab_ana = st.tabs(["Registrar Bajas", "Análisis de Bajas"])

        with tab_reg:
            MOTIVOS = ["Vencimiento natural","Venta del vehículo","Pago anticipado del cliente",
                       "Incumplimiento / morosidad","Siniestro (pérdida total/parcial)","Robo del activo",
                       "Jurídico / incumplimiento total","Refinanciamiento","Error administrativo","Otro"]

            st.subheader("Baja Masiva")
            sel = st.multiselect(
                "Contratos a dar de baja",
                options=act['ID_Contrato'].tolist() if not act.empty else [],
                key="sel_masiva"
            )
            mc1, mc2, mc3 = st.columns(3)
            fb_masiva  = mc1.date_input("Fecha de baja", value=date.today(), key="fb_masiva")
            tipo_masiva = mc2.selectbox("Tipo de baja",["ANTICIPADA","NATURAL","OTRO"], key="tipo_masiva")
            motivo_masiva = mc3.selectbox("Motivo", MOTIVOS, key="motivo_masiva")

            if sel:
                st.info(f"Se darán de baja **{len(sel)}** contrato(s) con fecha **{fb_masiva.strftime('%d/%m/%Y')}**")
                if st.button("Confirmar Baja Masiva", type="primary", key="btn_masiva"):
                    conn = get_db()
                    fecha_str = fb_masiva.isoformat()
                    exito = 0
                    for id_ in sel:
                        try:
                            conn.execute(
                                """UPDATE contratos
                                   SET Estatus=?,Fecha_Baja=?,Tipo_Baja=?,Motivo_Baja=?
                                   WHERE ID_Contrato=?""",
                                ('BAJA', fecha_str, tipo_masiva, motivo_masiva, id_)
                            )
                            exito += 1
                        except Exception as e:
                            st.error(f"Error en {id_}: {e}")
                    conn.commit()
                    st.success(f"{exito} contratos dados de baja con fecha {fb_masiva.strftime('%d/%m/%Y')}.")
                    st.session_state['_refresh'] = True
            else:
                st.warning("Selecciona al menos un contrato de la lista.")

            st.divider()

            st.subheader("Baja Individual")
            if act.empty:
                st.info("No hay contratos activos.")
            else:
                with st.form("form_baja_individual", clear_on_submit=True):
                    ic1, ic2 = st.columns(2)
                    ib   = ic1.selectbox("Contrato", act['ID_Contrato'].tolist(), key="ib_sel")
                    fb_i = ic2.date_input("Fecha de baja", value=date.today(), key="fb_individual")
                    ic3, ic4 = st.columns(2)
                    tipo_i   = ic3.selectbox("Tipo", ["ANTICIPADA","NATURAL","OTRO"], key="tipo_individual")
                    motivo_i = ic4.selectbox("Motivo", MOTIVOS, key="motivo_individual")
                    obs_i = st.text_input("Observaciones (opcional)", key="obs_individual")
                    if st.form_submit_button("Confirmar Baja Individual"):
                        conn = get_db()
                        conn.execute(
                            """UPDATE contratos
                               SET Estatus=?,Fecha_Baja=?,Tipo_Baja=?,Motivo_Baja=?
                               WHERE ID_Contrato=?""",
                            ('BAJA', fb_i.isoformat(), tipo_i,
                             f"{motivo_i}{' — '+obs_i if obs_i else ''}", ib)
                        )
                        conn.commit()
                        st.success(f"Contrato **{ib}** dado de baja el {fb_i.strftime('%d/%m/%Y')}.")
                        st.session_state['_refresh'] = True

            st.divider()
            st.subheader("Histórico de Bajas")
            if bj.empty:
                st.info("No hay contratos dados de baja.")
            else:
                cols_show=['ID_Contrato','Cliente','Vehiculo','Fecha_Alta','Fecha_Vencimiento',
                           'Fecha_Baja','Tipo_Baja','Motivo_Baja','Valor_Sin_IVA','Mensualidad_Sin_IVA','Plazo']
                cols_show=[c for c in cols_show if c in bj.columns]
                st.markdown('<span class="section-label">Todos los contratos dados de baja</span>',
                            unsafe_allow_html=True)
                st.dataframe(bj[cols_show].sort_values('Fecha_Baja',ascending=False).style.format({
                    'Valor_Sin_IVA':'${:,.2f}','Mensualidad_Sin_IVA':'${:,.2f}'
                }), width='stretch',
        key="df_010")
                buf = excel_con_formato({'Bajas': bj[cols_show].sort_values('Fecha_Baja',ascending=False)},
                    currency_cols=['Valor_Sin_IVA','Mensualidad_Sin_IVA'])
                st.download_button("Descargar historial",buf,"bajas_historico.xlsx")

        with tab_ana:
            if bj.empty:
                st.info("Aún no hay contratos dados de baja. Los análisis aparecerán aquí una vez que registres bajas.")
            else:
                df_b = bj.copy()
                df_b['Dias_Activo'] = (df_b['Fecha_Baja'] - df_b['Fecha_Alta']).dt.days.fillna(0).astype(int)
                df_b['Meses_Activo'] = (df_b['Dias_Activo'] / 30.44).round(1)
                df_b['Pct_Plazo_Cumplido'] = ((df_b['Meses_Activo'] / df_b['Plazo']) * 100).clip(0,100).round(1)
                df_b['Mes_Baja'] = df_b['Fecha_Baja'].dt.to_period('M').astype(str)
                df_b['Año_Baja']  = df_b['Fecha_Baja'].dt.year.astype('Int64').astype(str)
                df_b['Tipo_Baja'] = df_b['Tipo_Baja'].replace('','SIN_CLASIFICAR').fillna('SIN_CLASIFICAR')
                df_b['Meses_Cobrados'] = df_b[['Meses_Activo','Plazo']].min(axis=1).apply(int)
                df_b['Ingresos_Reales'] = df_b['Mensualidad_Sin_IVA'] * df_b['Meses_Cobrados']
                df_b['Ingreso_Esperado'] = df_b['Mensualidad_Sin_IVA'] * df_b['Plazo'] + df_b['Residual_Monto']
                df_b['Ingreso_Perdido']  = (df_b['Ingreso_Esperado'] - df_b['Ingresos_Reales']).clip(lower=0)

                anticipadas = df_b[df_b['Tipo_Baja']=='ANTICIPADA']
                naturales   = df_b[df_b['Tipo_Baja']=='NATURAL']
                tot_baj     = len(df_b)
                pct_ant     = len(anticipadas)/tot_baj*100 if tot_baj else 0
                ing_perd     = anticipadas['Ingreso_Perdido'].sum()

                k1,k2,k3,k4,k5,k6 = st.columns(6)
                k1.metric("Total Bajas",          f"{tot_baj:,}")
                k2.metric("Anticipadas",        f"{len(anticipadas)} ({pct_ant:.0f}%)")
                k3.metric("Naturales",          f"{len(naturales)}")
                k4.metric("Duración media", f"{df_b['Meses_Activo'].mean():.1f} m")
                k5.metric("% Plazo cumplido",      f"{df_b['Pct_Plazo_Cumplido'].mean():.0f}%")
                k6.metric("Ingreso Perdido",    f"${ing_perd:,.0f}")

                st.markdown("---")

                r1c1, r1c2 = st.columns(2)
                with r1c1:
                    tipo_cnt = df_b['Tipo_Baja'].value_counts().reset_index()
                    tipo_cnt.columns=['Tipo','Cantidad']
                    COLOR_TIPOS={'ANTICIPADA':C_PASTEL['accent'],'NATURAL':C_PASTEL['success'],'OTRO':C_PASTEL['warning'],'SIN_CLASIFICAR':'#9AA3AC','SIN_FECHA':C_PASTEL['info']}
                    fig1 = px.pie(tipo_cnt, values='Cantidad', names='Tipo', hole=0.5,
                                  title="Composición de Bajas por Tipo",
                                  color='Tipo', color_discrete_map=COLOR_TIPOS)
                    fig1 = sfig(fig1, h=280)
                    fig1.update_traces(textposition='inside', textinfo='percent+label',
                                       hovertemplate='<b>%{label}</b><br>%{value} contratos (%{percent})')
                    fig1.add_annotation(text=f"<b>{tot_baj}</b><br>bajas", x=.5, y=.5,
                                        showarrow=False, font=dict(size=14,color=C['primary'],family='Segoe UI, Helvetica Neue, Arial, sans-serif'))
                    st.plotly_chart(fig1, width='stretch', key="pc_009")
                    explain("Anticipadas vs Naturales",
                        "Verde: el contrato terminó su plazo normal. Rojo: se canceló antes de tiempo.")

                with r1c2:
                    fig2 = px.bar(df_b.groupby('Tipo_Baja')['Ingreso_Perdido'].sum().reset_index(),
                                  x='Tipo_Baja', y='Ingreso_Perdido', color='Tipo_Baja',
                                  text_auto='.2s',
                                  title="Ingreso Potencial No Cobrado por Tipo de Baja",
                                  color_discrete_map=COLOR_TIPOS,
                                  labels={'Tipo_Baja':'Tipo','Ingreso_Perdido':'Ingreso Perdido (MXN)'})
                    fig2 = sfig(fig2, h=280)
                    fig2.update_layout(showlegend=False)
                    fig2.update_traces(hovertemplate='<b>%{x}</b><br>$%{y:,.2f} no cobrados')
                    st.plotly_chart(fig2, width='stretch', key="pc_010")
                    explain("Impacto económico de cada tipo de baja",
                        "Cuánto ingreso se dejó de cobrar por cada tipo de baja.")

                r2c1, r2c2 = st.columns(2)
                with r2c1:
                    by_mes = df_b.groupby('Mes_Baja').agg(
                        Anticipadas=('Tipo_Baja', lambda x: (x=='ANTICIPADA').sum()),
                        Naturales  =('Tipo_Baja', lambda x: (x=='NATURAL').sum()),
                        Total      =('ID_Contrato','count')
                    ).reset_index().sort_values('Mes_Baja')
                    fig3 = px.bar(by_mes, x='Mes_Baja', y=['Anticipadas','Naturales'],
                                  barmode='stack', title="Bajas por Mes: Anticipadas vs Naturales",
                                  color_discrete_map={'Anticipadas':C['accent'],'Naturales':C['success']},
                                  labels={'value':'Contratos','variable':'Tipo','Mes_Baja':'Mes'})
                    fig3 = sfig(fig3, h=280)
                    fig3.update_traces(hovertemplate='%{x}<br>%{y} contratos')
                    st.plotly_chart(fig3, width='stretch', key="pc_011")
                    explain("Evolución mensual de las bajas",
                        "Cuántas bajas hubo cada mes y de qué tipo.")

                with r2c2:
                    fig4 = px.histogram(df_b, x='Pct_Plazo_Cumplido', nbins=20,
                                        color='Tipo_Baja', barmode='overlay',
                                        color_discrete_map=COLOR_TIPOS,
                                        title="¿En qué punto del plazo ocurren las bajas?",
                                        labels={'Pct_Plazo_Cumplido':'% del plazo cumplido al dar de baja'})
                    fig4 = sfig(fig4, h=280)
                    fig4.update_traces(opacity=0.75)
                    fig4.add_vline(x=df_b['Pct_Plazo_Cumplido'].mean(), line_dash='dash',
                                   line_color=C['primary'],
                                   annotation_text=f"Prom: {df_b['Pct_Plazo_Cumplido'].mean():.0f}%",
                                   annotation_font=dict(color=C['primary'], size=11))
                    st.plotly_chart(fig4, width='stretch', key="pc_012")
                    explain("¿Cuándo en el ciclo de vida ocurre la baja?",
                        "Qué tan pronto, respecto al plazo total, se cancelan los contratos.")

                r3c1, r3c2 = st.columns(2)
                with r3c1:
                    df_b['_sz_b'] = sz(df_b['Mensualidad_Sin_IVA'])
                    fig5 = px.scatter(df_b, x='Meses_Activo', y='Valor_Sin_IVA',
                                      size='_sz_b', color='Tipo_Baja',
                                      color_discrete_map=COLOR_TIPOS,
                                      hover_name='ID_Contrato',
                                      hover_data={'Cliente':True,'_sz_b':False,
                                                  'Pct_Plazo_Cumplido':True,'Motivo_Baja':True},
                                      title="Duración Activa vs Valor del Contrato",
                                      labels={'Meses_Activo':'Meses activo','Valor_Sin_IVA':'Valor (MXN)'})
                    fig5 = sfig(fig5, h=290)
                    fig5.update_traces(hovertemplate='<b>%{hovertext}</b><br>Activo: %{x:.1f}m<br>Valor: $%{y:,.2f}')
                    st.plotly_chart(fig5, width='stretch', key="pc_013")
                    explain("¿Qué contratos duraron más vs su valor?",
                        "Compara el valor del contrato contra cuánto tiempo duró antes de darse de baja.")

                with r3c2:
                    motivos = df_b['Motivo_Baja'].replace('','Sin registrar').value_counts().head(8).reset_index()
                    motivos.columns=['Motivo','Cantidad']
                    fig6 = px.bar(motivos, y='Motivo', x='Cantidad', orientation='h',
                                  color='Cantidad',
                                  color_continuous_scale=[[0,C['light']],[1,C['accent']]],
                                  title="Top 8 Motivos de Baja",
                                  text_auto=True,
                                  labels={'Cantidad':'Contratos','Motivo':'Motivo de baja'})
                    fig6 = sfig(fig6, h=290)
                    fig6.update_layout(yaxis={'categoryorder':'total ascending'}, coloraxis_showscale=False)
                    fig6.update_traces(hovertemplate='<b>%{y}</b><br>%{x} contratos')
                    st.plotly_chart(fig6, width='stretch', key="pc_014")
                    explain("¿Por qué se dan de baja los contratos?",
                        "Los motivos más comunes por los que se cancelan los contratos.")

                r4c1, r4c2 = st.columns(2)
                with r4c1:
                    df_all = obtener()
                    if not df_all.empty:
                        churn_rows=[]
                        for mes_str in sorted(by_mes['Mes_Baja']):
                            try:
                                mes_ts = pd.Timestamp(mes_str + '-01')
                                tot_ese_mes = df_all[(df_all['Fecha_Alta']<=mes_ts)&
                                                      ((df_all['Fecha_Baja'].isna())|(df_all['Fecha_Baja']>=mes_ts))].shape[0]
                                bajas_ese_mes = df_b[df_b['Mes_Baja']==mes_str].shape[0]
                                tasa = (bajas_ese_mes/tot_ese_mes*100) if tot_ese_mes>0 else 0
                                churn_rows.append({'Mes':mes_str,'Bajas':bajas_ese_mes,'Total':tot_ese_mes,'Churn_Pct':round(tasa,2)})
                            except Exception:
                                pass
                        if churn_rows:
                            df_churn=pd.DataFrame(churn_rows)
                            fig7=px.line(df_churn,x='Mes',y='Churn_Pct',markers=True,
                                          title="Tasa de Churn Mensual (%)",
                                          color_discrete_sequence=[C['accent']],
                                          labels={'Churn_Pct':'Churn (%)'})
                            fig7=sfig(fig7,h=280)
                            fig7.update_traces(line_width=2.5,marker_size=8,
                                               hovertemplate='%{x}<br>Churn: %{y:.2f}%')
                            fig7.add_hrect(y0=0,y1=2,fillcolor="green",opacity=0.06,line_width=0,annotation_text="Zona saludable (≤2%)")
                            fig7.add_hrect(y0=2,y1=5,fillcolor="orange",opacity=0.06,line_width=0)
                            st.plotly_chart(fig7,width='stretch', key="pc_015")
                            explain("Tasa de cancelación (churn) por mes",
                        "Porcentaje de contratos que se cancela cada mes. Mientras más bajo, mejor.")

                with r4c2:
                    if len(df_b['Tipo_Baja'].unique()) >= 2:
                        fig8=px.violin(df_b,x='Tipo_Baja',y='Pct_Plazo_Cumplido',
                                        color='Tipo_Baja',box=True,points='all',
                                        color_discrete_map=COLOR_TIPOS,
                                        title="Distribución del % de Plazo Cumplido por Tipo",
                                        labels={'Tipo_Baja':'Tipo','Pct_Plazo_Cumplido':'% plazo cumplido'})
                        fig8=sfig(fig8,h=280)
                        fig8.update_traces(hovertemplate='<b>%{x}</b><br>%{y:.1f}% cumplido')
                        st.plotly_chart(fig8,width='stretch', key="pc_016")
                        explain("Comparativa de cuánto se cumplió según tipo de baja",
                        "Compara qué tanto del plazo se cumplió antes de cada tipo de baja.")
                    else:
                        fig8=px.box(df_b,x='Tipo_Baja',y='Pct_Plazo_Cumplido',
                                     color='Tipo_Baja',color_discrete_map=COLOR_TIPOS,points='all',
                                     title="% de Plazo Cumplido al dar de Baja")
                        st.plotly_chart(sfig(fig8,h=280),width='stretch', key="pc_017")

                st.markdown("---")
                st.markdown('<span class="section-label">Detalle Completo de Contratos Dados de Baja</span>',
                            unsafe_allow_html=True)
                df_tabla = df_b[['ID_Contrato','Cliente','Vehiculo','Fecha_Alta','Fecha_Baja',
                                   'Tipo_Baja','Motivo_Baja','Meses_Activo','Pct_Plazo_Cumplido',
                                   'Valor_Sin_IVA','Mensualidad_Sin_IVA','Plazo',
                                   'Ingresos_Reales','Ingreso_Esperado','Ingreso_Perdido']].copy()
                st.dataframe(
                    df_tabla.sort_values('Fecha_Baja',ascending=False).style.format({
                        'Valor_Sin_IVA':'${:,.2f}','Mensualidad_Sin_IVA':'${:,.2f}',
                        'Ingresos_Reales':'${:,.2f}','Ingreso_Esperado':'${:,.2f}',
                        'Ingreso_Perdido':'${:,.2f}','Meses_Activo':'{:.1f}','Pct_Plazo_Cumplido':'{:.1f}%'
                    }).background_gradient(subset=['Ingreso_Perdido'],cmap=CMAP_CORAL),
                    width='stretch',
        key="df_011")
                buf = excel_con_formato({'Analisis_Bajas': df_tabla},
                    currency_cols=['Valor_Sin_IVA','Mensualidad_Sin_IVA','Ingresos_Reales','Ingreso_Esperado','Ingreso_Perdido'],
                    pct_cols=['Pct_Plazo_Cumplido'])
                st.download_button("Descargar análisis completo",buf,"analisis_bajas.xlsx")

    elif menu=="Eventos Especiales":
        st.title("Siniestros, Robos y Procesos Jurídicos")
        st.caption("Registra siniestros o procesos legales y calcula sus ajustes contables.")
        df_all = obtener()
        tab_nuevo, tab_seg, tab_calc = st.tabs(["Registrar evento", "Estatus de seguro", "Calcular ajuste"])

        with tab_nuevo:
            if df_all.empty:
                st.info("Sin contratos.")
            else:
                c1,c2 = st.columns(2)
                id_sel = c1.selectbox("Contrato", df_all['ID_Contrato'])
                tipo_ev = c2.selectbox("Tipo de evento", ["SINIESTRO","ROBO","JURIDICO"],
                            format_func=lambda x:{"SINIESTRO":"Siniestro (daño)","ROBO":"Robo / pérdida total","JURIDICO":"Proceso jurídico"}[x])
                c3,c4 = st.columns(2)
                fecha_ev = c3.date_input("Fecha del evento", value=date.today())
                val_rec = c4.number_input("Valor recuperable estimado (salvamento)", min_value=0.0, value=0.0, step=1000.0)
                monto_recl = st.number_input("Monto reclamado a la aseguradora (si aplica)", min_value=0.0, value=0.0, step=1000.0)
                obs = st.text_area("Observaciones")
                st.markdown("---")
                TIPOS_EXCL_EVT = ["SINIESTRO","ROBO","JURIDICO"]
                excluir_de_poliza = tipo_ev in TIPOS_EXCL_EVT
                if excluir_de_poliza:
                    st.caption(f"Los eventos de tipo {tipo_ev} excluyen el contrato de las pólizas de parcialidades.")
                    fecha_excl_evt_reg = st.date_input("Fecha de inicio de exclusión", value=fecha_ev, key="fexcl_evt_reg")
                else:
                    fecha_excl_evt_reg = None
                if st.button("Registrar evento"):
                    registrar_evento_especial(id_sel, tipo_ev, fecha_ev.isoformat(), val_rec, monto_recl, obs)
                    if excluir_de_poliza and fecha_excl_evt_reg:
                        get_db().execute(
                            "UPDATE contratos SET Fecha_Excl_Poliza=?, Motivo_Excl_Poliza=? WHERE ID_Contrato=?",
                            (fecha_excl_evt_reg.isoformat(), f"Evento especial: {tipo_ev}", id_sel))
                        get_db().commit()
                        _tocar_datos()
                        agregar_anotacion(id_sel, f"Contrato excluido de pólizas de parcialidades por evento {tipo_ev} a partir de {fecha_excl_evt_reg}.", "Alerta")
                    st.success("Evento registrado." + (f" Contrato excluido de pólizas desde {fecha_excl_evt_reg}." if excluir_de_poliza and fecha_excl_evt_reg else "") + " Ve a la pestaña 'Calcular ajuste' para ver el monto a reconocer.")
                    st.session_state['_refresh'] = True

            st.divider()
            st.caption("Eventos registrados:")
            dfe_prev = obtener_eventos_especiales()
            if dfe_prev.empty:
                st.info("Aún no hay eventos registrados.")
            else:
                st.dataframe(dfe_prev[['id','ID_Contrato','Tipo_Evento','Fecha_Evento','Estatus_Seguro',
                            'Monto_Reclamado','Monto_Confirmado','Ajuste_Registrado']], width='stretch',
        key="df_012")

        with tab_seg:
            dfe = obtener_eventos_especiales()
            if dfe.empty:
                st.info("Registra un evento primero en la pestaña 'Registrar evento'.")
            else:
                st.caption("Hasta confirmar el reembolso de la aseguradora, el cálculo no incluirá ingresos por seguro.")
                ev_sel = st.selectbox("Evento", dfe['id'],
                            format_func=lambda i: f"#{i} — {dfe[dfe['id']==i]['ID_Contrato'].values[0]} ({dfe[dfe['id']==i]['Tipo_Evento'].values[0]}) — estatus actual: {dfe[dfe['id']==i]['Estatus_Seguro'].values[0]}")
                ec1,ec2,ec3 = st.columns(3)
                nuevo_est = ec1.selectbox("Nuevo estatus", ["PENDIENTE","CONFIRMADO","RECHAZADO","COBRADO"])
                monto_conf = ec2.number_input("Monto confirmado por la aseguradora", min_value=0.0, value=0.0, step=1000.0)
                fecha_conf = ec3.date_input("Fecha de confirmación", value=date.today())
                if st.button("Actualizar estatus de seguro"):
                    actualizar_estatus_seguro(ev_sel, nuevo_est, monto_conf, fecha_conf.isoformat())
                    st.success("Estatus actualizado.")
                    st.session_state['_refresh'] = True

        with tab_calc:
            dfe = obtener_eventos_especiales()
            if dfe.empty:
                st.info("Registra un evento primero en la pestaña 'Registrar evento'.")
            else:
                ev_sel2 = st.selectbox("Selecciona el evento a calcular", dfe['id'],
                            format_func=lambda i: f"#{i} — {dfe[dfe['id']==i]['ID_Contrato'].values[0]} ({dfe[dfe['id']==i]['Tipo_Evento'].values[0]}) — {dfe[dfe['id']==i]['Fecha_Evento'].values[0]}",
                            key="sel_calc_evento")
                ev_row = dfe[dfe['id']==ev_sel2].iloc[0]
                con_match = df_all[df_all['ID_Contrato']==ev_row['ID_Contrato']]
                if con_match.empty:
                    st.error("El contrato de este evento ya no existe en la base.")
                else:
                    con_row = con_match.iloc[0]
                    seguro_confirmado = ev_row['Monto_Confirmado'] if ev_row['Estatus_Seguro'] in ('CONFIRMADO','COBRADO') else 0.0
                    calc = calc_ajuste_evento_especial(con_row, ev_row['Fecha_Evento'], ev_row['Valor_Recuperable'], seguro_confirmado)

                    if ev_row['Estatus_Seguro']=='PENDIENTE':
                        st.warning("El reembolso de la aseguradora sigue pendiente de confirmación, así que este cálculo no incluye ningún ingreso por seguro.")

                    k1,k2,k3 = st.columns(3)
                    k1.metric("Saldo contable a la fecha del evento", f"${calc['saldo_evento']:,.2f}")
                    k2.metric("Saldo que muestra el sistema hoy", f"${calc['saldo_hoy_sistema']:,.2f}")
                    k3.metric("Meses con registro indebido", calc['mes_hoy']-calc['mes_evento'])

                    if calc['es_retroactivo']:
                        st.info(f"Evento registrado con retraso ({calc['mes_hoy']-calc['mes_evento']} meses). Se revertirán las amortizaciones e intereses generados en ese período.")

                    if abs(calc['loss_plug_anterior'])>0.001:
                        st.warning(f"${calc['loss_plug_anterior']:,.2f} corresponde a un ejercicio fiscal cerrado. Debe registrarse contra utilidades acumuladas con apoyo de tu contador.")

                    with st.expander("Ver detalle mes por mes del ajuste"):
                        if not calc['detalle'].empty:
                            st.dataframe(calc['detalle'].style.format({'Interes_Leasing':'${:,.2f}','Capital':'${:,.2f}',
                                'Interes_Residual':'${:,.2f}','Total':'${:,.2f}'}), width='stretch',
        key="df_013")
                        else:
                            st.caption("No hay meses pendientes de revertir: el ajuste se está registrando en el mismo mes del evento.")

                    with st.expander("Ver saldos por cuenta que se van a cancelar"):
                        sl=calc['saldos_hoy']
                        filas_sl=[("115 — CXC Corto Plazo","CXC_CP"),("125 — CXC Largo Plazo","CXC_LP"),
                                  ("126 — VP Residual LP","RESIDUAL"),("114-01 — VP Residual CP","RESIDUAL_CP"),
                                  ("126 — Interés Residual acum. LP","CXC_RESIDUAL_INTERESES"),
                                  ("114-01 — Interés Residual acum. CP","CXC_RESIDUAL_INTERESES_CP"),
                                  ("208 — Interés por devengar CP","INT_CP"),("228 — Interés por devengar LP","INT_LP"),
                                  ("204 — Comisión apertura pendiente","COMISION_PASIVO")]
                        st.dataframe(pd.DataFrame([{'Cuenta':n,'Saldo a cancelar':sl[k]} for n,k in filas_sl])
                                     .style.format({'Saldo a cancelar':'${:,.2f}'}), width='stretch',
        key="df_014")
                        st.caption("Saldos actuales del contrato que serán cancelados por el ajuste.")

                    st.divider()
                    st.subheader("Póliza de ajuste sugerida")
                    pol_df = pol_ajuste_evento(con_row, calc, CAT, ev_row['Tipo_Evento'])
                    if pol_df.empty:
                        st.info("No hay movimientos que registrar (pérdida neta = 0).")
                    else:
                        tc=pol_df['Cargo'].sum(); ta=pol_df['Abono'].sum()
                        m1,m2,m3 = st.columns(3)
                        m1.metric("Total Cargos", f"${tc:,.2f}"); m2.metric("Total Abonos", f"${ta:,.2f}"); m3.metric("Diferencia", f"${ta-tc:,.2f}")
                        st.dataframe(pol_df, width='stretch', key="df_015")
                        buf = excel_con_formato({'Ajuste': pol_df}, currency_cols=['Cargo','Abono'])
                        cd1,cd2,cd3 = st.columns(3)
                        cd1.download_button("Excel", buf, f"ajuste_{ev_row['ID_Contrato']}.xlsx")
                        cd2.download_button("PDF", pdf_poliza(pol_df, f"Ajuste por {ev_row['Tipo_Evento']}", ev_row['ID_Contrato']), f"ajuste_{ev_row['ID_Contrato']}.pdf")
                        if cd3.button("Marcar como registrado en contabilidad"):
                            marcar_ajuste_registrado(ev_sel2)
                            agregar_anotacion(ev_row['ID_Contrato'], f"Póliza de ajuste por {ev_row['Tipo_Evento']} registrada en contabilidad por ${tc:,.2f}.", "Ajuste contable")
                            st.success("Marcado como registrado.")
                            st.session_state['_refresh'] = True

    elif menu=="Anotaciones":
        st.title("Anotaciones")

        if 'anot_edit_id'    not in st.session_state: st.session_state['anot_edit_id']    = None
        if 'anot_edit_texto' not in st.session_state: st.session_state['anot_edit_texto'] = ""
        if 'anot_edit_tipo'  not in st.session_state: st.session_state['anot_edit_tipo']  = "General"
        if 'anot_confirm_del'not in st.session_state: st.session_state['anot_confirm_del']= None

        TIPOS_ANOT = ["General","Ajuste contable","Nota legal","Seguimiento","Alerta","Acuerdo con cliente","Otro"]
        TIPO_COLORS= {"General":"#3E6FA6","Ajuste contable":"#B3261E","Nota legal":"#6E5A9C",
                      "Seguimiento":"#1E5C4F","Alerta":"#96660C","Acuerdo con cliente":"#1C7A4D","Otro":"#8A6D2F"}

        df_anot = obtener_anotaciones()
        dfc     = obtener()[['ID_Contrato','Cliente','Estatus']]

        col_form, col_kpi = st.columns([2, 1], gap="large")

        with col_form:
            st.markdown("#### Nueva anotación")
            df_contratos = obtener()[['ID_Contrato','Cliente']]
            opts_c = df_contratos.apply(lambda r: f"{r['ID_Contrato']} — {r['Cliente']}", axis=1).tolist()
            sel_nueva = st.selectbox("Contrato", opts_c, key="anot_new_contrato")
            id_nueva  = sel_nueva.split(" — ")[0] if sel_nueva else ""
            nc1, nc2  = st.columns([2,1])
            texto_nueva = nc1.text_area("Texto", height=90, placeholder="Escribe la anotación…", key="anot_new_texto")
            tipo_nueva  = nc2.selectbox("Tipo", TIPOS_ANOT, key="anot_new_tipo")
            if st.button("Guardar anotación", width='stretch'):
                if texto_nueva.strip() and id_nueva:
                    agregar_anotacion(id_nueva, texto_nueva.strip(), tipo_nueva)
                    st.success("Anotación guardada.")
                    st.session_state['_refresh'] = True
                else:
                    st.warning("Completa el texto antes de guardar.")

        with col_kpi:
            st.markdown("#### Resumen")
            if not df_anot.empty:
                df_anot_k = df_anot.merge(dfc, on='ID_Contrato', how='left')
                total_anot   = len(df_anot_k)
                contratos_k  = df_anot_k['ID_Contrato'].nunique()
                alertas_k    = int((df_anot_k['Tipo']=='Alerta').sum())
                ajustes_k    = int((df_anot_k['Tipo']=='Ajuste contable').sum())
                st.metric("Total anotaciones",   total_anot)
                st.metric("Contratos con notas", contratos_k)
                a1,a2 = st.columns(2)
                a1.metric("Alertas",   alertas_k)
                a2.metric("Ajustes",   ajustes_k)
            else:
                st.info("Sin anotaciones aún.")

        st.divider()

        if not df_anot.empty:
            df_anot = df_anot.merge(dfc, on='ID_Contrato', how='left')
            st.markdown("#### Filtrar y gestionar")
            fa1,fa2,fa3,fa4 = st.columns(4)
            clientes_op = ["Todos"] + sorted(df_anot['Cliente'].dropna().unique().tolist())
            cli_f    = fa1.selectbox("Cliente",   clientes_op, key="af_cli")
            tipos_op = ["Todos"] + sorted(df_anot['Tipo'].dropna().unique().tolist())
            tipo_f   = fa2.selectbox("Tipo",      tipos_op,    key="af_tipo")
            busq_f   = fa3.text_input("Buscar texto", placeholder="palabra clave…", key="af_busq")
            orden_f  = fa4.selectbox("Ordenar por",   ["Más reciente","Más antigua","Contrato A-Z","Tipo"], key="af_orden")

            dfx = df_anot.copy()
            if cli_f  != "Todos": dfx = dfx[dfx['Cliente']==cli_f]
            if tipo_f != "Todos": dfx = dfx[dfx['Tipo']==tipo_f]
            if busq_f.strip():    dfx = dfx[dfx['Texto'].str.contains(busq_f.strip(), case=False, na=False) |
                                            dfx['ID_Contrato'].str.contains(busq_f.strip(), case=False, na=False)]
            if orden_f == "Más reciente":  dfx = dfx.sort_values('Fecha', ascending=False)
            elif orden_f == "Más antigua": dfx = dfx.sort_values('Fecha', ascending=True)
            elif orden_f == "Contrato A-Z":dfx = dfx.sort_values('ID_Contrato')
            elif orden_f == "Tipo":        dfx = dfx.sort_values('Tipo')

            st.caption(f"Mostrando **{len(dfx)}** anotación(es)")

            for _, row in dfx.iterrows():
                rid      = int(row['id'])
                tipo_r   = str(row.get('Tipo','General') or 'General')
                color    = TIPO_COLORS.get(tipo_r,'#0369A1')
                cliente  = str(row.get('Cliente','') or '')
                estatus  = str(row.get('Estatus','') or '')
                est_col  = "#1C7A4D" if estatus=='ACTIVO' else "#B3261E"

                editing  = (st.session_state['anot_edit_id'] == rid)
                confirm_del = (st.session_state['anot_confirm_del'] == rid)

                with st.container(key=f"anotpage_card_{rid}"):
                    st.markdown(f"""
                    <div style="border:1px solid #DCE0E5;border-left:4px solid {color};background:#FFFFFF;border-radius:0 4px 4px 0;
                                padding:12px 16px;margin-bottom:8px;">
                      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;">
                        <span style="font-weight:700;color:{color};font-size:.92rem;">{tipo_r}</span>
                        <span style="font-size:.78rem;color:#8A929C;">{row['Fecha']}</span>
                      </div>
                      <div style="font-size:.82rem;color:#565E68;margin-bottom:4px;">
                        <b>{row['ID_Contrato']}</b> · {cliente}
                        <span style="background:{est_col};color:#fff;border-radius:4px;
                                     padding:1px 7px;font-size:.72rem;margin-left:6px;">{estatus}</span>
                      </div>
                      {"" if editing else f'<div style="color:#20242B;font-size:.9rem;white-space:pre-wrap;">{row["Texto"]}</div>'}
                    </div>
                    """, unsafe_allow_html=True)

                    if editing:
                        te1, te2 = st.columns([3,1])
                        nuevo_txt  = te1.text_area("Editar texto", value=st.session_state['anot_edit_texto'],
                                                   height=80, key=f"edit_txt_{rid}")
                        nuevo_tipo = te2.selectbox("Tipo", TIPOS_ANOT,
                                                   index=TIPOS_ANOT.index(st.session_state['anot_edit_tipo'])
                                                   if st.session_state['anot_edit_tipo'] in TIPOS_ANOT else 0,
                                                   key=f"edit_tipo_{rid}")
                        ea1,ea2,ea3 = st.columns([1,1,3])
                        if ea1.button("Guardar", key=f"save_{rid}", width='stretch'):
                            if nuevo_txt.strip():
                                get_db().execute("UPDATE anotaciones SET Texto=?, Tipo=? WHERE id=?",
                                                 (nuevo_txt.strip(), nuevo_tipo, rid))
                                get_db().commit()
                                st.session_state['anot_edit_id'] = None
                                st.success("Anotación actualizada.")
                                st.session_state['_refresh'] = True
                            else:
                                st.warning("El texto no puede estar vacío.")
                        if ea2.button("Cancelar", key=f"cancel_{rid}", width='stretch'):
                            st.session_state['anot_edit_id'] = None
                            st.session_state['_refresh'] = True
                    else:
                        bc1, bc2, bc3 = st.columns([1,1,6])
                        if bc1.button("Editar", key=f"ed_{rid}", width='stretch'):
                            st.session_state['anot_edit_id']    = rid
                            st.session_state['anot_edit_texto'] = str(row['Texto'])
                            st.session_state['anot_edit_tipo']  = str(row.get('Tipo','General') or 'General')
                            st.session_state['anot_confirm_del']= None
                            st.session_state['_refresh'] = True
                        if confirm_del:
                            bc2.warning("¿Eliminar?")
                            cc1,cc2 = st.columns(2)
                            if cc1.button("Sí, borrar", key=f"confirm_yes_{rid}", width='stretch'):
                                eliminar_anotacion(rid)
                                st.session_state['anot_confirm_del'] = None
                                st.success("Eliminada.")
                                st.session_state['_refresh'] = True
                            if cc2.button("No", key=f"confirm_no_{rid}", width='stretch'):
                                st.session_state['anot_confirm_del'] = None
                                st.session_state['_refresh'] = True
                        else:
                            if bc2.button("Eliminar", key=f"del_{rid}", width='stretch'):
                                st.session_state['anot_confirm_del'] = rid
                                st.session_state['anot_edit_id']     = None
                                st.session_state['_refresh'] = True

            st.divider()
            if len(dfx) > 1:
                dist = dfx['Tipo'].value_counts().reset_index()
                dist.columns = ['Tipo','Cantidad']
                dist['Color'] = dist['Tipo'].map(TIPO_COLORS).fillna('#aaa')
                fig_d = px.bar(dist, x='Cantidad', y='Tipo', orientation='h',
                               color='Tipo', color_discrete_map=TIPO_COLORS,
                               title="Distribución por tipo de anotación", text='Cantidad')
                fig_d.update_traces(textposition='outside')
                fig_d = sfig(fig_d, h=max(180, len(dist)*44))
                st.plotly_chart(fig_d, width='stretch', key="pc_018")

            cols_export = ['Fecha','ID_Contrato','Cliente','Estatus','Tipo','Texto']
            cols_export = [c for c in cols_export if c in dfx.columns]
            buf_a = excel_con_formato({'Anotaciones': dfx[cols_export]})
            st.download_button("Exportar Excel", buf_a, "anotaciones.xlsx")
        else:
            st.info("Aún no hay anotaciones. Usa el formulario de arriba para agregar la primera.")

    elif menu=="Pólizas Contables":
        st.title("Generador de Pólizas Contables")
        df_c=obtener(); opts=["Todos"]+df_c['ID_Contrato'].tolist()
        tipo=st.radio("Tipo",["Inicial","Parcialidad (mensual)","Comisión por apertura","Reglas Personalizadas"])

        if tipo=="Reglas Personalizadas":
            st.caption("Configura manualmente las cuentas de cargo y abono para armar tus pólizas.")
            with st.expander("Configurar reglas", expanded=obtener_reglas_poliza_custom().empty):
                df_reglas = obtener_reglas_poliza_custom()
                if not df_reglas.empty:
                    for _, rg in df_reglas.iterrows():
                        rc1, rc2, rc3, rc4, rc5 = st.columns([2.2, 1.3, 1, 0.8, 0.8])
                        rc1.markdown(f"**{CONCEPTOS_POLIZA_CUSTOM.get(rg['concepto'], rg['concepto'])}**")
                        rc2.markdown(f"Cuenta: `{rg['cuenta_base']}`")
                        rc3.markdown(f"{'Cargo' if rg['tipo']=='CARGO' else 'Abono'}")
                        _activa_nueva = rc4.checkbox("Activa", value=bool(rg['activa']), key=f"rpc_act_{rg['id']}")
                        if _activa_nueva != bool(rg['activa']):
                            activar_regla_poliza_custom(int(rg['id']), _activa_nueva)
                            st.session_state['_refresh'] = True
                        if rc5.button("Borrar", key=f"rpc_del_{rg['id']}"):
                            eliminar_regla_poliza_custom(int(rg['id']))
                            st.session_state['_refresh'] = True
                    st.divider()
                else:
                    st.info("Todavía no tienes reglas configuradas — agrega la primera abajo.")

                st.markdown("**Agregar regla nueva**")
                nc1, nc2, nc3, nc4 = st.columns([2, 1.3, 1.6, 1])
                _concepto_nuevo = nc1.selectbox(
                    "Concepto", list(CONCEPTOS_POLIZA_CUSTOM.keys()),
                    format_func=lambda k: CONCEPTOS_POLIZA_CUSTOM[k], key="rpc_nuevo_concepto"
                )
                _cuenta_nueva = nc2.text_input("Cuenta base (ej. 110-00)", key="rpc_nueva_cuenta")
                _nombre_nuevo = nc3.text_input("Nombre de la cuenta (opcional)", key="rpc_nuevo_nombre")
                _tipo_nuevo = nc4.selectbox("Cargo/Abono", ["CARGO","ABONO"], key="rpc_nuevo_tipo")
                if st.button("Agregar regla"):
                    if not _cuenta_nueva.strip():
                        st.error("Captura la cuenta base.")
                    else:
                        guardar_regla_poliza_custom(_concepto_nuevo, _cuenta_nueva.strip(), _nombre_nuevo.strip(), _tipo_nuevo)
                        st.success("Regla agregada.")
                        st.session_state['_refresh'] = True

            st.divider()
            sel_c_rp = st.selectbox("Filtrar contrato", opts, key="rpc_sel_contrato")
            crp1, crp2 = st.columns(2)
            mes_rp = crp1.selectbox("Mes", range(1,13), index=datetime.now().month-1, format_func=lambda m:MN[m-1], key="rpc_mes")
            anio_rp = crp2.number_input("Año", value=datetime.now().year, key="rpc_anio")

            if st.button("Generar Póliza Personalizada", type="primary"):
                df_f_rp = df_c if sel_c_rp=="Todos" else df_c[df_c['ID_Contrato']==sel_c_rp]
                with st.spinner("Generando con tus reglas…"):
                    df_fin_rp = generar_poliza_custom(df_f_rp, mes_rp, int(anio_rp))
                if df_fin_rp.empty:
                    st.info("Sin movimientos para el período con las reglas activas — revisa que tengas al menos una regla activa y que el período caiga dentro del plazo de algún contrato.")
                else:
                    tc_rp, ta_rp = df_fin_rp['Cargo'].sum(), df_fin_rp['Abono'].sum()
                    rp1, rp2, rp3 = st.columns(3)
                    rp1.metric("Total Cargos", f"${tc_rp:,.2f}")
                    rp2.metric("Total Abonos", f"${ta_rp:,.2f}")
                    rp3.metric("Diferencia", f"${ta_rp-tc_rp:,.2f}")
                    st.dataframe(df_fin_rp, width='stretch', key="df_reglas_custom")
                    buf_rp = excel_con_formato({'Poliza_Personalizada': df_fin_rp}, currency_cols=['Cargo','Abono'])
                    rpd1, rpd2, rpd3 = st.columns(3)
                    rpd1.download_button("Excel", buf_rp, f"poliza_custom_{mes_rp}_{anio_rp}.xlsx")
                    rpd2.download_button("PDF", pdf_poliza(df_fin_rp, "Reglas Personalizadas", f"{mes_rp}/{anio_rp}"), f"poliza_custom_{mes_rp}_{anio_rp}.pdf")
                    rpd3.download_button(
                        "TXT (Contpaqi)",
                        poliza_txt_contpaqi(df_fin_rp, "Reglas Personalizadas", pd.Timestamp(year=int(anio_rp), month=mes_rp, day=1)),
                        f"poliza_custom_{mes_rp}_{anio_rp}.txt",
                        help="Archivo en formato de texto para importar en Contpaqi."
                    )

        else:
            sel_c=st.selectbox("Filtrar contrato",opts)
            c1_,c2_=st.columns(2)
            mes=c1_.selectbox("Mes",range(1,13),index=datetime.now().month-1,format_func=lambda m:MN[m-1])
            anio=c2_.number_input("Año",value=datetime.now().year)
            inc_pm=inc_am=False
            if tipo=="Parcialidad (mensual)": inc_pm=st.checkbox("Incluir mes de alta (primer mes en firma)", value=True)
            elif tipo=="Comisión por apertura": inc_am=st.checkbox("Incluir amortización en mes de alta")

            if st.button("Generar Póliza"):
                df_f=df_c if sel_c=="Todos" else df_c[df_c['ID_Contrato']==sel_c]
                with st.spinner("Generando…"):
                    if tipo=="Inicial":
                        mask=(df_f['Fecha_Alta'].dt.year==anio)&(df_f['Fecha_Alta'].dt.month==mes)
                        pols=[pol_inicial(r,CAT) for _,r in df_f[mask].iterrows()]
                    elif tipo=="Parcialidad (mensual)":
                        periodo_dt = pd.Timestamp(year=int(anio), month=mes, day=1)
                        def primer_mes_siguiente(fecha_str):
                            if not fecha_str: return None
                            try:
                                fd = pd.Timestamp(fecha_str)
                                return (fd + pd.offsets.MonthBegin(1))
                            except: return None
                        excl_desde = df_f['Fecha_Excl_Poliza'].apply(primer_mes_siguiente) if 'Fecha_Excl_Poliza' in df_f.columns else pd.Series([None]*len(df_f), index=df_f.index)
                        motivo_excl = df_f['Motivo_Excl_Poliza'] if 'Motivo_Excl_Poliza' in df_f.columns else pd.Series(['']*len(df_f), index=df_f.index)
                        mask_excl = excl_desde.apply(lambda d: d is not None and periodo_dt >= d)
                        df_f_par  = df_f[~mask_excl]
                        if mask_excl.any():
                            df_excl = df_f[mask_excl][['ID_Contrato','Cliente','Vehiculo']].copy()
                            df_excl['Motivo']       = motivo_excl[mask_excl].values
                            df_excl['Excluido desde'] = excl_desde[mask_excl].apply(lambda d: d.strftime('%Y-%m') if d else '').values
                            st.warning(f"**{mask_excl.sum()}** contrato(s) excluido(s) de esta póliza por morosidad o evento especial:")
                            with st.expander("Ver contratos excluidos"):
                                st.dataframe(df_excl, width='stretch', key="df_016")
                        pols=[pol_parcialidad(r,mes,anio,CAT,inc_pm) for _,r in df_f_par.iterrows()]
                    else: pols=[pol_comision(r,mes,anio,CAT,inc_am) for _,r in df_f.iterrows()]
                    pols=[p for p in pols if not p.empty]
                if pols:
                    df_fin=pd.concat(pols,ignore_index=True); tc=df_fin['Cargo'].sum(); ta=df_fin['Abono'].sum()
                    c1,c2,c3=st.columns(3)
                    c1.metric("Total Cargos",f"${tc:,.2f}"); c2.metric("Total Abonos",f"${ta:,.2f}"); c3.metric("Diferencia",f"${ta-tc:,.2f}")
                    st.dataframe(df_fin,width='stretch', key="df_017")
                    buf = excel_con_formato({'Poliza': df_fin}, currency_cols=['Cargo','Abono'])
                    cd1,cd2,cd3=st.columns(3)
                    cd1.download_button("Excel",buf,f"poliza_{mes}_{anio}.xlsx")
                    cd2.download_button("PDF",pdf_poliza(df_fin,tipo,f"{mes}/{anio}"),f"poliza_{mes}_{anio}.pdf")
                    cd3.download_button(
                        "TXT (Contpaqi)",
                        poliza_txt_contpaqi(df_fin, tipo, pd.Timestamp(year=int(anio), month=mes, day=1)),
                        f"poliza_{mes}_{anio}.txt",
                        help="Layout de ancho fijo verificado carácter por carácter contra un TXT real "
                             "exportado de Contpaqi. Si algún día tu Contpaqi trae otro layout, manda un "
                             "TXT de muestra de ese caso para ajustarlo."
                    )
                else: st.info("Sin movimientos para el período.")

    elif menu=="Reportes por Cliente":
        st.title("Reportes por Cliente")
        df_t=obtener()
        if df_t.empty: st.info("Sin contratos.")
        else:
            cli_sel=st.selectbox("Cliente",sorted(df_t['Cliente'].unique()),key="reporte_cliente_sel")
            if cli_sel:
                df2=df_t[df_t['Cliente']==cli_sel].copy()
                c1,c2,c3=st.columns(3)
                c1.metric("Contratos",len(df2)); c2.metric("Valor total",f"${df2['Valor_Sin_IVA'].sum():,.2f}"); c3.metric("Renta total",f"${df2['Mensualidad_Sin_IVA'].sum():,.2f}")
                st.dataframe(df2[['ID_Contrato','Vehiculo','Fecha_Alta','Valor_Sin_IVA','Mensualidad_Sin_IVA','Plazo','Tasa_Calculada','Estatus','Nivel_Morosidad']],width='stretch', key=f"df_018_{cli_sel}")
                if len(df2)>1:
                    f=px.bar(df2,x='ID_Contrato',y='Valor_Sin_IVA',color='Plazo',color_continuous_scale='Blues',title=f"Valor por Contrato — {cli_sel}")
                    st.plotly_chart(sfig(f),width='stretch', key=f"pc_019_{cli_sel}")
                st.download_button("Exportar Excel",export_excel(),f"reporte_{cli_sel}.xlsx")

    elif menu=="Intereses del Mes":
        st.title("Reporte de Intereses del Mes")
        st.caption("Cálculo de ingresos financieros de la cartera para el período.")
        cc1,cc2,cc3=st.columns(3)
        mes_i=cc1.selectbox("Mes",range(1,13),index=datetime.now().month-1,format_func=lambda m:MN[m-1])
        anio_i=cc2.number_input("Año",min_value=2020,max_value=2050,value=datetime.now().year,step=1,key="yi")
        mes_a=mes_i-1 if mes_i>1 else 12; anio_a=anio_i if mes_i>1 else anio_i-1
        if cc3.button("Calcular"):
            with st.spinner("Calculando intereses de la cartera…"):
                df_i,m=calc_int_mes(mes_i,anio_i); _,m_ant=calc_int_mes(mes_a,anio_a)
            if df_i.empty: st.info("Sin contratos activos para el período.")
            else:
                d_il=m['il']-m_ant.get('il',0); d_tot=m['tot']-m_ant.get('tot',0)
                c1,c2,c3,c4,c5=st.columns(5)
                c1.metric("Interés Leasing",f"${m['il']:,.2f}",f"${d_il:+,.2f}"); c2.metric("Interés Residual",f"${m['ir']:,.2f}")
                c3.metric("Total Intereses",f"${m['tot']:,.2f}",f"${d_tot:+,.2f}"); c4.metric("Capital Amortizado",f"${m['cap']:,.2f}"); c5.metric("Amort. Comisión",f"${m['com']:,.2f}")
                st.caption(f"Mayor aportación: **{m['top']}** · {m['n']} contratos activos")
                st.divider()
                ta,tb,tc_=st.tabs(["Por Contrato","Por Cliente","Tabla"])
                with ta:
                    t20=df_i.head(20)
                    fb=px.bar(t20,x='ID_Contrato',y=['Int_Leasing','Int_Residual'],barmode='stack',
                               title=f"Intereses por Contrato — {MN[mes_i-1]} {anio_i} (Top 20)",
                               color_discrete_map={'Int_Leasing':C['primary'],'Int_Residual':C['info']},labels={'value':'Interés (MXN)','variable':'Tipo'})
                    fb=sfig(fb); st.plotly_chart(fb,width='stretch', key="pc_020")
                    explain("¿Cuánto genera cada contrato en intereses este mes?",
                        "Cuánto interés genera cada contrato este mes.")
                    t10=df_i.head(10)
                    fw=go.Figure(go.Waterfall(orientation="v",measure=["relative"]*len(t10)+["total"],
                        x=list(t10['ID_Contrato'])+["TOTAL"],y=list(t10['Total_Int'])+[0],
                        text=[f"${v:,.0f}" for v in t10['Total_Int']]+[f"${t10['Total_Int'].sum():,.0f}"],
                        textposition="outside",connector=dict(line=dict(color="rgba(55,48,163,.18)")),
                        increasing=dict(marker_color=C_PASTEL['primary']),totals=dict(marker_color=C_PASTEL['accent'])))
                    fw.update_layout(title="Cascada: Acumulación de Intereses (Top 10)",paper_bgcolor="rgba(0,0,0,0)",
                                     plot_bgcolor="rgba(0,0,0,0)",font=dict(family="Segoe UI, Helvetica Neue, Arial, sans-serif"),
                                     margin=dict(t=50,r=16,b=36,l=16),height=300)
                    st.plotly_chart(fw,width='stretch', key="pc_021")
                    explain("Acumulación progresiva de los 10 mayores contratos",
                        "Los 10 contratos que más interés generan, sumados uno por uno.")
                with tb:
                    dc=df_i.groupby('Cliente').agg(IL=('Int_Leasing','sum'),IR=('Int_Residual','sum'),Tot=('Total_Int','sum'),N=('ID_Contrato','count')).reset_index().sort_values('Tot',ascending=False)
                    ec1,ec2=st.columns(2)
                    with ec1:
                        fc=px.bar(dc,y='Cliente',x='Tot',orientation='h',color='Tot',color_continuous_scale=[[0,C['light']],[1,C['primary']]],title="Total Intereses por Cliente")
                        fc=sfig(fc); fc.update_layout(yaxis={'categoryorder':'total ascending'},coloraxis_showscale=False)
                        st.plotly_chart(fc,width='stretch', key="pc_022")
                        explain("¿Qué cliente aporta más intereses?",
                        "Los clientes que más interés generan este mes.")
                    with ec2:
                        fp=px.pie(dc,values='Tot',names='Cliente',hole=0.45,title="Participación por Cliente",color_discrete_sequence=PAL_PASTEL)
                        fp=sfig(fp); fp.update_traces(textposition='inside',textinfo='percent+label')
                        st.plotly_chart(fp,width='stretch', key="pc_023")
                        explain("¿Qué tan concentrado está el ingreso?",
                        "Si un solo cliente aporta mucho del ingreso, hay más riesgo si deja de pagar.")
                    dm=dc.melt(id_vars='Cliente',value_vars=['IL','IR'],var_name='Tipo',value_name='Monto')
                    dm['Tipo']=dm['Tipo'].map({'IL':'Leasing (208)','IR':'Residual (126/114)'})
                    fs=px.bar(dm,x='Cliente',y='Monto',color='Tipo',barmode='stack',title="Desglose Leasing vs Residual por Cliente",
                               color_discrete_map={'Leasing (208)':C['primary'],'Residual (126/114)':C['info']})
                    fs=sfig(fs); st.plotly_chart(fs,width='stretch', key="pc_024")
                    explain("¿Cuánto viene del leasing y cuánto del residual?",
                        "Divide el ingreso entre el interés del leasing y el interés del residual.")
                with tc_:
                    st.dataframe(df_i.style.format({'Renta':'${:,.2f}','Int_Leasing':'${:,.2f}','Capital':'${:,.2f}',
                        'Saldo_Cap':'${:,.2f}','Int_Residual':'${:,.2f}','Amort_Com':'${:,.2f}',
                        'Total_Int':'${:,.2f}','Tasa_Anual_Pct':'{:.4f}%'}).background_gradient(subset=['Total_Int'],cmap=CMAP_INDIGO),width='stretch',
        key="df_019")
                    buf = excel_con_formato({'Intereses': df_i},
                        currency_cols=['Renta','Int_Leasing','Capital','Saldo_Cap','Int_Residual','Amort_Com','Total_Int'],
                        pct_cols=['Tasa_Anual_Pct'])
                    st.download_button("Excel",buf,f"intereses_{MN[mes_i-1]}_{anio_i}.xlsx")

    elif menu=="Proyección Financiera":
        st.title("Proyección Financiera de Intereses")
        df_act=obtener('ACTIVO')
        if df_act.empty: st.info("Sin contratos activos.")
        else:
            mp=st.slider("Meses a proyectar",6,24,12)
            if st.button("Generar Proyección"):
                with st.spinner("Calculando proyección…"): dfp=proy_intereses(mp)
                if not dfp.empty:
                    st.caption(f"Proyección calculada a partir de hoy ({fecha_larga(date.today())}): cubre de **{dfp['Label'].iloc[0]}** a **{dfp['Label'].iloc[-1]}** ({mp} meses).")
                    c1,c2,c3,c4=st.columns(4)
                    c1.metric("Total Intereses",f"${dfp['Total'].sum():,.0f}"); c2.metric("Promedio Mensual",f"${dfp['Total'].mean():,.0f}")
                    c3.metric("Máximo Mensual",f"${dfp['Total'].max():,.0f}"); c4.metric("Contratos Prom.",f"{dfp['N'].mean():.0f}")
                    dm=dfp.melt(id_vars=['Mes','Label'],value_vars=['IL','IR'],var_name='Tipo',value_name='Monto')
                    dm['Tipo']=dm['Tipo'].map({'IL':'Interés Leasing','IR':'Interés Residual'})
                    fa=px.area(dm,x='Mes',y='Monto',color='Tipo',title=f"Proyección de Intereses — {mp} meses",
                                color_discrete_map={'Interés Leasing':C['primary'],'Interés Residual':C['info']})
                    fa=sfig(fa,h=300); fa.update_traces(line_width=2)
                    st.plotly_chart(fa,width='stretch', key="pc_025")
                    explain("Intereses proyectados mes a mes",
                        "Cuánto interés se espera cobrar cada mes en los próximos meses.")
                    ec1,ec2=st.columns(2)
                    with ec1:
                        fb=px.bar(dfp,x='Mes',y='Total',color='Total',text_auto='.2s',color_continuous_scale=[[0,C['light']],[1,C['primary']]],title="Total Intereses por Mes")
                        fb=sfig(fb); fb.update_layout(coloraxis_showscale=False)
                        st.plotly_chart(fb,width='stretch', key="pc_026")
                        explain("Monto total de intereses por mes",
                        "Suma de todo el interés esperado cada mes.")
                    with ec2:
                        fc=px.line(dfp,x='Mes',y='N',markers=True,title="Contratos Activos por Mes Proyectado",color_discrete_sequence=[C['success']])
                        fc=sfig(fc); fc.update_traces(line_width=2.5,marker_size=8)
                        st.plotly_chart(fc,width='stretch', key="pc_027")
                        explain("¿Cuántos contratos estarán activos?",
                        "Cuántos contratos seguirán activos mes a mes.")
                    ri=proy_rentas(df_act,mp)
                    if not ri.empty and len(ri)==len(dfp):
                        dfc=dfp.copy(); dfc['Rentas']=ri['Rentas'].values; dfc['Intereses']=dfc['Total']
                        fd=make_subplots(specs=[[{"secondary_y":True}]])
                        fd.add_trace(go.Bar(x=dfc['Mes'],y=dfc['Rentas'],name="Rentas (izq.)",marker_color=C_PASTEL['primary'],opacity=0.82),secondary_y=False)
                        fd.add_trace(go.Scatter(x=dfc['Mes'],y=dfc['Intereses'],mode='lines+markers',name="Intereses (der.)",line=dict(color=C_PASTEL['accent'],width=3),marker_size=7),secondary_y=True)
                        fd.update_layout(title="Rentas vs Intereses — Proyección Dual",font=dict(family="Segoe UI, Helvetica Neue, Arial, sans-serif"),paper_bgcolor="rgba(0,0,0,0)",plot_bgcolor="rgba(0,0,0,0)",height=300,margin=dict(t=50,r=16,b=36,l=16))
                        fd.update_yaxes(title_text="Rentas (MXN)",secondary_y=False); fd.update_yaxes(title_text="Intereses (MXN)",secondary_y=True)
                        st.plotly_chart(fd,width='stretch', key="pc_028")
                        explain("Rentas totales vs ingresos por intereses",
                        "Compara el total cobrado contra la parte que es solo interés.")
                    dfp['Acum']=dfp['Total'].cumsum()
                    fe=px.area(dfp,x='Mes',y='Acum',title="Intereses Acumulados Proyectados",color_discrete_sequence=[C['gold']])
                    fe=sfig(fe); fe.update_traces(line_width=3,fillcolor='rgba(161,98,7,.18)',hovertemplate='%{x}<br>$%{y:,.2f}')
                    st.plotly_chart(fe,width='stretch', key="pc_029")
                    explain("Total acumulado de intereses en el período",
                        "Suma de todo el interés esperado en el periodo elegido.")
                    st.dataframe(dfp.style.format({'IL':'${:,.2f}','IR':'${:,.2f}','Total':'${:,.2f}','Acum':'${:,.2f}'}),width='stretch', key="df_020")
                    buf = excel_con_formato({'Proyeccion': dfp}, currency_cols=['IL','IR','Total','Acum'])
                    st.download_button("Descargar",buf,"proyeccion_intereses.xlsx")

    elif menu == "Cotizador Comercial":
        st.title("Cotizador Comercial")
        st.caption("Simula pagos, rentas y genera cotizaciones en PDF para prospectos.")

        col1, col2 = st.columns([1, 1])
        with col1:
            st.subheader("1. Datos del Cliente y Vehículo")
            cot_cliente = st.text_input("Nombre del Prospecto / Cliente", value="PROSPECTO EJEMPLO, S.A. DE C.V.", key="cot_cliente")
            cot_vehiculo = st.text_input("Descripción del Vehículo / Equipo", value="NISSAN URVAN 2026", key="cot_vehiculo")
            cot_valor = st.number_input("Valor de la Unidad (Sin IVA)", min_value=10000.0, max_value=50000000.0, value=450000.0, step=10000.0, key="cot_valor")
            cot_incluir_iva = st.checkbox("Aplicar IVA del 16%", value=True, key="cot_incluir_iva")

        with col2:
            st.subheader("2. Condiciones Financieras")
            c_f1, c_f2 = st.columns(2)
            cot_anticipo = c_f1.number_input("% Anticipo a Capital", min_value=0.0, max_value=70.0, value=10.0, step=5.0, key="cot_anticipo")
            cot_residual = c_f2.number_input("% Valor Residual", min_value=0.0, max_value=60.0, value=20.0, step=5.0, key="cot_residual")
            
            c_f3, c_f4 = st.columns(2)
            cot_tasa = c_f3.number_input("Tasa Anual de Interés (%)", min_value=1.0, max_value=60.0, value=24.0, step=0.5, key="cot_tasa")
            cot_plazo = c_f4.selectbox("Plazo (Meses)", [12, 18, 24, 36, 48, 60], index=3, key="cot_plazo")
            
            cot_comision = st.number_input("% Comisión por Apertura", min_value=0.0, max_value=10.0, value=2.0, step=0.5, key="cot_comision")

        # Ejecutar Simulación
        sim = simular_cotizacion(
            valor_vehiculo_sin_iva=cot_valor,
            pct_anticipo=cot_anticipo,
            pct_residual=cot_residual,
            tasa_anual=cot_tasa,
            plazo_meses=cot_plazo,
            pct_comision=cot_comision,
            incluir_iva=cot_incluir_iva
        )

        st.markdown("---")
        st.subheader("Resumen de la Cotización")

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Renta Mensual (Sin IVA)", f"${sim['renta_mensual_sin_iva']:,.2f}")
        k2.metric("Renta Total (Con IVA)", f"${sim['renta_total_con_iva']:,.2f}")
        k3.metric("Pago Inicial Requerido", f"${sim['pago_inicial_con_iva']:,.2f}")
        k4.metric("TIR Anualizada Proyectada", f"{sim['tir_anual']:.2f}%")

        st.markdown("#### Desglose Detallado de Pagos")
        d1, d2 = st.columns(2)
        with d1:
            st.markdown(f"""
            * **Inversión Neta a Financiar:** `${sim['inversion_neta']:,.2f}`
            * **Anticipo a Capital ({sim['pct_anticipo']}%):** `${sim['monto_anticipo']:,.2f}` (+ IVA `${sim['monto_anticipo']*0.16:,.2f}`)
            * **Comisión por Apertura ({sim['pct_comision']}%):** `${sim['monto_comision']:,.2f}` (+ IVA `${sim['monto_comision']*0.16:,.2f}`)
            """)
        with d2:
            st.markdown(f"""
            * **Valor Residual Pactado ({sim['pct_residual']}%):** `${sim['monto_residual']:,.2f}` (+ IVA `${sim['monto_residual']*0.16:,.2f}`)
            * **Valor Presente del Residual:** `${sim['vp_residual']:,.2f}`
            * **Pago Inicial Total (Con IVA):** `${sim['pago_inicial_con_iva']:,.2f}`
            """)

        # Gráfica interactiva de Amortización Proyectada
        dfa_sim = sim['tabla_amortizacion']
        fig_sim = px.bar(
            dfa_sim,
            x='Mes',
            y=['Interes', 'Capital'],
            title=f"Amortización Proyectada de Renta Mensual (${sim['renta_mensual_sin_iva']:,.2f})",
            labels={'value': 'Monto (MXN)', 'variable': 'Concepto'},
            color_discrete_map={'Interes': '#1E5C4F', 'Capital': '#8FD9BE'}
        )
        fig_sim.add_trace(go.Scatter(
            x=dfa_sim['Mes'],
            y=dfa_sim['Saldo'],
            mode='lines+markers',
            name='Saldo Insoluto',
            line=dict(color='#B3261E', width=2.5)
        ))
        fig_sim = sfig(fig_sim, h=340)
        st.plotly_chart(fig_sim, width="stretch", key="pc_cotizador_sim")

        # Botones de exportación
        st.markdown("---")
        ex1, ex2 = st.columns(2)
        with ex1:
            pdf_bytes = generar_pdf_cotizacion(
                sim,
                cliente_nombre=cot_cliente,
                vehiculo_desc=cot_vehiculo,
                empresa_nombre=st.session_state.get('_nombre_empresa', 'O-Leasing')
            )
            st.download_button(
                "Descargar Cotización Formal en PDF",
                pdf_bytes,
                file_name=f"cotizacion_{cot_cliente.replace(' ','_')[:15]}.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True
            )
        with ex2:
            buf_sim = excel_con_formato(
                {"Amortización Proyectada": dfa_sim},
                currency_cols=["Renta", "Interes", "Capital", "Saldo", "Saldo_Fin"]
            )
            st.download_button(
                "Descargar Tabla en Excel",
                buf_sim,
                file_name=f"simulacion_amort_{cot_cliente.replace(' ','_')[:15]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        with st.expander("Ver Tabla Completa de Amortización Proyectada"):
            fmt_cols = {c: '${:,.2f}' for c in ['Renta', 'Interes', 'Capital', 'Saldo'] if c in dfa_sim.columns}
            st.dataframe(
                dfa_sim.style.format(fmt_cols),
                width="stretch"
            )

    elif menu == "Cierre y Conciliación Mensual":
        st.title("Cierre y Conciliación Mensual")
        st.caption("Reporte de cierre contable mensual y expediente descargable en PDF o Excel.")

        c1, c2 = st.columns([1, 2])
        with c1:
            mes_cierre = st.selectbox("Mes de Cierre", list(range(1, 13)), index=datetime.now().month - 1, format_func=lambda m: MN[m - 1], key="cierre_mes_sel")
            anio_cierre = st.number_input("Año de Cierre", min_value=2020, max_value=2050, value=datetime.now().year, key="cierre_anio_sel")

        # Cargar datos de cierre
        di, dr, dc, ds, err_c = tabla_mensual_conceptos(int(anio_cierre))
        
        if err_c:
            st.error(err_c)
        else:
            nombre_mes_sel = MN[mes_cierre - 1]
            
            # Totales del mes
            tot_leasing = float(di[nombre_mes_sel].sum()) if di is not None and nombre_mes_sel in di.columns else 0.0
            tot_residual = float(dr[nombre_mes_sel].sum()) if dr is not None and nombre_mes_sel in dr.columns else 0.0
            tot_comision = float(dc[nombre_mes_sel].sum()) if dc is not None and nombre_mes_sel in dc.columns else 0.0
            
            # Facturas conciliadas
            conn = get_db()
            per_str = f"{int(anio_cierre):04d}-{int(mes_cierre):02d}"
            r_fact = conn.execute("SELECT COUNT(*) as n, SUM(total) as tot FROM facturas WHERE periodo=? AND (cancelada IS NULL OR cancelada = 0)", (per_str,)).fetchone()
            tot_facturado = float(r_fact['tot'] or 0.0) if r_fact else 0.0
            n_facturas = int(r_fact['n'] or 0) if r_fact else 0
            
            r_conc = conn.execute("SELECT COUNT(*) FROM facturas WHERE periodo=? AND estatus='CONCILIADO' AND (cancelada IS NULL OR cancelada = 0)", (per_str,)).fetchone()
            r_disc = conn.execute("SELECT COUNT(*) FROM facturas WHERE periodo=? AND estatus='DISCREPANCIA' AND (cancelada IS NULL OR cancelada = 0)", (per_str,)).fetchone()
            n_conc = int(r_conc[0]) if r_conc else 0
            n_disc = int(r_disc[0]) if r_disc else 0

            st.markdown("---")
            st.subheader(f"Resumen del Cierre — {nombre_mes_sel} {anio_cierre}")

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Interés Leasing (Cta 208)", f"${tot_leasing:,.2f}")
            k2.metric("Interés Residual", f"${tot_residual:,.2f}")
            k3.metric("Amort. Comisión", f"${tot_comision:,.2f}")
            k4.metric("Total CFDI Facturado", f"${tot_facturado:,.2f}", f"{n_facturas} facturas")

            st.markdown("#### Estado de la Conciliación del Mes")
            kc1, kc2, kc3 = st.columns(3)
            kc1.metric("Facturas Conciliadas (100% OK)", f"{n_conc}")
            kc2.metric("Discrepancias / Pendientes", f"{n_disc}")
            kc3.metric("Cobertura de Facturación", f"{(tot_facturado / max(tot_leasing, 1) * 100):.1f}%")

            st.markdown("---")
            st.subheader("Tablas de Cierre Contable")

            t_c1, t_c2 = st.tabs(["Detalle por Contrato", "Póliza Consolidada del Mes"])
            with t_c1:
                if di is not None and not di.empty:
                    st.dataframe(di[[nombre_mes_sel]].style.format('{:,.2f}'), width="stretch")
            with t_c2:
                # Póliza consolidada mensual
                pol_rows = [
                    {"Cuenta": "1150-000-000", "Concepto": f"CXC Rentas {nombre_mes_sel}", "Cargo": round(tot_leasing, 2), "Abono": 0.0},
                    {"Cuenta": "1260-000-000", "Concepto": f"CXC VP Residual {nombre_mes_sel}", "Cargo": round(tot_residual, 2), "Abono": 0.0},
                    {"Cuenta": "2080-000-000", "Concepto": f"Intereses Leasing Devengados", "Cargo": 0.0, "Abono": round(tot_leasing, 2)},
                    {"Cuenta": "2280-000-000", "Concepto": f"Intereses Residual Devengados", "Cargo": 0.0, "Abono": round(tot_residual, 2)},
                ]
                df_pol_cons = pd.DataFrame(pol_rows)
                st.dataframe(df_pol_cons.style.format({'Cargo': '${:,.2f}', 'Abono': '${:,.2f}'}), width="stretch")

            # Botones de exportación del Expediente de Cierre
            st.markdown("---")
            kpis_cierre = {
                'interes_leasing': tot_leasing,
                'interes_residual': tot_residual,
                'comision': tot_comision,
                'facturado_total': tot_facturado,
                'contratos_activos': len(di) if di is not None else 0,
                'conciliadas': n_conc,
                'discrepancias': n_disc
            }
            
            ex1, ex2 = st.columns(2)
            with ex1:
                pdf_cierre_bytes = generar_pdf_cierre_mensual(
                    mes=int(mes_cierre),
                    anio=int(anio_cierre),
                    kpis=kpis_cierre,
                    df_conciliados=di,
                    empresa_nombre=st.session_state.get('_nombre_empresa', 'O-Leasing')
                )
                st.download_button(
                    "Descargar Informe de Cierre en PDF",
                    pdf_cierre_bytes,
                    file_name=f"informe_cierre_{nombre_mes_sel}_{anio_cierre}.pdf",
                    mime="application/pdf",
                    type="primary",
                    use_container_width=True
                )
            with ex2:
                buf_cierre = excel_con_formato(
                    {
                        "Intereses Leasing": di[[nombre_mes_sel]] if di is not None else pd.DataFrame(),
                        "Póliza Consolidada": df_pol_cons
                    },
                    currency_cols=[nombre_mes_sel, 'Cargo', 'Abono']
                )
                st.download_button(
                    "Descargar Expediente en Excel",
                    buf_cierre,
                    file_name=f"expediente_cierre_{nombre_mes_sel}_{anio_cierre}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    elif menu=="Cuentas (Macro)":
        st.title("Catálogo Maestro de Cuentas Contables")

        with st.expander("Codificación de cuentas por contrato", expanded=False):
            st.caption("Configura el formato de las subcuentas contables por contrato para esta empresa.")
            _cfg_cod = get_config_codificacion_cuentas()
            cc1, cc2, cc3 = st.columns(3)
            _dbase = cc1.number_input("Dígitos de la cuenta base", min_value=0, max_value=20, value=int(_cfg_cod['digitos_base']))
            _dcontr = cc2.number_input("Dígitos para el contrato", min_value=0, max_value=20, value=int(_cfg_cod['digitos_contrato']))
            _sufijo = cc3.text_input("Sufijo al final", value=_cfg_cod['sufijo'])
            _solo_anexo = st.checkbox(
                "Usar solo la segunda mitad del número de contrato (después del guion)",
                value=_cfg_cod['solo_anexo'],
                help="Usa únicamente el número anexo en lugar del folio completo."
            )
            _cfg_preview = {'digitos_base': _dbase, 'digitos_contrato': _dcontr, 'sufijo': _sufijo, 'solo_anexo': _solo_anexo}
            _ejemplo = cta_sg("117-00-00", "0524-0001", cfg=_cfg_preview)
            st.info(f"Con esta configuración, la cuenta base **117-00-00** para el contrato **0524-0001** quedaría: **`{_ejemplo}`**")
            if st.button("Guardar codificación de cuentas"):
                set_config_codificacion_cuentas(_cfg_preview)
                st.success("Guardado. Se aplicará a las próximas pólizas y al archivo de Contpaqi que generes.")

        st.caption("Asigna las cuentas contables por concepto. Las cuentas en blanco se omiten.")
        _ETIQUETAS_CUENTA = {
            'CXC_CP':                    'Cuentas por cobrar — corto plazo',
            'CXC_LP':                    'Cuentas por cobrar — largo plazo',
            'RESIDUAL':                  'CxC valor residual — largo plazo',
            'RESIDUAL_CP':               'CxC valor residual — corto plazo',
            'BAJA_INV':                  'Baja de inventario (al dar de baja un activo)',
            'INT_CP':                    'Intereses por devengar — corto plazo',
            'INT_LP':                    'Intereses por devengar — largo plazo',
            'ANT_CAP':                   'CxC anticipo a capital',
            'PERDIDA_CESION':            'Pérdida por cesión de contrato',
            'CXC_RESIDUAL_INTERESES':    'CxC intereses del residual — largo plazo',
            'CXC_RESIDUAL_INTERESES_CP': 'CxC intereses del residual — corto plazo',
            'ING_RESIDUAL':              'Ingresos por valor residual',
            'ING_INTERESES':             'Ingresos por intereses',
            'CANCELACION_CAPITAL':       'Cancelación de capital',
            'CANCELACION_ANTICIPO':      'Cancelación de anticipo',
            'COMISION_PASIVO':           'Pasivo por comisión de apertura',
            'COMISION_GASTO':            'Gasto por comisión de apertura',
            'AJUSTE_REDONDEO':           'Ajuste por redondeo',
            'PERDIDA_EVENTO_ESPECIAL':   'Pérdida por siniestro / robo / jurídico',
            'UTILIDADES_ACUMULADAS':     'Utilidades acumuladas (ajuste de ejercicios anteriores)',
            'CXC_ASEGURADORA':           'Cuentas por cobrar a la aseguradora',
            'SALVAMENTO':                'Activo recuperado / salvamento',
            'GASTO_DETERIORO':           'Gasto por deterioro de cartera',
            'ESTIMACION_INCOBRABLES':    'Estimación para cuentas incobrables',
        }
        n_activas   = sum(1 for d in CAT.values() if str(d.get('cuenta') or '').strip())
        n_inactivas = len(CAT) - n_activas
        mc1,mc2,mc3 = st.columns(3)
        mc1.metric("Cuentas en catálogo", len(CAT))
        mc2.metric("Activas (se exportan)", n_activas)
        mc3.metric("En blanco (se omiten)", n_inactivas)
        st.markdown("---")

        with st.form("cuentas"):
            nuevas={}
            for k,d in CAT.items():
                cuenta_actual = str(d.get('cuenta') or '').strip()
                c1,c2 = st.columns([1,2])
                cta=c1.text_input(_ETIQUETAS_CUENTA.get(k,k),cuenta_actual,key=f"c_{k}",placeholder="Deja en blanco para omitir")
                nom=c2.text_input(f"Nombre contable ({_ETIQUETAS_CUENTA.get(k,k)})",d['nombre'],key=f"n_{k}")
                nuevas[k]={'cuenta':cta,'nombre':nom}
            if st.form_submit_button("Guardar catálogo", width='stretch'):
                guardar_catalogo(nuevas)
                CAT.clear(); CAT.update(cargar_catalogo())
                st.success("Catálogo actualizado. Las cuentas en blanco quedaron marcadas como inactivas y no se exportarán a Contpaqi.")
                st.session_state['_refresh'] = True

    elif menu=="Contpaqi (Cuentas)":
        st.title("Exportador de Cuentas para Contpaqi")
        st.caption("Genera el archivo para importar el catálogo de cuentas a Contpaqi.")

        cat_actual = cargar_catalogo()
        activas  = [k for k,d in cat_actual.items() if str(d.get('cuenta') or '').strip()]
        blancas  = [k for k,d in cat_actual.items() if not str(d.get('cuenta') or '').strip()]
        gc1,gc2 = st.columns(2)
        gc1.metric("Claves que se exportarán", len(activas))
        gc2.metric("Claves omitidas (en blanco)", len(blancas))
        if blancas:
            with st.expander(f"Ver las {len(blancas)} cuentas que se omitirán"):
                st.write(", ".join(blancas))
                st.caption("Ve a 'Cuentas (Macro)' si alguna de éstas sí debería exportarse.")

        if st.button("Generar archivo Contpaqi", width='stretch'):
            csv,msg=gen_contpaqi()
            if csv:
                st.success(msg)
                st.download_button("Descargar CSV", csv, "cuentas_contpaqi.csv", "text/csv", width='stretch')
            else:
                st.error(msg)

    elif menu=="Análisis de Rentabilidad":
        st.title("Análisis de Rentabilidad")
        df_act=obtener('ACTIVO'); df_baj=obtener('BAJA')
        if df_act.empty: st.info("Sin contratos activos.")
        else:
            with st.expander("Ver fórmulas"):
                st.markdown("**Ingresos** = `Renta×Plazo + Residual + Comisión` · **Egreso** = `Valor − Anticipo` · **Ganancia** = `Ingresos − Egreso` · **TIR** sobre flujos netos")
            res=[]
            for _,row in df_act.iterrows():
                g,m,pe=rentabilidad(row); t=tir(row)
                res.append({'ID_Contrato':row['ID_Contrato'],'Cliente':row['Cliente'],
                            'Valor':row['Valor_Sin_IVA'],'Egreso':max(row['Valor_Sin_IVA']-row['Anticipo_Monto'],0.01),
                            'Renta':max(row['Mensualidad_Sin_IVA'],0.01),'Plazo':row['Plazo'],
                            'Residual':row['Residual_Monto'],'Comision':row['Comision_Monto'],'Ganancia':g,'Margen':m,'PE':pe,'TIR Anual':t})
            df_r=pd.DataFrame(res).sort_values('Ganancia',ascending=False)
            c1,c2,c3,c4=st.columns(4)
            c1.metric("Ganancia Total",f"${df_r['Ganancia'].sum():,.0f}"); c2.metric("Ganancia Prom",f"${df_r['Ganancia'].mean():,.0f}")
            c3.metric("Margen Prom",f"{df_r['Margen'].mean():.2f}%"); c4.metric("TIR Anual Prom",f"{df_r['TIR Anual'].mean():.2f}%" if df_r['TIR Anual'].notna().any() else "N/A")
            st.dataframe(df_r.style.format({'Valor':'${:,.2f}','Egreso':'${:,.2f}','Renta':'${:,.2f}','Residual':'${:,.2f}','Comision':'${:,.2f}','Ganancia':'${:,.2f}','Margen':'{:.2f}%','PE':'{:.1f}','TIR Anual':'{:.2f}%'}).background_gradient(subset=['Ganancia','Margen'],cmap=CMAP_TEAL),width='stretch', key="df_021")
            ta1,ta2,ta3,ta4,ta5=st.tabs(["Ganancia por Contrato","Ganancia vs Inversión","Margen por Plazo","P&L de un Contrato","Top / Bottom 5"])
            with ta1:
                f=px.bar(df_r,x='ID_Contrato',y='Ganancia',color='Margen',text_auto='.2s',color_continuous_scale='Viridis',title="Ganancia Neta por Contrato")
                st.plotly_chart(sfig(f),width='stretch', key="pc_030")
                explain("Ganancia neta de cada contrato",
                        "Cuánto gana la empresa con cada contrato.")
            with ta2:
                df_r['_sz_r']=sz(df_r['Renta'])
                f2=px.scatter(df_r,x='Egreso',y='Ganancia',size='_sz_r',color='Margen',hover_name='Cliente',color_continuous_scale='Plasma',title="Ganancia vs Inversión Neta")
                st.plotly_chart(sfig(f2),width='stretch', key="pc_031")
                explain("Eficiencia: ganancia por peso invertido",
                        "Qué contratos dan más ganancia por cada peso invertido.")
            with ta3:
                f3=px.box(df_r,x='Plazo',y='Margen',color_discrete_sequence=[C['warning']],title="Margen por Plazo",points='all')
                st.plotly_chart(sfig(f3),width='stretch', key="pc_032")
                explain("Dispersión del margen según el plazo",
                        "Compara qué tan rentables son los contratos según su plazo.")
            with ta4:
                csel=st.selectbox("Contrato para cascada P&L",df_r['ID_Contrato'],key="rentab_csel"); rs=df_r[df_r['ID_Contrato']==csel].iloc[0]
                f4=go.Figure(go.Waterfall(measure=["absolute","relative","relative","relative","total"],
                    x=["Inversión","Rentas","Residual","Comisión","Ganancia"],
                    y=[-rs['Egreso'],rs['Renta']*rs['Plazo'],rs['Residual'],rs['Comision'],0],
                    text=[f"${-rs['Egreso']:,.0f}",f"${rs['Renta']*rs['Plazo']:,.0f}",f"${rs['Residual']:,.0f}",f"${rs['Comision']:,.0f}",f"${rs['Ganancia']:,.0f}"],
                    textposition="outside",increasing=dict(marker_color=C_PASTEL['success']),decreasing=dict(marker_color=C_PASTEL['accent']),totals=dict(marker_color=C_PASTEL['primary'])))
                f4.update_layout(title=f"P&L Detallado — {csel}",paper_bgcolor="rgba(0,0,0,0)",plot_bgcolor="rgba(0,0,0,0)",font=dict(family="Segoe UI, Helvetica Neue, Arial, sans-serif"),margin=dict(t=50,r=16,b=36,l=16),height=300)
                st.plotly_chart(f4,width='stretch', key=f"pc_033_{csel}")
                explain("P&L completo de un contrato",
                        "Resumen de gastos e ingresos de un solo contrato.")
            with ta5:
                ec1,ec2=st.columns(2)
                with ec1:
                    st.write("**Top 5 Mayor Margen**")
                    st.dataframe(df_r.nlargest(5,'Margen')[['ID_Contrato','Cliente','Egreso','Ganancia','Margen','TIR Anual']].style.format({'Egreso':'${:,.2f}','Ganancia':'${:,.2f}','Margen':'{:.2f}%','TIR Anual':'{:.2f}%'}),width='stretch', key="df_022")
                with ec2:
                    st.write("**Bottom 5 Menor Margen**")
                    st.dataframe(df_r.nsmallest(5,'Margen')[['ID_Contrato','Cliente','Egreso','Ganancia','Margen','TIR Anual']].style.format({'Egreso':'${:,.2f}','Ganancia':'${:,.2f}','Margen':'{:.2f}%','TIR Anual':'{:.2f}%'}),width='stretch', key="df_023")
        if not df_baj.empty:
            st.subheader("Pérdidas por Cesión")
            dp=perdidas_cesion(df_baj,CAT)
            if not dp.empty:
                f=px.bar(dp,x='Mes',y='Perdida',color_discrete_sequence=[C['accent']],title="Pérdida por Cesión por Mes",text_auto='.2s')
                st.plotly_chart(sfig(f),width='stretch', key="pc_034")
                explain("¿Cuánto se perdió por cancelaciones anticipadas?",
                        "Cuánto dinero se dejó de ganar por cancelaciones antes de tiempo.")
                st.metric("Pérdida Total Acumulada",f"${dp['Perdida'].sum():,.2f}")

    elif menu=="Respaldo y Restauración":
        st.title("Respaldo y Restauración de la Base de Datos")
        c1,_=st.columns([1,3])
        with c1:
            if st.button("Descargar copia"):
                bk=backup_db()
                if bk: st.download_button("Guardar .db",bk,"leasing_backup.db","application/x-sqlite3")
        st.download_button("Exportar Excel",export_excel(),"contratos_backup.xlsx")
        uf=st.file_uploader("Restaurar desde .db",type=['db'])
        if uf and st.button("Restaurar"):
            if restore_db(uf): st.success("Restaurado. Recarga la app.")

        st.divider()
        st.subheader("Cifrado de los respaldos automáticos")
        st.caption("Opcional: Asigna una contraseña para proteger tus respaldos con cifrado.")
        _pass_actual = get_cfg('respaldo_password')
        if not _PYZIPPER_DISPONIBLE:
            st.warning("La librería de cifrado (pyzipper) no está instalada en este equipo — los respaldos "
                       "se seguirán guardando sin cifrar hasta que se instale.")
        elif _pass_actual:
            st.success("Los respaldos automáticos nuevos se están guardando cifrados.")
            if st.button("Quitar la contraseña (volver a respaldos sin cifrar)"):
                set_cfg('respaldo_password', '')
                st.success("Listo — los próximos respaldos ya no se cifrarán.")
                st.session_state['_refresh'] = True
        else:
            with st.form("form_pass_respaldo"):
                _p1 = st.text_input("Nueva contraseña para los respaldos", type="password")
                _p2 = st.text_input("Repite la contraseña", type="password")
                if st.form_submit_button("Activar cifrado de respaldos"):
                    if not _p1:
                        st.error("Escribe una contraseña.")
                    elif _p1 != _p2:
                        st.error("Las contraseñas no coinciden.")
                    else:
                        set_cfg('respaldo_password', _p1)
                        st.success("Listo — a partir del próximo respaldo automático, se guardará cifrado.")
                        st.session_state['_refresh'] = True

        st.divider()
        st.subheader("Diagnóstico y respaldos automáticos")
        st.caption("Copias de seguridad recientes de la base de datos disponibles para restaurar.")
        _db_path_diag = get_db_path()
        dc1, dc2 = st.columns(2)
        with dc1:
            if st.button("Verificar integridad ahora"):
                if bd_esta_sana(_db_path_diag):
                    st.success("La base de datos actual está sana.")
                else:
                    st.error("La base de datos actual reporta daño. Restaura un respaldo abajo.")
        with dc2:
            if st.button("Crear respaldo manual ahora"):
                crear_respaldo_automatico(_db_path_diag, password=(get_cfg('respaldo_password') or None))
                st.success("Respaldo creado (si la base estaba sana en este momento).")

        _resp = listar_respaldos_automaticos(_db_path_diag)
        if _resp:
            st.markdown(f"**{len(_resp)} respaldo(s) automático(s) disponibles:**")
            for _arch, _fecha, _cifrado in _resp:
                rc1, rc2 = st.columns([3,1])
                _etiqueta_cifrado = " [Cifrado]" if _cifrado else ""
                rc1.write(f"{_fecha.strftime('%Y-%m-%d %H:%M:%S')}{_etiqueta_cifrado}  ·  {os.path.getsize(_arch)/1024:.0f} KB")
                _clave_this = None
                if _cifrado:
                    _clave_this = rc1.text_input("Contraseña", type="password", key=f"pass_restore_{_arch}", label_visibility="collapsed", placeholder="Contraseña del respaldo")
                if rc2.button("Restaurar", key=f"restore_auto_{_arch}"):
                    try:
                        restaurar_respaldo_automatico(_db_path_diag, _arch, password=_clave_this)
                        st.success("Restaurado. Recarga la app (F5) para continuar con estos datos.")
                        st.session_state.clear()
                        st.stop()
                    except RuntimeError as e:
                        st.error(str(e))
        else:
            st.info("Todavía no hay respaldos automáticos para esta empresa — se crea el primero la próxima vez que abras la app.")

        st.divider(); st.subheader("Configuración de Marca")
        with st.expander("Configuración de Marca", expanded=True):
            with st.form("brand"):
                ne=st.text_input("Nombre empresa",value=get_cfg('nombre_empresa','Orange Leasing'))
                lf=st.file_uploader("Logo",type=['png','jpg','jpeg'])
                if st.form_submit_button("Guardar"):
                    if lf: set_cfg('logo',base64.b64encode(lf.read()).decode())
                    set_cfg('nombre_empresa',ne)
                    st.success("Guardado.")
                    st.session_state['_refresh'] = True
            if st.session_state.get('confirmar_reset_marca'):
                st.warning("¿Seguro que quieres quitar el logo y el nombre personalizado? Se vuelve a los valores por default.")
                _rc1, _rc2 = st.columns(2)
                if _rc1.button("Sí, restablecer", type="primary", key="reset_marca_ok"):
                    conn=get_db(); conn.execute("DELETE FROM configuracion WHERE clave IN ('logo','nombre_empresa')"); conn.commit()
                    st.session_state['confirmar_reset_marca'] = False
                    st.success("Restablecido.")
                    st.session_state['_refresh'] = True
                if _rc2.button("Cancelar", key="reset_marca_no"):
                    st.session_state['confirmar_reset_marca'] = False
                    st.session_state['_refresh'] = True
            else:
                if st.button("Restablecer nombre y logo", key="reset_marca_ask"):
                    st.session_state['confirmar_reset_marca'] = True
                    st.session_state['_refresh'] = True

    elif menu=="Gestión de Morosidad":
        st.title("Gestión de Morosidad")
        st.markdown("**0** Al corriente · **1** Atraso · **2** Convenio · **3** Devuelve sin pago · **4** Judicial")
        df_act=obtener('ACTIVO')
        if df_act.empty: st.info("Sin contratos activos.")
        else:
            tab1,tab2=st.tabs(["Análisis","Actualizar"])
            ML={0:"Al corriente",1:"Atraso",2:"Convenio",3:"Dev.s/pago",4:"Judicial"}
            with tab1:
                mora=df_act['Nivel_Morosidad'].value_counts().sort_index().reset_index()
                mora.columns=['Nivel','Cantidad']; mora['Label']=mora['Nivel'].map(ML)
                ec1,ec2=st.columns(2)
                with ec1:
                    f=px.bar(mora,x='Label',y='Cantidad',color='Nivel',text_auto=True,
                              color_continuous_scale=[[0,C['success']],[.5,C['warning']],[1,C['accent']]],title="Contratos por Nivel de Morosidad")
                    f=sfig(f); f.update_layout(coloraxis_showscale=False)
                    st.plotly_chart(f,width='stretch', key="pc_035")
                    explain("Distribución por salud crediticia",
                        "Cuántos contratos están al corriente y cuántos en mora.")
                with ec2:
                    vm=df_act.groupby('Nivel_Morosidad')['Valor_Sin_IVA'].sum().reset_index()
                    vm['Label']=vm['Nivel_Morosidad'].map(ML)
                    f2=px.pie(vm,values='Valor_Sin_IVA',names='Label',hole=0.45,title="Valor Expuesto por Nivel de Morosidad",
                               color_discrete_map={v:cl for v,cl in zip(ML.values(),[C['success'],C['info'],C['warning'],C['accent'],C['primary']])})
                    f2=sfig(f2); f2.update_traces(textposition='inside',textinfo='percent+label')
                    st.plotly_chart(f2,width='stretch', key="pc_036")
                    explain("¿Cuánto valor está en riesgo?",
                        "Cuánto dinero de la cartera está en riesgo por mora.")
                explain("¿Qué clientes tienen contratos en mora?",
                        "Qué clientes tienen contratos en mora y qué tan grave es.")
                nf=st.selectbox("Filtrar detalle por nivel",[0,1,2,3,4])
                st.dataframe(df_act[df_act['Nivel_Morosidad']==nf][['ID_Contrato','Cliente','Vehiculo','Valor_Sin_IVA','Mensualidad_Sin_IVA','Plazo']],width='stretch', key="df_024")
            with tab2:
                modo=st.radio("Actualizar por:",["Cliente (todos sus contratos)","Contrato individual"],horizontal=True)
                if modo=="Cliente (todos sus contratos)":
                    clientes=sorted(df_act['Cliente'].dropna().unique().tolist())
                    cli_sel=st.selectbox("Cliente",clientes)
                    df_cli=df_act[df_act['Cliente']==cli_sel].copy()
                    df_cli['Nivel']=df_cli['Nivel_Morosidad'].map(ML)
                    st.caption(f"{cli_sel} tiene {len(df_cli)} contrato(s) activo(s):")
                    st.dataframe(df_cli[['ID_Contrato','Vehiculo','Valor_Sin_IVA','Mensualidad_Sin_IVA','Nivel']],width='stretch', key="df_025")
                    nn_cli=st.selectbox("Nuevo nivel para TODOS sus contratos",[0,1,2,3,4],
                                format_func=lambda x:{0:"0-Al corriente",1:"1-Atraso",2:"2-Convenio",3:"3-Devuelve no paga",4:"4-Judicial"}[x],
                                key="nn_cliente")
                    fecha_excl_cli = None
                    if nn_cli in [3,4]:
                        st.markdown("---")
                        st.caption(f"Los contratos en nivel {nn_cli} se excluirán de pólizas de parcialidades a partir de la fecha indicada.")
                        fecha_excl_cli = st.date_input("Fecha de aplicación", value=date.today(), key="fexcl_mora_cli")
                    if st.button(f"Actualizar los {len(df_cli)} contrato(s) de {cli_sel}"):
                        if nn_cli in [3,4] and fecha_excl_cli:
                            get_db().execute(
                                "UPDATE contratos SET Nivel_Morosidad=?, Fecha_Excl_Poliza=?, Motivo_Excl_Poliza=? WHERE Cliente=? AND Estatus='ACTIVO'",
                                (nn_cli, fecha_excl_cli.isoformat(), f"Morosidad nivel {nn_cli}", cli_sel))
                            for idc in df_cli['ID_Contrato']:
                                agregar_anotacion(idc, f"Nivel de morosidad actualizado a {nn_cli} ({ML[nn_cli]}) — masivo por cliente. Excluido de pólizas desde {fecha_excl_cli}.", "Alerta")
                        elif nn_cli in [0,1,2]:
                            get_db().execute(
                                "UPDATE contratos SET Nivel_Morosidad=?, Fecha_Excl_Poliza=NULL, Motivo_Excl_Poliza=NULL WHERE Cliente=? AND Estatus='ACTIVO'",
                                (nn_cli, cli_sel))
                            for idc in df_cli['ID_Contrato']:
                                agregar_anotacion(idc, f"Nivel de morosidad restablecido a {nn_cli} ({ML[nn_cli]}) — masivo. Exclusión de pólizas eliminada.", "Nota legal")
                        else:
                            get_db().execute("UPDATE contratos SET Nivel_Morosidad=? WHERE Cliente=? AND Estatus='ACTIVO'",(nn_cli,cli_sel))
                            for idc in df_cli['ID_Contrato']:
                                agregar_anotacion(idc,f"Nivel de morosidad actualizado a {nn_cli} ({ML[nn_cli]}) — actualización masiva por cliente ({cli_sel}).","Nota general")
                        get_db().commit()
                        _tocar_datos()
                        st.success(f"Se actualizaron {len(df_cli)} contrato(s) de {cli_sel} al nivel {nn_cli} — {ML[nn_cli]}.")
                        st.session_state['_refresh'] = True
                else:
                    csel=st.selectbox("Contrato",df_act['ID_Contrato']); row=df_act[df_act['ID_Contrato']==csel].iloc[0]
                    nivel_actual = int(row['Nivel_Morosidad'])
                    excl_actual  = row.get('Fecha_Excl_Poliza','') or ''
                    st.info(f"**{row['Cliente']}** · Nivel actual: **{nivel_actual} — {ML.get(nivel_actual,'N/A')}**"
                            + (f"  ·  Excluido de pólizas desde: **{excl_actual}**" if excl_actual else ""))
                    nn=st.selectbox("Nuevo nivel",[0,1,2,3,4],format_func=lambda x:{0:"0-Al corriente",1:"1-Atraso",2:"2-Convenio",3:"3-Devuelve no paga",4:"4-Judicial"}[x])
                    fecha_excl_mora = None
                    if nn in [3,4]:
                        st.markdown("---")
                        st.caption(f"El contrato quedará excluido de pólizas de parcialidades a partir de la fecha indicada.")
                        fecha_excl_mora = st.date_input("Fecha de aplicación", value=date.today(), key="fexcl_mora_ind")
                    elif nn in [0,1,2] and excl_actual:
                        st.info("Al bajar el nivel, se eliminará la exclusión de pólizas de este contrato.")
                    if st.button("Actualizar Nivel"):
                        if nn in [3,4] and fecha_excl_mora:
                            get_db().execute(
                                "UPDATE contratos SET Nivel_Morosidad=?, Fecha_Excl_Poliza=?, Motivo_Excl_Poliza=? WHERE ID_Contrato=?",
                                (nn, fecha_excl_mora.isoformat(), f"Morosidad nivel {nn}", csel))
                            agregar_anotacion(csel, f"Nivel de morosidad actualizado a {nn} ({ML[nn]}). Excluido de pólizas de parcialidades desde {fecha_excl_mora}.", "Alerta")
                        elif nn in [0,1,2]:
                            get_db().execute(
                                "UPDATE contratos SET Nivel_Morosidad=?, Fecha_Excl_Poliza=NULL, Motivo_Excl_Poliza=NULL WHERE ID_Contrato=?",
                                (nn, csel))
                            agregar_anotacion(csel, f"Nivel de morosidad restablecido a {nn} ({ML[nn]}). Exclusión de pólizas eliminada.", "Nota legal")
                        else:
                            get_db().execute("UPDATE contratos SET Nivel_Morosidad=? WHERE ID_Contrato=?",(nn,csel))
                        get_db().commit()
                        _tocar_datos()
                        st.success(f"Actualizado a nivel {nn}." + (f" Excluido de pólizas desde {fecha_excl_mora}." if nn in [3,4] and fecha_excl_mora else ""))
                        st.session_state['_refresh'] = True

    elif menu=="Tabla Mensual por Contrato":
        st.title("Tabla Mensual de Conceptos Financieros")
        st.caption("Resumen mensual de intereses, comisiones y saldos por contrato.")
        anio_t=st.number_input("Año",min_value=2020,max_value=2050,value=datetime.now().year,step=1,key="yt")
        if st.button("Generar Tablas"):
            di,dr,dc,ds,err=tabla_mensual_conceptos(anio_t)
            if err: st.error(err)
            else:
                for titulo,dft,fname in [
                    ("1. Intereses Leasing — cta 208",di,f"int_208_{anio_t}.xlsx"),
                    ("2. Intereses Residual — cta 126/114",dr,f"int_residual_{anio_t}.xlsx"),
                    ("3. Amortización Comisión por Apertura",dc,f"amort_comision_{anio_t}.xlsx"),
                    ("4. Saldo del Valor Residual Activo",ds,f"saldo_residual_{anio_t}.xlsx"),
                ]:
                    st.subheader(titulo)
                    if dft is None or dft.empty: st.info("Sin datos.")
                    else:
                        st.dataframe(dft.style.format("{:,.2f}").highlight_null("lightgray"),width='stretch', key=f"df_026_{fname}")
                        _dft_export = dft.reset_index().rename(columns={'index': 'ID_Contrato'})
                        buf = excel_con_formato({titulo[:31]: _dft_export},
                            currency_cols=[c for c in _dft_export.columns if c != 'ID_Contrato'])
                        st.download_button(f"{titulo[:25]}",buf,fname,key=f"dl_{fname}")

    elif menu=="Reporte Maestro":
        st.title("Reporte Maestro: Capital e Intereses")
        st.caption("Tablas de amortización y resúmenes de la cartera activa.")
        anio_m=st.number_input("Año del reporte maestro",min_value=2020,max_value=2050,value=datetime.now().year,step=1,key="ym")
        if st.button("Generar Reporte Maestro", type="primary"):
            di, dcap, drenta, dr, dc, ds, err = reporte_maestro_mensual(anio_m)
            if err: st.error(err)
            else:
                MN = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
                
                # Calcular totales globales para KPIs y gráficas (limitado a 2 decimales)
                dcap_g = dcap[dcap.index != 'TOTAL_000']
                di_g = di[di.index != 'TOTAL_000']
                drenta_g = drenta[drenta.index != 'TOTAL_000']
                
                total_cap = dcap_g[MN].sum().round(2)
                total_int = di_g[MN].sum().round(2)
                total_renta = drenta_g[MN].sum().round(2)
                
                st.markdown("---")
                st.subheader(f"Resumen Ejecutivo {anio_m}")
                
                k1, k2, k3 = st.columns(3)
                k1.metric("Capital a Recuperar", f"${total_cap.sum():,.2f}")
                k2.metric("Intereses a Devengar", f"${total_int.sum():,.2f}")
                k3.metric("Flujo Total (Renta Neta)", f"${total_renta.sum():,.2f}")
                
                # Gráfica Coqueta (Composición mensual)
                fig_exec = go.Figure()
                fig_exec.add_trace(go.Bar(x=MN, y=total_cap.values, name='Amortización de Capital', marker_color=C['primary'], opacity=0.85))
                fig_exec.add_trace(go.Bar(x=MN, y=total_int.values, name='Intereses', marker_color=C['accent'], opacity=0.85))
                fig_exec.add_trace(go.Scatter(x=MN, y=total_renta.values, name='Flujo Total', mode='lines+markers+text',
                                              text=[f"${v/1000:,.0f}k" if v > 0 else "" for v in total_renta.values],
                                              textposition="top center",
                                              line=dict(color=C['success'], width=3),
                                              marker=dict(size=8, color=C['success'], line=dict(width=2, color='white'))))
                
                fig_exec.update_layout(
                    barmode='stack',
                    title=f"Evolución del Flujo de Efectivo en {anio_m}",
                    hovermode="x unified",
                    xaxis=dict(showgrid=False),
                    yaxis=dict(showgrid=True, gridcolor='rgba(200,200,200,0.2)'),
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    margin=dict(l=10, r=10, t=40, b=10),
                    legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1)
                )
                fig_exec = sfig(fig_exec, h=380)
                st.plotly_chart(fig_exec, width='stretch', key="pc_maestro_exec")
                
                # Generar Excel completo con Graficas
                buf_completo = exportar_maestro_completo(anio_m, di, dcap, drenta, dr, dc, ds, total_cap, total_int, total_renta)
                st.download_button("Descargar Todo en un Solo Excel (con Gráficas)", buf_completo, f"reporte_maestro_completo_{anio_m}.xlsx", type="primary", use_container_width=True)
                
                st.markdown("---")
                
                for titulo,dft,fname in [
                    ("1. Capital Leasing (Amortización Principal)", dcap, f"capital_leasing_{anio_m}.xlsx"),
                    ("2. Intereses Leasing (Devengados)", di, f"intereses_leasing_{anio_m}.xlsx"),
                    ("3. Renta Neta (Capital + Interés)", drenta, f"renta_neta_{anio_m}.xlsx"),
                    ("4. Intereses Residual (Acumulación)", dr, f"int_residual_{anio_m}.xlsx"),
                    ("5. Amortización Comisión por Apertura", dc, f"amort_comision_{anio_m}.xlsx"),
                    ("6. Saldo del Valor Residual Activo", ds, f"saldo_residual_{anio_m}.xlsx"),
                ]:
                    st.subheader(titulo)
                    if dft is None or dft.empty: st.info("Sin datos.")
                    else:
                        fmt_dict = {m: "{:,.2f}" for m in MN}
                        fmt_dict['Valor_Sin_IVA'] = "{:,.2f}"
                        fmt_dict['Total Año'] = "{:,.2f}"
                        
                        # Limitar a 2 decimales explícitamente en el dataframe visual
                        dft_vis = dft.copy()
                        for col in MN + ['Valor_Sin_IVA', 'Tasa_Anual_%', 'Total Año']:
                            if col in dft_vis.columns:
                                dft_vis[col] = dft_vis[col].round(2)
                                
                        st.dataframe(dft_vis.style.format(fmt_dict).highlight_null("lightgray"), width='stretch', key=f"df_026_m_{fname}")
                        _dft_export = dft_vis.reset_index().rename(columns={'index': 'ID_Contrato'})
                        
                        skip_currency = ['ID_Contrato', 'Cliente', 'Vehiculo', 'Estatus', 'Fecha_Alta', 'Plazo', 'Fecha_Baja', 'Tasa_Anual_%']
                        curr_cols = [c for c in _dft_export.columns if c not in skip_currency]
                        
                        buf = excel_con_formato({titulo[:31]: _dft_export}, currency_cols=curr_cols, pct_cols=['Tasa_Anual_%'])
                        st.download_button(f"{titulo[:25]}", buf, fname, key=f"dl_m_{fname}")

    elif menu=="Reporte Maestro Saldos":
        st.title("Reporte Maestro: Saldos Insolutos")
        st.caption("Tablas de saldos insolutos mensuales de la cartera activa.")
        anio_s=st.number_input("Año del reporte de saldos",min_value=2020,max_value=2050,value=datetime.now().year,step=1,key="ys")
        if st.button("Generar Reporte de Saldos", type="primary"):
            dcap, dint, dtot, dres, err = reporte_maestro_saldos(anio_s)
            if err: st.error(err)
            else:
                MN = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
                
                # Omitimos la fila TOTAL_000 para calcular totales de gráfica
                dcap_g = dcap[dcap.index != 'TOTAL_000']
                dint_g = dint[dint.index != 'TOTAL_000']
                dtot_g = dtot[dtot.index != 'TOTAL_000']
                
                total_cap = dcap_g[MN].sum().round(2)
                total_int = dint_g[MN].sum().round(2)
                total_tot = dtot_g[MN].sum().round(2)
                
                st.markdown("---")
                st.subheader(f"Resumen de Cartera (Saldos a fin de mes) - {anio_s}")
                
                k1, k2, k3 = st.columns(3)
                k1.metric("Saldo Capital (Diciembre)", f"${total_cap.iloc[-1]:,.2f}")
                k2.metric("Saldo Intereses (Diciembre)", f"${total_int.iloc[-1]:,.2f}")
                k3.metric("Saldo Total (Diciembre)", f"${total_tot.iloc[-1]:,.2f}")
                
                fig_exec = go.Figure()
                fig_exec.add_trace(go.Scatter(x=MN, y=total_cap.values, name='Saldo Capital', fill='tonexty', mode='lines', line=dict(color=C['primary'], width=3)))
                fig_exec.add_trace(go.Scatter(x=MN, y=total_int.values, name='Saldo Intereses', fill='tonexty', mode='lines', line=dict(color=C['accent'], width=3)))
                fig_exec.add_trace(go.Scatter(x=MN, y=total_tot.values, name='Saldo Total', mode='lines+markers+text',
                                              text=[f"${v/1000:,.0f}k" if v > 0 else "" for v in total_tot.values],
                                              textposition="top center",
                                              line=dict(color=C['success'], width=3, dash='dot'),
                                              marker=dict(size=8, color=C['success'])))
                
                fig_exec.update_layout(
                    title=f"Curva de Abatimiento de Saldos de la Cartera en {anio_s}",
                    hovermode="x unified",
                    xaxis=dict(showgrid=False),
                    yaxis=dict(showgrid=True, gridcolor='rgba(200,200,200,0.2)'),
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    margin=dict(l=10, r=10, t=40, b=10),
                    legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1)
                )
                fig_exec = sfig(fig_exec, h=380)
                st.plotly_chart(fig_exec, width='stretch', key="pc_maestro_saldos")
                
                buf_completo = exportar_saldos_completo(anio_s, dcap, dint, dtot, dres, total_cap, total_int, total_tot)
                st.download_button("Descargar Todo en un Solo Excel (con Gráficas)", buf_completo, f"reporte_saldos_completo_{anio_s}.xlsx", type="primary", use_container_width=True)
                
                st.markdown("---")
                
                for titulo,dft,fname in [
                    ("1. Saldo Capital (Insoluto al fin de mes)", dcap, f"saldo_capital_{anio_s}.xlsx"),
                    ("2. Saldo Intereses (Por devengar al fin de mes)", dint, f"saldo_intereses_{anio_s}.xlsx"),
                    ("3. Saldo Total (Lo que debe en total)", dtot, f"saldo_total_{anio_s}.xlsx"),
                    ("4. Saldo Residual", dres, f"saldo_residual_{anio_s}.xlsx"),
                ]:
                    st.subheader(titulo)
                    if dft is None or dft.empty: st.info("Sin datos.")
                    else:
                        fmt_dict = {m: "{:,.2f}" for m in MN}
                        fmt_dict['Valor_Sin_IVA'] = "{:,.2f}"
                        fmt_dict['Total Año'] = "{:,.2f}"
                        
                        dft_vis = dft.copy()
                        for col in MN + ['Valor_Sin_IVA', 'Tasa_Anual_%', 'Total Año']:
                            if col in dft_vis.columns:
                                dft_vis[col] = pd.to_numeric(dft_vis[col], errors='ignore').round(2)
                                
                        st.dataframe(dft_vis.style.format(fmt_dict).highlight_null("lightgray"), width='stretch', key=f"df_026_s_{fname}")
                        _dft_export = dft_vis.reset_index().rename(columns={'index': 'ID_Contrato'})
                        
                        skip_currency = ['ID_Contrato', 'Cliente', 'Vehiculo', 'Estatus', 'Fecha_Alta', 'Plazo', 'Fecha_Baja', 'Tasa_Anual_%']
                        from reports.excel import excel_con_formato
                        curr_cols = [c for c in _dft_export.columns if c not in skip_currency]
                        buf = excel_con_formato({titulo[:31]: _dft_export}, currency_cols=curr_cols, pct_cols=['Tasa_Anual_%'])
                        st.download_button(f"{titulo[:25]}", buf, fname, key=f"dl_s_{fname}")

    elif menu=="Multiempresa":
        st.title("Gestión Multiempresa")
        st.caption("Administra empresas registradas y su configuración independiente.")
        empresas=load_empresas(); emp_act=get_empresa_actual()

        _busq_emp = st.text_input(
            "Buscar empresa", placeholder="Escribe un nombre para filtrar…",
            key="multiemp_busqueda", label_visibility="collapsed"
        )
        _empresas_vista = [e for e in empresas if _busq_emp.strip().lower() in e['nombre'].lower()] if _busq_emp.strip() else empresas

        if not _empresas_vista:
            st.info("Ninguna empresa coincide con esa búsqueda.")

        _cols_emp = st.columns(3)
        for _i_emp, e in enumerate(_empresas_vista):
            activa = e['id'] == emp_act['id']
            with _cols_emp[_i_emp % 3]:
                _inicial = (e['nombre'] or "?").strip()[0].upper()
                _borde = "var(--accent)" if activa else "var(--line)"
                st.markdown(f"""
                <div style="border:1.5px solid {_borde};border-radius:8px;padding:14px;margin-bottom:12px;
                            background:var(--paper-hi);">
                  <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">
                    <div style="width:38px;height:38px;border-radius:8px;background:var(--accent);
                                color:#fff;display:flex;align-items:center;justify-content:center;
                                font-weight:700;font-size:1.05rem;">{_inicial}</div>
                    <div>
                      <div style="font-weight:700;color:var(--ink);font-size:.92rem;">{e['nombre']}</div>
                      <div style="font-size:.68rem;color:var(--ink-faint);">{'EMPRESA ACTIVA' if activa else e['db_path']}</div>
                    </div>
                  </div>
                </div>
                """, unsafe_allow_html=True)
                if activa:
                    st.success("Estás viendo esta empresa", icon=None)
                else:
                    if st.button("Cambiar a esta empresa", key=f"sw_{e['id']}", width='stretch'):
                        st.session_state['empresa_id'] = e['id']
                        limpiar_seleccion_contrato()
                        st.session_state['_refresh'] = True
                with st.expander("Detalle / eliminar"):
                    st.caption(f"**ID:** `{e['id']}`  ·  **DB:** `{e['db_path']}`")
                    if not activa and len(empresas) > 1:
                        _conf_key = f"del_conf_{e['id']}"
                        _confirmar = st.checkbox(
                            f"Confirmo que quiero eliminar '{e['nombre']}' (esto no borra su base de datos en disco)",
                            key=_conf_key
                        )
                        if st.button("Eliminar empresa", key=f"del_{e['id']}", type="primary", disabled=not _confirmar):
                            empresas = [x for x in empresas if x['id'] != e['id']]
                            save_empresas(empresas)
                            st.success(f"'{e['nombre']}' eliminada de la lista.")
                            st.session_state['_refresh'] = True

        st.divider(); st.subheader("Agregar Nueva Empresa")
        with st.form("nueva"):
            ne_n=st.text_input("Nombre","Nueva Leasing S.A."); ne_i=st.text_input("ID único","empresa2"); ne_d=st.text_input("Ruta DB","data/empresa2.db")
            if st.form_submit_button("Crear Empresa"):
                if ne_i in [e['id'] for e in empresas]: st.error(f"El ID '{ne_i}' ya existe.")
                elif not ne_n or not ne_i: st.error("Completa todos los campos.")
                else:
                    empresas.append({"id":ne_i,"nombre":ne_n,"db_path":ne_d}); save_empresas(empresas)
                    old=st.session_state.get('empresa_id'); st.session_state['empresa_id']=ne_i; init_db(); st.session_state['empresa_id']=old
                    st.success(f"Empresa '{ne_n}' creada.")
                    st.session_state['_refresh'] = True

        st.divider()
        df_info=obtener(); ec1,ec2,ec3,ec4=st.columns(4)
        ec1.metric("ID",emp_act['id']); ec2.metric("DB",emp_act['db_path'])
        ec3.metric("Total contratos",len(df_info)); ec4.metric("Activos",len(df_info[df_info['Estatus']=='ACTIVO']) if not df_info.empty else 0)

    elif menu=="Usuarios y Roles":
        from models.auth import listar_grupos, crear_grupo, eliminar_grupo, obtener_empresas_de_grupo
        st.title("Usuarios, Grupos y Roles")
        st.caption("Administra usuarios, grupos y permisos por rol en el sistema.")
        
        tab_usrs, tab_grps = st.tabs(["Usuarios", "Grupos de Acceso"])
        
        with tab_grps:
            st.subheader("Grupos Existentes")
            _grupos = listar_grupos(AUTH_DB)
            if not _grupos:
                st.info("No hay grupos definidos.")
            for _g in _grupos:
                with st.expander(f"Grupo: {_g['nombre_grupo']}"):
                    st.write(f"**Descripción:** {_g['descripcion']}")
                    st.write(f"**Empresas asignadas:** {', '.join(_g['empresas']) if _g['empresas'] else 'Ninguna'}")
                    if st.button("Eliminar Grupo", key=f"del_grp_{_g['id']}", type="primary"):
                        ok, msg = eliminar_grupo(AUTH_DB, _g['id'])
                        (st.success if ok else st.error)(msg)
                        if ok: st.session_state['_refresh'] = True
            
            st.divider()
            st.subheader("Crear Nuevo Grupo")
            with st.form("nuevo_grupo"):
                _ng_nombre = st.text_input("Nombre del grupo")
                _ng_desc = st.text_input("Descripción")
                _todas_emps = load_empresas()
                _ng_emps = st.multiselect("Empresas a las que tendrá acceso", options=[e['id'] for e in _todas_emps], format_func=lambda x: next((e['nombre'] for e in _todas_emps if e['id']==x), x))
                if st.form_submit_button("Crear grupo"):
                    ok, msg = crear_grupo(AUTH_DB, _ng_nombre, _ng_desc, _ng_emps)
                    (st.success if ok else st.error)(msg)
                    if ok: st.session_state['_refresh'] = True
                    
        with tab_usrs:
            st.subheader("Usuarios Registrados")
            _usrs = listar_usuarios(AUTH_DB)
            _grupos_lista = listar_grupos(AUTH_DB)
            _grupo_opts = [0] + [g['id'] for g in _grupos_lista]
            _grupo_fmt = lambda x: "Ninguno (Solo Admins)" if x == 0 else next((g['nombre_grupo'] for g in _grupos_lista if g['id'] == x), str(x))
            
            for _u in _usrs:
                with st.expander(
                    f"{_u['nombre_completo']}  —  @{_u['username']}  —  {ROL_LABELS.get(_u['rol'], _u['rol'])}"
                    + ("" if _u['activo'] else "  —  DESACTIVADO"),
                    expanded=False,
                ):
                    _uc1, _uc2, _uc3 = st.columns([2, 1, 1])
                    _nuevo_rol = _uc1.selectbox(
                        "Rol", ROLES, index=ROLES.index(_u['rol']), key=f"rol_{_u['id']}",
                        format_func=lambda r: ROL_LABELS.get(r, r),
                    )
                    _nuevo_grp = _uc1.selectbox(
                        "Grupo de acceso", _grupo_opts, index=_grupo_opts.index(_u['grupo_id'] if _u['grupo_id'] else 0),
                        key=f"grp_{_u['id']}", format_func=_grupo_fmt, disabled=(_nuevo_rol in ("admin", "super_usuario"))
                    )
                    if _nuevo_rol != _u['rol'] or _nuevo_grp != (_u['grupo_id'] if _u['grupo_id'] else 0):
                        if _uc1.button("Guardar cambios", key=f"guardar_rol_{_u['id']}"):
                            grp_val = _nuevo_grp if _nuevo_grp != 0 else None
                            ok, msg = cambiar_rol_usuario(AUTH_DB, _u['id'], _nuevo_rol, grp_val)
                            (st.success if ok else st.error)(msg)
                            if ok: st.session_state['_refresh'] = True
                    
                    if _u['activo']:
                        if _uc2.button("Desactivar", key=f"desact_{_u['id']}",
                                        disabled=(_u['username'] == AUTH_USER.get('username'))):
                            cambiar_estado_usuario(AUTH_DB, _u['id'], False)
                            st.session_state['_refresh'] = True
                    else:
                        if _uc2.button("Reactivar", key=f"react_{_u['id']}"):
                            cambiar_estado_usuario(AUTH_DB, _u['id'], True)
                            st.session_state['_refresh'] = True
                    with _uc3.popover("Restablecer contraseña"):
                        _npw = st.text_input("Nueva contraseña", type="password", key=f"npw_{_u['id']}")
                        if st.button("Guardar", key=f"npw_btn_{_u['id']}"):
                            ok, msg = resetear_password(AUTH_DB, _u['id'], _npw)
                            (st.success if ok else st.error)(msg)
                    st.caption(f"Último acceso: {_u['ultimo_login'] or 'nunca'}")

                    # --- Pantallas permitidas (palomitas una por una) ---
                    if _u['rol'] in ("admin", "super_usuario"):
                        st.info("El administrador tiene acceso completo a todas las pantallas.")
                    else:
                        st.markdown("**Pantallas permitidas**")
                        st.caption("Selecciona las pantallas que este usuario puede consultar.")
                        _grupos_ref = {
                            "Cartera": [
                                "Dashboard & Cartera", "Estado de Cuenta", "Carga Masiva y Altas",
                                "Editar / Eliminar", "Gestor de Bajas", "Tabla Mensual por Contrato",
                                "Reporte Maestro", "Reporte Maestro Saldos",
                            ],
                            "Gestión de Riesgo": ["Gestión de Morosidad", "Eventos Especiales", "Anotaciones"],
                            "Finanzas & Contabilidad": [
                                "Pólizas Contables", "Intereses del Mes", "Facturación de Intereses",
                                "Conciliación de Facturas", "Comparar Analíticas", "Tablas de Amortización",
                            ],
                            "Análisis": [
                                "Proyección Financiera", "Análisis de Rentabilidad",
                                "Punto de Equilibrio", "Reportes por Cliente",
                            ],
                        }
                        try:
                            _perm_actual = obtener_pantallas_usuario(AUTH_DB, int(_u['id']))
                        except Exception:
                            _perm_actual = None
                        _todas_flat = [x for xs in _grupos_ref.values() for x in xs]
                        if _perm_actual is None:
                            _default_on = set(_todas_flat)
                        else:
                            _default_on = set(_perm_actual)
                        _nuevas = []
                        for _gname, _items in _grupos_ref.items():
                            st.markdown(f"*{_gname}*")
                            _cols = st.columns(2)
                            for _ii, _pname in enumerate(_items):
                                _chk = _cols[_ii % 2].checkbox(
                                    _pname,
                                    value=(_pname in _default_on),
                                    key=f"pant_{_u['id']}_{_pname}",
                                )
                                if _chk:
                                    _nuevas.append(_pname)
                        _bc1, _bc2 = st.columns(2)
                        if _bc1.button("Guardar pantallas", key=f"save_pant_{_u['id']}", type="primary"):
                            ok, msg = guardar_pantallas_usuario(AUTH_DB, int(_u['id']), _nuevas)
                            (st.success if ok else st.error)(msg)
                        if _bc2.button("Quitar restricción (ver todo su rol)", key=f"clear_pant_{_u['id']}"):
                            ok, msg = guardar_pantallas_usuario(AUTH_DB, int(_u['id']), [])
                            (st.success if ok else st.error)(
                                "Restricción eliminada: verá el menú completo de su rol." if ok else msg
                            )

            st.divider(); st.subheader("Agregar Nuevo Usuario")
            with st.form("nuevo_usuario"):
                _nu_user = st.text_input("Usuario (para iniciar sesión)")
                _nu_nombre = st.text_input("Nombre completo")
                _nu_rol = st.selectbox("Rol", ROLES, format_func=lambda r: ROL_LABELS.get(r, r))
                _nu_grp = st.selectbox("Grupo de acceso (Obligatorio si no es admin)", _grupo_opts, format_func=_grupo_fmt)
                _nu_pw1 = st.text_input("Contraseña", type="password")
                _nu_pw2 = st.text_input("Confirmar contraseña", type="password")
                if st.form_submit_button("Crear usuario"):
                    if _nu_pw1 != _nu_pw2:
                        st.error("Las contraseñas no coinciden.")
                    else:
                        grp_val = _nu_grp if _nu_grp != 0 else None
                        ok, msg = crear_usuario(AUTH_DB, _nu_user, _nu_nombre, _nu_pw1, _nu_rol, grp_val)
                        (st.success if ok else st.error)(msg)
                        if ok:
                            st.session_state['_refresh'] = True

    elif menu=="Punto de Equilibrio":
        st.title("Punto de Equilibrio")
        st.caption("Estimación del período de recuperación de inversión por contrato y cartera.")

        with st.spinner("Calculando el punto de equilibrio de cada contrato…"):
            df_pe=calcular_punto_equilibrio()
        if df_pe.empty:
            estado_vacio("Sin contratos activos para analizar",
                         "El análisis de punto de equilibrio necesita al menos un contrato ACTIVO con inversión neta positiva.")
        else:
            recuperados=df_pe['PE_Recuperado'].sum()
            pe_prom=df_pe['Mes_PE_Rentas'].mean()
            plazo_prom=df_pe['Plazo'].mean()
            pct_prom=df_pe['Pct_PE'].mean()
            inv_total=df_pe['Inversion_Neta'].sum()
            gan_total=df_pe['Ganancia'].sum()
            k1,k2,k3,k4,k5,k6=st.columns(6)
            k1.metric("Mes PE Promedio",f"{pe_prom:.1f}m")
            k2.metric("% del Plazo",    f"{pct_prom:.1f}%")
            k3.metric("Contratos c/PE", f"{recuperados}/{len(df_pe)}")
            k4.metric("Inversión Total",f"${inv_total:,.0f}")
            k5.metric("Ganancia Total", f"${gan_total:,.0f}")
            k6.metric("ROI Promedio",   f"{(gan_total/inv_total*100):.1f}%" if inv_total>0 else "N/A")

            st.markdown("---")
            ya_hoy = int(df_pe['PE_Alcanzado_Hoy'].sum())
            aun_no = len(df_pe) - ya_hoy
            pct_ya = ya_hoy/len(df_pe)*100 if len(df_pe) else 0
            hc1, hc2 = st.columns([1,1.4])
            with hc1:
                st.markdown('<span class="section-label">Contratos en punto de equilibrio hoy</span>', unsafe_allow_html=True)
                st.caption("Contratos que ya han cubierto su costo a la fecha actual.")
                hh1,hh2,hh3=st.columns(3)
                hh1.metric("Ya lo alcanzaron", f"{ya_hoy}", f"{pct_ya:.0f}% de la cartera activa")
                hh2.metric("Aún no", f"{aun_no}")
                _prom_falta = df_pe.loc[~df_pe['PE_Alcanzado_Hoy'],'Meses_Para_PE'].mean() if aun_no>0 else 0
                hh3.metric("Promedio de meses que faltan (los que aún no)", f"{_prom_falta:.1f}m" if aun_no>0 else "—")
            with hc2:
                fig_pe_hoy = px.pie(
                    pd.DataFrame({'Estado':['Ya alcanzado','Aún no'],'Contratos':[ya_hoy,aun_no]}),
                    names='Estado', values='Contratos', hole=.6,
                    color='Estado', color_discrete_map={'Ya alcanzado':C_PASTEL['success'],'Aún no':C_PASTEL['warning']},
                    title="Cartera activa por estado de equilibrio"
                )
                fig_pe_hoy.update_traces(pull=[0.06 if ya_hoy>=aun_no else 0, 0.06 if aun_no>ya_hoy else 0])
                fig_pe_hoy = sfig(fig_pe_hoy, h=260)
                fig_pe_hoy.add_annotation(
                    text=f"<b>{ya_hoy}/{ya_hoy+aun_no}</b>",
                    x=0.5, y=0.5, showarrow=False, font=dict(size=15, color="#20242B")
                )
                st.plotly_chart(fig_pe_hoy, width='stretch', key="pc_pe_hoy_pie")
            with st.expander(f"Ver los {aun_no} contratos que aún no llegan a su punto de equilibrio, ordenados por cercanía"):
                st.dataframe(
                    df_pe.loc[~df_pe['PE_Alcanzado_Hoy']].sort_values('Meses_Para_PE')
                        [['ID_Contrato','Cliente','Mes_PE_Capital','Meses_Transcurridos','Meses_Para_PE']]
                        .rename(columns={'Mes_PE_Capital':'Mes en que llega al PE','Meses_Transcurridos':'Meses ya transcurridos','Meses_Para_PE':'Meses que faltan'}),
                    width='stretch', height=220, key="df_pe_pendientes"
                )

            st.markdown("---")
            tab1,tab2,tab3,tab4=st.tabs(["Por Contrato","Curvas de Recuperación","Tabla","Análisis Estratégico"])

            with tab1:
                r1c1,r1c2=st.columns(2)
                with r1c1:
                    df_pe_s=df_pe.sort_values('Mes_PE_Rentas')
                    fig1=px.bar(df_pe_s,y='ID_Contrato',x='Mes_PE_Rentas',orientation='h',
                                 color='Mes_PE_Rentas',color_continuous_scale=[[0,C['success']],[.5,C['warning']],[1,C['accent']]],
                                 title="Mes de Punto de Equilibrio por Contrato",
                                 labels={'Mes_PE_Rentas':'Mes PE','ID_Contrato':'Contrato'},text_auto=True)
                    fig1=sfig(fig1,h=max(350,len(df_pe_s)*28))
                    fig1.update_layout(yaxis={'categoryorder':'total ascending'},coloraxis_showscale=False)
                    fig1.update_traces(hovertemplate='<b>%{y}</b><br>PE en mes %{x}')
                    st.plotly_chart(fig1,width='stretch', key="pc_044")
                    explain("¿Cuándo se recupera cada contrato?",
                        "En qué mes cada contrato recupera lo invertido.")

                with r1c2:
                    df_pe['_sz_pe']=sz(df_pe['Inversion_Neta'])
                    fig2=px.scatter(df_pe,x='Mes_PE_Rentas',y='Margen',size='_sz_pe',
                                     color='Tasa_Anual',color_continuous_scale=[[0,C['light']],[1,C['primary']]],
                                     hover_name='ID_Contrato',hover_data={'Cliente':True,'_sz_pe':False,'Plazo':True},
                                     title="PE vs Margen (tamaño = inversión neta)",
                                     labels={'Mes_PE_Rentas':'Mes de equilibrio','Margen':'Margen (%)'})
                    fig2=sfig(fig2,h=290)
                    fig2.update_traces(hovertemplate='<b>%{hovertext}</b><br>PE: mes %{x}<br>Margen: %{y:.2f}%')
                    st.plotly_chart(fig2,width='stretch', key="pc_045")
                    explain("Relación entre velocidad de equilibrio y rentabilidad",
                        "Compara qué tan rápido se recupera la inversión contra qué tan rentable es el contrato.")

                r2c1,r2c2=st.columns(2)
                with r2c1:
                    fig3=px.histogram(df_pe,x='Mes_PE_Rentas',nbins=15,color_discrete_sequence=[C['primary']],
                                       title="Distribución del Mes de Equilibrio (todos los contratos)",
                                       labels={'Mes_PE_Rentas':'Mes PE'})
                    fig3=sfig(fig3,h=280)
                    fig3.update_traces(marker_line_color='white',marker_line_width=1.5,
                                       hovertemplate='Mes %{x}: %{y} contratos')
                    fig3.add_vline(x=pe_prom,line_dash="dash",line_color=C['accent'],annotation_text=f"Prom: {pe_prom:.0f}m",
                                   annotation_font=dict(color=C['accent'],size=12))
                    st.plotly_chart(fig3,width='stretch', key="pc_046")
                    explain("Distribución de los puntos de equilibrio",
                        "En qué mes, en promedio, se recupera la inversión de los contratos.")

                with r2c2:
                    fig4=px.bar(df_pe.sort_values('Pct_PE'),y='ID_Contrato',x='Pct_PE',orientation='h',
                                 color='Pct_PE',color_continuous_scale=[[0,C['success']],[.5,C['warning']],[1,C['accent']]],
                                 title="PE como % del Plazo Total",
                                 labels={'Pct_PE':'% del plazo al PE','ID_Contrato':'Contrato'},text_auto='.0f')
                    fig4=sfig(fig4,h=max(350,len(df_pe)*26))
                    fig4.update_layout(yaxis={'categoryorder':'total ascending'},coloraxis_showscale=False)
                    fig4.add_vline(x=50,line_dash="dot",line_color=C['primary'],annotation_text="50% del plazo",
                                   annotation_font=dict(color=C['primary'],size=11))
                    fig4.update_traces(hovertemplate='<b>%{y}</b><br>PE al %{x:.0f}% del plazo')
                    st.plotly_chart(fig4,width='stretch', key="pc_047")
                    explain("¿Qué tan temprano en el plazo se recupera la inversión?",
                        "Si la inversión se recupera antes de la mitad del plazo, el resto del contrato es ganancia.")

            with tab2:
                st.subheader("Curvas de Recuperación Individual")
                contratos_sel=st.multiselect("Selecciona contratos para ver su curva (máx 5):",
                                              df_pe['ID_Contrato'].tolist(),
                                              default=df_pe['ID_Contrato'].tolist()[:min(3,len(df_pe))])
                if contratos_sel:
                    fig_curvas=go.Figure()
                    colores_l=[C['primary'],C['accent'],C['success'],C['warning'],C['purple']]
                    df_act=obtener('ACTIVO')
                    for i,id_c in enumerate(contratos_sel[:5]):
                        row=df_act[df_act['ID_Contrato']==id_c]
                        if row.empty: continue
                        row=row.iloc[0]; inv=row['Valor_Sin_IVA']-row['Anticipo_Monto']
                        if inv<=0: continue
                        dfa,_,_,_,_=calc_amort(round(inv,4),round(row['Mensualidad_Sin_IVA'],4),round(row['Residual_Monto'],4),int(row['Plazo']),round(row['Tasa_Calculada'],8))
                        acum=dfa['Capital'].cumsum().values
                        color_i=colores_l[i%len(colores_l)]
                        pe_idx=next((j for j,v in enumerate(acum) if v>=inv),len(acum)-1)
                        fig_curvas.add_trace(go.Scatter(
                            x=list(range(1,len(acum)+1)),y=acum,mode='lines',name=f"{id_c} ({row['Cliente'][:12]})",
                            line=dict(color=color_i,width=2.5),
                            hovertemplate=f'<b>{id_c}</b><br>Mes %{{x}}<br>Capital recuperado: $%{{y:,.2f}}'))
                        if pe_idx<len(acum):
                            fig_curvas.add_scatter(x=[pe_idx+1],y=[acum[pe_idx]],mode='markers',
                                                   marker=dict(size=12,color=color_i,symbol='star',line=dict(width=2,color='white')),
                                                   showlegend=False,hovertemplate=f'PE mes {pe_idx+1}<br>${{acum[pe_idx]:,.0f}}')
                    max_inv=df_pe.loc[df_pe['ID_Contrato'].isin(contratos_sel),'Inversion_Neta'].max()
                    fig_curvas.add_hline(y=max_inv,line_dash="dash",line_color="rgba(55,48,163,.35)",
                                         annotation_text="Inversión máx.",annotation_font=dict(size=11))
                    fig_curvas.update_layout(title="Curvas de Recuperación de Capital — Contratos Seleccionados",
                                              xaxis_title="Mes del contrato",yaxis_title="Capital amortizado acumulado (MXN)",
                                              paper_bgcolor="rgba(0,0,0,0)",plot_bgcolor="rgba(0,0,0,0)",
                                              font=dict(family="Segoe UI, Helvetica Neue, Arial, sans-serif"),margin=dict(t=50,r=16,b=36,l=16),height=320)
                    st.plotly_chart(fig_curvas,width='stretch', key="pc_048")
                    explain("Curvas de recuperación comparativas",
                        "Compara qué tan rápido distintos contratos recuperan la inversión.")

            with tab3:
                df_pe_tabla = df_pe.copy()
                df_pe_tabla['¿Ya llegó a su PE?'] = df_pe_tabla['PE_Alcanzado_Hoy'].map({True:'Sí', False:'Aún no'})
                titled_table("Tabla de Punto de Equilibrio por Contrato",
                    df_pe_tabla[['ID_Contrato','Cliente','Vehiculo','Inversion_Neta','Renta','Plazo',
                            'Mes_PE_Rentas','Pct_PE','¿Ya llegó a su PE?','Meses_Transcurridos','Ganancia','Margen','TIR Anual','Tasa_Anual']],
                    fmt_dict={'Inversion_Neta':'${:,.2f}','Renta':'${:,.2f}','Ganancia':'${:,.2f}',
                              'Margen':'{:.2f}%','TIR Anual':'{:.2f}%','Tasa_Anual':'{:.2f}%','Pct_PE':'{:.1f}%'},
                    cmap_col='Mes_PE_Rentas')
                buf = excel_con_formato({'Punto_Equilibrio': df_pe},
                    currency_cols=['Inversion_Neta','Renta','Ganancia'],
                    pct_cols=['Margen','TIR Anual','Tasa_Anual','Pct_PE'])
                st.download_button("Descargar Excel",buf,"punto_equilibrio.xlsx")

            with tab4:
                st.subheader("Análisis Estratégico de Equilibrio")
                r1c1,r1c2=st.columns(2)
                with r1c1:
                    tot_inv=df_pe['Inversion_Neta'].sum()
                    tot_rentas=sum(r['Renta']*r['Plazo'] for _,r in df_pe.iterrows())
                    tot_res=df_pe['Residual'].sum(); tot_gan=df_pe['Ganancia'].sum()
                    fw=go.Figure(go.Waterfall(
                        measure=["absolute","relative","relative","total"],
                        x=["Inversión Total","Ingresos x Rentas","+ Residuales","= Ganancia Neta"],
                        y=[-tot_inv, tot_rentas, tot_res, 0],
                        text=[f"${-tot_inv:,.0f}",f"${tot_rentas:,.0f}",f"${tot_res:,.0f}",f"${tot_gan:,.0f}"],
                        textposition="outside",textfont=dict(family="Segoe UI, Helvetica Neue, Arial, sans-serif",size=11),
                        increasing=dict(marker_color=C_PASTEL['success']),decreasing=dict(marker_color=C_PASTEL['accent']),
                        totals=dict(marker_color=C_PASTEL['primary'])))
                    fw.update_layout(title="Cascada P&L Global de la Cartera",paper_bgcolor="rgba(0,0,0,0)",
                                     plot_bgcolor="rgba(0,0,0,0)",font=dict(family="Segoe UI, Helvetica Neue, Arial, sans-serif"),
                                     margin=dict(t=50,r=16,b=36,l=16),height=300)
                    st.plotly_chart(fw,width='stretch', key="pc_049")
                    explain("P&L simplificado de toda la cartera",
                        "Resumen de inversión, ingresos y ganancia de toda la cartera.")

                with r1c2:
                    df_tir2=df_pe.dropna(subset=['TIR Anual'])
                    if not df_tir2.empty:
                        df_tir2['_sz5']=sz(df_tir2['Inversion_Neta'])
                        fig5=px.scatter(df_tir2,x='TIR Anual',y='Mes_PE_Rentas',size='_sz5',
                                         color='Margen',color_continuous_scale=[[0,C['light']],[1,C['primary']]],
                                         hover_name='ID_Contrato',hover_data={'Cliente':True,'_sz5':False},
                                         title="TIR vs Mes de Equilibrio (ideal: derecha-abajo)",
                                         labels={'TIR Anual':'TIR Anual (%)','Mes_PE_Rentas':'Mes PE'})
                        fig5=sfig(fig5,h=300)
                        fig5.update_traces(hovertemplate='<b>%{hovertext}</b><br>TIR: %{x:.2f}%<br>PE mes: %{y}')
                        tir_med=df_tir2['TIR Anual'].median(); pe_med=df_tir2['Mes_PE_Rentas'].median()
                        fig5.add_vline(x=tir_med,line_dash="dot",line_color="rgba(55,48,163,.3)",annotation_text="TIR med.",annotation_font_size=10)
                        fig5.add_hline(y=pe_med,line_dash="dot",line_color="rgba(55,48,163,.3)",annotation_text="PE med.",annotation_font_size=10)
                        st.plotly_chart(fig5,width='stretch', key="pc_050")
                        explain("Cuadrante estratégico: TIR vs Velocidad de equilibrio",
                        "Compara qué tan rentable es un contrato contra qué tan rápido se recupera la inversión.")

                df_resumen=pd.DataFrame([{
                    'Indicador': 'Inversión neta total',         'Valor': f"${tot_inv:,.2f}"},
                    {'Indicador': 'Ganancia neta proyectada',    'Valor': f"${tot_gan:,.2f}"},
                    {'Indicador': 'ROI de la cartera',           'Valor': f"{tot_gan/tot_inv*100:.2f}%"},
                    {'Indicador': 'Mes PE promedio (rentas)',     'Valor': f"{pe_prom:.1f} meses"},
                    {'Indicador': '% del plazo al PE',           'Valor': f"{pct_prom:.1f}%"},
                    {'Indicador': 'Contratos con PE dentro plazo','Valor':f"{recuperados} de {len(df_pe)}"},
                    {'Indicador': 'TIR Anual promedio de cartera', 'Valor': f"{df_pe['TIR Anual'].dropna().mean():.2f}%"},
                    {'Indicador': 'Margen promedio',             'Valor': f"{df_pe['Margen'].mean():.2f}%"},
                ])
                titled_table("Resumen Ejecutivo de Rentabilidad y Equilibrio",df_resumen)

    # Tablas de amortización (leasing + residual, todos los contratos)
    elif menu == "Tablas de Amortización":
        st.title("Tablas de Amortización")
        st.caption("Consulta la amortización de leasing y acumulación de residual por contrato.")

        df_act = obtener()
        if df_act.empty:
            st.info("No hay contratos registrados.")
        else:
            opciones = ["— Todos los contratos —"] + df_act['ID_Contrato'].tolist()
            sel = st.selectbox(
                "Contrato",
                opciones,
                format_func=lambda x: x if x.startswith("—") else
                    f"{x}  —  {df_act.loc[df_act['ID_Contrato']==x,'Cliente'].values[0]}  |  "
                    f"{df_act.loc[df_act['ID_Contrato']==x,'Vehiculo'].values[0]}"
                    + ("  · BAJA" if df_act.loc[df_act['ID_Contrato']==x,'Estatus'].values[0] == 'BAJA' else ""),
                key="amort_sel_contrato"
            )

            tab_lea, tab_res = st.tabs(["Amortización Leasing", "Acumulación Residual"])

            contratos_sel = df_act if sel.startswith("—") else df_act[df_act['ID_Contrato'] == sel]

            with tab_lea:
                st.caption("Amortización mensual de capital e intereses.")
                frames_lea = []
                for _, row in contratos_sel.iterrows():
                    inv = row['Valor_Sin_IVA'] - row['Anticipo_Monto']
                    if inv <= 0:
                        continue
                    pl = int(row['Plazo']); t = round(row['Tasa_Calculada'], 8)
                    r  = round(row['Mensualidad_Sin_IVA'], 4)
                    res= round(row['Residual_Monto'], 4)
                    fa = pd.to_datetime(row['Fecha_Alta'])
                    dfa, _, _, _, _ = calc_amort(round(inv, 4), r, res, pl, t)
                    dfa = dfa.copy()
                    dfa.insert(0, 'ID_Contrato', row['ID_Contrato'])
                    dfa.insert(1, 'Cliente',     row['Cliente'])
                    dfa.insert(2, 'Vehiculo',    row['Vehiculo'])
                    dfa['Fecha'] = dfa['Mes'].apply(lambda m: (fa + relativedelta(months=m)).strftime('%Y-%m'))
                    # Si el contrato ya se dio de baja, la tabla se corta en ese mes —
                    # no tiene sentido seguir mostrando amortización de meses en los
                    # que el contrato ya no estaba vigente.
                    if str(row.get('Estatus','')).upper() == 'BAJA' and pd.notna(row.get('Fecha_Baja')):
                        _fb = pd.to_datetime(row['Fecha_Baja'])
                        mes_baja = max(0, (_fb.year - fa.year) * 12 + (_fb.month - fa.month))
                        dfa = dfa[dfa['Mes'] <= max(mes_baja, 0)]
                    if dfa.empty:
                        continue
                    saldos_ini = [round(inv, 4)] + list(dfa['Saldo'].iloc[:-1].round(4))
                    dfa.insert(dfa.columns.get_loc('Interes'), 'Saldo_Ini', saldos_ini)
                    dfa.rename(columns={'Saldo': 'Saldo_Fin'}, inplace=True)
                    frames_lea.append(dfa)

                if frames_lea:
                    df_lea = pd.concat(frames_lea, ignore_index=True)
                    cols_show_lea = ['ID_Contrato','Cliente','Vehiculo','Fecha','Mes',
                                     'Saldo_Ini','Interes','Capital','Saldo_Fin']
                    df_lea = df_lea[cols_show_lea]

                    if len(contratos_sel) == 1:
                        fig_lea = go.Figure()
                        fig_lea.add_trace(go.Bar(
                            x=df_lea['Fecha'], y=df_lea['Capital'],
                            name='Capital', marker_color=C_PASTEL['primary']))
                        fig_lea.add_trace(go.Bar(
                            x=df_lea['Fecha'], y=df_lea['Interes'],
                            name='Interés', marker_color=C_PASTEL['accent']))
                        fig_lea.add_trace(go.Scatter(
                            x=df_lea['Fecha'], y=df_lea['Saldo_Fin'],
                            name='Saldo', mode='lines+markers',
                            line=dict(color=C_PASTEL['success'], width=2.5),
                            yaxis='y2'))
                        fig_lea.update_layout(
                            barmode='stack', title='Amortización Leasing — Capital e Interés por Mes',
                            xaxis_title='Mes', yaxis_title='Monto (MXN)',
                            yaxis2=dict(title='Saldo (MXN)', overlaying='y', side='right', showgrid=False),
                            legend=dict(orientation='h', y=1.08))
                        fig_lea = sfig(fig_lea, h=340)
                        st.plotly_chart(fig_lea, width='stretch', key=f"pc_051_{sel}")

                    fmt_lea = {c: '${:,.4f}' for c in ['Saldo_Ini','Interes','Capital','Saldo_Fin']}
                    st.dataframe(
                        df_lea.style.format(fmt_lea),
                        width='stretch', height=420,
        key=f"df_027_{sel}")

                    k1, k2, k3, k4 = st.columns(4)
                    k1.metric("Total Capital Amortizado", f"${df_lea['Capital'].sum():,.2f}")
                    k2.metric("Total Intereses Leasing",  f"${df_lea['Interes'].sum():,.2f}")
                    k3.metric("Contratos incluidos",       str(df_lea['ID_Contrato'].nunique()))
                    k4.metric("Total mensualidades",       str(len(df_lea)))

                    buf_lea = excel_con_formato({'Amort_Leasing': df_lea},
                        currency_cols=['Saldo_Ini','Interes','Capital','Saldo_Fin'])
                    st.download_button(
                        "Descargar Excel — Amortización Leasing",
                        buf_lea,
                        f"amort_leasing{'_'+sel if not sel.startswith('—') else '_todos'}.xlsx"
                    )
                else:
                    st.info("Sin datos para los contratos seleccionados.")

            with tab_res:
                st.caption("Evolución mensual del valor residual hasta el vencimiento.")
                frames_res = []
                for _, row in contratos_sel.iterrows():
                    inv = row['Valor_Sin_IVA'] - row['Anticipo_Monto']
                    if inv <= 0:
                        continue
                    pl  = int(row['Plazo']); t = round(row['Tasa_Calculada'], 8)
                    res = round(row['Residual_Monto'], 4)
                    fa  = pd.to_datetime(row['Fecha_Alta'])
                    vpr = vp_res(res, t, pl)
                    dfr = calc_res_amort(round(vpr, 4), t, pl).copy()
                    dfr.insert(0, 'ID_Contrato', row['ID_Contrato'])
                    dfr.insert(1, 'Cliente',     row['Cliente'])
                    dfr.insert(2, 'Vehiculo',    row['Vehiculo'])
                    dfr['Fecha']           = dfr['Mes'].apply(lambda m: (fa + relativedelta(months=m)).strftime('%Y-%m'))
                    dfr['Residual_Pactado']= res
                    dfr['VP_Residual']     = round(vpr, 4)
                    if str(row.get('Estatus','')).upper() == 'BAJA' and pd.notna(row.get('Fecha_Baja')):
                        _fb2 = pd.to_datetime(row['Fecha_Baja'])
                        mes_baja2 = max(0, (_fb2.year - fa.year) * 12 + (_fb2.month - fa.month))
                        dfr = dfr[dfr['Mes'] <= max(mes_baja2, 0)]
                    if dfr.empty:
                        continue
                    frames_res.append(dfr)

                if frames_res:
                    df_res = pd.concat(frames_res, ignore_index=True)
                    cols_show_res = ['ID_Contrato','Cliente','Vehiculo','Fecha','Mes',
                                     'VP_Residual','Saldo_Ini','Interes','Saldo_Fin','Residual_Pactado']
                    df_res = df_res[cols_show_res]

                    if len(contratos_sel) == 1:
                        fig_res = go.Figure()
                        fig_res.add_trace(go.Scatter(
                            x=df_res['Fecha'], y=df_res['Saldo_Fin'],
                            name='Saldo acumulado', fill='tozeroy',
                            mode='lines', line=dict(color=C_PASTEL['gold'], width=2.5)))
                        fig_res.add_hline(
                            y=df_res['Residual_Pactado'].iloc[0],
                            line_dash='dash', line_color=C['accent'],
                            annotation_text=f"Residual pactado ${df_res['Residual_Pactado'].iloc[0]:,.2f}",
                            annotation_font=dict(size=11))
                        fig_res.add_hline(
                            y=df_res['VP_Residual'].iloc[0],
                            line_dash='dot', line_color=C['info'],
                            annotation_text=f"VP inicial ${df_res['VP_Residual'].iloc[0]:,.4f}",
                            annotation_font=dict(size=11))
                        fig_res.update_layout(
                            title='Acumulación del Residual — VP a Residual Pactado',
                            xaxis_title='Mes', yaxis_title='Saldo Residual (MXN)')
                        fig_res = sfig(fig_res, h=320)
                        st.plotly_chart(fig_res, width='stretch', key=f"pc_052_{sel}")

                    fmt_res = {c: '${:,.4f}' for c in ['VP_Residual','Saldo_Ini','Interes','Saldo_Fin','Residual_Pactado']}
                    st.dataframe(
                        df_res.style.format(fmt_res),
                        width='stretch', height=420,
        key=f"df_028_{sel}")

                    k1r, k2r, k3r, k4r = st.columns(4)
                    k1r.metric("Total Intereses Residual", f"${df_res['Interes'].sum():,.2f}")
                    k2r.metric("VP Residual promedio",     f"${df_res.groupby('ID_Contrato')['VP_Residual'].first().mean():,.4f}")
                    k3r.metric("Contratos incluidos",       str(df_res['ID_Contrato'].nunique()))
                    k4r.metric("Residual pactado promedio", f"${df_res.groupby('ID_Contrato')['Residual_Pactado'].first().mean():,.2f}")

                    _hojas_res = {'Acum_Residual': df_res}
                    if frames_lea:
                        _hojas_res['Amort_Leasing'] = df_lea
                    buf_res = excel_con_formato(
                        _hojas_res,
                        currency_cols={
                            'Acum_Residual': ['VP_Residual','Saldo_Ini','Interes','Saldo_Fin','Residual_Pactado'],
                            'Amort_Leasing': ['Saldo_Ini','Interes','Capital','Saldo_Fin'],
                        },
                    )
                    st.download_button(
                        "Descargar Excel — Residual (+ Leasing en hoja 2)",
                        buf_res,
                        f"amort_residual{'_'+sel if not sel.startswith('—') else '_todos'}.xlsx"
                    )
                else:
                    st.info("Sin datos para los contratos seleccionados.")

    # CONCILIACIÓN DE FACTURAS (NUEVO)
    elif menu == "Conciliación de Facturas":
        try:
            _render_conciliacion()
        except Exception as _e_conc:
            st.error("Error al abrir Conciliación de Facturas. Detalle técnico abajo.")
            st.exception(_e_conc)

    elif menu == "Comparar Analíticas":
        _render_analiticas()
except Exception as _e_fatal:
    _pantalla_error_amigable(_e_fatal, contexto=str(st.session_state.get("menu_item","?")))
