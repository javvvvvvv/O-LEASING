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
from core.cfdi import clasificar_concepto
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
)
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
            fa=row['Fecha_Alta']; fv=row['Fecha_Vencimiento']
            if fp<pd.Timestamp(fa.year,fa.month,1) or fp>pd.Timestamp(fv.year,fv.month,1): continue
            mc=(cursor.year-fa.year)*12+(cursor.month-fa.month)+1
            if mc<1 or mc>row['Plazo']: continue
            inv=row['Valor_Sin_IVA']-row['Anticipo_Monto']
            dfa,_,_,_,_=calc_amort(round(inv,4),round(row['Mensualidad_Sin_IVA'],4),round(row['Residual_Monto'],4),int(row['Plazo']),round(row['Tasa_Calculada'],8))
            dfr=calc_res_amort(round(float(row['VP_Residual']),4),round(row['Tasa_Calculada'],8),int(row['Plazo']))
            ti+=dfa.iloc[mc-1]['Interes']; tr+=dfr.iloc[mc-1]['Interes']
            tc+=row['Comision_Monto']/row['Plazo']; cnt+=1
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
    """Calcula cuánto vale un concepto para este contrato en (mes, anio),
    usando las mismas fórmulas de amortización que ya usa el resto de la
    app. Regresa 0.0 si el concepto no aplica ese mes (p. ej. Anticipo
    solo aplica en el mes de alta)."""
    fa = pd.to_datetime(con['Fecha_Alta'])
    pl = int(con['Plazo'])
    mc = (anio - fa.year) * 12 + (mes - fa.month) + 1
    inv = con['Valor_Sin_IVA'] - con['Anticipo_Monto']
    r, t, res = con['Mensualidad_Sin_IVA'], con['Tasa_Calculada'], con['Residual_Monto']

    if concepto in ('ANTICIPO_INICIAL', 'RESIDUAL_PACTADO'):
        if mc != 1:
            return 0.0
        return float(con['Anticipo_Monto']) if concepto == 'ANTICIPO_INICIAL' else float(res)

    if mc < 1 or mc > pl:
        return 0.0

    if concepto == 'COMISION_MES':
        com = con.get('Comision_Monto', 0) or 0
        return round(com / pl, 2) if pl else 0.0

    if concepto in ('INTERES_MES', 'CAPITAL_MES', 'RENTA_TOTAL_MES'):
        dfa, _, _, _, _ = calc_amort(round(inv, 4), round(r, 4), round(res, 4), pl, round(t, 8))
        fila = dfa.iloc[mc - 1]
        if concepto == 'INTERES_MES': return round(float(fila['Interes']), 2)
        if concepto == 'CAPITAL_MES': return round(float(fila['Capital']), 2)
        return round(float(fila['Interes']) + float(fila['Capital']), 2)

    if concepto == 'RESIDUAL_INTERES_MES':
        vpr = vp_res(res, t, pl)
        dfr = calc_res_amort(round(vpr, 4), round(t, 8), pl)
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
        df['Fecha_Alta']=pd.to_datetime(df['Fecha_Alta'])
        df['Fecha_Vencimiento']=pd.to_datetime(df['Fecha_Vencimiento'])
        df['Fecha_Baja']=pd.to_datetime(df['Fecha_Baja'],errors='coerce')
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
    df=obtener('ACTIVO')
    if df.empty: return pd.DataFrame(),{}
    rows=[]; fp=pd.Timestamp(anio,mes,1)
    for _,row in df.iterrows():
        fa=row['Fecha_Alta']; fv=row['Fecha_Vencimiento']
        if fp<pd.Timestamp(fa.year,fa.month,1) or fp>pd.Timestamp(fv.year,fv.month,1): continue
        mc=(anio-fa.year)*12+(mes-fa.month)+1
        if mc<1 or mc>row['Plazo']: continue
        inv=row['Valor_Sin_IVA']-row['Anticipo_Monto']
        dfa,_,_,_,_=calc_amort(round(inv,4),round(row['Mensualidad_Sin_IVA'],4),
                                round(row['Residual_Monto'],4),int(row['Plazo']),round(row['Tasa_Calculada'],8))
        dfr=calc_res_amort(round(float(row['VP_Residual']),4),round(row['Tasa_Calculada'],8),int(row['Plazo']))
        fa_=dfa.iloc[mc-1]; fr_=dfr.iloc[mc-1]
        rows.append({'ID_Contrato':row['ID_Contrato'],'Cliente':row['Cliente'],'Vehiculo':row['Vehiculo'],
                     'Mes_Cont':mc,'Plazo':int(row['Plazo']),'Renta':row['Mensualidad_Sin_IVA'],
                     'Int_Leasing':round(fa_['Interes'],2),'Capital':round(fa_['Capital'],2),
                     'Saldo_Cap':round(fa_['Saldo'],2),'Int_Residual':round(fr_['Interes'],2),
                     'Amort_Com':round(row['Comision_Monto']/row['Plazo'],2),
                     'Tasa_Anual_Pct':round(row['Tasa_Calculada']*1200,4)})
    df2=pd.DataFrame(rows)
    if df2.empty: return df2,{}
    df2['Total_Int']=df2['Int_Leasing']+df2['Int_Residual']
    m={'il':df2['Int_Leasing'].sum(),'ir':df2['Int_Residual'].sum(),
       'tot':df2['Total_Int'].sum(),'cap':df2['Capital'].sum(),
       'com':df2['Amort_Com'].sum(),'n':len(df2),
       'top':df2.groupby('Cliente')['Total_Int'].sum().idxmax() if len(df2)>0 else 'N/A'}
    return df2.sort_values('Total_Int',ascending=False),m

def proy_intereses(meses=12):
    df=obtener('ACTIVO')
    if df.empty: return pd.DataFrame()
    hoy=hoy_ref(); rows=[]; bar=st.progress(0,"Proyectando intereses…")
    for i in range(meses):
        fd=hoy+relativedelta(months=i+1); mp,ap=fd.month,fd.year; fp=pd.Timestamp(ap,mp,1)
        ti=tr=0.0; cnt=0
        for _,row in df.iterrows():
            fa=row['Fecha_Alta']; fv=row['Fecha_Vencimiento']
            if fp<pd.Timestamp(fa.year,fa.month,1) or fp>pd.Timestamp(fv.year,fv.month,1): continue
            mc=(ap-fa.year)*12+(mp-fa.month)+1
            if mc<1 or mc>row['Plazo']: continue
            inv=row['Valor_Sin_IVA']-row['Anticipo_Monto']
            dfa,_,_,_,_=calc_amort(round(inv,4),round(row['Mensualidad_Sin_IVA'],4),
                                    round(row['Residual_Monto'],4),int(row['Plazo']),round(row['Tasa_Calculada'],8))
            dfr=calc_res_amort(round(float(row['VP_Residual']),4),round(row['Tasa_Calculada'],8),int(row['Plazo']))
            ti+=dfa.iloc[mc-1]['Interes']; tr+=dfr.iloc[mc-1]['Interes']; cnt+=1
        rows.append({'Mes':fd.strftime('%Y-%m'),'Label':fd.strftime('%b %Y'),'IL':round(ti,2),'IR':round(tr,2),'Total':round(ti+tr,2),'N':cnt})
        bar.progress((i+1)/meses,text=f"Mes {i+1}/{meses}…")
    bar.empty(); return pd.DataFrame(rows)

def tabla_rentas_mensuales(anio):
    df=obtener('ACTIVO')
    if df.empty: return pd.DataFrame()
    MN=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
    data=[]
    for _,row in df.iterrows():
        fila={'ID_Contrato':row['ID_Contrato'],'Cliente':row['Cliente']}
        for i,n in enumerate(MN,1):
            fm=pd.Timestamp(anio,i,1)
            fila[n]=row['Mensualidad_Sin_IVA'] if row['Fecha_Alta']<=fm<=row['Fecha_Vencimiento'] else 0.0
        data.append(fila)
    return pd.DataFrame(data).set_index('ID_Contrato')

def tabla_mensual_conceptos(anio):
    df=obtener()
    if df.empty: return None,None,None,None,"Sin contratos registrados"
    MN=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
    ids=df['ID_Contrato'].tolist()
    di=pd.DataFrame(index=ids,columns=range(1,13),dtype=float); dr=di.copy(); dc=di.copy(); ds=di.copy()
    bar=st.progress(0,"Calculando...")
    tot=len(df)
    for i,(_,row) in enumerate(df.iterrows()):
        id_c=row['ID_Contrato']; fa=row['Fecha_Alta']; pl=int(row['Plazo'])
        t=round(row['Tasa_Calculada'],8); inv=row['Valor_Sin_IVA']-row['Anticipo_Monto']
        try:
            dfa,_,_,_,_=calc_amort(round(inv,4),round(row['Mensualidad_Sin_IVA'],4),round(row['Residual_Monto'],4),pl,t)
            dfr=calc_res_amort(round(float(row['VP_Residual']),4),t,pl)
        except: continue
        es_baja = str(row.get('Estatus','')).upper() == 'BAJA'
        fecha_baja = pd.to_datetime(row['Fecha_Baja']) if es_baja and pd.notna(row.get('Fecha_Baja')) else None
        for mes in range(1,13):
            fm=pd.Timestamp(anio,mes,1)
            if fm<pd.Timestamp(fa.year,fa.month,1) or fm>row['Fecha_Vencimiento']: continue
            if fecha_baja is not None and fm>pd.Timestamp(fecha_baja.year,fecha_baja.month,1): continue
            mc=(anio-fa.year)*12+(mes-fa.month)+1
            if mc<1 or mc>pl: continue
            di.at[id_c,mes]=round(dfa.iloc[mc-1]['Interes'],2)
            dr.at[id_c,mes]=round(dfr.iloc[mc-1]['Interes'],2)
            dc.at[id_c,mes]=round(row['Comision_Monto']/pl,2)
            ds.at[id_c,mes]=round(dfr.iloc[mc-1]['Saldo_Fin'],2)
        bar.progress((i+1)/tot,text=f"Procesando {i+1}/{tot}...")
    bar.empty()
    for d in [di,dr,dc,ds]: 
        d.columns=MN
        d.dropna(how='all',inplace=True)
    return di,dr,dc,ds,None

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
        id_c=row['ID_Contrato']; fa=row['Fecha_Alta']; pl=int(row['Plazo'])
        t=round(row['Tasa_Calculada'],8); inv=row['Valor_Sin_IVA']-row['Anticipo_Monto']
        try:
            dfa,_,_,_,_=calc_amort(round(inv,4),round(row['Mensualidad_Sin_IVA'],4),round(row['Residual_Monto'],4),pl,t)
            dfr=calc_res_amort(round(float(row['VP_Residual']),4),t,pl)
        except: continue
        es_baja = str(row.get('Estatus','')).upper() == 'BAJA'
        fecha_baja = pd.to_datetime(row['Fecha_Baja']) if es_baja and pd.notna(row.get('Fecha_Baja')) else None
        
        interes_array = dfa['Interes'].values
        rem_int = np.zeros(pl)
        for idx in range(pl):
            rem_int[idx] = np.sum(interes_array[idx+1:])
            
        for mes in range(1,13):
            fm=pd.Timestamp(anio,mes,1)
            if fm < pd.Timestamp(fa.year,fa.month,1): continue
                
            if fecha_baja is not None and fm > pd.Timestamp(fecha_baja.year,fecha_baja.month,1):
                d_cap.at[id_c,mes] = 0; d_int.at[id_c,mes] = 0; d_tot.at[id_c,mes] = 0; d_res.at[id_c,mes] = 0
                continue
                
            mc = (anio-fa.year)*12 + (mes-fa.month) + 1
            if mc > pl:
                d_cap.at[id_c,mes] = 0; d_int.at[id_c,mes] = 0; d_tot.at[id_c,mes] = 0; d_res.at[id_c,mes] = 0
                continue
                
            idx = mc - 1
            saldo_cap = dfa.iloc[idx]['Saldo']
            saldo_int = rem_int[idx]
            saldo_res = dfr.iloc[idx]['Saldo_Fin']
            
            d_cap.at[id_c,mes] = round(saldo_cap, 2)
            d_int.at[id_c,mes] = round(saldo_int, 2)
            d_tot.at[id_c,mes] = round(saldo_cap + saldo_int, 2)
            d_res.at[id_c,mes] = round(saldo_res, 2)
            
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
        id_c=row['ID_Contrato']; fa=row['Fecha_Alta']; pl=int(row['Plazo'])
        t=round(row['Tasa_Calculada'],8); inv=row['Valor_Sin_IVA']-row['Anticipo_Monto']
        try:
            dfa,_,_,_,_=calc_amort(round(inv,4),round(row['Mensualidad_Sin_IVA'],4),round(row['Residual_Monto'],4),pl,t)
            dfr=calc_res_amort(round(float(row['VP_Residual']),4),t,pl)
        except: continue
        es_baja = str(row.get('Estatus','')).upper() == 'BAJA'
        fecha_baja = pd.to_datetime(row['Fecha_Baja']) if es_baja and pd.notna(row.get('Fecha_Baja')) else None
        for mes in range(1,13):
            fm=pd.Timestamp(anio,mes,1)
            if fm<pd.Timestamp(fa.year,fa.month,1) or fm>row['Fecha_Vencimiento']: continue
            if fecha_baja is not None and fm>pd.Timestamp(fecha_baja.year,fecha_baja.month,1): continue
            mc=(anio-fa.year)*12+(mes-fa.month)+1
            if mc<1 or mc>pl: continue
            
            interes_leasing = round(dfa.iloc[mc-1]['Interes'],2)
            capital_leasing = round(dfa.iloc[mc-1]['Capital'],2)
            
            di.at[id_c,mes] = interes_leasing
            dcap.at[id_c,mes] = capital_leasing
            drenta.at[id_c,mes] = interes_leasing + capital_leasing
            dr.at[id_c,mes] = round(dfr.iloc[mc-1]['Interes'],2)
            dc.at[id_c,mes] = round(row['Comision_Monto']/pl,2)
            ds.at[id_c,mes] = round(dfr.iloc[mc-1]['Saldo_Fin'],2)
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
    parse_cfdi, clasificar_concepto, normalizar_contrato,
    REGLAS_CONCEPTO_DEFAULT, MAX_XML_BYTES, NS,
)
from core.alias import resolver_numero_contrato, decidir_accion_alias, sugerir_contrato_similar

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
        return _res('NO_APLICA',
                    f"Este CFDI es de tipo {_nombres_comp.get(tipo_comp, tipo_comp)}, no una factura de "
                    f"ingreso — no se compara automáticamente contra la renta esperada.",
                    esperado=0.0, facturado=fact.get('total', 0.0), dif=0.0, detalle=[])

    # Blindaje: si el CFDI no está en pesos, comparar los montos tal cual
    # contra la renta (que sí está en MXN) daría una discrepancia falsa o,
    # peor, un "cuadre" casual sin sentido. Se manda a revisión manual.
    moneda_cfdi = (fact.get('moneda') or 'MXN').upper()
    if moneda_cfdi not in ('MXN', ''):
        return _res('ERROR',
                    f"Este CFDI está emitido en {moneda_cfdi}, no en pesos (MXN) — no se puede comparar "
                    f"directamente contra la renta pactada sin convertir primero. Revísalo a mano.",
                    esperado=0.0, facturado=fact.get('total', 0.0), dif=0.0, detalle=[])

    id_c = fact.get('id_contrato')
    if not id_c:
        return _res('SIN_CONTRATO', 'No se pudo extraer el número de contrato del XML')

    # Si la factura es de tipo OTRO, no se concilia
    if fact.get('tipo') == 'OTRO':
        return _res('NO_APLICA', 'Factura de servicios no relacionados con leasing (seguro, gestoría, etc.)',
                    esperado=0.0, facturado=fact.get('total', 0.0), dif=0.0, detalle=[])

    conn = get_db()
    con = _cache.get('contratos', {}).get(id_c)
    if con is None:
        row = conn.execute("SELECT * FROM contratos WHERE ID_Contrato=?", (id_c,)).fetchone()
        if not row:
            return _res('SIN_CONTRATO', f'Contrato {id_c} no encontrado en la base de datos')
        con = dict(row)
        _cache.setdefault('contratos', {})[id_c] = con

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

    mes = fact.get('mes_contrato')
    if not mes or mes < 1 or mes > pl:
        return _res('ERROR', f'Mes del contrato ({mes}) fuera de rango 1-{pl}')

    # Blindaje: ¿el "mes X de Y" que trae el texto del concepto coincide con
    # el mes que le tocaría según la fecha real de la factura? Si alguien se
    # equivocó al capturar el folio/concepto (p. ej. puso "14/60" cuando por
    # fecha le tocaba el mes 15), la comparación de montos usaría la fila
    # equivocada de la tabla de amortización sin que nadie se diera cuenta.
    aviso_mes = None
    if fecha_factura is not None and fact.get('tipo') == 'MENSUAL':
        try:
            fa_ref = pd.to_datetime(con['Fecha_Alta'])
            mes_por_fecha = (fecha_factura.year - fa_ref.year) * 12 + (fecha_factura.month - fa_ref.month) + 1
            if 1 <= mes_por_fecha <= pl and abs(mes_por_fecha - mes) >= 1:
                aviso_mes = (
                    f"El concepto de la factura dice que es el mes {mes} de {pl}, pero por su fecha "
                    f"({fecha_factura.strftime('%m/%Y')}) debería ser el mes {mes_por_fecha} — revisa si el "
                    f"folio o el concepto tienen un error de captura."
                )
        except Exception:
            pass

    conceptos = fact['conceptos']
    fila = dfa.iloc[mes - 1]
    renta_pura_esp = round(fila['Interes'] + fila['Capital'], 2)
    renta_total_esp = round(con['Mensualidad_Sin_IVA'], 2)

    if fact['tipo'] == 'MENSUAL':
        renta_fac   = round(conceptos['RENTA'] + conceptos['ADMIN'] + conceptos['GEOLOC'], 2)
        dif         = round(renta_fac - renta_total_esp, 2)
        errores     = []
        if abs(dif) > TOLERANCIA:
            errores.append(
                f"Renta+Admin+Geoloc: esperado ${renta_total_esp:,.2f}, "
                f"facturado ${renta_fac:,.2f} (dif ${dif:+,.2f})"
            )
        if aviso_mes:
            errores.append(aviso_mes)
        status = 'CONCILIADO' if not errores else 'DISCREPANCIA'
        return _res(status, '; '.join(errores) if errores else 'OK',
                    esperado=renta_total_esp, facturado=renta_fac, dif=dif, detalle=errores)
    else:  # ANTICIPO
        errores = []
        ant_esp = round(con['Anticipo_Monto'], 2)
        ant_fac = round(conceptos['ANTICIPO'], 2)
        dif_ant = round(ant_fac - ant_esp, 2)
        if abs(dif_ant) > TOLERANCIA:
            errores.append(f"Anticipo: esperado ${ant_esp:,.2f}, facturado ${ant_fac:,.2f} (dif ${dif_ant:+,.2f})")

        com_esp = round(con['Comision_Monto'], 2)
        com_fac = round(conceptos['COMISION'], 2)
        dif_com = round(com_fac - com_esp, 2)
        if abs(dif_com) > TOLERANCIA:
            errores.append(f"Comisión: esperado ${com_esp:,.2f}, facturado ${com_fac:,.2f} (dif ${dif_com:+,.2f})")

        renta_mes1_esp = renta_total_esp
        renta_mes1_fac = round(conceptos['RENTA'] + conceptos['ADMIN'] + conceptos['GEOLOC'], 2)
        dif_r1 = round(renta_mes1_fac - renta_mes1_esp, 2)
        if abs(dif_r1) > TOLERANCIA:
            errores.append(f"Renta mes 1: esperado ${renta_mes1_esp:,.2f}, facturado ${renta_mes1_fac:,.2f} (dif ${dif_r1:+,.2f})")

        total_esp = round(ant_esp + com_esp + renta_mes1_esp, 2)
        total_fac = round(ant_fac + com_fac + renta_mes1_fac, 2)
        dif_total = round(total_fac - total_esp, 2)

        status = 'CONCILIADO' if not errores else 'DISCREPANCIA'
        return _res(status, '; '.join(errores) if errores else 'OK',
                    esperado=total_esp, facturado=total_fac, dif=dif_total, detalle=errores,
                    ant_esp=ant_esp, ant_fac=ant_fac,
                    com_esp=com_esp, com_fac=com_fac,
                    renta_esp=renta_total_esp, renta_fac=renta_mes1_fac)

def conciliar_factura(fact: dict, _cache: dict | None = None) -> dict:
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
                folios = ', '.join(o['folio'] or o['uuid'][:8] for o in otras)
                _agregar_aviso(f"Ojo: ya existe otra factura para este contrato en {fact['periodo']} (folio(s): {folios}) — revisa que no sea una doble facturación.")
        except Exception:
            pass
    if res['status'] in ('CONCILIADO', 'DISCREPANCIA') and fact.get('folio') and fact.get('uuid'):
        try:
            conn = get_db()
            mismo_folio = conn.execute(
                """SELECT uuid, id_contrato FROM facturas
                   WHERE folio=? AND uuid != ? AND (cancelada IS NULL OR cancelada=0)""",
                (fact['folio'], fact['uuid'])
            ).fetchall()
            if mismo_folio:
                otros_c = ', '.join(f"{m['id_contrato'] or '—'} ({m['uuid'][:8]}…)" for m in mismo_folio)
                _agregar_aviso(f"Ojo: el folio {fact['folio']} ya existe en otra factura con distinto UUID ({otros_c}) — revisa que no sea un folio reutilizado o una factura duplicada.")
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
        meses_facturados = facturas_dict.get(con['ID_Contrato'], set())
    else:
        conn = get_db()
        rows = conn.execute(
            """SELECT DISTINCT mes_contrato FROM facturas
               WHERE id_contrato=? AND tipo='MENSUAL' AND estatus IN ('CONCILIADO','DISCREPANCIA')
                 AND (cancelada IS NULL OR cancelada=0)""",
            (con['ID_Contrato'],)
        ).fetchall()
        meses_facturados = {r['mes_contrato'] for r in rows if r['mes_contrato']}
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
    """UUIDs de la lista que YA existen en la base de datos — para avisar
    antes de sobrescribir una factura que ya había sido procesada (y que tal
    vez ya se resolvió manualmente: contrato asignado, comisión ajustada)."""
    if not lista:
        return set()
    conn = get_db()
    uuids = [r['uuid'] for r in lista if r.get('uuid')]
    if not uuids:
        return set()
    placeholders = ','.join('?' * len(uuids))
    rows = conn.execute(f"SELECT uuid FROM facturas WHERE uuid IN ({placeholders})", uuids).fetchall()
    return {r[0] for r in rows}

def guardar_facturas_batch(lista: list, sobrescribir_existentes: bool = True):
    conn = get_db()
    cur  = conn.cursor()
    existentes = detectar_duplicados(lista) if not sobrescribir_existentes else set()
    rows_fact = []
    uuids_guardadas = []
    conceptos_por_uuid = {}
    omitidas = []
    for r in lista:
        if r['uuid'] in existentes:
            omitidas.append(r.get('folio', r['uuid']))
            continue
        rows_fact.append((
            r['uuid'],
            r['id_contrato'],
            r['fecha'],
            r.get('folio', ''),
            r['subtotal'],
            r['total'],
            r['tipo'],
            r['periodo'],
            r['mes_contrato'],
            r['status'],
            r.get('msg', ''),
            r.get('rfc_emisor'),
            r.get('rfc_receptor'),
            r.get('esperado', 0) or 0,
            r.get('facturado', 0) or 0,
            r.get('dif', 0) or 0,
            r.get('id_contrato_detectado') or r.get('id_contrato'),
            1 if r.get('alias_aplicado') else 0,
        ))
        uuids_guardadas.append(r['uuid'])
        conceptos_por_uuid[r['uuid']] = r.get('conceptos_raw', [])
    if rows_fact:
        cur.executemany("""
            INSERT OR REPLACE INTO facturas
            (uuid, id_contrato, fecha_emision, folio, subtotal, total,
             tipo, periodo, mes_contrato, estatus, observaciones,
             rfc_emisor, rfc_receptor, esperado, facturado, diferencia,
             id_contrato_detectado, alias_aplicado)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
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
        conn.execute(
            "UPDATE facturas SET estatus=?, observaciones=?, tipo=?, esperado=?, facturado=?, diferencia=? WHERE uuid=?",
            (res['status'], res['msg'], fact['tipo'],
             res.get('esperado', 0) or 0, res.get('facturado', 0) or 0, res.get('dif', 0) or 0, uuid)
        )
        conn.commit()
        return res
    except Exception as e:
        return {'status':'ERROR', 'msg':str(e)}

# --- Pantalla de conciliación de facturas ---
def _render_conciliacion():
    st.title("Conciliación de Facturas CFDI")
    st.caption(
        "Sube los XMLs del mes. El sistema detecta el contrato, clasifica los conceptos, "
        "compara contra la tabla de amortización y marca diferencias. "
        "Puedes resolver discrepancias en comisiones y asignar contratos a facturas sin contrato. "
        "Las facturas de seguro, gestoría y otros servicios no relacionados con leasing se registran como 'NO APLICA'."
    )

    tab_carga, tab_historial, tab_resolver, tab_conceptos, tab_faltantes, tab_reporte = st.tabs([
        "Carga y Conciliación", "Historial", "Resolver Pendientes",
        "Conceptos sin Reconocer", "Avance de Pago", "Reporte"
    ])

    # ===================================================================
    # TAB 1 – CARGA Y CONCILIACIÓN
    # ===================================================================
    with tab_carga:
        with st.expander("Configuración de blindaje (RFC de tu arrendadora)", expanded=not (get_cfg('rfc_arrendadora', '') or '').strip()):
            st.caption(
                "Si capturas aquí el RFC de tu empresa (la que emite las facturas de renta), el sistema "
                "rechaza automáticamente cualquier XML cuyo RFC emisor no coincida — así una factura de "
                "otra empresa, o un lote mezclado por error, nunca se compara contra un contrato que no "
                "le corresponde. Déjalo en blanco si no quieres esta validación."
            )
            _rfc_actual = get_cfg('rfc_arrendadora', '') or ''
            _rfc_nuevo = st.text_input("RFC de tu arrendadora", value=_rfc_actual, max_chars=13,
                                        placeholder="Ej. ABC010101AB1").strip().upper()
            if st.button("Guardar RFC"):
                set_cfg('rfc_arrendadora', _rfc_nuevo)
                st.success("RFC guardado. Se aplicará a partir de la próxima carga.")

        col1, col2 = st.columns([3, 1])
        with col1:
            archivos = st.file_uploader(
                "Sube XMLs de CFDI o un Excel (.xlsx) con el listado de facturas",
                type=['xml', 'xlsx', 'csv'],
                accept_multiple_files=True,
                help="Puedes subir múltiples XMLs, o un Layout en Excel con columnas: UUID, FECHA, FOLIO, SUBTOTAL, TOTAL, CONTRATO, CONCEPTO"
            )
        with col2:
            periodo_ref = st.date_input(
                "Mes de referencia",
                value=hoy_ref().replace(day=1),
                help="Se usa como período si no se detecta en el XML"
            )
            sobrescribir = st.checkbox("Sobrescribir facturas ya cargadas", value=False,
                                       help="Si una factura con el mismo UUID ya existe (por ejemplo, ya la resolviste manualmente), "
                                            "por default se OMITE para no perder ese trabajo. Actívalo solo si de verdad quieres reemplazarla.")

        if not archivos:
            st.info("Sube los XMLs para comenzar.")
        else:
            st.markdown(f"**{len(archivos)} archivo(s) seleccionados**")

            if st.button("Procesar y Conciliar", type="primary"):
                resultados = []
                errores_parse = []
                barra = st.progress(0, text="Procesando XMLs...")
                total = len(archivos)
                uuids_en_lote = {}  # detecta duplicados DENTRO del mismo lote subido

                # Se resuelven una sola vez ANTES del ciclo, no archivo por
                # archivo: `reglas` viene de la configuración de la empresa
                # (antes se releía de la BD en cada XML) y `cache_lote`
                # guarda el contrato + tabla de amortización ya calculada
                # de cada ID_Contrato que se repite en el lote (renta +
                # comisión del mismo anticipo, varios meses del mismo
                # contrato, etc.). Con lotes de cientos de XMLs esto es la
                # diferencia entre segundos y minutos.
                reglas = get_reglas_concepto()
                cache_lote = {}
                alias_map = cargar_alias_contrato()
                usos_alias = {}
                # Refrescar la barra en cada archivo, uno por uno, es
                # innecesario en lotes grandes (cada refresco es un viaje
                # de ida y vuelta al navegador); con más de 40 archivos se
                # actualiza cada 1% para no ser, ella misma, un cuello de
                # botella.
                paso_barra = max(1, total // 100)

                for i, arch in enumerate(archivos):
                    if i % paso_barra == 0 or i == total - 1:
                        barra.progress((i + 1) / total, text=f"Procesando {arch.name} ({i+1}/{total})")
                    try:
                        name_lower = arch.name.lower()
                        facts = []
                        if name_lower.endswith('.xml'):
                            raw = arch.read()
                            facts = [parse_cfdi(raw, reglas)]
                        elif name_lower.endswith(('.xlsx', '.csv')):
                            df_arch = pd.read_excel(arch) if name_lower.endswith('.xlsx') else pd.read_csv(arch)
                            df_arch.columns = [str(col).strip().upper() for col in df_arch.columns]
                            for idx, row in df_arch.iterrows():
                                _uuid = str(row.get('UUID', f'VIRTUAL-{arch.name}-{idx}')).strip()
                                _fecha = str(row.get('FECHA', '')).split(' ')[0]
                                _folio = str(row.get('FOLIO', ''))
                                try: _subt = float(row.get('SUBTOTAL', 0.0) or 0)
                                except: _subt = 0.0
                                try: _tot = float(row.get('TOTAL', 0.0) or 0)
                                except: _tot = 0.0
                                _con = str(row.get('CONTRATO', '')).strip()
                                _desc = str(row.get('CONCEPTO', 'Renta'))
                                
                                c_clave = clasificar_concepto(_desc, reglas)
                                fact = {
                                    'uuid': _uuid, 'fecha': _fecha, 'folio': _folio,
                                    'subtotal': _subt, 'total': _tot,
                                    'id_contrato': _con, 'mes_contrato': None,
                                    'tipo': c_clave, 'periodo': _fecha[:7] if len(_fecha)>=7 else '',
                                    'conceptos': {c_clave: _subt}, 'conceptos_raw': [{'descripcion': _desc, 'importe': _subt, 'clave': c_clave}],
                                    'rfc_emisor': '', 'rfc_receptor': '', 'tipo_comprobante': 'I', 'moneda': 'MXN'
                                }
                                facts.append(fact)
                        
                        for fact in facts:
                            if fact.get('uuid') and fact['uuid'] in uuids_en_lote:
                                errores_parse.append(
                                    f"**{arch.name}**: UUID duplicado omitido."
                                )
                                continue
                            if fact.get('uuid'):
                                uuids_en_lote[fact['uuid']] = arch.name
                            if not fact['periodo']:
                                fact['periodo'] = periodo_ref.strftime('%Y-%m')
                                
                            fact['id_contrato_detectado'] = fact.get('id_contrato')
                            detectado = fact.get('id_contrato')
                            id_final, se_aplico = resolver_numero_contrato(detectado, alias_map)
                            fact['id_contrato'] = id_final
                            fact['alias_aplicado'] = 1 if se_aplico else 0
                            if se_aplico:
                                usos_alias[detectado] = usos_alias.get(detectado, 0) + 1
                            res = conciliar_factura(fact, cache_lote)
                            merged = {**fact, **res}
                            merged['status'] = res['status']
                            if fact['alias_aplicado']:
                                merged['msg'] = (
                                    f"{merged.get('msg', '')} -> Se aplicó alias ({detectado} -> {id_final})"
                                ).strip(' ->')
                            resultados.append(merged)
                    except Exception as e:
                        errores_parse.append(f"**{arch.name}**: {e}")

                barra.empty()
                marcar_alias_usado_lote(usos_alias)
                st.session_state['conciliacion_errores_parse'] = errores_parse

                if not resultados:
                    st.error("No se procesó ningún XML correctamente.")
                    st.session_state['conciliacion_resultados'] = None
                else:
                    with st.spinner("Guardando en base de datos..."):
                        omitidas = guardar_facturas_batch(resultados, sobrescribir_existentes=sobrescribir)
                    registrar_corrida_conciliacion(periodo_ref.strftime('%Y-%m'), resultados, omitidas=len(omitidas))
                    # Se guarda en session_state (no en variables locales) para que la
                    # tabla y las acciones de resolución sigan visibles aunque
                    # resuelvas una discrepancia, cambies de pestaña, o Streamlit
                    # vuelva a correr el script por cualquier otro motivo — antes
                    # todo esto vivía SOLO dentro de este "if", así que en cuanto
                    # dabas clic en cualquier otro botón (como "Asignar contrato")
                    # la tabla entera desaparecía de golpe hasta recargar la página.
                    st.session_state['conciliacion_resultados'] = resultados
                    st.session_state['conciliacion_omitidas'] = omitidas
                    st.session_state['conciliacion_periodo_lote'] = periodo_ref.strftime('%Y-%m')

        # --- Resultados guardados: todos los meses y años, siempre desde la
        # base de datos, nunca solo en memoria de la sesión. ---
        _errores_parse = st.session_state.get('conciliacion_errores_parse') or []
        if _errores_parse:
            with st.expander(f"{len(_errores_parse)} archivo(s) con error de parseo o duplicado"):
                for msg in _errores_parse:
                    st.markdown(msg)

        _omitidas_ultima = st.session_state.get('conciliacion_omitidas') or []
        if _omitidas_ultima:
            st.warning(
                f"**{len(_omitidas_ultima)} factura(s) ya existían** y se omitieron para no perder resoluciones "
                f"previas: {', '.join(str(x) for x in _omitidas_ultima[:15])}"
                + (f" y {len(_omitidas_ultima)-15} más…" if len(_omitidas_ultima) > 15 else "")
                + " — marca 'Sobrescribir facturas ya cargadas' arriba si de verdad quieres reemplazarlas."
            )

        st.divider()
        st.markdown("**Resultados guardados**")
        st.caption(
            "Aquí se ve el detalle completo de cualquier mes o año que ya hayas procesado, no solo el "
            "último lote — queda guardado de forma permanente en la base de datos, así que nunca "
            "desaparece ni depende de dejar la pantalla abierta."
        )
        conn_res = get_db()
        _periodos_todos = [r[0] for r in conn_res.execute(
            "SELECT DISTINCT periodo FROM facturas WHERE periodo IS NOT NULL ORDER BY periodo DESC"
        ).fetchall()]

        if not _periodos_todos:
            st.info("Aún no has procesado ningún XML.")
        else:
            _periodo_default = st.session_state.get('conciliacion_periodo_lote')
            _idx_default = _periodos_todos.index(_periodo_default) if _periodo_default in _periodos_todos else 0
            periodo_ver = st.selectbox("Mes a revisar", _periodos_todos, index=_idx_default, key="periodo_ver_resultados")

            filas = conn_res.execute(
                """SELECT uuid, folio, id_contrato, tipo, mes_contrato, subtotal, total,
                          estatus, observaciones, esperado, facturado, diferencia, cancelada
                   FROM facturas WHERE periodo=? ORDER BY estatus, id_contrato""",
                (periodo_ver,)
            ).fetchall()

            if not filas:
                st.info(f"No hay facturas guardadas para {periodo_ver}.")
            else:
                n_ok   = sum(1 for f in filas if f['estatus'] == 'CONCILIADO')
                n_disc = sum(1 for f in filas if f['estatus'] == 'DISCREPANCIA')
                n_sin  = sum(1 for f in filas if f['estatus'] == 'SIN_CONTRATO')
                n_err  = sum(1 for f in filas if f['estatus'] == 'ERROR')
                n_noa  = sum(1 for f in filas if f['estatus'] == 'NO_APLICA')
                n_rfc  = sum(1 for f in filas if f['estatus'] == 'RFC_INCORRECTO')

                k1, k2, k3, k4, k5, k6 = st.columns(6)
                k1.metric("Conciliados",   n_ok)
                k2.metric("Discrepancia",  n_disc)
                k3.metric("Sin contrato",  n_sin)
                k4.metric("Error",         n_err)
                k5.metric("No aplica",     n_noa)
                k6.metric("RFC incorrecto", n_rfc)

                # Tabla de resultados
                df_res = pd.DataFrame([{
                    'Folio':        f['folio'] or '-',
                    'UUID':         (f['uuid'] or '')[:8] + '…',
                    'Contrato':     f['id_contrato'] or '—',
                    'Tipo':         f['tipo'] or '-',
                    'Mes':          str(int(f['mes_contrato'])) if f['mes_contrato'] is not None else '-',
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

                # Desplegar acciones de resolución para cada resultado
                with st.expander("Acciones de resolución (discrepancias y sin contrato)", expanded=False):
                    for f in filas:
                        fr = dict(f)
                        if fr['estatus'] == 'DISCREPANCIA' and fr.get('tipo') == 'ANTICIPO':
                            # Verificar si la discrepancia es en comisión
                            if 'Comisión' in (fr.get('observaciones') or ''):
                                _conc_fila = _totales_desde_conceptos(fr['uuid'])
                                com_fac = _conc_fila.get('COMISION', 0)
                                _con_row = conn_res.execute(
                                    "SELECT Comision_Monto FROM contratos WHERE ID_Contrato=?", (fr['id_contrato'],)
                                ).fetchone()
                                com_esp = _con_row['Comision_Monto'] if _con_row else 0
                                st.divider()
                                st.markdown(f"**Contrato {fr.get('id_contrato','')}** · Folio {fr.get('folio','')}")
                                st.write(f"Comisión esperada: **${com_esp:,.2f}** (del contrato)")
                                st.write(f"Comisión facturada: **${com_fac:,.2f}**")
                                if st.button(f"Actualizar comisión del contrato a ${com_fac:,.2f}", key=f"upd_com_{fr['uuid']}"):
                                    conn_res.execute("UPDATE contratos SET Comision_Monto=? WHERE ID_Contrato=?", (com_fac, fr['id_contrato']))
                                    conn_res.commit()
                                    agregar_anotacion(fr['id_contrato'], f"Comisión actualizada de ${com_esp:,.2f} a ${com_fac:,.2f} por conciliación de factura {fr.get('folio','')}", "Ajuste contable")
                                    res2 = re_conciliar_factura(fr['uuid'])
                                    if res2:
                                        st.success(f"Comisión actualizada y factura reconciliada. Nuevo estatus: {res2['status']}")
                                    else:
                                        st.warning("Comisión actualizada, pero no se encontró la factura para re-conciliarla.")
                                    st.session_state['_refresh'] = True
                        elif fr['estatus'] == 'SIN_CONTRATO':
                            st.divider()
                            st.markdown(f"**Factura sin contrato** · Folio {fr.get('folio','')} · UUID {fr.get('uuid','')}")
                            df_todos = obtener()  # todos los contratos, no solo ACTIVO, para poder encontrar cualquiera
                            if not df_todos.empty:
                                df_todos_sorted = df_todos.sort_values('ID_Contrato')
                                etiquetas = (df_todos_sorted['ID_Contrato'] + ' — ' + df_todos_sorted['Cliente'].fillna('') +
                                             ' (' + df_todos_sorted['Estatus'].fillna('') + ')').tolist()
                                mapa_etiqueta_id = dict(zip(etiquetas, df_todos_sorted['ID_Contrato']))
                                etiqueta_sel = st.selectbox(
                                    "Selecciona un contrato (escribe para buscar por número o cliente)",
                                    etiquetas, key=f"asign_{fr['uuid']}"
                                )
                                sel_contrato = mapa_etiqueta_id[etiqueta_sel]
                                if st.button(f"Asignar contrato {sel_contrato} y reconciliar", key=f"asign_btn_{fr['uuid']}"):
                                    _detectado = fr.get('id_contrato_detectado') or fr.get('id_contrato')
                                    res2 = asignar_contrato_manual(fr['uuid'], sel_contrato, _detectado)
                                    if res2:
                                        if _detectado and _detectado != sel_contrato:
                                            st.success(f"Contrato asignado. Nuevo estatus: {res2['status']}. "
                                                       f"Se guardó como regla: la próxima factura que llegue con "
                                                       f"'{_detectado}' se asignará sola a {sel_contrato}.")
                                        else:
                                            st.success(f"Contrato asignado. Nuevo estatus: {res2['status']}")
                                    else:
                                        st.warning("Contrato asignado, pero no se encontró la factura para re-conciliarla.")
                                    st.session_state['_refresh'] = True
                            else:
                                st.warning("No hay contratos disponibles en la base de datos.")

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
    with tab_historial:
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
                           f.subtotal, f.total, f.estatus, f.observaciones, f.fecha_emision
                    FROM facturas f
                    WHERE f.periodo=? AND f.estatus IN ({placeholders})
                    ORDER BY f.estatus, f.id_contrato""",
                [per_sel] + estatus_fil
            ).fetchall()

            if rows:
                df_hist = pd.DataFrame(rows, columns=[
                    'UUID', 'Folio', 'Contrato', 'Tipo', 'Mes',
                    'Subtotal', 'Total', 'Estatus', 'Observación', 'Fecha'
                ])
                df_hist['UUID'] = df_hist['UUID'].str[:8] + '…'
                st.dataframe(df_hist, width='stretch', height=500, key="df_003")

                all_rows = conn.execute(
                    "SELECT estatus, COUNT(*) as n FROM facturas WHERE periodo=? GROUP BY estatus",
                    (per_sel,)
                ).fetchall()
                kk = {r[0]: r[1] for r in all_rows}
                c1, c2, c3, c4, c5 = st.columns(5)
                c1.metric("Conciliados",  kk.get('CONCILIADO', 0))
                c2.metric("Discrepancia", kk.get('DISCREPANCIA', 0))
                c3.metric("Sin contrato", kk.get('SIN_CONTRATO', 0))
                c4.metric("Error",        kk.get('ERROR', 0))
                c5.metric("No aplica",    kk.get('NO_APLICA', 0))
            else:
                st.info("No hay facturas con esos filtros para ese período.")

        st.divider()
        st.subheader("Historial de Corridas")
        st.caption(
            "Un renglón por cada vez que se procesó un lote de XMLs — a diferencia de la tabla de "
            "arriba (que se puede sobreescribir si vuelves a cargar el mismo período), esto siempre "
            "suma un registro nuevo, para tener a la mano cuándo se concilió cada lote."
        )
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
    # TAB — CONCEPTOS SIN RECONOCER
    # ===================================================================
    with tab_conceptos:
        st.subheader("Conceptos sin Reconocer")
        st.caption(
            "Aquí aparece cualquier concepto de un CFDI cuyo texto no trae ninguna de las palabras "
            "que ya está estipulado que debe decir (Renta, Administración, Geolocalización, Anticipo "
            "o Comisión). Mientras un concepto esté aquí, su importe NO se toma en cuenta en la "
            "comparación contra la póliza. Puedes reclasificarlo a mano si en realidad sí corresponde "
            "a alguna de esas categorías, o dejarlo fuera a propósito (por ejemplo, un cargo que de "
            "verdad no debe compararse). Las facturas que ya se clasificaron como 'servicios no "
            "relacionados con leasing' (seguro, gestoría, etc.) no aparecen aquí, porque para esas es "
            "normal y esperado que ningún concepto coincida — revísalas en la pestaña Historial si "
            "crees que alguna en realidad sí era de renta."
        )
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
    # TAB 3 – CONTRATOS SIN FACTURA + EMPAREJAMIENTO
    # ===================================================================
    with tab_faltantes:
        st.subheader("Avance de Pago por Contrato")
        st.caption(
            "Con base en las facturas ya conciliadas (no en el nivel de morosidad capturado a mano), "
            "identifica qué contratos van adelantados con sus mensualidades, cuáles van al corriente, y "
            "cuáles tienen meses vencidos sin factura — con el detalle de cuáles meses faltan."
        )
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
                    facturas_dict.setdefault(r['id_contrato'], set()).add(r['mes_contrato'])
                
                filas_avance = []
                for _, _con_av in df_avance_base.iterrows():
                    _avp = calcular_avance_pago(_con_av, facturas_dict)
                    filas_avance.append({
                        'ID_Contrato': _con_av['ID_Contrato'], 'Cliente': _con_av['Cliente'],
                        'Estado': _avp['estado'],
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
            ka1.metric("Atrasados", n_atr)
            ka2.metric("Al corriente", n_cor)
            ka3.metric("Adelantados", n_ade)

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
                        "FROM facturas WHERE periodo=? AND estatus='SIN_CONTRATO'",
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
    # TAB 4 – RESOLVER PENDIENTES (vista general)
    # ===================================================================
    with tab_resolver:
        st.subheader("Resolver Pendientes")
        conn = get_db()

        with st.expander("🧠 Reglas aprendidas (contratos facturados con otro número)"):
            st.caption(
                "Cada vez que resuelves a mano una factura 'Sin contrato', si el número que traía el XML "
                "es distinto al contrato real que le asignaste, el sistema guarda esa equivalencia aquí. "
                "La próxima factura que llegue con ese mismo número se conciliará sola, sin volver a pedir "
                "que la asignes — y queda anotado en el resultado que se resolvió por regla."
            )
            reglas_alias, historial_alias = obtener_alias_contrato_con_historial()
            if not reglas_alias:
                st.info("Todavía no hay ninguna regla aprendida. En cuanto resuelvas a mano tu primera factura "
                        "'Sin contrato' con un número distinto al del contrato real, aparecerá aquí.")
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
                if historial_alias:
                    st.markdown("**Historial de cambios** (últimos 200)")
                    df_hist = pd.DataFrame([dict(h) for h in historial_alias])
                    st.dataframe(
                        df_hist[['fecha', 'numero_facturado', 'id_contrato_real', 'accion', 'detalle']],
                        width='stretch', key="df_hist_alias"
                    )

        pendientes = conn.execute(
            "SELECT uuid, folio, id_contrato, id_contrato_detectado, tipo, estatus, observaciones, periodo, mes_contrato, "
            "fecha_emision, fecha_registro, rfc_emisor, total FROM facturas "
            "WHERE estatus IN ('DISCREPANCIA','SIN_CONTRATO','RFC_INCORRECTO','FUERA_DE_VIGENCIA') AND (cancelada IS NULL OR cancelada=0) "
            "ORDER BY periodo DESC, estatus"
        ).fetchall()
        if not pendientes:
            st.success("¡No hay pendientes! Todo está conciliado.")
        else:
            st.markdown(f"**{len(pendientes)} facturas pendientes de resolver**")
            try:
                _dias_prom = (pd.Timestamp(hoy_ref()) - pd.to_datetime([p['fecha_registro'] for p in pendientes]).min()).days
            except Exception:
                _dias_prom = None
            if _dias_prom and _dias_prom > 30:
                st.warning(f"Hay pendientes desde hace más de {_dias_prom} días — entre más tiempo pase, más difícil es rastrear el origen de la discrepancia.")
            for p in pendientes:
                with st.expander(f"{p['folio']} | {p['tipo']} | {p['estatus']} | Contrato: {p['id_contrato'] or '—'} | ${p['total']:,.2f}"):
                    st.write(f"**UUID:** {p['uuid']}")
                    st.write(f"**Período:** {p['periodo']}, Mes contrato: {p['mes_contrato']}")
                    st.write(f"**Observaciones:** {p['observaciones']}")
                    if p['estatus'] == 'RFC_INCORRECTO':
                        st.error(f"RFC emisor detectado: **{p['rfc_emisor'] or '—'}**. No coincide con el RFC configurado de tu arrendadora.")
                        st.caption("Revisa si subiste el lote de la empresa equivocada, o si el RFC configurado en 'Carga y Conciliación' ya no es el correcto.")
                    if p['estatus'] == 'FUERA_DE_VIGENCIA':
                        st.error("El contrato ya estaba dado de baja cuando se emitió esta factura.")
                        st.caption(
                            "Revisa si de verdad se siguió cobrando renta después de la baja (y hay que "
                            "investigar por qué), si la fecha de baja capturada en el contrato está mal, o si "
                            "esta factura en realidad es de otro contrato y se identificó mal el número."
                        )
                    if p['estatus'] == 'DISCREPANCIA' and p['tipo'] == 'ANTICIPO':
                        if 'Comisión' in p['observaciones']:
                            st.info("Esta discrepancia es por comisión. Puedes actualizar la comisión del contrato.")
                            if st.button(f"Actualizar comisión para {p['id_contrato']}", key=f"res_com_{p['uuid']}"):
                                fila_com = conn.execute(
                                    "SELECT COALESCE(SUM(importe),0) AS total FROM factura_conceptos WHERE uuid=? AND clave='COMISION'",
                                    (p['uuid'],)
                                ).fetchone()
                                com_fac = fila_com['total'] if fila_com else 0.0
                                try:
                                    conn.execute("UPDATE contratos SET Comision_Monto=? WHERE ID_Contrato=?", (com_fac, p['id_contrato']))
                                    conn.commit()
                                    agregar_anotacion(p['id_contrato'], f"Comisión actualizada por resolución manual de factura {p['folio']}", "Ajuste contable")
                                    res2 = re_conciliar_factura(p['uuid'])
                                    st.success("Comisión actualizada y factura reconciliada.")
                                    st.session_state['_refresh'] = True
                                except Exception as e:
                                    st.error(f"Error: {e}")
                    elif p['estatus'] == 'SIN_CONTRATO':
                        df_todos = obtener()  # todos los contratos, no solo ACTIVO
                        if not df_todos.empty:
                            df_todos_sorted = df_todos.sort_values('ID_Contrato')
                            etiquetas = (df_todos_sorted['ID_Contrato'] + ' — ' + df_todos_sorted['Cliente'].fillna('') +
                                         ' (' + df_todos_sorted['Estatus'].fillna('') + ')').tolist()
                            mapa_etiqueta_id = dict(zip(etiquetas, df_todos_sorted['ID_Contrato']))
                            _detectado_previo = p['id_contrato_detectado'] or p['id_contrato']
                            _sugerencias = sugerir_contrato_similar(_detectado_previo, df_todos_sorted['ID_Contrato'].tolist())
                            if _sugerencias:
                                st.caption(
                                    f"El número '{_detectado_previo}' se parece a: {', '.join(_sugerencias)} "
                                    f"— revisa si es un typo antes de buscarlo a mano."
                                )
                            etiqueta_sel = st.selectbox(
                                "Asignar contrato (escribe para buscar por número o cliente)",
                                etiquetas, key=f"res_asign_{p['uuid']}"
                            )
                            sel_contrato = mapa_etiqueta_id[etiqueta_sel]
                            if st.button(f"Asignar {sel_contrato}", key=f"res_asign_btn_{p['uuid']}"):
                                _detectado = p['id_contrato_detectado'] or p['id_contrato']
                                res2 = asignar_contrato_manual(p['uuid'], sel_contrato, _detectado)
                                if res2:
                                    if _detectado and _detectado != sel_contrato:
                                        st.success(f"Asignado. Nuevo estatus: {res2['status']}. Se guardó como "
                                                   f"regla: '{_detectado}' se asignará solo a {sel_contrato} "
                                                   f"de ahora en adelante.")
                                    else:
                                        st.success(f"Asignado. Nuevo estatus: {res2['status']}")
                                else:
                                    st.warning("Asignado pero no se pudo re-conciliar.")
                                st.session_state['_refresh'] = True
                        else:
                            st.warning("No hay contratos en la base de datos.")
                    st.divider()
                    st.caption("¿Esta factura en realidad fue cancelada ante el SAT? Márcala para sacarla de los pendientes y de los totales de control, sin borrar su historial.")
                    if st.button("Marcar como cancelada", key=f"cancel_res_{p['uuid']}"):
                        marcar_factura_cancelada(p['uuid'], True)
                        st.success("Factura marcada como cancelada.")
                        st.session_state['_refresh'] = True

    # ===================================================================
    # TAB 5 – REPORTE DE CONCILIACIÓN (NUEVO)
    # ===================================================================
    with tab_reporte:
        st.subheader("Reporte Consolidado de Conciliación")
        conn = get_db()
        periodos = conn.execute(
            "SELECT DISTINCT periodo FROM facturas WHERE periodo IS NOT NULL ORDER BY periodo DESC"
        ).fetchall()
        lista_per = [r[0] for r in periodos]

        if not lista_per:
            st.info("Aún no hay facturas procesadas. Carga algunos XMLs primero.")
        else:
            per_rep = st.selectbox("Selecciona un período", lista_per, key="per_reporte")
            if st.button("Generar Reporte", width='stretch'):
                with st.spinner("Generando reporte..."):
                    rows = conn.execute(
                        """SELECT uuid, folio, id_contrato, tipo, mes_contrato,
                                  subtotal, total, estatus, observaciones, fecha_emision, fecha_registro, cancelada
                           FROM facturas WHERE periodo=? ORDER BY fecha_emision""",
                        (per_rep,)
                    ).fetchall()
                    if not rows:
                        st.warning("No hay facturas para este período.")
                    else:
                        df_rep = pd.DataFrame(rows, columns=[
                            'UUID', 'Folio', 'Contrato', 'Tipo', 'Mes',
                            'Subtotal', 'Total', 'Estatus', 'Observación', 'Fecha Emisión', 'Fecha Registro', 'Cancelada'
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
                        st.markdown('<span class="section-label">Control global del período — ¿cuadra lo facturado contra lo que debía facturarse?</span>', unsafe_allow_html=True)
                        st.caption(
                            "Compara la renta mensual pactada de TODOS los contratos activos contra lo que "
                            "realmente se facturó (sin contar facturas canceladas) en este período. Una "
                            "diferencia grande suele significar contratos sin facturar, facturas duplicadas, "
                            "o facturas todavía sin contrato asignado."
                        )
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
    st.markdown(
        "Sube el Excel de analíticas que te entrega el contador (las hojas típicas son 'Ingresos por "
        "Intereses', 'Ingresos por Valor Residual' y 'Arrendamiento de Comisión Apertura', un renglón "
        "por contrato con lo registrado mes a mes). El sistema calcula lo que debería haberse "
        "registrado para cada contrato y mes, compara contra lo que trae el Excel, y te dice "
        "exactamente en dónde hay diferencia y por qué — contrato dado de baja que se sigue "
        "contabilizando, renta capturada distinta, contrato que no encuentra en el sistema, etc."
    )
    arch_analitica = st.file_uploader("Excel de analíticas", type=["xlsx", "xls"], key="analiticas_uploader")

    if arch_analitica is not None:
        try:
            df_tidy, anio_detectado, hojas_no_rec = parse_analiticas_excel(arch_analitica.read())
        except Exception as e:
            st.error(f"No se pudo leer el archivo: {e}")
            df_tidy, anio_detectado, hojas_no_rec = pd.DataFrame(), None, []

        if hojas_no_rec:
            st.info(
                f"No reconocí de qué son estas hojas (no las comparé): {', '.join(hojas_no_rec)}. "
                f"Reconozco hojas cuyo nombre contenga 'interés', 'residual' o 'comisión'."
            )

        if df_tidy.empty:
            st.warning("No encontré renglones con datos en ninguna hoja reconocida del archivo.")
        else:
            anio_cmp = st.number_input(
                "Año que cubre esta analítica", min_value=2015, max_value=2100,
                value=anio_detectado or hoy_ref().year, step=1,
                help="Se detectó automáticamente del archivo si venía escrito ahí (ej. 'Anexos al 30 de Junio de 2026'); ajústalo si no es correcto."
            )
            n_sin_id = df_tidy['ID_Contrato_Detectado'].isna().sum()
            if n_sin_id:
                st.caption(f"{n_sin_id} renglón(es) del Excel no traían un número de contrato reconocible en el texto — se marcarán como 'no identificado'.")

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
            st.caption(
                "Suma la diferencia neta por contrato y por tipo de ingreso, y propone una línea de "
                "Cargo/Abono para dejar la cuenta contable alineada con lo que dice el sistema. "
                "Es una sugerencia para que la revises — no se contabiliza sola."
            )
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
        "El archivo de esta empresa se dañó a nivel de disco "
        "(`database disk image is malformed`). No es un error de tus datos "
        "ni de captura — es el archivo .db en sí. Esto casi siempre pasa por "
        "guardar la carpeta del programa dentro de OneDrive/Dropbox/Google Drive, "
        "o por un antivirus escaneando el archivo mientras se escribe."
    )
    respaldos = listar_respaldos_automaticos(db_path)
    if respaldos:
        st.success(f"Buenas noticias: hay {len(respaldos)} respaldo(s) automático(s) de esta misma base de datos.")
        opciones = {
            f"{fecha.strftime('%Y-%m-%d %H:%M:%S')}" + (" 🔒 (cifrado)" if cifrado else ""): archivo
            for archivo, fecha, cifrado in respaldos
        }
        elegido = st.selectbox("Elige el respaldo a restaurar (el más reciente arriba):", list(opciones.keys()))
        archivo_elegido = opciones[elegido]
        clave_restaurar = None
        if archivo_elegido.endswith(".zip"):
            clave_restaurar = st.text_input("Contraseña del respaldo", type="password")
        st.caption("El archivo dañado NO se borra: se renombra y se guarda junto a los demás, por si necesitas soporte técnico.")
        if st.button("Restaurar este respaldo y continuar", type="primary"):
            try:
                restaurar_respaldo_automatico(db_path, archivo_elegido, password=clave_restaurar)
                st.session_state.clear()
                st.success("Restaurado. Vuelve a abrir la aplicación.")
                st.stop()
            except RuntimeError as e:
                st.error(str(e))
    else:
        st.warning(
            "No se encontró ningún respaldo automático todavía para este archivo "
            "(se empiezan a crear a partir de esta versión, uno por cada vez que abras "
            "la app con la base de datos sana). Si tienes una copia manual descargada "
            "desde 'Respaldo y Restauración', puedes restaurarla ahí una vez que "
            "vuelvas a poder abrir la app con una base de datos nueva."
        )
    st.markdown("---")
    _carpeta_ejemplo = r"C:\O-Leasing" if sys.platform.startswith("win") else "/home/tuusuario/O-Leasing"
    st.markdown("**Para que no vuelva a pasar:** mueve la carpeta del programa fuera de "
                f"cualquier carpeta sincronizada por OneDrive/Dropbox/Google Drive (por ejemplo, "
                f"a `{_carpeta_ejemplo}` directamente), y agrega esa carpeta como excepción en tu antivirus "
                f"{'(si usas uno en Windows)' if sys.platform.startswith('win') else ''}.")
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
        Algo no salió como se esperaba al preparar esta vista — no es que se haya perdido información,
        el resto del sistema sigue intacto. Prueba recargar la página o cambiar de contrato/filtro;
        si vuelve a pasar en el mismo lugar, avísale a soporte con el detalle técnico de abajo.
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
        "Comparar Analíticas",
        "Tablas de Amortización",
    ],
    "Análisis": [
        "Proyección Financiera",
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
    "Dashboard & Cartera": "Panorama general de tu cartera: KPIs, alertas automáticas y gráficas de morosidad, vencimientos y flujo.",
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
    "Comparar Analíticas": "Sube las analíticas contables del contador y compáralas contra lo que el sistema esperaba, con sugerencia de póliza de ajuste.",
    "Tablas de Amortización": "Consulta la tabla de amortización completa de cualquier contrato.",
    "Proyección Financiera": "Proyecta el comportamiento financiero futuro de tu cartera.",
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

if ROL_ACTUAL == "lectura" and menu in PANTALLAS_SOLO_ESCRITURA:
    st.warning(
        f"Tu usuario tiene rol de **solo lectura** y esta pantalla ({menu}) es para capturar o modificar "
        "información. Pide a un usuario con rol de captura o administrador que haga este cambio."
    )
    st.stop()

# Dashboard
try:
    if menu=="Dashboard & Cartera":
        # Dashboard principal
        hoy = hoy_ref()
        df_all  = obtener()
        df      = obtener('ACTIVO')
        df_anot = obtener_anotaciones()
        df_evts = obtener_eventos_especiales()

        # Buscador universal: ver cualquier contrato al instante
        # Va arriba a propósito — es la forma más rápida de "ver lo que
        # quieras, cuando lo quieras" sin tener que navegar el menú ni
        # bajar hasta la tabla.
        # (Antes había aquí un banner verde grande repitiendo "O-Leasing —
        # Dashboard" y la fecha — información que ya está arriba, en la
        # barra de ruta, y abajo, en el KPI de "Contratos activos". Quitarlo
        # deja ver el contenido real de inmediato en vez de tres bloques de
        # encabezado seguidos antes de llegar a algo útil.)
        busq_top = st.text_input("Buscar contrato o cliente y saltar directo a su estado de cuenta",
                                  placeholder="Escribe un ID o un nombre…", key="dash_busq_top")
        if busq_top.strip() and not df_all.empty:
            _match_top = df_all[
                df_all['ID_Contrato'].str.contains(busq_top, case=False, na=False) |
                df_all['Cliente'].str.contains(busq_top, case=False, na=False)
            ]
            if not _match_top.empty:
                for _, _mr in _match_top.head(5).iterrows():
                    if st.button(f"{_mr['ID_Contrato']} — {_mr['Cliente']}", key=f"jump_{_mr['ID_Contrato']}", width='stretch'):
                        st.session_state['ec_contrato'] = _mr['ID_Contrato']
                        st.session_state['menu_item']   = "Estado de Cuenta"
                        for g, its in GRUPOS.items():
                            if "Estado de Cuenta" in its:
                                st.session_state['menu_grupo'] = g
                        st.rerun()
            else:
                st.caption("Sin coincidencias.")

        # KPIs — el semáforo: solo lo que responde "¿estamos bien hoy?"
        tot_v   = df['Valor_Sin_IVA'].sum()       if not df.empty else 0
        tot_r   = df['Mensualidad_Sin_IVA'].sum() if not df.empty else 0
        tot_res = df['Residual_Monto'].sum()      if not df.empty else 0
        tot_ant = df['Anticipo_Monto'].sum()      if not df.empty else 0
        tot_com = df['Comision_Monto'].sum()      if not df.empty else 0
        n_mora  = int((df['Nivel_Morosidad']>0).sum()) if not df.empty else 0
        n_excl  = int(df['Fecha_Excl_Poliza'].notna().sum()) if ('Fecha_Excl_Poliza' in df.columns and not df.empty) else 0
        try:
            _df_pe_dash = calcular_punto_equilibrio()
            n_pe_hoy = int(_df_pe_dash['PE_Alcanzado_Hoy'].sum()) if not _df_pe_dash.empty else 0
        except Exception:
            n_pe_hoy = 0

        k1,k2,k3,k4 = st.columns(4)
        k1.metric("Contratos activos",  f"{len(df):,}")
        k2.metric("Valor cartera",      f"${tot_v/1e6:.2f}M")
        k3.metric("Con morosidad",   f"{n_mora}", delta=f"-{n_mora} contratos" if n_mora else None,
                  delta_color="inverse")
        k4.metric("Ya en punto de equilibrio", f"{n_pe_hoy}/{len(df)}" if len(df) else "0/0",
                  help="Contratos activos que ya recuperaron su capital invertido, según la tabla de amortización real. Ver detalle en 'Punto de Equilibrio'.")

        with st.expander("Ver más detalle financiero (rentas, residual, anticipo, comisión, excluidos)"):
            j1,j2,j3,j4,j5 = st.columns(5)
            j1.metric("Rentas / mes",        f"${tot_r:,.0f}")
            j2.metric("Residual total",      f"${tot_res:,.0f}")
            j3.metric("Anticipo total",      f"${tot_ant:,.0f}")
            j4.metric("Comisión total",      f"${tot_com:,.0f}")
            j5.metric("Excluidos póliza", f"{n_excl}", delta=f"-{n_excl}" if n_excl else None,
                      delta_color="inverse")

        st.markdown("---")

        # ALERTAS AUTOMÁTICAS
        alertas = []

        # Vencimientos próximos 60 días
        if not df.empty and 'Fecha_Vencimiento' in df.columns:
            df_venc60 = df[
                (df['Fecha_Vencimiento'].dt.date >= hoy) &
                (df['Fecha_Vencimiento'].dt.date <= hoy + relativedelta(days=60))
            ]
            if not df_venc60.empty:
                alertas.append(('warning',
                    f"{len(df_venc60)} contrato(s) vencen en los próximos 60 días",
                    df_venc60[['ID_Contrato','Cliente','Fecha_Vencimiento']].copy()))

        # Contratos en morosidad alta
        if not df.empty:
            df_mora_alta = df[df['Nivel_Morosidad'].fillna(0).astype(int) >= 3]
            if not df_mora_alta.empty:
                alertas.append(('error',
                    f"{len(df_mora_alta)} contrato(s) en Morosidad nivel 3 o 4",
                    df_mora_alta[['ID_Contrato','Cliente','Nivel_Morosidad']].copy()))

        # Eventos especiales pendientes
        if not df_evts.empty:
            df_evts_pend = df_evts[df_evts['Estatus_Seguro']=='PENDIENTE']
            if not df_evts_pend.empty:
                alertas.append(('warning',
                    f"{len(df_evts_pend)} evento(s) especial(es) pendiente(s) de resolución",
                    df_evts_pend[['ID_Contrato','Tipo_Evento','Fecha_Evento']].head(10).copy()))

        # Contratos sin anotaciones en >90 días (posibles "fantasma")
        if not df.empty and not df_anot.empty:
            ultima = df_anot.groupby('ID_Contrato')['Fecha'].max().reset_index()
            ultima['Fecha'] = pd.to_datetime(ultima['Fecha'])
            df_sin_anot = df[~df['ID_Contrato'].isin(ultima['ID_Contrato'])]
            df_viejas = df.merge(ultima, on='ID_Contrato', how='left')
            df_viejas = df_viejas[df_viejas['Fecha'] < pd.Timestamp(hoy - relativedelta(days=90))]
            n_sin_seguimiento = len(df_sin_anot) + len(df_viejas)
            if n_sin_seguimiento > 0:
                alertas.append(('info',
                    f"{n_sin_seguimiento} contrato(s) sin anotaciones en los últimos 90 días",
                    None))

        # Contratos vencidos que siguen marcados ACTIVO
        if not df.empty and 'Fecha_Vencimiento' in df.columns:
            df_venc_ya = df[df['Fecha_Vencimiento'].dt.date < hoy]
            if not df_venc_ya.empty:
                alertas.append(('warning',
                    f"{len(df_venc_ya)} contrato(s) ya vencieron y siguen marcados ACTIVO — revisa si hay que renovarlos o darlos de baja",
                    df_venc_ya[['ID_Contrato','Cliente','Fecha_Vencimiento']].copy()))

        # Inconsistencias entre Estatus y Fecha_Baja (datos que no cuadran
        # entre sí) — se revisan sobre TODOS los contratos, no solo activos.
        if not df_all.empty:
            df_baja_sin_fecha = df_all[(df_all['Estatus'].astype(str).str.upper()=='BAJA') &
                                        (df_all['Fecha_Baja'].isna() | (df_all['Fecha_Baja'].astype(str).str.strip()==''))]
            if not df_baja_sin_fecha.empty:
                alertas.append(('warning',
                    f"{len(df_baja_sin_fecha)} contrato(s) marcados BAJA sin fecha de baja capturada",
                    df_baja_sin_fecha[['ID_Contrato','Cliente','Estatus']].copy()))

            df_activo_con_baja = df_all[(df_all['Estatus'].astype(str).str.upper()!='BAJA') &
                                         df_all['Fecha_Baja'].notna() & (df_all['Fecha_Baja'].astype(str).str.strip()!='')]
            if not df_activo_con_baja.empty:
                alertas.append(('warning',
                    f"{len(df_activo_con_baja)} contrato(s) tienen fecha de baja capturada pero su estatus no es BAJA",
                    df_activo_con_baja[['ID_Contrato','Cliente','Estatus','Fecha_Baja']].copy()))

            df_baja_antes_alta = df_all[df_all['Fecha_Baja'].notna() &
                                         (pd.to_datetime(df_all['Fecha_Baja'], errors='coerce') < df_all['Fecha_Alta'])]
            if not df_baja_antes_alta.empty:
                alertas.append(('error',
                    f"{len(df_baja_antes_alta)} contrato(s) con fecha de baja ANTERIOR a su fecha de alta — dato imposible, seguro es un error de captura",
                    df_baja_antes_alta[['ID_Contrato','Cliente','Fecha_Alta','Fecha_Baja']].copy()))

        # eso se avisa arriba con las demás alertas (el detalle y la gráfica
        # viven en la pestaña "Morosidad y Concentración" más abajo).
        _top1_pct = 0.0
        if not df.empty and tot_v > 0:
            _conc_alerta = df.groupby('Cliente')['Valor_Sin_IVA'].sum().sort_values(ascending=False)
            _top1_pct = float(_conc_alerta.iloc[0] / tot_v * 100)
            if _top1_pct > 20:
                alertas.append(('warning',
                    f"Un solo cliente ({_conc_alerta.index[0]}) concentra el {_top1_pct:.1f}% de tu cartera — "
                    f"revisa el detalle en la pestaña 'Morosidad y Concentración'",
                    None))

        if alertas:
            # Reorganizar visualmente como "Centro de Atención" (D.1)
            # Priorizamos 'error' (datos imposibles), luego 'warning', luego 'info'
            orden_severidad = {'error': 0, 'warning': 1, 'info': 2}
            alertas.sort(key=lambda x: orden_severidad.get(x[0], 99))
            
            st.markdown(f'''
                <div style="display:flex; justify-content:space-between; align-items:baseline; margin-bottom:10px;">
                    <h3 style="margin:0; color:var(--text-primary) !important;">Centro de Atención</h3>
                    <span style="color:var(--text-secondary); font-size:14px; font-weight:600;">{len(alertas)} pendientes</span>
                </div>
            ''', unsafe_allow_html=True)
            
            for _idx_alerta, (tipo, msg, detalle) in enumerate(alertas):
                if tipo == 'error':
                    icono = "🔴"
                    bg = "var(--danger-soft)"
                    color = "var(--danger)"
                elif tipo == 'warning':
                    icono = "🟡"
                    bg = "var(--warning-soft)"
                    color = "var(--warning)"
                else:
                    icono = "🔵"
                    bg = "var(--info-soft)"
                    color = "var(--info)"
                
                with st.container(key=f"ca_{_idx_alerta}"):
                    st.markdown(f'''
                        <div style="background:{bg}; border-left:4px solid {color}; border-radius:4px; padding:12px; margin-bottom:10px;">
                            <span style="font-size:16px; margin-right:8px;">{icono}</span>
                            <span style="font-weight:500; color:var(--text-primary);">{msg}</span>
                        </div>
                    ''', unsafe_allow_html=True)
                    if detalle is not None:
                        with st.expander("Ver detalle", key=f"exp_alerta_{_idx_alerta}"):
                            st.dataframe(detalle, width='stretch', key=f"df_006_{_idx_alerta}")
            st.markdown("---")

        # El resto va en pestañas: antes estas cuatro secciones (morosidad,
        # vencimientos, flujo, anotaciones) estaban todas apiladas y visibles
        # siempre, y era demasiado para ver de un vistazo. Ahora eliges cuál
        # ver; las demás siguen ahí, a un clic.
        tab_mor, tab_venc, tab_flujo, tab_act = st.tabs([
            "Morosidad y Concentración", "Vencimientos", "Proyección de Flujo", "Actividad Reciente"
        ])

        with tab_mor:
            mc1, mc2 = st.columns([1.1, 1])
            with mc1:
                st.markdown('<span class="section-label">Cartera por nivel de morosidad</span>', unsafe_allow_html=True)
                if not df.empty:
                    ML_labels = {0:"0-Al corriente",1:"1-Atraso",2:"2-Convenio",3:"3-Devuelve no paga",4:"4-Judicial"}
                    MC = {0:C_PASTEL['success'],1:C_PASTEL['gold'],2:C_PASTEL['warning'],3:C_PASTEL['accent'],4:"#B4607E"}
                    df_mor = df.copy()
                    df_mor['Nivel_Morosidad'] = df_mor['Nivel_Morosidad'].fillna(0).astype(int)
                    dist_m = df_mor.groupby('Nivel_Morosidad').agg(
                        Contratos=('ID_Contrato','count'),
                        Valor=('Valor_Sin_IVA','sum')
                    ).reset_index()
                    dist_m['Label'] = dist_m['Nivel_Morosidad'].map(ML_labels)
                    dist_m['Color'] = dist_m['Nivel_Morosidad'].map(MC)
                    fig_mor = go.Figure()
                    fig_mor.add_trace(go.Bar(
                        x=dist_m['Label'], y=dist_m['Contratos'],
                        marker_color=dist_m['Color'].tolist(),
                        text=dist_m['Contratos'], textposition='outside',
                        customdata=dist_m['Valor'],
                        hovertemplate='<b>%{x}</b><br>Contratos: %{y}<br>Valor: $%{customdata:,.0f}<extra></extra>'
                    ))
                    fig_mor.update_layout(showlegend=False, xaxis_tickangle=-20)
                    fig_mor = sfig(fig_mor, h=240)
                    st.plotly_chart(fig_mor, width='stretch', key="pc_004")
            with mc2:
                st.markdown('<span class="section-label">Concentración por cliente</span>', unsafe_allow_html=True)
                if not df.empty and tot_v > 0:
                    conc = df.groupby('Cliente')['Valor_Sin_IVA'].sum().sort_values(ascending=False).reset_index()
                    conc['% de la cartera'] = conc['Valor_Sin_IVA'] / tot_v * 100
                    top5_pct = conc.head(5)['% de la cartera'].sum()
                    fconc = px.bar(conc.head(6), x='% de la cartera', y='Cliente', orientation='h',
                                    text_auto='.1f', color='% de la cartera',
                                    color_continuous_scale=["#5E9587", "#D9A75C", "#C4796C"])
                    fconc = sfig(fconc, h=240)
                    fconc.update_layout(yaxis={'categoryorder':'total ascending'}, coloraxis_showscale=False)
                    fconc.update_traces(textposition='outside', texttemplate='%{x:.1f}%')
                    st.plotly_chart(fconc, width='stretch', key="pc_concentracion")
                    st.caption(f"Top 1: **{_top1_pct:.1f}%** · Top 5: **{top5_pct:.1f}%** de la cartera. "
                               "Regla común: ningún cliente debería superar 15–20%.")

        with tab_venc:
            st.markdown('<span class="section-label">Vencimientos próximos 12 meses</span>', unsafe_allow_html=True)
            if not df.empty and 'Fecha_Vencimiento' in df.columns:
                meses12 = [(hoy + relativedelta(months=i)).strftime('%Y-%m') for i in range(0,13)]
                df_vt = df.copy()
                df_vt['Mes_Venc'] = df_vt['Fecha_Vencimiento'].dt.to_period('M').astype(str)
                dv12 = df_vt[df_vt['Mes_Venc'].isin(meses12)].groupby('Mes_Venc').agg(
                    Contratos=('ID_Contrato','count'), Valor=('Valor_Sin_IVA','sum')
                ).reset_index().sort_values('Mes_Venc')
                if not dv12.empty:
                    fig_venc = px.bar(dv12, x='Mes_Venc', y='Contratos',
                        color='Valor', color_continuous_scale=[[0,"#F3F4F6"],[1,"#5E9587"]],
                        labels={'Mes_Venc':'Mes','Contratos':'Contratos','Valor':'Valor (MXN)'},
                        text='Contratos')
                    fig_venc.update_traces(textposition='outside',
                        hovertemplate='<b>%{x}</b><br>Contratos: %{y}<br>Valor: $%{marker.color:,.0f}<extra></extra>')
                    fig_venc.update_layout(coloraxis_showscale=False, showlegend=False)
                    fig_venc = sfig(fig_venc, h=230)
                    st.plotly_chart(fig_venc, width='stretch', key="pc_005")
                else:
                    st.info("Sin vencimientos en los próximos 12 meses.")

        with tab_flujo:
            st.markdown('<span class="section-label">Flujo esperado — próximos 12 meses</span>', unsafe_allow_html=True)
            if not df.empty:
                ri12 = proy_rentas(df, 12)
                rf12 = proy_residual(df)
                if not ri12.empty:
                    ri12_c = ri12.copy()
                    if not rf12.empty:
                        rf_d = {r['Mes']:r['Residual_Monto'] for _,r in rf12.iterrows()}
                        ri12_c['Residual'] = ri12_c['Mes'].map(rf_d).fillna(0)
                    else:
                        ri12_c['Residual'] = 0
                    fig_fl = go.Figure()
                    fig_fl.add_trace(go.Bar(x=ri12_c['Mes'], y=ri12_c['Rentas'],
                        name='Rentas', marker_color=C_PASTEL['primary'], opacity=0.85))
                    fig_fl.add_trace(go.Bar(x=ri12_c['Mes'], y=ri12_c['Residual'],
                        name='Residual', marker_color=C_PASTEL['accent'], opacity=0.85))
                    fig_fl.update_layout(barmode='stack', legend=dict(orientation='h', y=1.05))
                    fig_fl = sfig(fig_fl, h=280)
                    st.plotly_chart(fig_fl, width='stretch', key="pc_006")
                    total_12 = ri12_c['Rentas'].sum() + ri12_c['Residual'].sum()
                    st.caption(f"Total proyectado 12 meses: **${total_12:,.0f}**")

        with tab_act:
            st.markdown('<span class="section-label">Últimas anotaciones</span>', unsafe_allow_html=True)
            TIPO_COLORS_D = {"General":"#3E6FA6","Ajuste contable":"#B3261E","Nota legal":"#6E5A9C",
                             "Seguimiento":"#1E5C4F","Alerta":"#96660C","Acuerdo con cliente":"#1C7A4D","Otro":"#8A6D2F"}
            if not df_anot.empty:
                ult = df_anot.sort_values('Fecha', ascending=False).head(6)
                for _, an in ult.iterrows():
                    tipo_a = str(an.get('Tipo','General') or 'General')
                    col_a  = TIPO_COLORS_D.get(tipo_a,'#0369A1')
                    texto_corto = str(an['Texto'])[:80] + ('…' if len(str(an['Texto']))>80 else '')
                    with st.container(key=f"dash_anot_{int(an['id'])}"):
                        st.markdown(f"""
                        <div style="border:1px solid #DCE0E5;border-left:3px solid {col_a};padding:5px 10px;
                                    margin-bottom:5px;background:#FFFFFF;border-radius:0 4px 4px 0;">
                          <span style="font-size:.72rem;color:{col_a};font-weight:700;">{tipo_a}</span>
                          <span style="font-size:.7rem;color:#8A929C;float:right;">{an['Fecha']} · {an['ID_Contrato']}</span><br>
                          <span style="font-size:.82rem;color:#20242B;">{texto_corto}</span>
                        </div>
                    """, unsafe_allow_html=True)
            else:
                st.info("Sin anotaciones registradas.")

        st.markdown("---")

        # TABLA DETALLE + acceso rápido a estado de cuenta
        # Colapsada por default a propósito: el buscador de arriba ya cubre
        # "encuentra un contrato rápido"; esto es para cuando de verdad
        # quieres ver/exportar la cartera completa de un vistazo.
        with st.expander(f"Ver tabla completa de contratos activos ({len(df)})"):
            busq_dash = st.text_input("Filtrar esta tabla por contrato o cliente", placeholder="ID, nombre…", key="dash_busq")
            df_show = df.copy()
            if busq_dash.strip():
                mask_b = (df_show['ID_Contrato'].str.contains(busq_dash, case=False, na=False) |
                          df_show['Cliente'].str.contains(busq_dash, case=False, na=False))
                df_show = df_show[mask_b]

            if not df_show.empty:
                cols_dash = ['ID_Contrato','Cliente','Vehiculo','Fecha_Alta','Fecha_Vencimiento',
                             'Valor_Sin_IVA','Mensualidad_Sin_IVA','Residual_Monto','Nivel_Morosidad',
                             'Fecha_Excl_Poliza','Motivo_Excl_Poliza']
                cols_dash = [c for c in cols_dash if c in df_show.columns]
                fmt_d = {}
                for c in ['Valor_Sin_IVA','Mensualidad_Sin_IVA','Residual_Monto']:
                    if c in df_show.columns: fmt_d[c] = '${:,.2f}'
                st.dataframe(
                    df_show[cols_dash].style.format(fmt_d)
                        .map(lambda v: 'background-color:#FFE8E8' if v and str(v) not in ['0','0.0','nan','None',''] else '',
                                  subset=[c for c in ['Fecha_Excl_Poliza'] if c in cols_dash]),
                    width='stretch', height=340,
            key="df_007")
                # Botón acceso rápido a estado de cuenta
                ids_show = df_show['ID_Contrato'].tolist()
                if len(ids_show) == 1:
                    if st.button(f"Ver estado de cuenta — {ids_show[0]}", width='stretch'):
                        st.session_state['ec_contrato'] = ids_show[0]
                        st.session_state['menu_item']   = "Estado de Cuenta"
                        for g, its in GRUPOS.items():
                            if "Estado de Cuenta" in its:
                                st.session_state['menu_grupo'] = g
                        st.session_state['_refresh'] = True
            else:
                st.info("Sin contratos que coincidan.")


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
                        st.caption("Capital + interés + residual pactado deben cuadrar contra la suma de rentas más el valor residual del contrato.")
                    else:
                        st.info("No se puede calcular la composición: inversión neta no positiva.")

                    st.markdown("---")
                    st.markdown('<span class="section-label">Cuentas contables asignadas a este contrato</span>', unsafe_allow_html=True)
                    st.caption(
                        "Estas son, cuenta por cuenta, las claves de tu catálogo maestro que aplican a este "
                        "contrato para Contpaqi. Las que dejaste en blanco en 'Cuentas (Macro)' no aparecen "
                        "aquí ni se exportan."
                    )
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
                        st.warning("No hay cuentas activas en el catálogo maestro. Ve a 'Cuentas (Macro)' y captura al menos una.")

                    st.markdown("---")
                    st.markdown('<span class="section-label">Otros contratos del mismo cliente</span>', unsafe_allow_html=True)
                    _otros = df_todos[(df_todos['Cliente']==row['Cliente']) & (df_todos['ID_Contrato']!=sel_ec)]
                    if not _otros.empty:
                        st.dataframe(
                            _otros[['ID_Contrato','Vehiculo','Estatus','Plazo','Mensualidad_Sin_IVA','Nivel_Morosidad']]
                                .rename(columns={'Mensualidad_Sin_IVA':'Renta mensual'}),
                            width='stretch', height=180, key=f"df_otros_{sel_ec}"
                        )
                        st.caption("Útil para ver la exposición total de este cliente antes de aprobar un nuevo contrato o negociar una renovación.")
                    else:
                        st.caption("Este cliente no tiene otros contratos registrados en esta empresa.")

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
                    else: st.error("Completa los campos obligatorios.")

    elif menu=="Editar / Eliminar":
        st.title("Edición y Eliminación de Contratos")
        dfa=obtener()
        if dfa.empty: st.info("Sin contratos.")
        else:
            ids=st.selectbox("Contrato",dfa['ID_Contrato'],key="edit_sel_contrato"); row=dfa[dfa['ID_Contrato']==ids].iloc[0]

            with st.expander(f"Cambiar el número de contrato (ahora mismo es {ids})"):
                st.caption(
                    "Renumera el contrato completo: se actualiza también en sus facturas, anotaciones y "
                    "eventos especiales, para que nada se quede apuntando al número viejo."
                )
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
            n_fact_rel = get_db().execute("SELECT COUNT(*) FROM facturas WHERE id_contrato=?", (ids,)).fetchone()[0]
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
            st.caption("Deja notas sobre ajustes contables u observaciones de este contrato en particular.")
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
        st.caption("Los contratos que llegan al final de su plazo se dan de baja por terminación natural automáticamente, cada vez que abres la app.")
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
            st.caption("Selecciona múltiples contratos y registra la baja con un solo clic.")
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
        st.caption("Registra estos eventos y calcula el ajuste contable correcto, incluso si ya pasaron meses o años y la cartera siguió corriendo amortizaciones e intereses normales mientras tanto.")
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
                fecha_ev = c3.date_input("Fecha real en que ocurrió el evento", value=date.today(),
                            help="Aunque haya pasado hace meses o años, captura la fecha real del hecho — el sistema calcula el ajuste a partir de ahí.")
                val_rec = c4.number_input("Valor recuperable estimado (salvamento)", min_value=0.0, value=0.0, step=1000.0)
                monto_recl = st.number_input("Monto reclamado a la aseguradora (si aplica)", min_value=0.0, value=0.0, step=1000.0)
                obs = st.text_area("Observaciones")
                st.markdown("---")
                TIPOS_EXCL_EVT = ["SINIESTRO","ROBO","JURIDICO"]
                excluir_de_poliza = tipo_ev in TIPOS_EXCL_EVT
                if excluir_de_poliza:
                    st.markdown(f"Los eventos de tipo **{tipo_ev}** excluyen automáticamente el contrato de las pólizas de parcialidades.")
                    fecha_excl_evt_reg = st.date_input("¿Desde qué fecha aplica la exclusión?",
                                                        value=fecha_ev, key="fexcl_evt_reg",
                                                        help="El contrato quedará excluido a partir del primer día del mes siguiente a esta fecha.")
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
            st.caption("Eventos ya registrados:")
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
                st.markdown("Mientras no confirmes el reembolso, el cálculo del ajuste **no** incluirá ningún ingreso por seguro — así evitas reconocer una recuperación que aún no es segura (regla del activo contingente, IAS 37 / NIF C-9).")
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
                        st.info(f"Este evento ocurrió hace {calc['mes_hoy']-calc['mes_evento']} mes(es) y el sistema siguió generando amortización e interés como si el contrato estuviera sano. Ese ingreso indebido se revierte junto con la pérdida en el ajuste de abajo.")

                    if abs(calc['loss_plug_anterior'])>0.001:
                        st.error(f"${calc['loss_plug_anterior']:,.2f} de este ajuste corresponde a un ejercicio fiscal ya cerrado. Bajo NIF B-1 / IAS 8 esto se corrige contra **Utilidades Acumuladas**, no contra el resultado del año en curso, y normalmente implica reexpresar los estados comparativos del periodo afectado. Coordínalo con tu auditor.")

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
                        st.caption("Estos son los saldos que tu sistema muestra hoy para este contrato; la póliza de abajo los lleva a cero.")

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
            st.caption(
                "Para empresas cuya forma de contabilizar no encaja con las pólizas estándar de arriba: tú "
                "dictas, concepto por concepto, en qué cuenta va y si es cargo o abono. El sistema calcula el "
                "monto de cada concepto con las mismas fórmulas de amortización de siempre, y arma la póliza."
            )
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
                        help="Layout de ancho fijo verificado carácter por carácter contra un TXT real "
                             "exportado de Contpaqi. Si algún día tu Contpaqi trae otro layout, manda un "
                             "TXT de muestra de ese caso para ajustarlo."
                    )

        else:
            sel_c=st.selectbox("Filtrar contrato",opts)
            c1_,c2_=st.columns(2)
            mes=c1_.selectbox("Mes",range(1,13),index=datetime.now().month-1,format_func=lambda m:MN[m-1])
            anio=c2_.number_input("Año",value=datetime.now().year)
            inc_pm=inc_am=False
            if tipo=="Parcialidad (mensual)": inc_pm=st.checkbox("Incluir mes de alta")
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
        st.markdown("Calcula exactamente cuántos ingresos financieros generará tu cartera en el período seleccionado.")
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

    elif menu=="Cuentas (Macro)":
        st.title("Catálogo Maestro de Cuentas Contables")

        with st.expander("Codificación de cuentas (cómo se arma el número completo por contrato)", expanded=False):
            st.caption(
                "Cada empresa puede tener un esquema distinto para armar el número de cuenta específico de "
                "un contrato (cuenta base + identificador del contrato + sufijo). Esto es **por empresa** — "
                "cambiarlo aquí no afecta a las demás empresas que manejes en Multiempresa."
            )
            _cfg_cod = get_config_codificacion_cuentas()
            cc1, cc2, cc3 = st.columns(3)
            _dbase = cc1.number_input("Dígitos de la cuenta base", min_value=0, max_value=20, value=int(_cfg_cod['digitos_base']))
            _dcontr = cc2.number_input("Dígitos para el contrato", min_value=0, max_value=20, value=int(_cfg_cod['digitos_contrato']))
            _sufijo = cc3.text_input("Sufijo al final", value=_cfg_cod['sufijo'])
            _solo_anexo = st.checkbox(
                "Usar solo la segunda mitad del número de contrato (después del guion)",
                value=_cfg_cod['solo_anexo'],
                help="Actívalo si tus contratos internamente son '0000-0008' pero para la cuenta contable solo "
                     "quieres usar el '8', no las dos mitades juntas."
            )
            _cfg_preview = {'digitos_base': _dbase, 'digitos_contrato': _dcontr, 'sufijo': _sufijo, 'solo_anexo': _solo_anexo}
            _ejemplo = cta_sg("117-00-00", "0524-0001", cfg=_cfg_preview)
            st.info(f"Con esta configuración, la cuenta base **117-00-00** para el contrato **0524-0001** quedaría: **`{_ejemplo}`**")
            if st.button("Guardar codificación de cuentas"):
                set_config_codificacion_cuentas(_cfg_preview)
                st.success("Guardado. Se aplicará a las próximas pólizas y al archivo de Contpaqi que generes.")

        st.caption(
            "Aquí defines, cuenta por cuenta, cómo se traduce cada movimiento del leasing a tu contabilidad. "
            "**Deja una cuenta en blanco si no la usas** — se marcará como inactiva y jamás se incluirá al "
            "generar el archivo de Contpaqi, sin importar cuántos contratos tengas."
        )
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
        st.caption("Genera el archivo de importación de catálogo de cuentas (formato CONTPAQi Contabilidad) a partir de tus contratos vigentes y del catálogo maestro.")

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
        st.caption(
            "Los respaldos automáticos normalmente son una copia plana del .db — cualquiera con acceso a "
            "esa carpeta (una laptop robada, una carpeta de la nube mal compartida, un USB perdido) puede "
            "abrirla. Si pones una contraseña aquí, los respaldos nuevos se guardan cifrados con AES-256 "
            "(.zip en vez de .bak) y nadie puede abrirlos sin ella — ni tú, así que guárdala en un lugar "
            "seguro; si la pierdes, esos respaldos ya no se pueden recuperar."
        )
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
        st.caption(
            "Cada vez que abres la app con la base de datos sana, se guarda automáticamente una copia en "
            "la carpeta 'respaldos_auto' junto a tu base de datos (se conservan las últimas 8). Si algún día "
            "ves el error 'database disk image is malformed', puedes restaurar el más reciente desde aquí."
        )
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
                _etiqueta_cifrado = " 🔒" if _cifrado else ""
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
        st.caption(
            "Esta sección ya está protegida por tu inicio de sesión (solo la ve el rol administrador) — "
            "ya no hace falta un código adicional."
        )
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
                        st.markdown(f"Al asignar **Nivel {nn_cli}** a todos los contratos de **{cli_sel}**, quedarán excluidos de las pólizas de parcialidades a partir del mes siguiente a la fecha que indiques.")
                        fecha_excl_cli = st.date_input("¿Desde qué fecha aplica esta morosidad?",
                                                        value=date.today(), key="fexcl_mora_cli",
                                                        help="El sistema excluirá estos contratos a partir del primer día del mes siguiente a esta fecha.")
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
                        st.markdown(f"Al asignar **Nivel {nn}**, el contrato quedará **excluido de las pólizas de parcialidades** a partir del mes siguiente a la fecha que indiques.")
                        fecha_excl_mora = st.date_input("¿Desde qué fecha aplica esta morosidad?",
                                                         value=date.today(), key="fexcl_mora_ind",
                                                         help="El sistema excluirá este contrato a partir del primer día del mes siguiente a esta fecha.")
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
        st.markdown("Genera 4 tablas: **Intereses leasing** • **Intereses residual** • **Amort. comisión** • **Saldo residual activo**")
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
        st.markdown("Genera las tablas completas de amortización de toda la cartera activa, incluyendo metadatos de los contratos y resúmenes ejecutivos.")
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
                st.subheader(f"📊 Resumen Ejecutivo {anio_m}")
                
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
                st.download_button("📥 Descargar Todo en un Solo Excel (con Gráficas)", buf_completo, f"reporte_maestro_completo_{anio_m}.xlsx", type="primary", use_container_width=True)
                
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
        st.markdown("Genera las tablas de **saldos pendientes** (lo que los clientes aún deben al final de cada mes) de toda la cartera activa.")
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
                st.subheader(f"📊 Resumen de Cartera (Saldos a fin de mes) - {anio_s}")
                
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
                st.download_button("📥 Descargar Todo en un Solo Excel (con Gráficas)", buf_completo, f"reporte_saldos_completo_{anio_s}.xlsx", type="primary", use_container_width=True)
                
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
        st.markdown("Administra múltiples empresas leasing, cada una con su propia base de datos, catálogo y configuración.")
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
        st.markdown(
            "Administra los grupos de acceso y las cuentas de usuario. "
            "Cada persona debe entrar con su propio usuario. El rol define qué puede hacer una vez adentro: "
            "**admin** ve y cambia todo, **captura** da de alta y edita contratos pero no toca la configuración, "
            "**lectura** solo consulta."
        )
        
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
        st.title("Análisis de Punto de Equilibrio")
        st.markdown("¿En qué mes recupera la empresa su inversión en cada contrato? ¿Cuántos meses tarda la cartera en ser rentable?")

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
                st.markdown('<span class="section-label">¿Cuántos contratos ya llegaron a su punto de equilibrio HOY?</span>', unsafe_allow_html=True)
                st.caption("A diferencia de las gráficas de arriba (que miran todo el plazo), esto compara el mes de "
                           "equilibrio contra los meses que realmente ya han transcurrido desde hoy.")
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
        st.markdown(
            "Consulta la tabla de amortización de **leasing** y la tabla de **acumulación del residual** "
            "(VP → intereses acumulados → residual pactado) para cada contrato, incluyendo los que ya "
            "se dieron de baja (su tabla se corta en el mes de la baja, no sigue hasta el plazo original)."
        )

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
                st.markdown("Tabla clásica de amortización: saldo inicial, interés del mes, capital amortizado, saldo final.")
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
                st.markdown(
                    "Tabla de acumulación del residual: parte del **Valor Presente** (VP) del residual "
                    "y cada mes se le agregan intereses hasta llegar al **residual pactado** al vencimiento."
                )
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
        _render_conciliacion()

    elif menu == "Comparar Analíticas":
        _render_analiticas()
except Exception as _e_fatal:
    _pantalla_error_amigable(_e_fatal, contexto=str(st.session_state.get("menu_item","?")))
