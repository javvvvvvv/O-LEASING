# ============================================================================
# PROPIEDAD INTELECTUAL Y LICENCIA COMERCIAL CERRADA
# ============================================================================
# Autor Legal y Titular de Derechos: JAVIER ILLAN GONZALEZ
# Organizacion: ORANGE CREW
# Contacto: ILLANJAVIER9@GMAIL.COM
#
# ADVERTENCIA LEGAL (MEXICO Y GLOBAL):
# Este codigo fuente y su arquitectura son propiedad intelectual exclusiva de
# JAVIER ILLAN GONZALEZ. Queda estrictamente prohibida su reproduccion,
# distribucion, modificacion, ingenieria inversa, copia o uso comercial sin la
# autorizacion expresa y por escrito del autor. Obra protegida conforme a la
# Ley Federal del Derecho de Autor y tratados internacionales aplicables.
# ============================================================================
"""
theme.py - Sistema de Diseno y Gestion de Temas (Light / Dark) para O-Leasing.

Contiene las clases y diccionarios de configuracion de color solicitados:
  - 'Light': Background #F8FAFC, Surface #FFFFFF, Primary #1E293B, Accent #FF6B35
  - 'Dark':  Background #0F172A, Surface #1E293B, Primary #F1F5F9, Accent #FF8B60

Permite alternar dinamicamente entre ambos temas sin reiniciar la aplicacion,
leyendo la preferencia de sistema/URL o un valor por defecto.
"""

from typing import Dict, Any, Optional
import streamlit as st


# ============================================================================
# 1. CLASES DE CONFIGURACION DE COLORES
# ============================================================================

class Light:
    """Configuracion de colores para Modo Claro (Light)."""
    Background: str = "#F8FAFC"
    Surface: str    = "#FFFFFF"
    Primary: str    = "#1E293B"
    Accent: str     = "#FF6B35"

    # Alias en minusculas para maxima compatibilidad
    background: str = "#F8FAFC"
    surface: str    = "#FFFFFF"
    primary: str    = "#1E293B"
    accent: str     = "#FF6B35"


class Dark:
    """Configuracion de colores para Modo Oscuro (Dark)."""
    Background: str = "#0F172A"
    Surface: str    = "#1E293B"
    Primary: str    = "#F1F5F9"
    Accent: str     = "#FF8B60"

    # Alias en minusculas para maxima compatibilidad
    background: str = "#0F172A"
    surface: str    = "#1E293B"
    primary: str    = "#F1F5F9"
    accent: str     = "#FF8B60"


# ============================================================================
# 2. DICCIONARIOS DE CONFIGURACION DE COLORES Y TOKENS ARMONIZADOS
# ============================================================================

LIGHT_CONFIG: Dict[str, Any] = {
    # Tokens nucleares requeridos
    "Background": "#F8FAFC",
    "Surface":    "#FFFFFF",
    "Primary":    "#1E293B",
    "Accent":     "#FF6B35",

    # Alias en minusculas
    "background": "#F8FAFC",
    "surface":    "#FFFFFF",
    "primary":    "#1E293B",
    "accent":     "#FF6B35",

    # Elevaciones y fondos
    "bg_0": "#F8FAFC",
    "bg_1": "#FFFFFF",
    "bg_2": "#F1F5F9",
    "bg_3": "#E2E8F0",

    # Textos
    "text_primary":   "#1E293B",
    "text_secondary": "#64748B",
    "text_disabled":  "#94A3B8",
    "text_link":      "#FF6B35",

    # Marca e interaccion
    "brand":        "#FF6B35",
    "brand_hover":  "#E85924",
    "brand_soft":   "rgba(255, 107, 53, 0.12)",
    "brand_border": "rgba(255, 107, 53, 0.35)",

    # Superficies especificas
    "sidebar_bg":       "rgba(255, 255, 255, 0.90)",
    "topbar_bg":        "rgba(255, 255, 255, 0.92)",
    "card_bg":          "#FFFFFF",
    "card_border":      "rgba(30, 41, 59, 0.08)",
    "table_header_bg":  "#F1F5F9",
    "table_row_hover":  "rgba(30, 41, 59, 0.04)",

    # Bordes
    "border":       "rgba(30, 41, 59, 0.12)",
    "border_hover": "rgba(30, 41, 59, 0.22)",
    "border_focus": "rgba(255, 107, 53, 0.40)",

    # Semanticos
    "success":      "#16A34A",
    "success_soft": "rgba(22, 163, 74, 0.12)",
    "warning":      "#D97706",
    "warning_soft": "rgba(217, 119, 6, 0.12)",
    "danger":       "#DC2626",
    "danger_soft":  "rgba(220, 38, 38, 0.12)",
    "info":         "#2563EB",
    "info_soft":    "rgba(37, 99, 235, 0.12)",
}

DARK_CONFIG: Dict[str, Any] = {
    # Tokens nucleares requeridos
    "Background": "#0F172A",
    "Surface":    "#1E293B",
    "Primary":    "#F1F5F9",
    "Accent":     "#FF8B60",

    # Alias en minusculas
    "background": "#0F172A",
    "surface":    "#1E293B",
    "primary":    "#F1F5F9",
    "accent":     "#FF8B60",

    # Elevaciones y fondos
    "bg_0": "#0F172A",
    "bg_1": "#1E293B",
    "bg_2": "#283548",
    "bg_3": "#334155",

    # Textos
    "text_primary":   "#F1F5F9",
    "text_secondary": "#94A3B8",
    "text_disabled":  "#64748B",
    "text_link":      "#FF8B60",

    # Marca e interaccion
    "brand":        "#FF8B60",
    "brand_hover":  "#FF9F7A",
    "brand_soft":   "rgba(255, 139, 96, 0.15)",
    "brand_border": "rgba(255, 139, 96, 0.40)",

    # Superficies especificas
    "sidebar_bg":       "rgba(30, 41, 59, 0.85)",
    "topbar_bg":        "rgba(30, 41, 59, 0.85)",
    "card_bg":          "#1E293B",
    "card_border":      "rgba(241, 245, 249, 0.08)",
    "table_header_bg":  "#243044",
    "table_row_hover":  "rgba(241, 245, 249, 0.04)",

    # Bordes
    "border":       "rgba(241, 245, 249, 0.12)",
    "border_hover": "rgba(241, 245, 249, 0.24)",
    "border_focus": "rgba(255, 139, 96, 0.45)",

    # Semanticos
    "success":      "#22C55E",
    "success_soft": "rgba(34, 197, 94, 0.15)",
    "warning":      "#F59E0B",
    "warning_soft": "rgba(245, 158, 11, 0.15)",
    "danger":       "#EF4444",
    "danger_soft":  "rgba(239, 68, 68, 0.15)",
    "info":         "#38BDF8",
    "info_soft":    "rgba(56, 189, 248, 0.15)",
}

# Mapeo global de temas
THEMES: Dict[str, Dict[str, Any]] = {
    "Light": LIGHT_CONFIG,
    "Dark":  DARK_CONFIG,
    "light": LIGHT_CONFIG,
    "dark":  DARK_CONFIG,
}

THEME_CLASSES = {
    "Light": Light,
    "Dark":  Dark,
    "light": Light,
    "dark":  Dark,
}


# ============================================================================
# 3. METODOS GLOBALES DE GESTION Y ALTERNANCIA DINAMICA
# ============================================================================

def normalizar_nombre_tema(nombre: Optional[str]) -> str:
    """Devuelve 'Light' o 'Dark' garantizando consistencia."""
    if not nombre:
        return "Dark"
    s = str(nombre).strip().capitalize()
    return "Light" if s == "Light" else "Dark"


def inicializar_tema(default: str = "Dark") -> str:
    """
    Inicializa el tema leyendo primero la preferencia del sistema o URL,
    y si no existe, toma el valor por defecto.
    """
    default_norm = normalizar_nombre_tema(default)
    try:
        if "app_theme" not in st.session_state:
            # 1. Leer parametro URL ?theme=Light|Dark
            qp_tema = None
            if hasattr(st, "query_params"):
                qp_tema = st.query_params.get("theme")
            if qp_tema and qp_tema.capitalize() in ("Light", "Dark"):
                st.session_state["app_theme"] = qp_tema.capitalize()
            else:
                st.session_state["app_theme"] = default_norm
        return st.session_state["app_theme"]
    except Exception:
        return default_norm


def get_current_theme_name() -> str:
    """Devuelve el nombre del tema activo ('Light' o 'Dark')."""
    try:
        if "app_theme" in st.session_state:
            return st.session_state["app_theme"]
        if hasattr(st, "query_params"):
            qp = st.query_params.get("theme")
            if qp and qp.capitalize() in ("Light", "Dark"):
                return qp.capitalize()
    except Exception:
        pass
    return "Dark"


def get_current_theme() -> Dict[str, Any]:
    """Devuelve el diccionario de configuracion del tema activo."""
    nombre = get_current_theme_name()
    return THEMES.get(nombre, DARK_CONFIG)


def get_theme_class(theme_name: Optional[str] = None):
    """Devuelve la clase de configuracion ('Light' o 'Dark')."""
    nombre = normalizar_nombre_tema(theme_name or get_current_theme_name())
    return THEME_CLASSES.get(nombre, Dark)


def set_theme(theme_name: str) -> None:
    """
    Establece el tema activo ('Light' o 'Dark') de manera global e inmediata.
    Sincroniza el session_state y los query_params.
    """
    tema_norm = normalizar_nombre_tema(theme_name)
    try:
        st.session_state["app_theme"] = tema_norm
        if hasattr(st, "query_params"):
            st.query_params["theme"] = tema_norm
    except Exception:
        pass


def toggle_theme() -> str:
    """
    Alterna eficientemente entre Light y Dark sin reiniciar la aplicacion.
    Devuelve el nuevo tema activo.
    """
    actual = get_current_theme_name()
    nuevo = "Dark" if actual == "Light" else "Light"
    set_theme(nuevo)
    return nuevo


def render_theme_toggle(label: str = "Tema visual", key: str = "theme_toggle_widget") -> None:
    """
    Renderiza un componente accesible para alternar entre Light y Dark en tiempo real.
    """
    tema_act = get_current_theme_name()
    opciones = ["Modo Oscuro", "Modo Claro"]
    idx_actual = 0 if tema_act == "Dark" else 1

    try:
        sel = st.segmented_control(
            label,
            options=opciones,
            default=opciones[idx_actual],
            key=key,
            label_visibility="visible"
        )
        if sel:
            nuevo = "Dark" if sel == "Modo Oscuro" else "Light"
            if nuevo != tema_act:
                set_theme(nuevo)
                st.rerun()
    except Exception:
        # Respaldo si segmented_control no estuviera disponible
        nuevo_str = "Modo Claro" if tema_act == "Dark" else "Modo Oscuro"
        if st.button(f"Cambiar a {nuevo_str}", key=key, width="stretch"):
            toggle_theme()
            st.rerun()


def render_theme_toggle_button(key: str = "theme_toggle_btn", width: str = "stretch") -> None:
    """Boton directo de 1 clic para alternar entre Modo Claro y Modo Oscuro."""
    tema_act = get_current_theme_name()
    label = "Cambiar a Modo Claro" if tema_act == "Dark" else "Cambiar a Modo Oscuro"
    if st.button(label, key=key, width=width):
        toggle_theme()
        st.rerun()


# ============================================================================
# 4. GENERADOR E INYECTOR DE ESTILOS CSS
# ============================================================================

def generar_css(theme_name: Optional[str] = None) -> str:
    """Genera las reglas CSS inyectando los tokens del tema seleccionado."""
    t_name = normalizar_nombre_tema(theme_name or get_current_theme_name())
    cfg = THEMES.get(t_name, DARK_CONFIG)

    return f"""
<style>
/* ================================================================================
   SISTEMA DE DISENO O-LEASING - TEMA {t_name.upper()}
   ================================================================================ */

@font-face {{
  font-family: 'Inter';
  font-style: normal;
  font-weight: 100 900;
  font-display: swap;
  src: local('Inter'), url('https://fonts.gstatic.com/s/inter/v12/UcCO3FwrK3iLTeHuS_fvQtMwCp50KnMw2boKoduKmMEVuLyfMZhrib2Bg-4.ttf') format('truetype');
}}

:root {{
  /* Nucleo requerido */
  --bg-color:       {cfg["Background"]};
  --surface-color:  {cfg["Surface"]};
  --primary-color:  {cfg["Primary"]};
  --accent-color:   {cfg["Accent"]};

  /* Elevaciones y Fondos */
  --bg-0: {cfg["bg_0"]};
  --bg-1: {cfg["bg_1"]};
  --bg-2: {cfg["bg_2"]};
  --bg-3: {cfg["bg_3"]};

  /* Texto */
  --text-primary:   {cfg["text_primary"]};
  --text-secondary: {cfg["text_secondary"]};
  --text-disabled:  {cfg["text_disabled"]};
  --text-link:      {cfg["text_link"]};

  /* Marca e Interaccion */
  --brand:        {cfg["brand"]};
  --brand-hover:  {cfg["brand_hover"]};
  --brand-soft:   {cfg["brand_soft"]};
  --brand-border: {cfg["brand_border"]};

  /* Superficies */
  --sidebar-bg:       {cfg["sidebar_bg"]};
  --topbar-bg:        {cfg["topbar_bg"]};
  --card-bg:          {cfg["card_bg"]};
  --table-header-bg:  {cfg["table_header_bg"]};
  --table-row-hover:  {cfg["table_row_hover"]};

  /* Bordes */
  --border:       {cfg["border"]};
  --border-hover: {cfg["border_hover"]};
  --border-focus: {cfg["border_focus"]};

  /* Semanticos */
  --success:      {cfg["success"]};
  --success-soft: {cfg["success_soft"]};
  --warning:      {cfg["warning"]};
  --warning-soft: {cfg["warning_soft"]};
  --danger:       {cfg["danger"]};
  --danger-soft:  {cfg["danger_soft"]};
  --info:         {cfg["info"]};
  --info-soft:    {cfg["info_soft"]};

  /* Radios */
  --radius-sm: 6px;
  --radius-md: 10px;
  --radius-lg: 16px;

  /* Tipografia y Transicion */
  --app-font: "Inter", "Segoe UI", -apple-system, sans-serif;
  --ease-micro: 120ms cubic-bezier(.2,.8,.2,1);
  --ease-std:   200ms cubic-bezier(.2,.8,.2,1);
  --ease-panel: 280ms cubic-bezier(.2,.8,.2,1);
}}

/* Base Streamlit */
.stApp {{
  background-color: var(--bg-0) !important;
  font-family: var(--app-font);
  color: var(--text-primary) !important;
  transition: background-color var(--ease-std), color var(--ease-std);
}}

.main .block-container {{
  padding: 2rem !important;
  max-width: 1440px !important;
  overflow-x: hidden !important;
}}

/* Tipografia */
h1 {{ color: var(--text-primary) !important; font-weight: 600 !important; font-size: 22px !important; line-height: 28px !important; margin-bottom: 0.5rem !important; }}
h2 {{ color: var(--text-primary) !important; font-weight: 600 !important; font-size: 16px !important; line-height: 22px !important; }}
h3 {{ color: var(--text-primary) !important; font-weight: 600 !important; font-size: 14px !important; line-height: 20px !important; }}
.main p, .main .stMarkdown p {{ color: var(--text-secondary) !important; font-size: 14px; line-height: 20px; }}
.main strong, .stMarkdown strong {{ color: var(--text-primary) !important; }}
a {{ color: var(--text-link) !important; text-decoration: none; }}
a:hover {{ text-decoration: underline; }}

/* Numeros tabulares para finanzas */
[data-testid="stMetricValue"], [data-testid="stMetricDelta"], .stDataFrame, table, .tabular-nums {{
  font-variant-numeric: tabular-nums;
}}
[data-testid="stMetricValue"] {{
  font-size: 1.4rem !important;
  word-wrap: break-word !important;
  color: var(--text-primary) !important;
}}
[data-testid="stMetricLabel"] {{
  font-size: 0.85rem !important;
  white-space: normal !important;
  overflow: visible !important;
  color: var(--text-secondary) !important;
}}

/* Sidebar y Top-Bar */
[data-testid="stSidebar"] {{
  background: var(--sidebar-bg) !important;
  backdrop-filter: blur(14px);
  -webkit-backdrop-filter: blur(14px);
  border-right: 1px solid var(--border) !important;
}}
[data-testid="stSidebar"] p, [data-testid="stSidebar"] span, [data-testid="stSidebar"] label {{
  color: var(--text-primary);
}}

.top-header-bar {{
  position: sticky; top: 0; z-index: 999;
  background: var(--topbar-bg);
  backdrop-filter: blur(10px); -webkit-backdrop-filter: blur(10px);
  border: 1px solid var(--border);
  border-radius: var(--radius-md);
  padding: .6rem 1rem; margin-bottom: 1rem;
  display: flex; align-items: center; justify-content: space-between; gap: 10px; flex-wrap: wrap;
  box-shadow: 0 4px 15px rgba(0,0,0,0.08);
}}
.top-header-bar .thb-empresa {{ color: var(--text-primary) !important; font-weight: 600; }}
.top-header-bar .thb-tag {{ display: block; font-size: 12px; color: var(--text-secondary) !important; text-transform: uppercase; letter-spacing: 0.04em; }}

/* Tarjeta orientativa / breadcrumb adaptable */
.nav-breadcrumb-card {{
  background: var(--bg-1) !important;
  border: 1px solid var(--border) !important;
  border-radius: var(--radius-sm);
  padding: .65rem 1.1rem;
  margin-bottom: 1.1rem;
  display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 8px;
}}

/* Contenedores y Tarjetas */
.section-header {{
  background: var(--bg-1);
  border: 1px solid var(--border);
  border-left: 3px solid var(--brand);
  border-radius: var(--radius-md);
  padding: 10px 16px;
  margin: 12px 0 16px 0;
}}
.section-header h3 {{ color: var(--brand) !important; }}

div[data-testid="metric-container"],
.stExpander,
div[data-testid="stForm"],
div[data-testid="stVerticalBlockBorderWrapper"] > div {{
  background: var(--bg-1) !important;
  border: 1px solid var(--border) !important;
  border-radius: var(--radius-md) !important;
  padding: 20px !important;
  box-shadow: none !important;
  transition: border-color var(--ease-micro) !important;
}}
div[data-testid="metric-container"]:hover {{ border-color: var(--border-hover) !important; }}

/* Inputs */
.stTextInput input, .stDateInput input, .stNumberInput input, .stTextArea textarea,
.stSelectbox > div > div, .stMultiSelect > div > div {{
  border-radius: var(--radius-sm) !important;
  border: 1px solid var(--border) !important;
  background-color: var(--bg-1) !important;
  color: var(--text-primary) !important;
  transition: all var(--ease-micro) !important;
}}
.stTextInput input:focus, .stDateInput input:focus, .stNumberInput input:focus, .stTextArea textarea:focus,
.stSelectbox > div > div:focus-within, .stMultiSelect > div > div:focus-within {{
  border-color: var(--border-focus) !important;
  box-shadow: 0 0 0 2px var(--brand-soft) !important;
}}
label[data-testid="stWidgetLabel"] p, label[data-testid="stWidgetLabel"] span {{
  color: var(--text-primary) !important;
  font-weight: 500 !important;
}}

/* Popovers y menus */
div[data-baseweb="popover"], div[data-baseweb="menu"], ul[role="listbox"] {{
  background-color: var(--bg-1) !important;
  border: 1px solid var(--border) !important;
  color: var(--text-primary) !important;
}}
li[role="option"] {{
  color: var(--text-primary) !important;
}}
li[role="option"]:hover, li[role="option"][aria-selected="true"] {{
  background-color: var(--bg-2) !important;
}}

/* Botones */
.stButton button {{
  background: var(--bg-1) !important;
  border: 1px solid var(--border) !important;
  color: var(--text-primary) !important;
  border-radius: var(--radius-md) !important;
  transition: all var(--ease-micro) !important;
  font-weight: 500 !important;
}}
.stButton button:hover {{
  border-color: var(--border-hover) !important;
  background: var(--bg-2) !important;
}}
.stButton button:focus {{
  border-color: var(--border-focus) !important;
  box-shadow: 0 0 0 2px var(--brand-soft) !important;
}}
.stButton button[kind="primary"] {{
  background: var(--brand) !important;
  border: 1px solid var(--brand) !important;
  color: #FFFFFF !important;
}}
.stButton button[kind="primary"]:hover {{
  background: var(--brand-hover) !important;
  border-color: var(--brand-hover) !important;
  transform: translateY(-1px);
}}
.stButton button[kind="primary"]:active {{
  transform: translateY(1px);
}}
.stButton button:disabled {{
  opacity: 0.45 !important;
  cursor: not-allowed !important;
  transform: none !important;
}}

/* Segmented Control y Tabs */
div[data-testid="stSegmentedControl"] button {{
  background: var(--bg-1) !important;
  color: var(--text-secondary) !important;
  border: 1px solid var(--border) !important;
}}
div[data-testid="stSegmentedControl"] button[aria-checked="true"] {{
  background: var(--brand-soft) !important;
  color: var(--brand) !important;
  border-color: var(--brand-border) !important;
  font-weight: 600 !important;
}}

/* Tablas y DataFrames */
.stDataFrame, [data-testid="stDataFrame"], table {{
  background-color: var(--bg-1) !important;
  color: var(--text-primary) !important;
}}
th {{
  background-color: var(--table-header-bg) !important;
  color: var(--text-primary) !important;
  border-bottom: 1px solid var(--border) !important;
}}
td {{
  border-bottom: 1px solid var(--border) !important;
  color: var(--text-primary) !important;
}}
tr:hover td {{
  background-color: var(--table-row-hover) !important;
}}

/* Badges */
.badge-success {{ background-color: var(--success-soft); color: var(--success); padding: 2px 8px; border-radius: var(--radius-sm); font-size: 12px; font-weight: 600; }}
.badge-warning {{ background-color: var(--warning-soft); color: var(--warning); padding: 2px 8px; border-radius: var(--radius-sm); font-size: 12px; font-weight: 600; }}
.badge-danger  {{ background-color: var(--danger-soft);  color: var(--danger);  padding: 2px 8px; border-radius: var(--radius-sm); font-size: 12px; font-weight: 600; }}
.badge-info    {{ background-color: var(--info-soft);    color: var(--info);    padding: 2px 8px; border-radius: var(--radius-sm); font-size: 12px; font-weight: 600; }}
.badge-neutral {{ background-color: var(--bg-2);         color: var(--text-secondary); padding: 2px 8px; border-radius: var(--radius-sm); font-size: 12px; font-weight: 600; }}

/* Alertas y explicaciones */
.chart-explain {{
  background: var(--bg-1);
  border: 1px solid var(--border);
  border-radius: var(--radius-sm);
  padding: 8px 12px;
  margin: 4px 0 10px 0;
}}
.ce-title {{ color: var(--text-primary); font-weight: 600; font-size: 13px; }}
.ce-body  {{ color: var(--text-secondary); font-size: 12px; line-height: 16px; }}

.empty-state {{
  background: var(--bg-1);
  border: 1px dashed var(--border);
  border-radius: var(--radius-md);
  padding: 32px 20px;
  text-align: center;
}}
.es-title {{ color: var(--text-primary); font-weight: 600; font-size: 15px; display: block; margin-bottom: 6px; }}
.es-msg   {{ color: var(--text-secondary); font-size: 13px; display: block; }}

/* Animaciones y utilitarias */
@keyframes pulseDanger {{
  0% {{ box-shadow: 0 0 0 0 rgba(218,54,51,0.4); }}
  70% {{ box-shadow: 0 0 0 6px rgba(218,54,51,0); }}
  100% {{ box-shadow: 0 0 0 0 rgba(218,54,51,0); }}
}}
.pulse-dot {{
  display: inline-block;
  width: 8px; height: 8px;
  background-color: var(--danger);
  border-radius: 50%;
  animation: pulseDanger 2s infinite;
  margin-right: 6px;
}}

.label-caps {{
  font-size: 11px;
  line-height: 14px;
  font-weight: 600;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: var(--text-secondary);
}}
</style>
<script>
(function() {{
  try {{
    document.documentElement.setAttribute('data-theme', '{t_name}');
    const params = new URLSearchParams(window.location.search);
    if (!params.has('theme') && !sessionStorage.getItem('oleasing_theme_init')) {{
      sessionStorage.setItem('oleasing_theme_init', '1');
      const prefersDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
      const detected = prefersDark ? 'Dark' : 'Light';
      if (detected !== '{t_name}') {{
        params.set('theme', detected);
        window.location.search = params.toString();
      }}
    }}
  }} catch(e) {{}}
}})();
</script>
"""


def inyectar_css(theme_name: Optional[str] = None) -> None:
    """Inyecta el bloque de estilo CSS del tema actual en la interfaz de Streamlit."""
    css_content = generar_css(theme_name)
    st.markdown(css_content, unsafe_allow_html=True)
