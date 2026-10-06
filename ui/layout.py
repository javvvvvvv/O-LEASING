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
ui/layout.py - Componente de Layout principal para O-LEASING.

Estructura:
  1. Sidebar (menu lateral a la izquierda):
     - Moderno, con bordes redondeados y altura completa (100vh).
     - Botones de navegacion nativos con iconos vectoriales para todas las
       pestanas del sistema (Dashboard, Contratos, Clientes, Configuracion, etc.).
     - Transicion de 300ms ease-in-out para colapso y expansion fluida.

  2. TopBar (barra superior limpia):
     - Muestra contexto empresarial, seccion actual y migas de pan.
     - Incluye boton estilizado en color de acento naranja (#FF6B35 en Light,
       #FF8B60 en Dark) para alternar dinamicamente entre Modo Claro y Oscuro.

  3. Container (area de contenido principal a la derecha):
     - Padding general de 24px en toda la vista.
     - Fondo de profundidad (--bg-0) que genera contraste visual con las
       tarjetas y formularios (--bg-1).

Aplica directamente los tokens de color definidos en theme.py:
  Light: Background #F8FAFC, Surface #FFFFFF, Primary #1E293B, Accent #FF6B35
  Dark:  Background #0F172A, Surface #1E293B, Primary #F1F5F9, Accent #FF8B60
"""

from typing import Dict, List, Optional, Tuple, Any
from contextlib import contextmanager
import streamlit as st
import theme


# ============================================================================
# MAPA DE ICONOS DE NAVEGACION (Material Symbols Nativos - Sin Emojis)
# ============================================================================

GRUPO_ICONOS: Dict[str, str] = {
    "Cartera":                 ":material/folder_open:",
    "Gestión de Riesgo":       ":material/shield:",
    "Finanzas & Contabilidad": ":material/payments:",
    "Análisis":                ":material/insights:",
    "Configuración":           ":material/settings:",
}

ITEM_ICONOS: Dict[str, str] = {
    # Cartera
    "Dashboard & Cartera":         ":material/dashboard:",
    "Estado de Cuenta":            ":material/receipt_long:",
    "Carga Masiva y Altas":        ":material/upload_file:",
    "Editar / Eliminar":           ":material/edit_note:",
    "Gestor de Bajas":             ":material/archive:",
    "Tabla Mensual por Contrato":  ":material/table_rows:",
    "Reporte Maestro":             ":material/table_chart:",
    "Reporte Maestro Saldos":      ":material/account_balance_wallet:",
    # Riesgo
    "Gestión de Morosidad":        ":material/warning:",
    "Eventos Especiales":          ":material/notifications:",
    "Anotaciones":                 ":material/note_alt:",
    # Finanzas & Contabilidad
    "Pólizas Contables":           ":material/article:",
    "Intereses del Mes":           ":material/percent:",
    "Facturación de Intereses":    ":material/receipt:",
    "Conciliación de Facturas":    ":material/fact_check:",
    "Cierre y Conciliación Mensual": ":material/verified:",
    "Comparar Analíticas":         ":material/compare_arrows:",
    "Tablas de Amortización":      ":material/calendar_month:",
    # Analisis & Clientes
    "Proyección Financiera":       ":material/trending_up:",
    "Cotizador Comercial":         ":material/calculate:",
    "Análisis de Rentabilidad":    ":material/query_stats:",
    "Punto de Equilibrio":         ":material/balance:",
    "Reportes por Cliente":        ":material/group:",
    # Configuracion
    "Cuentas (Macro)":             ":material/format_list_bulleted:",
    "Contpaqi (Cuentas)":          ":material/sync_alt:",
    "Respaldo y Restauración":     ":material/cloud_download:",
    "Multiempresa":                ":material/domain:",
    "Usuarios y Roles":            ":material/manage_accounts:",
}


# ============================================================================
# ESTILOS DEL LAYOUT (Sidebar 300ms, TopBar con acento naranja, Container 24px)
# ============================================================================

def inyectar_layout_css() -> None:
    """Inyecta reglas CSS para Sidebar redondeado, animacion 300ms y TopBar estilizada."""
    cfg = theme.get_current_theme()
    accent = cfg.get("Accent", "#FF6B35")

    st.markdown(f"""
<style>
/* ================================================================================
   REGLAS DEL COMPONENTE LAYOUT (O-LEASING)
   ================================================================================ */

/* 1. SIDEBAR: Altura completa (100vh), borde redondeado derecho y transicion 300ms */
[data-testid="stSidebar"] {{
  height: 100vh !important;
  border-top-right-radius: 16px !important;
  border-bottom-right-radius: 16px !important;
  box-shadow: 4px 0 24px rgba(0, 0, 0, 0.07) !important;
  transition: all 300ms ease-in-out !important;
}}

/* Transicion suave en el boton nativo de expandir/colapsar */
[data-testid="stSidebarCollapseButton"] button,
button[kind="header"] {{
  transition: all 300ms ease-in-out !important;
}}

/* 2. CONTAINER PRINCIPAL: Padding exacto de 24px y fondo de profundidad */
.main .block-container {{
  padding: 24px !important;
  max-width: 1440px !important;
  background-color: var(--bg-0) !important;
  min-height: 100vh !important;
  transition: padding 300ms ease-in-out, background-color var(--ease-std) !important;
}}

/* Contraste de tarjetas (--bg-1) sobre el fondo general (--bg-0) */
div[data-testid="metric-container"],
.stExpander,
div[data-testid="stForm"],
div[data-testid="stVerticalBlockBorderWrapper"] > div {{
  background: var(--bg-1) !important;
  border: 1px solid var(--border) !important;
  border-radius: var(--radius-md) !important;
  box-shadow: 0 2px 10px rgba(0, 0, 0, 0.03) !important;
}}

/* 3. TOPBAR SUPERIOR: Limpia con sombras suaves */
.oleasing-topbar {{
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 10px 18px;
  background: var(--topbar-bg);
  backdrop-filter: blur(14px);
  -webkit-backdrop-filter: blur(14px);
  border: 1px solid var(--border);
  border-radius: 12px;
  margin-bottom: 20px;
  box-shadow: 0 4px 18px rgba(0, 0, 0, 0.04);
  transition: background-color 300ms ease-in-out, border-color 300ms ease-in-out;
}}

.oleasing-topbar-brand {{
  display: flex;
  align-items: center;
  gap: 12px;
}}

.oleasing-topbar-info {{
  display: flex;
  flex-direction: column;
}}

.oleasing-topbar-company {{
  font-size: 14px;
  font-weight: 700;
  color: var(--text-primary);
  line-height: 18px;
}}

.oleasing-topbar-sub {{
  font-size: 11px;
  color: var(--text-secondary);
  text-transform: uppercase;
  letter-spacing: 0.04em;
  font-weight: 600;
}}

.oleasing-topbar-breadcrumbs {{
  display: flex;
  align-items: center;
  font-size: 13px;
  gap: 6px;
}}

/* Boton de acento naranja en la TopBar */
.stButton > button[key="btn_topbar_theme_toggle"],
div[data-testid="stButton"] > button:has(div:contains("Modo")) {{
  background-color: {accent} !important;
  color: #FFFFFF !important;
  border: none !important;
  border-radius: 20px !important;
  font-weight: 600 !important;
  box-shadow: 0 2px 10px rgba(255, 107, 53, 0.3) !important;
  transition: all 200ms ease-in-out !important;
}}
</style>
""", unsafe_allow_html=True)


# ============================================================================
# COMPONENTE TOPBAR (Barra Superior)
# ============================================================================

def render_topbar(
    empresa_nombre: str = "O-Leasing",
    menu_actual: str = "Dashboard",
    grupo_actual: str = "Cartera",
    descripcion: str = "",
    fecha_texto: str = "",
    viendo_pasada: bool = False,
    logo_html: Optional[str] = None
) -> None:
    """
    Renderiza la TopBar superior limpia, con boton estilizado de acento
    naranja para alternar dinamicamente entre Modo Claro y Oscuro.
    """
    tema_act = theme.get_current_theme_name()
    icono_btn = ":material/light_mode:" if tema_act == "Dark" else ":material/dark_mode:"
    label_btn = "Modo Claro" if tema_act == "Dark" else "Modo Oscuro"

    # Logo fallback si no se pasa uno especifico
    if not logo_html:
        logo_html = '<div style="font-weight:800;font-size:18px;color:var(--brand);padding:2px 8px;border:1px solid var(--border);border-radius:6px;">O</div>'

    col_tb1, col_tb2 = st.columns([3.4, 1.0])

    with col_tb1:
        st.markdown(f"""
<div class="oleasing-topbar">
  <div class="oleasing-topbar-brand">
    {logo_html}
    <div class="oleasing-topbar-info">
      <span class="oleasing-topbar-sub">Empresa Activa</span>
      <span class="oleasing-topbar-company">{empresa_nombre}</span>
    </div>
    <div style="height:22px;width:1px;background:var(--border);margin:0 4px;"></div>
    <div class="oleasing-topbar-breadcrumbs">
      <span style="font-weight:600;color:var(--text-secondary);">{grupo_actual}</span>
      <span style="opacity:0.5;color:var(--text-secondary);">›</span>
      <span style="font-weight:700;color:var(--brand);">{menu_actual}</span>
      {f'<span style="font-size:12px;color:var(--text-secondary);margin-left:6px;opacity:0.85;">— {descripcion}</span>' if descripcion else ''}
    </div>
  </div>
  <div>
    {f'''<div style="background:{'#96660C' if viendo_pasada else 'var(--bg-2)'};color:{'#FFFFFF' if viendo_pasada else 'var(--text-secondary)'};font-size:11px;font-weight:600;padding:4px 10px;border-radius:6px;border:1px solid var(--border);white-space:nowrap;">{'Cierre al' if viendo_pasada else 'Cálculos al'} {fecha_texto}</div>''' if fecha_texto else ''}
  </div>
</div>
""", unsafe_allow_html=True)

    with col_tb2:
        if st.button(
            label_btn,
            key="btn_topbar_theme_toggle",
            icon=icono_btn,
            type="primary",
            use_container_width=True,
            help=f"Cambiar dinámicamente a {label_btn.lower()}"
        ):
            theme.toggle_theme()
            st.rerun()


# ============================================================================
# COMPONENTE SIDEBAR (Menu Lateral con Iconos)
# ============================================================================

def render_sidebar_navigation(
    grupos: Dict[str, List[str]],
    menu_actual: str,
    grupo_actual: str,
    pantallas_permitidas: Optional[List[str]] = None,
    descripciones: Optional[Dict[str, str]] = None,
    descripciones_grupo: Optional[Dict[str, str]] = None
) -> Tuple[str, str]:
    """
    Renderiza la navegacion del sidebar con iconos para cada pestana y seccion.
    Aplica una transicion suave de 300ms ease-in-out.
    """
    descripciones = descripciones or {}
    descripciones_grupo = descripciones_grupo or {}

    st.sidebar.markdown(
        '<div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.05em;color:var(--text-secondary);padding:6px 2px;">Navegación</div>',
        unsafe_allow_html=True
    )

    menu_seleccionado = menu_actual
    grupo_seleccionado = grupo_actual

    for grupo, items in grupos.items():
        items_visibles = [it for it in items if (pantallas_permitidas is None or it in pantallas_permitidas)]
        if not items_visibles:
            continue

        abierto = (grupo_actual == grupo)
        icono_grp = GRUPO_ICONOS.get(grupo, ":material/folder:")

        # Boton de seccion/grupo con icono
        if st.sidebar.button(
            grupo,
            key=f"nav_grp_{grupo}",
            icon=icono_grp,
            use_container_width=True,
            type="primary" if abierto else "secondary",
            help=descripciones_grupo.get(grupo, f"Sección {grupo}")
        ):
            grupo_seleccionado = grupo
            menu_seleccionado = items_visibles[0]
            st.session_state['menu_grupo'] = grupo_seleccionado
            st.session_state['menu_item'] = menu_seleccionado
            st.session_state['_refresh'] = True

        # Botones hijos si la seccion esta abierta
        if abierto:
            for item in items_visibles:
                activo = (menu_actual == item)
                icono_item = ITEM_ICONOS.get(item, ":material/arrow_right:")
                tipo_btn = "primary" if activo else "secondary"

                if st.sidebar.button(
                    item,
                    key=f"nav_item_{item}",
                    icon=icono_item,
                    use_container_width=True,
                    type=tipo_btn,
                    help=descripciones.get(item, "")
                ):
                    menu_seleccionado = item
                    grupo_seleccionado = grupo
                    st.session_state['menu_item'] = menu_seleccionado
                    st.session_state['menu_grupo'] = grupo_seleccionado
                    st.session_state['_refresh'] = True

    return menu_seleccionado, grupo_seleccionado


# ============================================================================
# COMPONENTE CONTAINER PRINCIPAL
# ============================================================================

@contextmanager
def main_container():
    """
    Context manager para estructurar el contenedor de contenido principal
    con padding general de 24px y fondo de profundidad (--bg-0 vs --bg-1).
    """
    st.markdown('<div class="oleasing-main-container">', unsafe_allow_html=True)
    try:
        yield
    finally:
        st.markdown('</div>', unsafe_allow_html=True)


# ============================================================================
# CLASE PRINCIPAL DEL LAYOUT
# ============================================================================

class Layout:
    """Clase unificada para invocar la arquitectura del Layout."""

    @staticmethod
    def inyectar_estilos():
        inyectar_layout_css()

    @staticmethod
    def render_topbar(*args, **kwargs):
        render_topbar(*args, **kwargs)

    @staticmethod
    def render_sidebar(*args, **kwargs):
        return render_sidebar_navigation(*args, **kwargs)

    @staticmethod
    def container():
        return main_container()
