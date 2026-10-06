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
ui_components.py - Modulo de Componentes UI Empresariales para O-LEASING.

Define clases personalizadas que heredan de los controles estandar de interfaz,
aplicando las reglas de diseno corporativo y respondiendo dinamicamente al cambio
de tema Claro / Oscuro:

1. PrimaryButton (hereda de BaseButton):
   - Fondo naranja de acento (#FF6B35 en Light, #FF8B60 en Dark)
   - Texto blanco (#FFFFFF)
   - border-radius de 8px
   - Padding interno comodo (10px 22px)
   - Efecto hover con transicion suave que cambia ligeramente el color

2. OutlinedButton (hereda de BaseButton):
   - Borde del color primario (#1E293B en Light, #F1F5F9 en Dark)
   - Fondo transparente
   - Texto del color primario
   - border-radius de 8px
   - Efecto hover sutil respetando transparencia

3. CardContainer (hereda de BaseContainer):
   - border-radius de 12px
   - Color de fondo 'Surface' (#FFFFFF en Light, #1E293B en Dark)
   - Sombra (elevation) muy sutil para dar profundidad visual
   - Borde sutil de definicion y padding interno comodo (22px 26px)
   - Soporte para protocolo context manager (uso con `with CardContainer():`)

4. FormInput (hereda de BaseInput):
   - Campos de texto con border-radius de 8px
   - Bordes en gris suave (#CBD5E1 en Light, #334155 en Dark)
   - Al hacer 'focus', el borde cambia al color naranja de acento con halo suave
"""

from typing import Dict, Any, Optional, Tuple, Callable
from contextlib import AbstractContextManager
import streamlit as st
import theme


# ============================================================================
# INYECCION DE ESTILOS CSS EMPRESARIALES PARA COMPONENTES
# ============================================================================

def inyectar_componentes_css(theme_name: Optional[str] = None) -> None:
    """
    Inyecta las reglas CSS empresariales para PrimaryButton, OutlinedButton,
    CardContainer y FormInput sincronizadas con el tema activo.
    """
    if theme_name:
        nombre = getattr(theme, "normalizar_nombre_tema", lambda x: "Dark")(theme_name)
        cfg = getattr(theme, "THEMES", {}).get(nombre, getattr(theme, "DARK_CONFIG", {}))
    else:
        cfg = theme.get_current_theme()
    accent = cfg.get("Accent", "#FF6B35")
    primary = cfg.get("Primary", "#1E293B")
    surface = cfg.get("Surface", "#FFFFFF")
    background = cfg.get("Background", "#F8FAFC")
    border = cfg.get("border", "rgba(30, 41, 59, 0.12)")
    brand_hover = cfg.get("brand_hover", "#E85924")
    brand_soft = cfg.get("brand_soft", "rgba(255, 107, 53, 0.15)")
    text_primary = cfg.get("text_primary", primary)
    text_secondary = cfg.get("text_secondary", "#64748B")

    # Configuracion de elevacion y bordes segun Light o Dark
    nombre = theme.normalizar_nombre_tema(theme_name or theme.get_current_theme_name())
    if nombre == "Dark":
        elevation = "0 4px 20px rgba(0, 0, 0, 0.32), 0 1px 4px rgba(0, 0, 0, 0.20)"
        input_border = "#334155"
        outlined_hover_bg = "rgba(241, 245, 249, 0.08)"
    else:
        elevation = "0 4px 18px rgba(0, 0, 0, 0.05), 0 1px 3px rgba(0, 0, 0, 0.03)"
        input_border = "#CBD5E1"
        outlined_hover_bg = "rgba(30, 41, 59, 0.05)"

    css = f"""
<style>
/* ============================================================================
   O-LEASING UI COMPONENTS - REGLAS DE DISENO EMPRESARIAL
   ============================================================================ */

/* 1. PRIMARY BUTTON: Naranja de acento, texto blanco, 8px radius, padding comodo, hover */
.stButton > button[kind="primary"],
div:has(> .stMarkdown .oleasing-primary-marker) + div .stButton button,
div:has(.oleasing-primary-marker) + div button {{
  background-color: {accent} !important;
  color: #FFFFFF !important;
  border: 1px solid {accent} !important;
  border-radius: 8px !important;
  padding: 10px 22px !important;
  font-weight: 600 !important;
  font-size: 14px !important;
  line-height: 1.4 !important;
  box-shadow: 0 2px 8px {brand_soft} !important;
  transition: all 200ms ease-in-out !important;
}}

.stButton > button[kind="primary"]:hover,
div:has(> .stMarkdown .oleasing-primary-marker) + div .stButton button:hover,
div:has(.oleasing-primary-marker) + div button:hover {{
  background-color: {brand_hover} !important;
  border-color: {brand_hover} !important;
  color: #FFFFFF !important;
  box-shadow: 0 4px 14px {brand_soft} !important;
  transform: translateY(-1px) !important;
}}

.stButton > button[kind="primary"]:active,
div:has(> .stMarkdown .oleasing-primary-marker) + div .stButton button:active,
div:has(.oleasing-primary-marker) + div button:active {{
  transform: translateY(1px) !important;
  box-shadow: 0 1px 4px {brand_soft} !important;
}}

/* 2. OUTLINED BUTTON: Borde primario, fondo transparente, texto primario, 8px radius, hover */
div:has(> .stMarkdown .oleasing-outlined-marker) + div .stButton button,
div:has(.oleasing-outlined-marker) + div button {{
  background-color: transparent !important;
  background: transparent !important;
  color: {primary} !important;
  border: 1.5px solid {primary} !important;
  border-radius: 8px !important;
  padding: 9px 20px !important;
  font-weight: 600 !important;
  font-size: 14px !important;
  line-height: 1.4 !important;
  box-shadow: none !important;
  transition: all 200ms ease-in-out !important;
}}

div:has(> .stMarkdown .oleasing-outlined-marker) + div .stButton button:hover,
div:has(.oleasing-outlined-marker) + div button:hover {{
  background-color: {outlined_hover_bg} !important;
  border-color: {primary} !important;
  color: {primary} !important;
  transform: translateY(-1px) !important;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05) !important;
}}

div:has(> .stMarkdown .oleasing-outlined-marker) + div .stButton button:active,
div:has(.oleasing-outlined-marker) + div button:active {{
  transform: translateY(1px) !important;
}}

/* 3. CARD CONTAINER: Border-radius 12px, fondo 'Surface', elevacion/sombra sutil, margen interno 20px */
div[data-testid="stVerticalBlockBorderWrapper"]:has(.oleasing-card-marker) > div,
.oleasing-card-container {{
  background-color: {surface} !important;
  border: 1px solid {border} !important;
  border-radius: 12px !important;
  box-shadow: {elevation} !important;
  padding: 20px !important;
  margin-bottom: 20px !important;
  transition: background-color 300ms ease-in-out, border-color 300ms ease-in-out, box-shadow 300ms ease-in-out !important;
}}

.oleasing-card-header {{
  border-bottom: 1px solid {border};
  padding-bottom: 10px;
  margin-bottom: 16px;
}}

.oleasing-card-title {{
  font-size: 16px;
  font-weight: 700;
  color: {text_primary};
  line-height: 22px;
}}

.oleasing-card-subtitle {{
  font-size: 12px;
  font-weight: 500;
  color: {text_secondary};
  margin-top: 3px;
}}

/* DATA TABLES EMPRESARIALES: Encabezado color primario, bordes sutiles y hover effect */
.stDataFrame, [data-testid="stDataFrame"], table, [data-testid="stTable"] {{
  background-color: {surface} !important;
  border: 1px solid {border} !important;
  border-radius: 8px !important;
  overflow: hidden !important;
}}

thead th, th, [data-testid="stTable"] th, .stDataFrame [role="columnheader"] {{
  background-color: {primary} !important;
  color: #FFFFFF !important;
  font-weight: 600 !important;
  font-size: 13px !important;
  letter-spacing: 0.02em !important;
  padding: 10px 14px !important;
  border-bottom: 1px solid {border} !important;
  border-right: 1px solid {border} !important;
}}

thead th p, th p, thead th span, th span {{
  color: #FFFFFF !important;
}}

tbody td, td, [data-testid="stTable"] td, .stDataFrame [role="gridcell"] {{
  border-bottom: 1px solid {border} !important;
  border-right: 1px solid {border} !important;
  color: {text_primary} !important;
  padding: 10px 14px !important;
  font-size: 13px !important;
  transition: background-color 150ms ease-in-out !important;
}}

tbody tr:hover td, tr:hover td, [data-testid="stTable"] tr:hover td {{
  background-color: {outlined_hover_bg} !important;
}}

/* 4. FORM INPUT: Border-radius 8px, borde gris suave, focus al color naranja de acento */
.stTextInput input, .stDateInput input, .stNumberInput input, .stTextArea textarea,
.stSelectbox > div > div, .stMultiSelect > div > div,
div:has(> .stMarkdown .oleasing-input-marker) + div .stTextInput input,
div:has(.oleasing-input-marker) + div input {{
  border-radius: 8px !important;
  border: 1px solid {input_border} !important;
  background-color: {surface} !important;
  color: {text_primary} !important;
  padding: 8px 12px !important;
  font-size: 14px !important;
  transition: border-color 200ms ease-in-out, box-shadow 200ms ease-in-out !important;
}}

.stTextInput input:focus, .stDateInput input:focus, .stNumberInput input:focus, .stTextArea textarea:focus,
.stSelectbox > div > div:focus-within, .stMultiSelect > div > div:focus-within,
div:has(> .stMarkdown .oleasing-input-marker) + div .stTextInput input:focus,
div:has(.oleasing-input-marker) + div input:focus {{
  border-color: {accent} !important;
  box-shadow: 0 0 0 3px {brand_soft} !important;
  outline: none !important;
}}

.stTextInput input::placeholder, .stTextArea textarea::placeholder {{
  color: {cfg.get('text_disabled', '#94A3B8')} !important;
}}

/* LETRAS Y TIPOGRAFIA GENERAL */
label[data-testid="stWidgetLabel"] p, label[data-testid="stWidgetLabel"] span {{
  color: {text_primary} !important;
  font-weight: 600 !important;
  font-size: 13px !important;
}}
.stCaption, [data-testid="stCaptionContainer"] p {{
  color: {text_secondary} !important;
}}

/* Marcadores ocultos para vinculacion DOM sin alterar el layout */
.oleasing-primary-marker,
.oleasing-outlined-marker,
.oleasing-card-marker,
.oleasing-input-marker {{
  display: none !important;
  height: 0 !important;
  width: 0 !important;
}}
</style>
"""
    st.markdown(css, unsafe_allow_html=True)


# ============================================================================
# CLASE BASE ESTANDAR (UIComponent)
# ============================================================================

class UIComponent:
    """Clase base estandar para todos los componentes de interfaz en O-LEASING."""

    @classmethod
    def get_theme_config(cls) -> Dict[str, Any]:
        """Obtiene la configuracion de tokens del tema activo."""
        return theme.get_current_theme()

    @classmethod
    def get_theme_name(cls) -> str:
        """Obtiene el nombre del tema activo ('Light' o 'Dark')."""
        return theme.get_current_theme_name()

    @classmethod
    def inyectar_estilos(cls) -> None:
        """Inyecta los estilos CSS de los componentes en la sesion actual."""
        inyectar_componentes_css()


# ============================================================================
# CLASE CONTROL ESTANDAR DE BOTON (BaseButton)
# ============================================================================

class BaseButton(UIComponent):
    """
    Control estandar de boton.
    Proporciona la interfaz base para botones interactivos en Streamlit.
    """

    def __init__(
        self,
        label: str,
        key: Optional[str] = None,
        help: Optional[str] = None,
        on_click: Optional[Callable[..., Any]] = None,
        args: Optional[Tuple[Any, ...]] = None,
        kwargs: Optional[Dict[str, Any]] = None,
        disabled: bool = False,
        use_container_width: bool = True,
        icon: Optional[str] = None
    ):
        self.label = label
        self.key = key or f"btn_{abs(hash((label, self.__class__.__name__)))}"
        self.help = help
        self.on_click = on_click
        self.args = args
        self.kwargs = kwargs
        self.disabled = disabled
        self.use_container_width = use_container_width
        self.icon = icon
        self._clicked = False

    def render(self) -> bool:
        """Renderiza el boton en Streamlit y retorna True si fue pulsado."""
        raise NotImplementedError("Las subclases deben implementar el metodo render().")

    def __call__(self) -> bool:
        """Permite invocar la instancia directamente como una funcion."""
        return self.render()

    def __bool__(self) -> bool:
        """Permite evaluar la instancia directamente en sentencias if."""
        return self.render()


# ============================================================================
# 1. PRIMARY BUTTON
# ============================================================================

class PrimaryButton(BaseButton):
    """
    Boton primario con fondo naranja de acento, texto blanco,
    border-radius de 8px, padding interno comodo y efecto hover que
    cambia ligeramente el color. Responde dinamicamente a Light y Dark.
    """

    @classmethod
    def render_button(
        cls,
        label: str,
        key: Optional[str] = None,
        help: Optional[str] = None,
        on_click: Optional[Callable[..., Any]] = None,
        args: Optional[Tuple[Any, ...]] = None,
        kwargs: Optional[Dict[str, Any]] = None,
        disabled: bool = False,
        use_container_width: bool = True,
        icon: Optional[str] = None
    ) -> bool:
        """Metodo de clase para renderizar el PrimaryButton directamente."""
        instancia = cls(
            label=label,
            key=key,
            help=help,
            on_click=on_click,
            args=args,
            kwargs=kwargs,
            disabled=disabled,
            use_container_width=use_container_width,
            icon=icon
        )
        return instancia.render()

    # Alias de clase para maxima conveniencia
    create = render_button

    def render(self) -> bool:
        inyectar_componentes_css()
        marker_id = f"mk_prim_{self.key}"
        st.markdown(f'<span class="oleasing-primary-marker" id="{marker_id}"></span>', unsafe_allow_html=True)
        clicked = st.button(
            label=self.label,
            key=self.key,
            help=self.help,
            on_click=self.on_click,
            args=self.args,
            kwargs=self.kwargs,
            disabled=self.disabled,
            use_container_width=self.use_container_width,
            icon=self.icon,
            type="primary"
        )
        self._clicked = clicked
        return clicked


# ============================================================================
# 2. OUTLINED BUTTON
# ============================================================================

class OutlinedButton(BaseButton):
    """
    Boton outlined con borde del color primario, fondo transparente,
    texto del color primario, border-radius de 8px y efecto hover sutil.
    Responde dinamicamente a Light y Dark.
    """

    @classmethod
    def render_button(
        cls,
        label: str,
        key: Optional[str] = None,
        help: Optional[str] = None,
        on_click: Optional[Callable[..., Any]] = None,
        args: Optional[Tuple[Any, ...]] = None,
        kwargs: Optional[Dict[str, Any]] = None,
        disabled: bool = False,
        use_container_width: bool = True,
        icon: Optional[str] = None
    ) -> bool:
        """Metodo de clase para renderizar el OutlinedButton directamente."""
        instancia = cls(
            label=label,
            key=key,
            help=help,
            on_click=on_click,
            args=args,
            kwargs=kwargs,
            disabled=disabled,
            use_container_width=use_container_width,
            icon=icon
        )
        return instancia.render()

    create = render_button

    def render(self) -> bool:
        inyectar_componentes_css()
        marker_id = f"mk_outl_{self.key}"
        st.markdown(f'<span class="oleasing-outlined-marker" id="{marker_id}"></span>', unsafe_allow_html=True)
        clicked = st.button(
            label=self.label,
            key=self.key,
            help=self.help,
            on_click=self.on_click,
            args=self.args,
            kwargs=self.kwargs,
            disabled=self.disabled,
            use_container_width=self.use_container_width,
            icon=self.icon,
            type="secondary"
        )
        self._clicked = clicked
        return clicked


# ============================================================================
# CLASE CONTROL ESTANDAR DE CONTENEDOR (BaseContainer)
# ============================================================================

class BaseContainer(UIComponent, AbstractContextManager):
    """
    Control estandar de contenedor con protocolo context manager (soporte 'with').
    """

    def __init__(
        self,
        key: Optional[str] = None,
        padding: str = "20px"
    ):
        self.key = key or f"cont_{abs(hash(self.__class__.__name__))}"
        self.padding = padding
        self._delta_container = None

    def __enter__(self):
        raise NotImplementedError("Las subclases deben implementar __enter__().")

    def __exit__(self, exc_type, exc_val, exc_tb):
        raise NotImplementedError("Las subclases deben implementar __exit__().")


# ============================================================================
# 3. CARD CONTAINER
# ============================================================================

class CardContainer(BaseContainer):
    """
    Contenedor con border-radius de 12px, color de fondo 'Surface'
    (dependiendo del tema) y una sombra (elevation) muy sutil para dar profundidad.
    Margen interno de 20px exacto.
    Soporta utilizacion con bloque `with` (context manager) o renderizado directo.
    """

    def __init__(
        self,
        title: Optional[str] = None,
        subtitle: Optional[str] = None,
        key: Optional[str] = None,
        padding: str = "20px",
        border: bool = True
    ):
        super().__init__(key=key, padding=padding)
        self.title = title
        self.subtitle = subtitle
        self.border = border

    def __enter__(self):
        inyectar_componentes_css()
        self._delta_container = st.container(border=self.border, key=self.key)
        ctx = self._delta_container.__enter__()
        st.markdown('<span class="oleasing-card-marker"></span>', unsafe_allow_html=True)
        if self.title:
            st.markdown(f"""
<div class="oleasing-card-header">
  <div class="oleasing-card-title">{self.title}</div>
  {f'<div class="oleasing-card-subtitle">{self.subtitle}</div>' if self.subtitle else ''}
</div>
""", unsafe_allow_html=True)
        return ctx

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._delta_container is not None:
            return self._delta_container.__exit__(exc_type, exc_val, exc_tb)
        return False


# ============================================================================
# CLASE CONTROL ESTANDAR DE ENTRADA (BaseInput)
# ============================================================================

class BaseInput(UIComponent):
    """
    Control estandar de entrada de datos de formulario.
    """

    def __init__(
        self,
        label: str,
        value: str = "",
        placeholder: Optional[str] = None,
        key: Optional[str] = None,
        help: Optional[str] = None,
        disabled: bool = False,
        type: str = "default",
        max_chars: Optional[int] = None
    ):
        self.label = label
        self.initial_value = value
        self.placeholder = placeholder or ""
        self.key = key or f"inp_{abs(hash((label, self.__class__.__name__)))}"
        self.help = help
        self.disabled = disabled
        self.type = type
        self.max_chars = max_chars
        self._current_value = value

    @property
    def value(self) -> str:
        """Obtiene el valor actual del campo desde session_state o valor local."""
        if hasattr(st, "session_state") and self.key in st.session_state:
            return str(st.session_state[self.key])
        return self._current_value

    def render(self) -> str:
        """Renderiza el campo de entrada y retorna el valor ingresado."""
        raise NotImplementedError("Las subclases deben implementar el metodo render().")

    def __call__(self) -> str:
        return self.render()

    def __str__(self) -> str:
        return self.value


# ============================================================================
# 4. FORM INPUT
# ============================================================================

class FormInput(BaseInput):
    """
    Campo de texto con border-radius de 8px, bordes gris suave (#CBD5E1 en Light,
    #334155 en Dark) y que al hacer 'focus' el borde cambie al color naranja
    de acento (#FF6B35 en Light, #FF8B60 en Dark) con resplandor sutil.
    """

    @classmethod
    def render_input(
        cls,
        label: str,
        value: str = "",
        placeholder: Optional[str] = None,
        key: Optional[str] = None,
        help: Optional[str] = None,
        disabled: bool = False,
        type: str = "default",
        max_chars: Optional[int] = None
    ) -> str:
        """Metodo de clase para renderizar el FormInput directamente."""
        instancia = cls(
            label=label,
            value=value,
            placeholder=placeholder,
            key=key,
            help=help,
            disabled=disabled,
            type=type,
            max_chars=max_chars
        )
        return instancia.render()

    create = render_input

    def render(self) -> str:
        inyectar_componentes_css()
        marker_id = f"mk_inp_{self.key}"
        st.markdown(f'<span class="oleasing-input-marker" id="{marker_id}"></span>', unsafe_allow_html=True)
        val = st.text_input(
            label=self.label,
            value=self.initial_value,
            placeholder=self.placeholder,
            key=self.key,
            help=self.help,
            disabled=self.disabled,
            type=self.type,
            max_chars=self.max_chars
        )
        self._current_value = val
        return val


# ============================================================================
# EXPORTACIONES DEL MODULO
# ============================================================================

__all__ = [
    "inyectar_componentes_css",
    "UIComponent",
    "BaseButton",
    "BaseContainer",
    "BaseInput",
    "PrimaryButton",
    "OutlinedButton",
    "CardContainer",
    "FormInput",
]
