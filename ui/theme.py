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
ui/theme.py - Re-exporta el sistema unificado de temas desde theme.py.
Mantiene compatibilidad total con importaciones existentes.
"""

from theme import (
    Light,
    Dark,
    LIGHT_CONFIG,
    DARK_CONFIG,
    THEMES,
    THEME_CLASSES,
    normalizar_nombre_tema,
    inicializar_tema,
    get_current_theme_name,
    get_current_theme,
    get_theme_class,
    set_theme,
    toggle_theme,
    render_theme_toggle,
    render_theme_toggle_button,
    generar_css,
    inyectar_css,
)

__all__ = [
    "Light",
    "Dark",
    "LIGHT_CONFIG",
    "DARK_CONFIG",
    "THEMES",
    "THEME_CLASSES",
    "normalizar_nombre_tema",
    "inicializar_tema",
    "get_current_theme_name",
    "get_current_theme",
    "get_theme_class",
    "set_theme",
    "toggle_theme",
    "render_theme_toggle",
    "render_theme_toggle_button",
    "generar_css",
    "inyectar_css",
]
