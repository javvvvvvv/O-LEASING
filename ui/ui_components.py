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
ui/ui_components.py - Re-exporta los componentes de ui_components.py
para permitir importaciones tanto desde la raiz como desde el paquete ui.
"""

from ui_components import (
    inyectar_componentes_css,
    UIComponent,
    BaseButton,
    BaseContainer,
    BaseInput,
    PrimaryButton,
    OutlinedButton,
    CardContainer,
    FormInput,
)

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
