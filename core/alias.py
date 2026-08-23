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
core/alias.py — Reglas de "contrato facturado con otro número".

Lógica pura (sin Streamlit ni base de datos) para decidir:
  1. Si una factura recién leída debe reasignarse a otro contrato según las
     reglas ya aprendidas (`resolver_numero_contrato`).
  2. Qué acción tomar al guardar una regla nueva o corregida, garantizando
     que nunca se crea una regla duplicada ni se pisa una igual a la que ya
     existía (`decidir_accion_alias`).

app.py (models/db.py en realidad) es quien persiste esto en la tabla
`contrato_alias`; aquí solo vive la decisión, para poder probarla con
pytest sin depender de SQLite ni de una sesión de Streamlit.
"""


def resolver_numero_contrato(id_contrato_detectado, alias_map: dict):
    """Si `id_contrato_detectado` tiene una regla aprendida, regresa el
    contrato real al que debe asignarse la factura. Si no, regresa el mismo
    número sin tocar.

    Regresa (numero_final, se_aplico_alias: bool).
    """
    if not id_contrato_detectado:
        return id_contrato_detectado, False
    real = alias_map.get(id_contrato_detectado)
    if real and real != id_contrato_detectado:
        return real, True
    return id_contrato_detectado, False


def decidir_accion_alias(numero_facturado, id_contrato_real, id_contrato_real_existente=None):
    """Decide qué hacer al registrar numero_facturado -> id_contrato_real:

    - None            → no hay nada que guardar (no aplica o ya es idéntica
                         a la regla existente; así nunca se duplica).
    - 'CREADA'         → no existía ninguna regla para este número.
    - 'CORREGIDA'      → ya existía pero apuntaba a otro contrato.

    `id_contrato_real_existente` es el valor actual en la BD para ese
    `numero_facturado`, o None si no hay ninguna regla todavía.
    """
    if not numero_facturado or not id_contrato_real or numero_facturado == id_contrato_real:
        return None
    if id_contrato_real_existente is None:
        return 'CREADA'
    if id_contrato_real_existente == id_contrato_real:
        return None
    return 'CORREGIDA'


def sugerir_contrato_similar(numero_detectado, contratos_existentes, max_sugerencias=3):
    """Cuando un número de contrato no coincide exacto con ninguno real ni
    tiene una regla de alias aprendida, sugiere los contratos existentes
    cuyo número se parece más (typo de un dígito, dígitos invertidos, un
    guion en otro lugar, etc.), para no tener que buscarlo a mano en toda
    la lista. Es solo una sugerencia — nunca asigna nada por sí sola.

    Se compara texto contra texto con difflib (de la librería estándar de
    Python, sin dependencias nuevas). No usa la base de datos: recibe la
    lista de contratos ya cargada."""
    import difflib
    if not numero_detectado or not contratos_existentes:
        return []
    return difflib.get_close_matches(
        numero_detectado, list(contratos_existentes), n=max_sugerencias, cutoff=0.6
    )

