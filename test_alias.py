"""
test_alias.py — Pruebas de core/alias.py (reglas de "contrato facturado
con otro número").

Lo más importante que prueba este archivo es lo que el usuario pidió
explícitamente: que nunca se dupliquen reglas. `decidir_accion_alias` es
la única función que decide si hay algo que guardar, así que si esto está
bien probado, toda la app hereda esa garantía sin importar desde cuál de
las tres pantallas se haya resuelto la factura.

Para correrlas: pytest test_alias.py -v
"""
import pytest

from core.alias import resolver_numero_contrato, decidir_accion_alias, sugerir_contrato_similar


# --- resolver_numero_contrato ---------------------------------------------

def test_resolver_numero_contrato_aplica_regla_existente():
    alias_map = {"0012-0005": "0012-0007"}
    numero, aplicado = resolver_numero_contrato("0012-0005", alias_map)
    assert numero == "0012-0007"
    assert aplicado is True


def test_resolver_numero_contrato_sin_regla_no_cambia_nada():
    numero, aplicado = resolver_numero_contrato("0012-0001", {})
    assert numero == "0012-0001"
    assert aplicado is False


def test_resolver_numero_contrato_sin_numero_detectado():
    numero, aplicado = resolver_numero_contrato(None, {"0012-0005": "0012-0007"})
    assert numero is None
    assert aplicado is False


def test_resolver_numero_contrato_regla_que_apunta_a_si_misma_no_cuenta_como_aplicada():
    """Caso raro pero posible tras limpiar datos: si por algún motivo la
    regla apunta al mismo número, no debe marcarse como 'aplicada' — no
    hubo ninguna corrección real."""
    numero, aplicado = resolver_numero_contrato("0012-0001", {"0012-0001": "0012-0001"})
    assert numero == "0012-0001"
    assert aplicado is False


# --- decidir_accion_alias ---------------------------------------------------

def test_decidir_accion_alias_regla_nueva():
    assert decidir_accion_alias("0012-0005", "0012-0007", id_contrato_real_existente=None) == "CREADA"


def test_decidir_accion_alias_regla_corregida():
    assert decidir_accion_alias("0012-0005", "0012-0009", id_contrato_real_existente="0012-0007") == "CORREGIDA"


def test_decidir_accion_alias_regla_identica_no_duplica():
    """Si ya existe exactamente la misma regla, NO debe volver a crearla
    ni a tocar el historial — este es el caso central de 'que no duplique'."""
    assert decidir_accion_alias("0012-0005", "0012-0007", id_contrato_real_existente="0012-0007") is None


def test_decidir_accion_alias_sin_numero_detectado():
    assert decidir_accion_alias(None, "0012-0007") is None
    assert decidir_accion_alias("", "0012-0007") is None


def test_decidir_accion_alias_sin_contrato_real():
    assert decidir_accion_alias("0012-0005", None) is None


def test_decidir_accion_alias_numero_igual_al_contrato_real():
    """Si el número detectado YA es el contrato real, no hay nada que
    aprender (no era un alias, el número estaba bien desde un principio)."""
    assert decidir_accion_alias("0012-0007", "0012-0007", id_contrato_real_existente=None) is None


def test_decidir_accion_alias_llamadas_repetidas_con_la_misma_regla_no_acumulan():
    """Simula resolver la misma factura (o una igual) varias veces seguidas:
    la primera vez crea la regla, las siguientes no deben generar más
    entradas — protege contra doble clic o contra procesar el mismo lote
    dos veces."""
    id_existente = None
    acciones = []
    for _ in range(5):
        accion = decidir_accion_alias("0012-0005", "0012-0007", id_contrato_real_existente=id_existente)
        acciones.append(accion)
        if accion:
            id_existente = "0012-0007"  # así quedaría la BD real después de guardar
    assert acciones == ["CREADA", None, None, None, None]


# --- sugerir_contrato_similar ------------------------------------------------

def test_sugerir_contrato_similar_detecta_typo_de_un_digito():
    contratos = ["0012-0007", "0034-0001", "0099-0002"]
    sugerencias = sugerir_contrato_similar("0012-0009", contratos)
    assert "0012-0007" in sugerencias


def test_sugerir_contrato_similar_sin_parecido_no_sugiere_nada():
    contratos = ["0012-0007", "0034-0001", "0099-0002"]
    sugerencias = sugerir_contrato_similar("9999-9999", contratos)
    assert sugerencias == []


def test_sugerir_contrato_similar_sin_numero_detectado():
    assert sugerir_contrato_similar(None, ["0012-0007"]) == []
    assert sugerir_contrato_similar("", ["0012-0007"]) == []


def test_sugerir_contrato_similar_sin_contratos():
    assert sugerir_contrato_similar("0012-0007", []) == []


def test_sugerir_contrato_similar_respeta_el_maximo():
    contratos = ["0012-0001", "0012-0002", "0012-0003", "0012-0004"]
    sugerencias = sugerir_contrato_similar("0012-0000", contratos, max_sugerencias=2)
    assert len(sugerencias) <= 2
