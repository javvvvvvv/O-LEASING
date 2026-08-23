"""
test_cfdi.py — Pruebas de core/cfdi.py (parseo y clasificación de CFDI).

Antes esta lógica vivía dentro de app.py, mezclada con Streamlit, y no se
podía probar sin levantar la app completa. Ahora que es un módulo puro se
prueba directo, incluyendo los casos que antes solo se detectaban subiendo
un XML real a mano: archivo corrupto, demasiado grande, UUID mal formado,
RFC mal formado, y la clasificación de conceptos (que decide si una
factura se compara o no contra el contrato).

Para correrlas: pytest test_cfdi.py -v
"""
import pytest

from core.cfdi import (
    parse_cfdi, clasificar_concepto, normalizar_contrato,
    REGLAS_CONCEPTO_DEFAULT, MAX_XML_BYTES,
)


def _cfdi_xml(
    uuid="A1B2C3D4-E5F6-4A1B-9C2D-3E4F5A6B7C8D",
    rfc_emisor="ABC010101AB1",
    rfc_receptor="XYZ020202XY2",
    fecha="2026-03-15T10:00:00",
    folio="1001",
    subtotal="11800.00",
    total="13688.00",
    tipo_comprobante="I",
    moneda="MXN",
    conceptos=None,
):
    """Arma un CFDI 4.0 sintético mínimo, con la misma forma que el real
    (namespaces, TimbreFiscalDigital, Conceptos), para no depender de
    tener un XML de verdad en el repo."""
    if conceptos is None:
        conceptos = [
            ("Renta mensual contrato 0012-0001, mes 3/36", "10000.00"),
            ("Comisión por administración", "1800.00"),
        ]
    conceptos_xml = "".join(
        f'<cfdi:Concepto Descripcion="{desc}" Importe="{imp}" Descuento="0"/>'
        for desc, imp in conceptos
    )
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<cfdi:Comprobante xmlns:cfdi="http://www.sat.gob.mx/cfd/4"
    xmlns:tfd="http://www.sat.gob.mx/TimbreFiscalDigital"
    Fecha="{fecha}" Folio="{folio}" SubTotal="{subtotal}" Total="{total}"
    TipoDeComprobante="{tipo_comprobante}" Moneda="{moneda}">
  <cfdi:Emisor Rfc="{rfc_emisor}" Nombre="Arrendadora Ejemplo"/>
  <cfdi:Receptor Rfc="{rfc_receptor}" Nombre="Cliente Ejemplo"/>
  <cfdi:Conceptos>
    {conceptos_xml}
  </cfdi:Conceptos>
  <cfdi:Complemento>
    <tfd:TimbreFiscalDigital UUID="{uuid}"/>
  </cfdi:Complemento>
</cfdi:Comprobante>""".encode("utf-8")


REGLAS = {k: list(v) for k, v in REGLAS_CONCEPTO_DEFAULT.items()}


def test_parse_cfdi_extrae_los_datos_basicos():
    fact = parse_cfdi(_cfdi_xml(), REGLAS)
    assert fact['uuid'] == "A1B2C3D4-E5F6-4A1B-9C2D-3E4F5A6B7C8D"
    assert fact['folio'] == "1001"
    assert fact['subtotal'] == 11800.00
    assert fact['total'] == 13688.00
    assert fact['rfc_emisor'] == "ABC010101AB1"
    assert fact['rfc_receptor'] == "XYZ020202XY2"
    assert fact['periodo'] == "2026-03"
    assert fact['moneda'] == "MXN"


def test_parse_cfdi_no_regresa_el_xml_crudo():
    """El resultado ya no debe traer el XML original — O-Leasing dejó de
    guardarlo en la base de datos."""
    fact = parse_cfdi(_cfdi_xml(), REGLAS)
    assert 'xml_raw' not in fact


def test_parse_cfdi_detecta_contrato_y_mes_del_texto():
    fact = parse_cfdi(_cfdi_xml(), REGLAS)
    assert fact['id_contrato'] == "0012-0001"
    assert fact['mes_contrato'] == 3


def test_parse_cfdi_clasifica_y_suma_conceptos():
    fact = parse_cfdi(_cfdi_xml(), REGLAS)
    assert fact['conceptos']['RENTA'] == pytest.approx(10000.00)
    assert fact['conceptos']['COMISION'] == pytest.approx(1800.00)
    assert fact['tipo'] == 'ANTICIPO'  # trae un concepto de COMISION


def test_parse_cfdi_factura_solo_renta_es_mensual():
    xml = _cfdi_xml(
        subtotal="10000.00", total="11600.00",
        conceptos=[("Renta mensual contrato 0012-0002, mes 5/36", "10000.00")],
    )
    fact = parse_cfdi(xml, REGLAS)
    assert fact['tipo'] == 'MENSUAL'
    assert fact['mes_contrato'] == 5


def test_parse_cfdi_factura_sin_conceptos_de_leasing_es_otro():
    xml = _cfdi_xml(
        subtotal="500.00", total="580.00",
        conceptos=[("Servicio de gestoría vehicular", "500.00")],
    )
    fact = parse_cfdi(xml, REGLAS)
    assert fact['tipo'] == 'OTRO'


def test_parse_cfdi_rechaza_xml_corrupto():
    with pytest.raises(ValueError, match="XML inválido"):
        parse_cfdi(b"<esto no cierra", REGLAS)


def test_parse_cfdi_rechaza_archivo_demasiado_grande():
    xml_gigante = _cfdi_xml() + b" " * (MAX_XML_BYTES + 1)
    with pytest.raises(ValueError, match="pesa"):
        parse_cfdi(xml_gigante, REGLAS)


def test_parse_cfdi_rechaza_uuid_mal_formado():
    xml = _cfdi_xml(uuid="no-es-un-uuid-valido")
    with pytest.raises(ValueError, match="UUID"):
        parse_cfdi(xml, REGLAS)


def test_parse_cfdi_avisa_rfc_mal_formado_sin_reventar():
    xml = _cfdi_xml(rfc_emisor="RFC-INVALIDO")
    fact = parse_cfdi(xml, REGLAS)
    assert any("RFC emisor" in aviso for aviso in fact['avisos_xml'])


def test_parse_cfdi_avisa_si_la_suma_de_conceptos_no_cuadra_con_subtotal():
    xml = _cfdi_xml(subtotal="99999.00")  # no coincide con la suma real de conceptos
    fact = parse_cfdi(xml, REGLAS)
    assert any("SubTotal" in aviso for aviso in fact['avisos_xml'])


def test_parse_cfdi_quita_el_bom_utf8():
    xml_con_bom = b'\xef\xbb\xbf' + _cfdi_xml()
    fact = parse_cfdi(xml_con_bom, REGLAS)
    assert fact['uuid'] == "A1B2C3D4-E5F6-4A1B-9C2D-3E4F5A6B7C8D"


@pytest.mark.parametrize("descripcion,clave_esperada", [
    ("Renta mensual contrato 0001-0001", "RENTA"),
    ("Comisión por apertura", "COMISION"),
    ("ANTICIPO a capital", "ANTICIPO"),
    ("Servicio de GEOLOCALIZACIÓN mensual", "GEOLOC"),
    ("Gastos de administración y cobranza", "ADMIN"),
    ("Concepto que no coincide con ninguna regla", "OTRO"),
])
def test_clasificar_concepto(descripcion, clave_esperada):
    assert clasificar_concepto(descripcion, REGLAS) == clave_esperada


def test_clasificar_concepto_respeta_prioridad_cuando_hay_dos_coincidencias():
    """Si el texto trae tanto 'RENTA' como 'COMISION', debe ganar COMISION
    (primera en el orden de prioridad) — así una comisión no se cuela como
    si fuera renta normal."""
    assert clasificar_concepto("Renta y Comisión del mes 1", REGLAS) == "COMISION"


def test_normalizar_contrato_rellena_con_ceros():
    assert normalizar_contrato("12", "1") == "0012-0001"
    assert normalizar_contrato("0012", "0001") == "0012-0001"
