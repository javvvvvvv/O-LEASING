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
core/cfdi.py — Parseo y clasificación de CFDI (facturas XML) de O-Leasing.

Igual que finanzas.py: aquí NO se importa Streamlit ni se toca la base de
datos, solo lógica pura, para poder probarla con pytest (test_cfdi.py) sin
levantar la app y para poder reutilizar `reglas` una sola vez por lote en
vez de reconsultar la configuración archivo por archivo.

parse_cfdi() ya NO regresa el XML crudo: O-Leasing dejó de guardar el
contenido completo del CFDI en la base de datos (pesaba mucho sin
necesidad) y ahora solo persiste lo que de verdad se usa después —
folio, montos, RFCs, UUID y el detalle de conceptos ya clasificados,
en las tablas `facturas` y `factura_conceptos`.

Autor: Javier Illán
"""
import re
import xml.etree.ElementTree as ET

NS = {
    'cfdi': 'http://www.sat.gob.mx/cfd/4',
    'tfd':  'http://www.sat.gob.mx/TimbreFiscalDigital',
}

MAX_XML_BYTES = 2 * 1024 * 1024  # un CFDI real pesa unos cuantos KB; 2MB ya es sospechoso

_PAT_CONTRATO = re.compile(
    r'(?:CONTRATO\s*|RENTA\s*CONTRATO\s*|RENTA\s*)?(\d{1,4})-(\d{1,4})',
    re.IGNORECASE,
)
_PAT_MES  = re.compile(r'\b(\d{1,3})/(\d{1,3})\b')
_PAT_UUID = re.compile(r'^[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}$')
_PAT_RFC  = re.compile(r'^[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}$')

# Reglas de clasificación de conceptos: qué palabras debe traer el texto del
# CFDI para que el sistema lo reconozca como Renta, Administración, etc.
# Los valores por default viven aquí; la versión configurable por el usuario
# se guarda en la BD y la resuelve app.py (esta capa no sabe de eso).
REGLAS_CONCEPTO_DEFAULT = {
    "COMISION": ["COMISION", "COMISIÓN"],
    "ANTICIPO": ["ANTICIPO", "APORTACION", "APORTACIÓN", "MONTO FINANCIADO"],
    "GEOLOC":   ["GEOLOCAL", "GEOLOC"],
    "ADMIN":    ["ADMINISTR", "COBRANZA"],
    "RENTA":    ["RENTA"],
}
# Orden en que se revisan las claves — si un texto pudiera coincidir con más
# de una (p. ej. trae "RENTA" y "COMISION" a la vez), gana la primera de
# esta lista.
_ORDEN_CLAVES_CONCEPTO = ["COMISION", "ANTICIPO", "GEOLOC", "ADMIN", "RENTA"]


def normalizar_contrato(raw_num: str, raw_suf: str) -> str:
    return f"{raw_num.zfill(4)}-{raw_suf.zfill(4)}"


def clasificar_concepto(descripcion: str, reglas: dict) -> str:
    """Clasifica el texto de un concepto según las reglas vigentes.

    `reglas` es obligatorio: se resuelve una sola vez por lote (en app.py,
    antes del ciclo que procesa los XMLs) en vez de reconsultar la
    configuración de la empresa activa en cada archivo — con lotes de
    cientos de XMLs, esa sola consulta repetida era buena parte de la
    lentitud al cargar.
    """
    d = descripcion.upper()
    for clave in _ORDEN_CLAVES_CONCEPTO:
        for palabra in reglas.get(clave, []):
            if palabra and palabra.upper() in d:
                return clave
    return 'OTRO'


def parse_cfdi(xml_bytes: bytes, reglas: dict) -> dict:
    """Extrae de un CFDI 4.0 únicamente los datos que O-Leasing necesita
    para conciliar: identificación, montos, RFCs y conceptos clasificados.
    No conserva el XML original — eso ya no se guarda en ningún lado."""
    if len(xml_bytes) > MAX_XML_BYTES:
        raise ValueError(f"El archivo pesa {len(xml_bytes)/1024:.0f} KB — un CFDI normal pesa unos cuantos KB. "
                          f"Parece que no es un CFDI válido (límite: {MAX_XML_BYTES//1024//1024} MB).")
    if xml_bytes.startswith(b'\xef\xbb\xbf'):
        xml_bytes = xml_bytes[3:]
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError as e:
        raise ValueError(f"XML inválido: {e}")

    comp = root.attrib
    fecha    = comp.get('Fecha', '')
    folio    = comp.get('Folio', '')
    subtotal = float(comp.get('SubTotal', 0))
    total    = float(comp.get('Total', 0))
    tipo_comprobante = comp.get('TipoDeComprobante', 'I')
    moneda            = comp.get('Moneda', 'MXN')

    emisor   = root.find('.//cfdi:Emisor', NS)
    receptor = root.find('.//cfdi:Receptor', NS)
    rfc_emisor   = emisor.get('Rfc') if emisor is not None else None
    rfc_receptor = receptor.get('Rfc') if receptor is not None else None

    avisos_xml = []
    if rfc_emisor and not _PAT_RFC.match(rfc_emisor.upper()):
        avisos_xml.append(f"El RFC emisor ('{rfc_emisor}') no tiene un formato de RFC válido — revisa que el XML no esté corrupto.")
    if rfc_receptor and not _PAT_RFC.match(rfc_receptor.upper()):
        avisos_xml.append(f"El RFC receptor ('{rfc_receptor}') no tiene un formato de RFC válido — revisa que el XML no esté corrupto.")

    timbre = root.find('.//tfd:TimbreFiscalDigital', NS)
    uuid   = timbre.get('UUID') if timbre is not None else None
    if uuid and not _PAT_UUID.match(uuid):
        raise ValueError(f"El UUID del timbre fiscal ('{uuid}') no tiene el formato esperado — revisa que sea un CFDI timbrado y no un borrador.")

    conceptos_raw = []
    for c in root.findall('.//cfdi:Concepto', NS):
        desc      = c.get('Descripcion', '')
        importe   = float(c.get('Importe', 0))
        descuento = float(c.get('Descuento', 0) or 0)
        importe_neto = round(importe - descuento, 2)
        clave = clasificar_concepto(desc, reglas)
        conceptos_raw.append({'desc': desc, 'importe': importe_neto, 'clave': clave})

    suma_conceptos = round(sum(c['importe'] for c in conceptos_raw), 2)
    if abs(suma_conceptos - subtotal) > 1:
        avisos_xml.append(
            f"La suma de los conceptos (${suma_conceptos:,.2f}) no coincide con el SubTotal declarado en el "
            f"XML (${subtotal:,.2f}) — el CFDI podría estar mal formado."
        )
    descripciones = ' '.join(c['desc'] for c in conceptos_raw)

    id_contrato = None
    for c in conceptos_raw:
        m = _PAT_CONTRATO.search(c['desc'])
        if m:
            id_contrato = normalizar_contrato(m.group(1), m.group(2))
            break
    if not id_contrato:
        m = _PAT_CONTRATO.search(descripciones)
        if m:
            id_contrato = normalizar_contrato(m.group(1), m.group(2))

    mes_contrato = None
    m_per = _PAT_MES.search(descripciones)
    if m_per:
        mes_contrato = int(m_per.group(1))

    conceptos_claves = ['RENTA', 'ADMIN', 'GEOLOC', 'ANTICIPO', 'COMISION']
    tiene_concepto_lease = any(c['clave'] in conceptos_claves for c in conceptos_raw)
    if not tiene_concepto_lease:
        tipo = 'OTRO'
    else:
        es_anticipo = any(c['clave'] in ('ANTICIPO', 'COMISION') for c in conceptos_raw)
        tipo = 'ANTICIPO' if es_anticipo else 'MENSUAL'

    if tipo == 'ANTICIPO' and mes_contrato is None:
        mes_contrato = 1

    totales = {'RENTA': 0.0, 'ADMIN': 0.0, 'GEOLOC': 0.0,
               'ANTICIPO': 0.0, 'COMISION': 0.0, 'OTRO': 0.0}
    for c in conceptos_raw:
        totales[c['clave']] += c['importe']

    periodo = fecha[:7] if fecha else ''

    return {
        'uuid':             uuid,
        'fecha':            fecha,
        'folio':            folio,
        'subtotal':         subtotal,
        'total':            total,
        'id_contrato':      id_contrato,
        'mes_contrato':     mes_contrato,
        'tipo':             tipo,
        'periodo':          periodo,
        'conceptos':        totales,
        'conceptos_raw':    conceptos_raw,
        'rfc_emisor':       rfc_emisor,
        'rfc_receptor':     rfc_receptor,
        'tipo_comprobante': tipo_comprobante,
        'moneda':           moneda,
        'avisos_xml':       avisos_xml,
    }
