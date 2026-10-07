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
# CFDI 3.3 usa otro namespace; se prueba si el 4.0 no encuentra nodos.
NS_V3 = {
    'cfdi': 'http://www.sat.gob.mx/cfd/3',
    'tfd':  'http://www.sat.gob.mx/TimbreFiscalDigital',
}

MAX_XML_BYTES = 10 * 1024 * 1024  # Aumentado a 10MB para CFDIs grandes o con Addendas

_PAT_CONTRATO = re.compile(
    r'(?:[A-Za-zÑñÁÉÍÓÚáéíóú\s]*?)(\d{1,4})\s*-\s*(\d{1,4})',
    re.IGNORECASE,
)
_PAT_MES = re.compile(r'(?:^|[^\d/])(\d{1,3})\s*/\s*(\d{1,3})(?![/\d])')
_PAT_UUID = re.compile(r'^[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}$')
_PAT_RFC  = re.compile(r'^[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}$')


def extraer_contratos_robusto(texto: str, contratos_validos: set | list | None = None) -> list[str]:
    """Extrae numeros de contrato normalizados (formato XXXX-YYYY) de un texto o concepto.
    Maneja conceptos con palabras pegadas (ej. RENTA0506-06, FINANCIADO0506-06),
    meses pegados al sufijo (ej. RENTA0590-0318/24 -> 0590-0003),
    codigos de 6 digitos sin guion (ej. RENTA068401 -> 0684-0001),
    contratos sin sufijo (ej. RENTACONTRATO 700 -> 0700-0001),
    y sufijos en cero (ej. RENTA0544-0 -> 0544-0001).
    Valida contra contratos_validos si se proporciona.
    """
    if not texto:
        return []

    t_u = str(texto).upper()
    detectados = []

    if contratos_validos is not None and not isinstance(contratos_validos, (set, dict)):
        contratos_validos = set(contratos_validos)

    # 1. Caso de sufijo pegado a numerador de mes: e.g. 0590-0318/24 -> 0590-03 y mes 18/24
    m_glued_mes = re.search(r'(\d{1,4})-(\d{1,2})(\d{1,2})\s*/\s*(\d{1,2})', t_u)
    if m_glued_mes:
        pref, suf, mes_n, plazo_n = m_glued_mes.groups()
        c_cand = normalizar_contrato(pref, suf)
        if not contratos_validos or c_cand in contratos_validos:
            if c_cand not in detectados:
                detectados.append(c_cand)

    # 2. Busqueda exhaustiva con guion y prefijo de texto pegado o separado
    for m in _PAT_CONTRATO.finditer(t_u):
        pref, suf = m.group(1), m.group(2)
        c_cand = normalizar_contrato(pref, suf)

        if any(c.startswith(f"{pref.zfill(4)}-") for c in detectados):
            continue

        if contratos_validos and c_cand not in contratos_validos and len(suf) >= 3:
            for cut in (2, 1):
                alt_cand = normalizar_contrato(pref, suf[:cut])
                if alt_cand in contratos_validos:
                    c_cand = alt_cand
                    break

        if contratos_validos and c_cand not in contratos_validos and suf == '0':
            posibles = [c for c in contratos_validos if c.startswith(pref.zfill(4) + '-')]
            if len(posibles) == 1:
                c_cand = posibles[0]
            elif f"{pref.zfill(4)}-0001" in contratos_validos:
                c_cand = f"{pref.zfill(4)}-0001"

        if c_cand not in detectados:
            if not contratos_validos or c_cand in contratos_validos:
                detectados.append(c_cand)

    # 3. Caso sin guion con 6 digitos: e.g. RENTA068401 -> 0684-01 -> 0684-0001
    for m in re.finditer(r'(?:RENTA|CONTRATO)?\s*(\d{4})(\d{2})\b', t_u):
        pref, suf = m.group(1), m.group(2)
        c_cand = normalizar_contrato(pref, suf)
        if contratos_validos and c_cand in contratos_validos:
            if c_cand not in detectados:
                detectados.append(c_cand)

    # 4. Caso CONTRATO XXX sin sufijo ni guion posterior: e.g. RENTACONTRATO 700 1/36
    for m in re.finditer(r'CONTRATO\s*(\d{1,4})(?!-)\b', t_u):
        pref = m.group(1).zfill(4)
        c_cand = f"{pref}-0001"
        if contratos_validos and c_cand in contratos_validos:
            if c_cand not in detectados:
                detectados.append(c_cand)
        elif contratos_validos:
            posibles = [c for c in contratos_validos if c.startswith(pref + '-')]
            if len(posibles) == 1:
                if posibles[0] not in detectados:
                    detectados.append(posibles[0])

    return detectados


# Reglas de clasificación de conceptos: qué palabras debe traer el texto del
# CFDI para que el sistema lo reconozca como Renta, Administración, etc.
# Los valores por default viven aquí; la versión configurable por el usuario
# se guarda en la BD y la resuelve app.py (esta capa no sabe de eso).
REGLAS_CONCEPTO_DEFAULT = {
    "VENTA_VEHICULO": ["VENTA DE VEHICULO", "VENTA VEHICULO", "VENTA DE UNIDAD", "VENTA UNIDAD", "ENAJENACION"],
    "INDEMNIZACION":  ["INDEMNIZACION", "INDEMNIZACIÓN", "RECUPERACION SEGURO", "REEMBOLSO SEGURO", "SINIESTRO"],
    "COMISION":       ["COMISION", "COMISIÓN"],
    "ANTICIPO":       ["ANTICIPO", "APORTACION", "APORTACIÓN", "MONTO FINANCIADO"],
    "GEOLOC":         ["GEOLOCAL", "GEOLOC"],
    "ADMIN":          ["ADMINISTR", "COBRANZA"],
    "RENTA":          ["RENTA"],
}
# Orden en que se revisan las claves — si un texto pudiera coincidir con más
# de una (p. ej. trae "RENTA" y "COMISION" a la vez), gana la primera de
# esta lista.
_ORDEN_CLAVES_CONCEPTO = ["VENTA_VEHICULO", "INDEMNIZACION", "COMISION", "ANTICIPO", "GEOLOC", "ADMIN", "RENTA"]


def normalizar_contrato(raw_num: str, raw_suf: str) -> str:
    return f"{raw_num.zfill(4)}-{raw_suf.zfill(4)}"


def normalizar_folio(folio: str | None) -> str:
    """Normaliza un folio de factura para comparación y descarte de duplicados.
    Quita prefijos de serie como 'F-', 'F_', o letras precedentes cuando hay números.
    Ejemplo: 'F-25236' -> '25236', '25236' -> '25236', 'F 25236' -> '25236'.
    """
    if folio is None:
        return ""
    s = str(folio).strip()
    if s.endswith('.0'):
        s = s[:-2]
    s_clean = re.sub(r'^[A-Za-z]+[\s\-_]*', '', s)
    digits = re.sub(r'\D', '', s_clean)
    if digits:
        return digits
    all_digits = re.sub(r'\D', '', s)
    if all_digits:
        return all_digits
    return s.strip().upper()


def clasificar_concepto(descripcion: str, reglas: dict) -> str:
    d = descripcion.upper()
    for clave in _ORDEN_CLAVES_CONCEPTO:
        for palabra in reglas.get(clave, []):
            if palabra and palabra.upper() in d:
                return clave
    return 'OTRO'


def parse_cfdi(xml_bytes: bytes, reglas: dict, contratos_validos: set | list | None = None) -> dict:
    """Extrae de un CFDI 3.3 / 4.0 únicamente los datos que O-Leasing necesita.
    Tolerante a cualquier codificación, BOM, espacios, ampersands sin escapar y variantes de namespace.
    """
    if len(xml_bytes) > MAX_XML_BYTES:
        raise ValueError(f"El archivo pesa {len(xml_bytes)/1024:.0f} KB — límite de 10 MB excedido.")
    
    if xml_bytes.startswith(b'\xef\xbb\xbf'):
        xml_bytes = xml_bytes[3:]
    
    # Decodificar texto limpio para evitar errores de encoding en ElementTree
    xml_str = ""
    for enc in ('utf-8', 'latin-1', 'cp1252'):
        try:
            xml_str = xml_bytes.decode(enc)
            break
        except Exception:
            continue
    if not xml_str:
        xml_str = xml_bytes.decode('utf-8', errors='replace')
        
    xml_str = xml_str.strip()

    # Sanitizar ampersands sin escapar que rompen ElementTree
    xml_str = re.sub(r'&(?!(amp|lt|gt|apos|quot|#\d+|#x[0-9a-fA-F]+);)', '&amp;', xml_str)

    try:
        # Remover declaración xml si trae encoding incompatible con fromstring
        xml_clean = re.sub(r'^\s*<\?xml[^>]*\?>', '', xml_str).strip()
        root = ET.fromstring(xml_clean.encode('utf-8'))
    except Exception as e:
        try:
            root = ET.fromstring(xml_bytes)
        except Exception as ex:
            raise ValueError(f"XML no válido o corrupto: {ex}")

    comp = root.attrib
    fecha    = comp.get('Fecha', '')
    folio    = normalizar_folio(comp.get('Folio', ''))
    subtotal = float(comp.get('SubTotal', 0) or 0)
    total    = float(comp.get('Total', 0) or 0)
    tipo_comprobante = comp.get('TipoDeComprobante', 'I')
    moneda            = comp.get('Moneda', 'MXN')

    # Búsqueda flexible de nodos sin depender exclusivamente de namespaces prefijados
    def _find_node(tag_name):
        for elem in root.iter():
            if elem.tag.endswith(tag_name):
                return elem
        return None

    emisor = _find_node('Emisor')
    receptor = _find_node('Receptor')
    rfc_emisor = emisor.get('Rfc') if emisor is not None else None
    rfc_receptor = receptor.get('Rfc') if receptor is not None else None

    avisos_xml = []
    if rfc_emisor and not _PAT_RFC.match(rfc_emisor.upper()):
        avisos_xml.append(f"El RFC emisor ('{rfc_emisor}') no tiene formato de RFC estándar.")
    if rfc_receptor and not _PAT_RFC.match(rfc_receptor.upper()):
        avisos_xml.append(f"El RFC receptor ('{rfc_receptor}') no tiene formato de RFC estándar.")

    timbre = _find_node('TimbreFiscalDigital')
    uuid = timbre.get('UUID') if timbre is not None else None
    
    # Fallback por regex si el namespace de TimbreFiscalDigital varió
    if not uuid:
        m_u = re.search(r'UUID=["\']([0-9A-Fa-f-]{36})["\']', xml_str, re.IGNORECASE)
        if m_u:
            uuid = m_u.group(1)
            
    if not uuid:
        import uuid as uuid_mod
        uuid = str(uuid_mod.uuid4())
        avisos_xml.append("No se encontró Timbre Fiscal Digital (UUID); se asignó un identificador único sintético.")

    conceptos_raw = []
    for c in root.iter():
        if c.tag.endswith('Concepto'):
            desc = c.get('Descripcion', '')
            importe = float(c.get('Importe', 0) or 0)
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

    # Deteccion exhaustiva de contratos (soporta facturas multi-contrato y conceptos pegados)
    contratos_detectados = []
    for c in conceptos_raw:
        for cid in extraer_contratos_robusto(c['desc'], contratos_validos):
            if cid not in contratos_detectados:
                contratos_detectados.append(cid)
    if not contratos_detectados:
        for cid in extraer_contratos_robusto(descripciones, contratos_validos):
            if cid not in contratos_detectados:
                contratos_detectados.append(cid)

    id_contrato = ', '.join(contratos_detectados) if len(contratos_detectados) > 1 else (contratos_detectados[0] if contratos_detectados else None)

    meses_detectados = []
    # Priorizar conceptos de RENTA para no confundir con seguro/otros plazos
    for c in conceptos_raw:
        desc_c = str(c.get('desc') or '')
        clv_c = str(c.get('clave') or '')
        if clv_c == 'RENTA' or 'RENTA' in desc_c.upper():
            if 'SEGURO' in desc_c.upper() and 'RENTA' not in desc_c.upper()[:10]:
                continue
            for m_str, pl_str in _PAT_MES.findall(desc_c):
                try:
                    m_val = int(m_str)
                    if m_val not in meses_detectados:
                        meses_detectados.append(m_val)
                except Exception:
                    pass
            m_glued = re.search(r'(\d{1,4})-(\d{1,2})(\d{1,2})\s*/\s*(\d{1,2})', desc_c.upper())
            if m_glued:
                try:
                    m_val = int(m_glued.group(3))
                    if m_val not in meses_detectados:
                        meses_detectados.append(m_val)
                except Exception:
                    pass

    # Si no se detecto mes en conceptos de RENTA, buscar en otros conceptos que no sean seguro/gestoria
    if not meses_detectados:
        for c in conceptos_raw:
            desc_c = str(c.get('desc') or '')
            clv_c = str(c.get('clave') or '')
            if clv_c in ('OTRO', 'SEGURO', 'GESTORIA') and 'RENTA' not in desc_c.upper():
                continue
            for m_str, pl_str in _PAT_MES.findall(desc_c):
                try:
                    m_val = int(m_str)
                    if m_val not in meses_detectados:
                        meses_detectados.append(m_val)
                except Exception:
                    pass

    if not meses_detectados:
        for m_str, pl_str in _PAT_MES.findall(descripciones):
            try:
                m_val = int(m_str)
                if m_val not in meses_detectados:
                    meses_detectados.append(m_val)
            except Exception:
                pass
    meses_detectados = sorted(meses_detectados)
    mes_contrato = ', '.join(str(m) for m in meses_detectados) if len(meses_detectados) > 1 else (meses_detectados[0] if meses_detectados else None)

    es_venta = any(c['clave'] == 'VENTA_VEHICULO' for c in conceptos_raw)
    es_indemn = any(c['clave'] == 'INDEMNIZACION' for c in conceptos_raw)

    if es_venta:
        tipo = 'VENTA_VEHICULO'
    elif es_indemn:
        tipo = 'INDEMNIZACION'
    else:
        conceptos_claves = ['RENTA', 'ADMIN', 'GEOLOC', 'ANTICIPO', 'COMISION']
        tiene_concepto_lease = any(c['clave'] in conceptos_claves for c in conceptos_raw)
        if not tiene_concepto_lease:
            tipo = 'OTRO'
        else:
            es_anticipo = any(c['clave'] in ('ANTICIPO', 'COMISION') for c in conceptos_raw)
            tipo = 'ANTICIPO' if es_anticipo else 'MENSUAL'

    if tipo in ('ANTICIPO', 'VENTA_VEHICULO', 'INDEMNIZACION') and mes_contrato is None:
        mes_contrato = 1

    totales = {'RENTA': 0.0, 'ADMIN': 0.0, 'GEOLOC': 0.0,
               'ANTICIPO': 0.0, 'COMISION': 0.0, 'VENTA_VEHICULO': 0.0, 'INDEMNIZACION': 0.0, 'OTRO': 0.0}
    for c in conceptos_raw:
        if c['clave'] in totales:
            totales[c['clave']] += c['importe']
        else:
            totales['OTRO'] += c['importe']

    periodo = fecha[:7] if fecha else ''

    return {
        'uuid':             uuid,
        'fecha':            fecha,
        'folio':            folio,
        'subtotal':         subtotal,
        'total':            total,
        'id_contrato':          id_contrato,
        'contratos_detectados': contratos_detectados,
        'id_contrato_detectado': ', '.join(contratos_detectados) if contratos_detectados else None,
        'mes_contrato':     mes_contrato,
        'meses_detectados': meses_detectados,
        'tipo':             tipo,
        'periodo':          periodo,
        'conceptos':        totales,
        'conceptos_raw':    conceptos_raw,
        'rfc_emisor':       rfc_emisor,
        'rfc_receptor':     rfc_receptor,
        'tipo_comprobante': tipo_comprobante,
        'moneda':           moneda,
        'cancelada':        1 if ('cancelad' in root.tag.lower() or 'acuse' in root.tag.lower() or root.attrib.get('Estatus') == 'Cancelado') else 0,
        'fecha_cancelacion': root.attrib.get('FechaCancelacion') or (fecha if ('cancelad' in root.tag.lower() or 'acuse' in root.tag.lower() or root.attrib.get('Estatus') == 'Cancelado') else None),
        'status':           'CANCELADA' if ('cancelad' in root.tag.lower() or 'acuse' in root.tag.lower() or root.attrib.get('Estatus') == 'Cancelado') else 'PENDIENTE',
        'avisos_xml':       avisos_xml,
    }


def parse_sat_excel(file_or_bytes, reglas=None, rfc_emisor='MAR031024EZ3', contratos_validos: set | list | None = None) -> list:
    """Extrae facturas de archivos Excel (.xlsx) exportados del repositorio del SAT.
    Detecta automáticamente encabezados en fila 4 o similar, normaliza columnas y extrae
    fechas, folios, importes, contratos, meses y conceptos clasificados.
    """
    import os
    import io
    import uuid as _uuid_mod
    import unicodedata
    import pandas as pd

    if reglas is None:
        reglas = REGLAS_CONCEPTO_DEFAULT

    def _norm(s):
        if not s:
            return ''
        return unicodedata.normalize('NFKD', str(s)).encode('ASCII', 'ignore').decode('utf-8').strip().lower()

    # Determinar si es ruta o bytes
    if isinstance(file_or_bytes, (str, os.PathLike)):
        df_raw = pd.read_excel(file_or_bytes, header=None, nrows=15)
    else:
        raw_io = file_or_bytes if hasattr(file_or_bytes, 'read') else io.BytesIO(file_or_bytes)
        df_raw = pd.read_excel(raw_io, header=None, nrows=15)

    header_idx = None
    for idx, row in df_raw.iterrows():
        row_norms = [_norm(v) for v in row.dropna().values]
        if any('emision' in v or 'subtotal' in v for v in row_norms):
            header_idx = idx
            break

    if header_idx is None:
        header_idx = 4

    if isinstance(file_or_bytes, (str, os.PathLike)):
        df = pd.read_excel(file_or_bytes, header=header_idx)
    else:
        if hasattr(file_or_bytes, 'seek'):
            file_or_bytes.seek(0)
            df = pd.read_excel(file_or_bytes, header=header_idx)
        else:
            df = pd.read_excel(io.BytesIO(file_or_bytes), header=header_idx)

    col_map = {_norm(c): c for c in df.columns}
    c_emision = col_map.get('emision') or col_map.get('fecha')
    c_serie = col_map.get('serie')
    c_folio = col_map.get('folio')
    c_subtotal = col_map.get('subtotal')
    c_total = col_map.get('total')
    c_rfc_rec = col_map.get('receptor rfc')
    c_desc = col_map.get('conceptos descripcion') or col_map.get('descripcion') or col_map.get('conceptos')
    c_estatus = (
        col_map.get('estatus')
        or col_map.get('estado')
        or col_map.get('estado del comprobante')
        or col_map.get('situacion')
        or col_map.get('status')
    )
    c_estatus_cancel = (
        col_map.get('estatus cancelacion')
        or col_map.get('estatus de cancelacion')
        or col_map.get('cancelacion')
    )
    c_fecha_cancel = (
        col_map.get('fecha cancelacion')
        or col_map.get('fecha de cancelacion')
        or col_map.get('fecha cancel')
    )
    c_uuid = col_map.get('uuid') or col_map.get('folio fiscal')
    c_tipo = col_map.get('tipo')

    facturas = []
    for _, r in df.iterrows():
        emision_val = r.get(c_emision) if c_emision else None
        if pd.isna(emision_val) or not emision_val:
            continue

        if hasattr(emision_val, 'strftime'):
            fecha = emision_val.strftime('%Y-%m-%d')
        else:
            fecha = str(emision_val)[:10]

        periodo = fecha[:7] if len(fecha) >= 7 else ''

        serie = str(r.get(c_serie, '') or '').strip() if c_serie and pd.notna(r.get(c_serie)) else ''
        folio_raw = r.get(c_folio, '') if c_folio else ''
        folio = normalizar_folio(folio_raw)

        subtotal = float(r.get(c_subtotal, 0) or 0) if c_subtotal and pd.notna(r.get(c_subtotal)) else 0.0
        total = float(r.get(c_total, 0) or 0) if c_total and pd.notna(r.get(c_total)) else 0.0
        rfc_rec = str(r.get(c_rfc_rec, '') or '').strip().upper() if c_rfc_rec and pd.notna(r.get(c_rfc_rec)) else None

        uid = str(r.get(c_uuid, '') or '').strip() if c_uuid and pd.notna(r.get(c_uuid)) else ''
        if not uid or len(uid) < 10:
            uid = str(_uuid_mod.uuid5(_uuid_mod.NAMESPACE_DNS, f"{rfc_emisor}-{serie}-{folio}-{fecha}-{total:.2f}"))

        desc_text = str(r.get(c_desc, '') or '').strip() if c_desc and pd.notna(r.get(c_desc)) else ''

        # Detectar contrato (soporta conceptos pegados y multi-contrato)
        contratos_detectados = extraer_contratos_robusto(desc_text, contratos_validos)
        id_contrato = ', '.join(contratos_detectados) if len(contratos_detectados) > 1 else (contratos_detectados[0] if contratos_detectados else None)

        # Parsear conceptos
        partes = [p.strip() for p in desc_text.split('|') if p.strip()]
        meses_detectados = []
        # Priorizar partes de RENTA
        for p in partes:
            clv_temp = clasificar_concepto(p, reglas)
            if clv_temp == 'RENTA' or 'RENTA' in p.upper():
                if 'SEGURO' in p.upper() and 'RENTA' not in p.upper()[:10]:
                    continue
                for m_str, pl_str in _PAT_MES.findall(p):
                    try:
                        m_val = int(m_str)
                        if m_val not in meses_detectados:
                            meses_detectados.append(m_val)
                    except Exception:
                        pass
                m_glued = re.search(r'(\d{1,4})-(\d{1,2})(\d{1,2})\s*/\s*(\d{1,2})', p.upper())
                if m_glued:
                    try:
                        m_val = int(m_glued.group(3))
                        if m_val not in meses_detectados:
                            meses_detectados.append(m_val)
                    except Exception:
                        pass

        if not meses_detectados:
            for p in partes:
                clv_temp = clasificar_concepto(p, reglas)
                if clv_temp in ('OTRO', 'SEGURO', 'GESTORIA') and 'RENTA' not in p.upper():
                    continue
                for m_str, pl_str in _PAT_MES.findall(p):
                    try:
                        m_val = int(m_str)
                        if m_val not in meses_detectados:
                            meses_detectados.append(m_val)
                    except Exception:
                        pass

        if not meses_detectados:
            for m_str, pl_str in _PAT_MES.findall(desc_text):
                try:
                    m_val = int(m_str)
                    if m_val not in meses_detectados:
                        meses_detectados.append(m_val)
                except Exception:
                    pass
        meses_detectados = sorted(meses_detectados)
        mes_contrato = ', '.join(str(m) for m in meses_detectados) if len(meses_detectados) > 1 else (meses_detectados[0] if meses_detectados else None)

        conceptos_raw = []
        if not partes:
            conceptos_raw.append({'desc': 'RENTA', 'importe': subtotal, 'clave': 'RENTA'})
        elif len(partes) == 1:
            clv = clasificar_concepto(partes[0], reglas)
            conceptos_raw.append({'desc': partes[0], 'importe': subtotal, 'clave': clv})
        else:
            imp_part = round(subtotal / len(partes), 2)
            for p in partes:
                conceptos_raw.append({'desc': p, 'importe': imp_part, 'clave': clasificar_concepto(p, reglas)})

        totales = {
            'RENTA': 0.0, 'ADMIN': 0.0, 'GEOLOC': 0.0, 'ANTICIPO': 0.0,
            'COMISION': 0.0, 'VENTA_VEHICULO': 0.0, 'INDEMNIZACION': 0.0, 'OTRO': 0.0
        }
        for c in conceptos_raw:
            totales[c['clave'] if c['clave'] in totales else 'OTRO'] += c['importe']

        tipo_raw = str(r.get(c_tipo, 'Ingreso') or '').strip().upper() if c_tipo else 'INGRESO'
        tipo_comprobante = 'E' if 'EGRESO' in tipo_raw else ('P' if 'PAGO' in tipo_raw else 'I')

        es_venta = any(c['clave'] == 'VENTA_VEHICULO' for c in conceptos_raw)
        es_indem = any(c['clave'] == 'INDEMNIZACION' for c in conceptos_raw)
        if es_venta:
            tipo = 'VENTA_VEHICULO'
        elif es_indem:
            tipo = 'INDEMNIZACION'
        else:
            tiene_lease = any(c['clave'] in ['RENTA', 'ADMIN', 'GEOLOC', 'ANTICIPO', 'COMISION'] for c in conceptos_raw)
            if not tiene_lease:
                tipo = 'OTRO'
            else:
                es_ant = any(c['clave'] in ('ANTICIPO', 'COMISION') for c in conceptos_raw)
                tipo = 'ANTICIPO' if es_ant else 'MENSUAL'

        estatus_str = str(r.get(c_estatus, '') or '').strip().upper() if c_estatus else ''
        estatus_cancel_str = str(r.get(c_estatus_cancel, '') or '').strip().upper() if c_estatus_cancel else ''
        fec_cancel_raw = r.get(c_fecha_cancel) if c_fecha_cancel else None

        cancelada = 0
        fecha_cancelacion = None

        if 'CANCELAD' in estatus_str:
            cancelada = 1
        elif 'CANCELAD' in estatus_cancel_str:
            cancelada = 1
        elif 'PLAZO VENCIDO' in estatus_cancel_str:
            cancelada = 1
        elif pd.notna(fec_cancel_raw) and str(fec_cancel_raw).strip() not in ('', 'NaT', 'None'):
            cancelada = 1

        if cancelada and pd.notna(fec_cancel_raw) and str(fec_cancel_raw).strip() not in ('', 'NaT', 'None'):
            if hasattr(fec_cancel_raw, 'strftime'):
                fecha_cancelacion = fec_cancel_raw.strftime('%Y-%m-%d')
            else:
                fecha_cancelacion = str(fec_cancel_raw)[:10]

        folio_completo = folio if folio else (serie or '')

        facturas.append({
            'uuid': uid,
            'fecha': fecha,
            'folio': folio_completo,
            'subtotal': subtotal,
            'total': total,
            'id_contrato': id_contrato,
            'contratos_detectados': contratos_detectados,
            'id_contrato_detectado': ', '.join(contratos_detectados) if contratos_detectados else None,
            'mes_contrato': mes_contrato,
            'meses_detectados': meses_detectados,
            'tipo': tipo,
            'periodo': periodo,
            'conceptos': totales,
            'conceptos_raw': conceptos_raw,
            'rfc_emisor': rfc_emisor,
            'rfc_receptor': rfc_rec,
            'tipo_comprobante': tipo_comprobante,
            'moneda': 'MXN',
            'cancelada': cancelada,
            'fecha_cancelacion': fecha_cancelacion,
            'status': 'CANCELADA' if cancelada else 'PENDIENTE',
            'avisos_xml': []
        })

    return facturas
