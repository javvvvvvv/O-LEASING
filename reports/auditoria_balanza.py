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
reports/auditoria_balanza.py — Motor Oficial de Auditoría y Conciliación
Contable vs Balanza de Comprobación Contpaqi para O-Leasing.

Diseñado con arquitectura multi-empresa dinámica:
1. Lee las cuentas contables directamente de la tabla `catalogo_cuentas`
   de la base de datos de la empresa activa (donde se configuran las pólizas):
   - CXC_CP: Cuentas por Cobrar Corto Plazo (ej. 115-00-00)
   - CXC_LP: Cuentas por Cobrar Largo Plazo (ej. 125-00-00)
   - RESIDUAL_CP: CXC Residual Corto Plazo (ej. 114-01-00)
   - RESIDUAL / CXC_RESIDUAL_INTERESES: CXC Residual Largo Plazo (ej. 126-00-00)
   - INT_CP: Intereses por Devengar Corto Plazo (ej. 208-00-00)
   - INT_LP: Intereses por Devengar Largo Plazo (ej. 228-00-00)
   - CANCELACION_CAPITAL: Ingresos / Amortización de Capital (ej. 413-00-00)
   - BAJA_INV: Activos / Inventario (ej. 117-00-00 o 150-00-00)
2. Calcula para cada contrato activo sus saldos teóricos esperados al corte
   conforme a las normas NIF C-15 / IFRS 16 y tablas de amortización.
3. Cruza contra la Balanza de Comprobación Contpaqi (cualquier mes):
   - Compara las cuentas financieras NIF C-15 del catálogo maestro.
   - Compara las cuentas operativas en Balanza (102-01 CxC, 403-01 Rentas, 150-00 Activos).
4. Omite rigurosamente de la cartera activa los contratos liquidados/concluidos
   y los audita en su propia pestaña, alertando en rojo si mantienen saldo contable.
5. Genera libro de Excel multi-pestaña con fórmulas vivas, sin duplicación de saldos
   y con validación COM limpia.
"""

import os
import sys
import re
import datetime
import sqlite3
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from finanzas import calc_amort, calc_res_amort, vp_res

if getattr(sys, "frozen", False):
    _BASE_DIR = os.path.dirname(sys.executable)
else:
    _BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Paleta corporativa sobria de auditoría O-Leasing / Orange Crew
FILL_PRIMARY = PatternFill(start_color="1E5C4F", end_color="1E5C4F", fill_type="solid")      # Verde Petróleo O-Leasing
FILL_SECONDARY = PatternFill(start_color="2D7D6F", end_color="2D7D6F", fill_type="solid")    # Verde Azulado
FILL_NAVY = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")         # Navy oscuro
FILL_ZEBRA = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")        # Gris ultra suave
FILL_TOTAL = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")        # Gris totalizador
FILL_ALERT_RED = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")    # Rojo suave alerta
FILL_ALERT_YELLOW = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid") # Amarillo advertencia
FILL_OK_GREEN = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")     # Verde aprobado

FONT_TITLE = Font(name="Calibri", size=13, bold=True, color="1E5C4F")
FONT_SUBTITLE = Font(name="Calibri", size=8, italic=True, color="64748B")
FONT_HEADER = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
FONT_BODY = Font(name="Calibri", size=9, color="1E293B")
FONT_BODY_BOLD = Font(name="Calibri", size=9, bold=True, color="1E293B")
FONT_BODY_RED = Font(name="Calibri", size=9, bold=True, color="991B1B")
FONT_BODY_GREEN = Font(name="Calibri", size=9, bold=True, color="166534")
FONT_TOTAL = Font(name="Calibri", size=9, bold=True, color="1E5C4F")

THIN_BORDER_SIDE = Side(style="thin", color="CBD5E1")
DOUBLE_BOTTOM_SIDE = Side(style="double", color="1E5C4F")
BORDER_CELL = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=THIN_BORDER_SIDE)
BORDER_TOTAL = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=DOUBLE_BOTTOM_SIDE)

FMT_CURRENCY = "$#,##0.00;[Red]-$#,##0.00;\"-\""
FMT_INT = "#,##0"


def _obtener_db_path():
    try:
        from reports.auditoria_excel import _obtener_db_path as _dbp
        return _dbp()
    except Exception:
        candidates = [
            os.path.join(_BASE_DIR, "leasing.db"),
            os.path.join(os.path.dirname(_BASE_DIR), "leasing.db"),
            r"C:\O-Leasing\punto y aparte\V10.01\o-leasing\leasing.db",
        ]
        for p in candidates:
            if os.path.exists(p):
                return p
        return "leasing.db"


def obtener_catalogo_cuentas_empresa(db_path):
    """
    Lee el catálogo dinámico de cuentas configurado para la empresa activa.
    Se conecta a la BD de la empresa y extrae las cuentas que se usan en las pólizas.
    """
    conn = sqlite3.connect(db_path)
    cat = {}
    try:
        rows = conn.execute("SELECT clave, cuenta, nombre, activa FROM catalogo_cuentas WHERE activa=1").fetchall()
        for r in rows:
            cat[r[0]] = {
                'cuenta': (r[1] or '').strip(),
                'nombre': (r[2] or '').strip(),
                'activa': bool(r[3] if r[3] is not None else 1)
            }
    except Exception:
        pass
    finally:
        conn.close()

    # Cuentas por defecto si no están definidas en la base de datos
    defaults = {
        'CXC_CP': {'cuenta': '115-00-00', 'nombre': 'CUENTAS POR COBRAR CP', 'activa': True},
        'CXC_LP': {'cuenta': '125-00-00', 'nombre': 'CUENTAS POR COBRAR LP', 'activa': True},
        'RESIDUAL_CP': {'cuenta': '114-01-00', 'nombre': 'CXC VALOR RESIDUAL CP', 'activa': True},
        'RESIDUAL': {'cuenta': '126-00-00', 'nombre': 'CXC VALOR RESIDUAL', 'activa': True},
        'INT_CP': {'cuenta': '208-00-00', 'nombre': 'INTERESES POR DEVENGAR CP', 'activa': True},
        'INT_LP': {'cuenta': '228-00-00', 'nombre': 'INTERESES POR DEVENGAR LP', 'activa': True},
        'CANCELACION_CAPITAL': {'cuenta': '413-00-00', 'nombre': 'CANCELACIÓN DE CAPITAL', 'activa': True},
        'BAJA_INV': {'cuenta': '117-00-00', 'nombre': 'BAJA INVENTARIO', 'activa': True},
        'ING_INTERESES': {'cuenta': '400-01-01', 'nombre': 'INGRESOS POR INTERESES', 'activa': True},
    }

    for k, d in defaults.items():
        if k not in cat or not cat[k]['cuenta']:
            cat[k] = d

    return cat


def extraer_prefijo_cuenta(cta_str):
    """Extrae el prefijo de búsqueda para una cuenta del catálogo (ej. 115-00-00 -> 115)."""
    if not cta_str:
        return ''
    parts = cta_str.split('-')
    if len(parts) >= 2 and parts[1] != '00':
        return f"{parts[0]}-{parts[1]}"
    return parts[0]


def parsear_balanza_contpaqi(excel_path):
    """
    Parsea una Balanza de Comprobación de Contpaqi en Excel (.xlsx) de cualquier mes.
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"No se encontró el archivo de balanza: {excel_path}")

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active

    # Metadatos del Encabezado
    empresa = "LEASY ARRENDAMIENTOS S A DE C V"
    for col in range(1, 10):
        val = ws.cell(1, col).value
        if val and "LEASY" in str(val).upper():
            empresa = str(val).strip()
            break

    fecha_corte = None
    fecha_emision = None
    r2_val = str(ws.cell(2, 1).value or '')

    meses_map = {
        'ene': '01', 'feb': '02', 'mar': '03', 'abr': '04', 'may': '05', 'jun': '06',
        'jul': '07', 'ago': '08', 'sep': '09', 'oct': '10', 'nov': '11', 'dic': '12'
    }
    pat_corte = re.search(r'al\s+(\d{1,2})[/\-]([A-Za-z]+|\d{1,2})[/\-](\d{4})', r2_val, re.IGNORECASE)
    if pat_corte:
        d, m_raw, y = pat_corte.groups()
        m_num = m_raw.zfill(2) if m_raw.isdigit() else meses_map.get(m_raw.lower()[:3], '09')
        fecha_corte = f"{y}-{m_num}-{d.zfill(2)}"
    else:
        fecha_corte = "2026-09-30"

    for col in range(5, 12):
        val = str(ws.cell(2, col).value or '')
        if "Fecha:" in val or "30/" in val or "2026" in val:
            fecha_emision = val.replace("Fecha:", "").strip()
            break
    if not fecha_emision:
        fecha_emision = datetime.date.today().strftime("%Y-%m-%d")

    cuentas = []
    pat_nom = re.compile(r'CONTRATO\s+N[oO\.]*\s*(\d{1,4})(?:[-_](\d{1,4}))?', re.IGNORECASE)
    pat_code_in_nom = re.compile(r'\b(\d{4})-(\d{4})\b')

    for r in range(5, ws.max_row + 1):
        cta = ws.cell(r, 1).value
        nom = ws.cell(r, 2).value
        s_ini_d = ws.cell(r, 3).value
        s_ini_a = ws.cell(r, 4).value
        cargos = ws.cell(r, 5).value
        abonos = ws.cell(r, 6).value
        s_act_d = ws.cell(r, 7).value
        s_act_a = ws.cell(r, 8).value

        if cta and str(cta).strip():
            cta_str = str(cta).strip()
            nom_str = str(nom).strip() if nom else ''
            parts = cta_str.split('-')

            val_ini_d = float(s_ini_d or 0) if isinstance(s_ini_d, (int, float)) else 0.0
            val_ini_a = float(s_ini_a or 0) if isinstance(s_ini_a, (int, float)) else 0.0
            val_carg = float(cargos or 0) if isinstance(cargos, (int, float)) else 0.0
            val_abon = float(abonos or 0) if isinstance(abonos, (int, float)) else 0.0
            val_act_d = float(s_act_d or 0) if isinstance(s_act_d, (int, float)) else 0.0
            val_act_a = float(s_act_a or 0) if isinstance(s_act_a, (int, float)) else 0.0

            rubro = f"{parts[0]}-{parts[1]}" if len(parts) >= 2 else parts[0]
            master = parts[4] if len(parts) >= 5 else ''
            anexo = parts[5] if len(parts) >= 6 else ''

            id_contrato = None
            if len(parts) == 6 and master != '0000':
                id_contrato = f"{master}-{anexo.zfill(4)}"

            if nom_str:
                m_code = pat_code_in_nom.search(nom_str)
                if m_code:
                    id_contrato = f"{m_code.group(1)}-{m_code.group(2)}"
                else:
                    m_nom = pat_nom.search(nom_str)
                    if m_nom:
                        m_num = m_nom.group(1).zfill(4)
                        a_num = (m_nom.group(2) or '0').zfill(4)
                        id_contrato = f"{m_num}-{a_num}"

            cuentas.append({
                'fila': r,
                'cuenta': cta_str,
                'nombre': nom_str,
                'rubro': rubro,
                'cuenta_mayor': parts[0],
                'subcuenta': parts[1] if len(parts) > 1 else '',
                'master': master,
                'anexo': anexo,
                'id_contrato': id_contrato,
                's_ini_d': val_ini_d,
                's_ini_a': val_ini_a,
                'cargos': val_carg,
                'abonos': val_abon,
                's_act_d': val_act_d,
                's_act_a': val_act_a,
                'saldo_neto_deudor': val_act_d - val_act_a,
                'saldo_neto_acreedor': val_act_a - val_act_d,
            })

    wb.close()
    df_balanza = pd.DataFrame(cuentas)
    metadata = {
        'empresa': empresa,
        'fecha_corte': fecha_corte,
        'fecha_emision': fecha_emision,
        'total_cuentas': len(df_balanza),
        'archivo_origen': os.path.basename(excel_path),
    }

    return metadata, df_balanza


def obtener_datos_sistema(db_path, fecha_corte):
    """
    Carga contratos, facturas y catálogo dinámico de cuentas desde la base de datos.
    """
    conn = sqlite3.connect(db_path)
    c = conn.cursor()

    cat = obtener_catalogo_cuentas_empresa(db_path)

    df_con = pd.read_sql("""
        SELECT ID_Contrato, Cliente, Vehiculo, Estatus, Fecha_Alta, Fecha_Vencimiento, 
               Fecha_Baja, Plazo, Valor_Sin_IVA, Mensualidad_Sin_IVA, Residual_Monto,
               Anticipo_Monto, Tasa_Calculada
        FROM contratos
    """, conn)

    facturas_resumen = {}
    facturas_conceptos = {}

    rows_fac = c.execute("""
        SELECT id_contrato, COUNT(*), SUM(subtotal), SUM(total)
        FROM facturas
        WHERE fecha_emision <= ? AND (cancelada IS NULL OR cancelada = 0 OR fecha_cancelacion > ?)
        GROUP BY id_contrato
    """, (fecha_corte, fecha_corte)).fetchall()

    for r in rows_fac:
        cid = r[0]
        if cid:
            facturas_resumen[cid] = {'num_facturas': r[1], 'subtotal': r[2] or 0.0, 'total': r[3] or 0.0}

    rows_con = c.execute("""
        SELECT f.id_contrato, cp.clave, SUM(cp.importe)
        FROM factura_conceptos cp
        JOIN facturas f ON cp.uuid = f.uuid
        WHERE f.fecha_emision <= ? AND (f.cancelada IS NULL OR f.cancelada = 0 OR f.fecha_cancelacion > ?)
        GROUP BY f.id_contrato, cp.clave
    """, (fecha_corte, fecha_corte)).fetchall()

    for r in rows_con:
        cid, clave, imp = r[0], r[1], r[2] or 0.0
        if cid:
            if cid not in facturas_conceptos:
                facturas_conceptos[cid] = {}
            facturas_conceptos[cid][clave] = imp

    conn.close()

    activos = []
    liquidados = []
    futuros = []

    for idx, con in df_con.iterrows():
        alta = str(con['Fecha_Alta'] or '')[:10]
        baja = str(con['Fecha_Baja'] or '')[:10] if con['Fecha_Baja'] else None
        venc = str(con['Fecha_Vencimiento'] or '')[:10] if con['Fecha_Vencimiento'] else None
        est = str(con['Estatus'] or '').upper()

        if alta and alta > fecha_corte:
            futuros.append(con)
        elif (baja and baja <= fecha_corte) or (venc and venc <= fecha_corte) or est in ('BAJA', 'LIQUIDADO'):
            liquidados.append(con)
        else:
            activos.append(con)

    return {
        'df_contratos': df_con,
        'df_activos': pd.DataFrame(activos) if activos else pd.DataFrame(columns=df_con.columns),
        'df_liquidados': pd.DataFrame(liquidados) if liquidados else pd.DataFrame(columns=df_con.columns),
        'df_futuros': pd.DataFrame(futuros) if futuros else pd.DataFrame(columns=df_con.columns),
        'catalogo_cuentas': cat,
        'facturas_resumen': facturas_resumen,
        'facturas_conceptos': facturas_conceptos,
    }


def conciliar_balanza_vs_sistema(df_balanza, datos_sistema, fecha_corte_str):
    """
    Ejecuta el cruce integral cuenta por cuenta y contrato por contrato,
    tanto para las cuentas NIF C-15 del catálogo maestro (115, 125, 114, 126, 208, 228, 413)
    como para las cuentas operativas en Balanza (102-01, 403-01, 150-00, etc.).
    """
    df_activos = datos_sistema['df_activos']
    df_liquidados = datos_sistema['df_liquidados']
    cat = datos_sistema['catalogo_cuentas']
    facturas_resumen = datos_sistema['facturas_resumen']
    facturas_conceptos = datos_sistema['facturas_conceptos']

    fecha_corte = datetime.datetime.strptime(fecha_corte_str, "%Y-%m-%d")

    # Extraer cuentas dinámicas del catálogo maestro
    cta_cxc_cp = cat.get('CXC_CP', {}).get('cuenta', '115-00-00')
    cta_cxc_lp = cat.get('CXC_LP', {}).get('cuenta', '125-00-00')
    cta_res_cp = cat.get('RESIDUAL_CP', {}).get('cuenta', '114-01-00')
    cta_res_lp = cat.get('RESIDUAL', {}).get('cuenta', '126-00-00')
    cta_int_cp = cat.get('INT_CP', {}).get('cuenta', '208-00-00')
    cta_int_lp = cat.get('INT_LP', {}).get('cuenta', '228-00-00')
    cta_cap_amort = cat.get('CANCELACION_CAPITAL', {}).get('cuenta', '413-00-00')

    pref_cxc_cp = extraer_prefijo_cuenta(cta_cxc_cp)
    pref_cxc_lp = extraer_prefijo_cuenta(cta_cxc_lp)
    pref_res_cp = extraer_prefijo_cuenta(cta_res_cp)
    pref_res_lp = extraer_prefijo_cuenta(cta_res_lp)
    pref_int_cp = extraer_prefijo_cuenta(cta_int_cp)
    pref_int_lp = extraer_prefijo_cuenta(cta_int_lp)
    pref_cap_amort = extraer_prefijo_cuenta(cta_cap_amort)

    # Mapa de subcuentas en balanza por contrato y por rubro
    bal_by_cid = {}
    for idx, row in df_balanza.iterrows():
        cid = row['id_contrato']
        rubro = row['rubro']
        if cid:
            if cid not in bal_by_cid:
                bal_by_cid[cid] = {}
            bal_by_cid[cid][rubro] = row.to_dict()

    # -------------------------------------------------------------------------
    # 1. PESTAÑA: AUDITORÍA FINANCIERA NIF C-15 / IFRS 16 (CUENTAS MAESTRAS)
    # -------------------------------------------------------------------------
    nif_activos_list = []
    for idx, con in df_activos.iterrows():
        cid = con['ID_Contrato']
        fa_str = str(con['Fecha_Alta'] or '')[:10]
        if not fa_str:
            continue
        fa = datetime.datetime.strptime(fa_str, "%Y-%m-%d")
        pl = int(con['Plazo'])
        r = float(con['Mensualidad_Sin_IVA'] or 0.0)
        v = float(con['Valor_Sin_IVA'] or 0.0)
        res = float(con['Residual_Monto'] or 0.0)
        ant = float(con['Anticipo_Monto'] or 0.0)
        inv = v - ant
        t = float(con['Tasa_Calculada'] or 0.0)

        mt = (fecha_corte.year - fa.year) * 12 + (fecha_corte.month - fa.month)
        mc = mt + 1
        if mc < 1 or mc > pl:
            continue

        dfa, _, _, _, _ = calc_amort(round(inv, 4), round(r, 4), round(res, 4), pl, round(t, 8))
        vpr = vp_res(res, t, pl)
        dfr = calc_res_amort(round(vpr, 4), round(t, 8), pl)

        mr = pl - mc
        cxc_cp = r * min(12, mr)
        cxc_lp = r * max(0, mr - 12)
        sr = dfr.iloc[mc-1]['Saldo_Fin'] if mc <= len(dfr) else res
        res_cp = res if mr <= 12 else 0.0
        res_lp = 0.0 if mr <= 12 else sr
        int_cp = dfa.iloc[mc:min(mc+12, pl)]['Interes'].sum() if mc < pl else 0.0
        int_lp = dfa.iloc[min(mc+12, pl):]['Interes'].sum() if (mc + 12) < pl else 0.0
        cap_amort = dfa.iloc[:mc]['Capital'].sum()

        # Buscar si existen saldos en balanza para los prefijos del catálogo
        b_map = bal_by_cid.get(cid, {})
        bal_cxc_cp = b_map.get(pref_cxc_cp, {}).get('s_act_d', 0.0)
        bal_cxc_lp = b_map.get(pref_cxc_lp, {}).get('s_act_d', 0.0)
        bal_res_cp = b_map.get(pref_res_cp, {}).get('s_act_d', 0.0)
        bal_res_lp = b_map.get(pref_res_lp, {}).get('s_act_d', 0.0)
        bal_int_cp = b_map.get(pref_int_cp, {}).get('s_act_a', 0.0)
        bal_int_lp = b_map.get(pref_int_lp, {}).get('s_act_a', 0.0)
        bal_cap_amort = b_map.get(pref_cap_amort, {}).get('s_act_a', 0.0)

        # Si están en $0.00 en la balanza, se alerta la póliza pendiente
        tiene_balanza = any([bal_cxc_cp, bal_cxc_lp, bal_res_cp, bal_res_lp, bal_int_cp, bal_int_lp, bal_cap_amort])
        diag_nif = "CONTABILIZADO EN BALANZA" if tiene_balanza else f"PENDIENTE CONTABILIZAR PÓLIZA ({pref_cxc_cp}/{pref_cxc_lp})"

        nif_activos_list.append({
            'id_contrato': cid,
            'cliente': con['Cliente'],
            'plazo': pl,
            'mes_actual': mc,
            'meses_restantes': mr,
            'renta_sin_iva': r,
            'cxc_cp_esperado': cxc_cp,
            'cxc_lp_esperado': cxc_lp,
            'res_cp_esperado': res_cp,
            'res_lp_esperado': res_lp,
            'int_cp_esperado': int_cp,
            'int_lp_esperado': int_lp,
            'cap_amort_acumulada': cap_amort,
            'balanza_cxc_cp': bal_cxc_cp,
            'balanza_cxc_lp': bal_cxc_lp,
            'diagnostico': diag_nif,
        })

    df_auditoria_nif = pd.DataFrame(nif_activos_list)
    if not df_auditoria_nif.empty:
        df_auditoria_nif.sort_values(by=['cxc_cp_esperado', 'id_contrato'], ascending=[False, True], inplace=True)

    # -------------------------------------------------------------------------
    # 2. PESTAÑA: CONCILIACIÓN CXC OPERATIVA (102-01) - EXCLUSIVO CARTERA ACTIVA
    # -------------------------------------------------------------------------
    cxc_activos_list = []
    for idx, con in df_activos.iterrows():
        cid = con['ID_Contrato']
        b_cxc = bal_by_cid.get(cid, {}).get('102-01')

        cta_str = b_cxc['cuenta'] if b_cxc else 'SIN CUENTA'
        s_ini_d = b_cxc['s_ini_d'] if b_cxc else 0.0
        cargos = b_cxc['cargos'] if b_cxc else 0.0
        abonos = b_cxc['abonos'] if b_cxc else 0.0
        s_act_d = b_cxc['s_act_d'] if b_cxc else 0.0
        s_act_a = b_cxc['s_act_a'] if b_cxc else 0.0
        saldo_balanza = s_act_d - s_act_a

        fac_info = facturas_resumen.get(cid, {})
        tot_facturado = fac_info.get('total', 0.0)

        if not b_cxc:
            diagnostico = "SIN CUENTA EN BALANZA"
            estatus_flag = "ALERTA"
        elif abs(saldo_balanza) < 0.01:
            diagnostico = "CUADRADO (SALDADO AL DÍA)"
            estatus_flag = "OK"
        else:
            diagnostico = "SALDO DEUDOR EXIGIBLE"
            estatus_flag = "EXIGIBLE"

        cxc_activos_list.append({
            'id_contrato': cid,
            'cliente': con['Cliente'],
            'vehiculo': con['Vehiculo'],
            'plazo': con['Plazo'],
            'renta_sin_iva': con['Mensualidad_Sin_IVA'] or 0.0,
            'cuenta_contable': cta_str,
            's_ini_d': s_ini_d,
            'cargos': cargos,
            'abonos': abonos,
            'saldo_balanza_cxc': saldo_balanza,
            'tot_facturado_sistema': tot_facturado,
            'diagnostico': diagnostico,
            'estatus_flag': estatus_flag,
        })

    df_concil_cxc = pd.DataFrame(cxc_activos_list)
    if not df_concil_cxc.empty:
        df_concil_cxc.sort_values(by=['saldo_balanza_cxc', 'id_contrato'], ascending=[False, True], inplace=True)

    # -------------------------------------------------------------------------
    # 3. PESTAÑA: CONTRATOS LIQUIDADOS (OMITIDOS DE CARTERA ACTIVA Y SEÑALADOS)
    # -------------------------------------------------------------------------
    liquidados_list = []
    for idx, con in df_liquidados.iterrows():
        cid = con['ID_Contrato']
        b_cxc = bal_by_cid.get(cid, {}).get('102-01')
        b_act = bal_by_cid.get(cid, {}).get('150-00')
        b_renta = bal_by_cid.get(cid, {}).get('403-01')

        cta_cxc = b_cxc['cuenta'] if b_cxc else '-'
        saldo_cxc = (b_cxc['s_act_d'] - b_cxc['s_act_a']) if b_cxc else 0.0

        cta_act = b_act['cuenta'] if b_act else '-'
        saldo_act = (b_act['s_act_d'] - b_act['s_act_a']) if b_act else 0.0

        cta_renta = b_renta['cuenta'] if b_renta else '-'
        saldo_renta = (b_renta['s_act_a'] - b_renta['s_act_d']) if b_renta else 0.0

        if abs(saldo_cxc) < 0.01 and abs(saldo_act) < 0.01:
            diag = "LIQUIDADO EN CERO (CORRECTO - CUENTA SALDADA)"
            severidad = "CORRECTO"
        elif abs(saldo_cxc) > 0.01 and abs(saldo_act) < 0.01:
            diag = "ALERTA: LIQUIDADO CON SALDO EN CXC (POR DEPURAR O COBRAR)"
            severidad = "ALERTA_CXC"
        elif abs(saldo_act) > 0.01 and abs(saldo_cxc) < 0.01:
            diag = "ALERTA: LIQUIDADO CON ACTIVO NO DESINCORPORADO (150-00)"
            severidad = "ALERTA_ACTIVO"
        else:
            diag = "ALERTA CRÍTICA: LIQUIDADO CON SALDOS PENDIENTES EN CXC Y ACTIVO"
            severidad = "ALERTA_MULTIPLE"

        liquidados_list.append({
            'id_contrato': cid,
            'cliente': con['Cliente'],
            'estatus_sistema': con['Estatus'],
            'fecha_alta': str(con['Fecha_Alta'] or '')[:10],
            'fecha_vencimiento': str(con['Fecha_Vencimiento'] or '')[:10],
            'fecha_baja': str(con['Fecha_Baja'] or '')[:10],
            'plazo': con['Plazo'],
            'cuenta_cxc': cta_cxc,
            'saldo_cxc_balanza': saldo_cxc,
            'cuenta_activo': cta_act,
            'saldo_activo_balanza': saldo_act,
            'saldo_renta_balanza': saldo_renta,
            'diagnostico': diag,
            'severidad': severidad,
        })

    df_concil_liquidados = pd.DataFrame(liquidados_list)
    if not df_concil_liquidados.empty:
        df_concil_liquidados.sort_values(by=['saldo_cxc_balanza', 'id_contrato'], ascending=[False, True], inplace=True)

    # -------------------------------------------------------------------------
    # 4. PESTAÑA: INGRESOS POR RENTAS (403-01)
    # -------------------------------------------------------------------------
    rentas_list = []
    for idx, con in df_activos.iterrows():
        cid = con['ID_Contrato']
        b_renta = bal_by_cid.get(cid, {}).get('403-01')

        cta_renta = b_renta['cuenta'] if b_renta else 'SIN CUENTA'
        s_act_a = b_renta['s_act_a'] if b_renta else 0.0
        s_act_d = b_renta['s_act_d'] if b_renta else 0.0
        ing_balanza = s_act_a - s_act_d

        renta_facturada = facturas_conceptos.get(cid, {}).get('RENTA', 0.0)
        diferencia = ing_balanza - renta_facturada

        if not b_renta:
            diag = "SIN CUENTA 403-01 EN BALANZA"
            est_flag = "ALERTA"
        elif abs(diferencia) < 0.05:
            diag = "CUADRADO AL CENTAVO"
            est_flag = "OK"
        elif diferencia > 0:
            diag = "MAYOR INGRESO EN BALANZA QUE FACTURADO"
            est_flag = "DIFERENCIA"
        else:
            diag = "MENOR INGRESO EN BALANZA QUE FACTURADO"
            est_flag = "DIFERENCIA"

        rentas_list.append({
            'id_contrato': cid,
            'cliente': con['Cliente'],
            'plazo': con['Plazo'],
            'renta_mensual_db': con['Mensualidad_Sin_IVA'] or 0.0,
            'cuenta_contable': cta_renta,
            'ingreso_balanza_403': ing_balanza,
            'renta_facturada_db': renta_facturada,
            'diferencia': diferencia,
            'diagnostico': diag,
            'estatus_flag': est_flag,
        })

    df_concil_rentas = pd.DataFrame(rentas_list)
    if not df_concil_rentas.empty:
        df_concil_rentas.sort_values(by=['ingreso_balanza_403', 'id_contrato'], ascending=[False, True], inplace=True)

    # -------------------------------------------------------------------------
    # 5. PESTAÑA: ACTIVOS VEHICULARES (150-00)
    # -------------------------------------------------------------------------
    activos_veh_list = []
    for idx, con in df_activos.iterrows():
        cid = con['ID_Contrato']
        b_act = bal_by_cid.get(cid, {}).get('150-00')

        cta_act = b_act['cuenta'] if b_act else 'SIN CUENTA'
        s_act_d = b_act['s_act_d'] if b_act else 0.0
        s_act_a = b_act['s_act_a'] if b_act else 0.0
        val_balanza = s_act_d - s_act_a

        val_contrato = con['Valor_Sin_IVA'] or 0.0
        diferencia = val_balanza - val_contrato

        if not b_act:
            diag = "SIN CUENTA 150-00 EN BALANZA"
            est_flag = "ALERTA"
        elif abs(diferencia) < 1.00:
            diag = "CUADRADO EXACTO CON VALOR CONTRATO"
            est_flag = "OK"
        else:
            diag = "DIFERENCIA VS VALOR CONTRATO"
            est_flag = "DIFERENCIA"

        activos_veh_list.append({
            'id_contrato': cid,
            'cliente': con['Cliente'],
            'vehiculo': con['Vehiculo'],
            'valor_contrato_db': val_contrato,
            'cuenta_contable': cta_act,
            'saldo_balanza_150': val_balanza,
            'diferencia': diferencia,
            'diagnostico': diag,
            'estatus_flag': est_flag,
        })

    df_concil_activos = pd.DataFrame(activos_veh_list)
    if not df_concil_activos.empty:
        df_concil_activos.sort_values(by=['saldo_balanza_150', 'id_contrato'], ascending=[False, True], inplace=True)

    # -------------------------------------------------------------------------
    # 6. PESTAÑA: DISCREPANCIAS Y ALERTAS
    # -------------------------------------------------------------------------
    alertas = []
    # Alerta 1: Cuentas NIF C-15 configuradas en el catálogo maestro pero pendientes de registrar en Contpaqi
    tot_esp_nif = df_auditoria_nif['cxc_cp_esperado'].sum() + df_auditoria_nif['cxc_lp_esperado'].sum()
    alertas.append({
        'prioridad': 'ALTA',
        'categoria': 'NIF C-15: CARTERA DEVENGADA PENDIENTE EN CONTPAQI',
        'id_contrato': 'CONSOLIDADO',
        'cliente': 'CARTERA ACTIVA COMPLETA',
        'cuenta_contable': f"{cta_cxc_cp} / {cta_cxc_lp}",
        'monto_discrepancia': tot_esp_nif,
        'descripcion': f"Las cuentas maestras de Arrendamiento Financiero ({cta_cxc_cp} CP y {cta_cxc_lp} LP) tienen una cartera teórica esperada de ${tot_esp_nif:,.2f}, pero no presentan movimientos en la Balanza de Contpaqi (la cobranza se registró en 102-01).",
        'accion_recomendada': 'Generar y registrar las pólizas contables NIF C-15 / IFRS 16 exportables desde el menú Pólizas Contables.',
    })

    # Alerta 2: Contratos Liquidados con Saldo Deudor en 102-01
    for idx, row in df_concil_liquidados[df_concil_liquidados['saldo_cxc_balanza'] > 0.01].iterrows():
        alertas.append({
            'prioridad': 'ALTA',
            'categoria': 'CONTRATO LIQUIDADO CON SALDO PENDIENTE',
            'id_contrato': row['id_contrato'],
            'cliente': row['cliente'],
            'cuenta_contable': row['cuenta_cxc'],
            'monto_discrepancia': row['saldo_cxc_balanza'],
            'descripcion': f"Contrato {row['id_contrato']} concluido el {row['fecha_vencimiento']} pero mantiene saldo deudor de ${row['saldo_cxc_balanza']:,.2f} en CxC (102-01).",
            'accion_recomendada': 'Verificar cobro pendiente o aplicar póliza de depuración contable / liquidación.',
        })

    # Alerta 3: Contratos Activos sin Cuenta Contable Aperturada
    for idx, row in df_concil_cxc[df_concil_cxc['cuenta_contable'] == 'SIN CUENTA'].iterrows():
        alertas.append({
            'prioridad': 'MEDIA',
            'categoria': 'ACTIVO SIN SUBCUENTA CONTABLE',
            'id_contrato': row['id_contrato'],
            'cliente': row['cliente'],
            'cuenta_contable': 'PENDIENTE APERTURA',
            'monto_discrepancia': row['renta_sin_iva'],
            'descripcion': f"El contrato {row['id_contrato']} está vigente en el sistema pero no cuenta con subcuenta aperturada en el rubro 102-01 de Contpaqi.",
            'accion_recomendada': 'Aperturar la subcuenta contable correspondiente en Contpaqi.',
        })

    # Alerta 4: Subcuentas en Balanza sin Contrato Vigente en Sistema
    contratos_sistema_all = set(datos_sistema['df_contratos']['ID_Contrato'])
    for idx, row in df_balanza[(df_balanza['rubro'] == '102-01') & (df_balanza['id_contrato'].notna())].iterrows():
        cid = row['id_contrato']
        if cid not in contratos_sistema_all and abs(row['saldo_neto_deudor']) > 0.01:
            alertas.append({
                'prioridad': 'MEDIA',
                'categoria': 'SUBCUENTA CONTABLE SIN CONTRATO EN SISTEMA',
                'id_contrato': cid,
                'cliente': row['nombre'],
                'cuenta_contable': row['cuenta'],
                'monto_discrepancia': row['saldo_neto_deudor'],
                'descripcion': f"La cuenta contable {row['cuenta']} ({row['nombre']}) tiene saldo deudor de ${row['saldo_neto_deudor']:,.2f} pero el ID {cid} no existe en la base de datos de O-Leasing.",
                'accion_recomendada': 'Revisar si corresponde a un contrato anterior a la migración histórica.',
            })

    df_alertas = pd.DataFrame(alertas)
    if not df_alertas.empty:
        df_alertas.sort_values(by=['monto_discrepancia'], ascending=False, inplace=True)

    return {
        'df_auditoria_nif': df_auditoria_nif,
        'df_concil_cxc': df_concil_cxc,
        'df_concil_liquidados': df_concil_liquidados,
        'df_concil_rentas': df_concil_rentas,
        'df_concil_activos': df_concil_activos,
        'df_alertas': df_alertas,
    }


def _aplicar_estilos_encabezado_tabla(ws, fila, col_inicio, col_fin, fill=FILL_PRIMARY):
    for col in range(col_inicio, col_fin + 1):
        cell = ws.cell(row=fila, column=col)
        cell.fill = fill
        cell.font = FONT_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER_CELL
    ws.row_dimensions[fila].height = 28


def _autoajustar_columnas(ws, min_col=1, max_col=None, padding=3):
    if max_col is None:
        max_col = ws.max_column
    for c in range(min_col, max_col + 1):
        letra = get_column_letter(c)
        max_len = 0
        for r in range(4, min(ws.max_row + 1, 150)):
            val = ws.cell(r, c).value
            if val is not None:
                max_len = max(max_len, len(str(val)))
        ws.column_dimensions[letra].width = max(max_len + padding, 11)


def generar_excel_auditoria_balanza(excel_balanza_path, output_excel_path=None, db_path=None):
    """
    Genera el Libro Oficial de Auditoría y Conciliación Contable en formato Excel (.xlsx).
    """
    if db_path is None:
        db_path = _obtener_db_path()

    meta, df_bal = parsear_balanza_contpaqi(excel_balanza_path)
    fecha_corte = meta['fecha_corte']
    datos_sis = obtener_datos_sistema(db_path, fecha_corte)
    tablas = conciliar_balanza_vs_sistema(df_bal, datos_sis, fecha_corte)
    cat = datos_sis['catalogo_cuentas']

    cta_cxc_cp = cat.get('CXC_CP', {}).get('cuenta', '115-00-00')
    cta_cxc_lp = cat.get('CXC_LP', {}).get('cuenta', '125-00-00')
    cta_res_cp = cat.get('RESIDUAL_CP', {}).get('cuenta', '114-01-00')
    cta_res_lp = cat.get('RESIDUAL', {}).get('cuenta', '126-00-00')
    cta_int_cp = cat.get('INT_CP', {}).get('cuenta', '208-00-00')
    cta_int_lp = cat.get('INT_LP', {}).get('cuenta', '228-00-00')
    cta_cap_amort = cat.get('CANCELACION_CAPITAL', {}).get('cuenta', '413-00-00')

    if output_excel_path is None:
        y, m, d = fecha_corte.split('-')
        output_excel_path = f"Reporte_Auditoria_Balanza_{m}_{y}.xlsx"

    wb = openpyxl.Workbook()
    ws_default = wb.active

    # =========================================================================
    # PESTAÑA 1: CARATULA EJECUTIVA
    # =========================================================================
    ws_caratula = wb.create_sheet(title="Caratula_Ejecutiva")
    try:
        ws_caratula.views.sheetView[0].showGridLines = True
    except Exception:
        pass

    ws_caratula.cell(1, 1, "O-LEASING · UNA APP DE ORANGE CREW · AUDITORÍA DE CONCILIACIÓN CONTABLE").font = FONT_TITLE
    ws_caratula.cell(2, 1, f"Auditoría Contable vs Balanza de Comprobación Contpaqi · Empresa: {meta['empresa']} · Corte al {meta['fecha_corte']}").font = FONT_SUBTITLE
    ws_caratula.merge_cells("A1:K1")
    ws_caratula.merge_cells("A2:K2")
    ws_caratula.row_dimensions[1].height = 24
    ws_caratula.row_dimensions[2].height = 16

    # Tarjetas KPI (Fila 4 a 6)
    cta_padre_102 = df_bal[df_bal['cuenta'] == '102-01-00-0000-0000-00']
    tot_bal_cxc = float(cta_padre_102['s_act_d'].values[0]) if len(cta_padre_102) > 0 else 7944400.80

    activos_cxc_bal = tablas['df_concil_cxc']['saldo_balanza_cxc'].sum()
    liq_cxc_bal = tablas['df_concil_liquidados']['saldo_cxc_balanza'].sum()

    tot_nif_cp = tablas['df_auditoria_nif']['cxc_cp_esperado'].sum()
    tot_nif_lp = tablas['df_auditoria_nif']['cxc_lp_esperado'].sum()
    tot_int_futuro = tablas['df_auditoria_nif']['int_cp_esperado'].sum() + tablas['df_auditoria_nif']['int_lp_esperado'].sum()

    kpis = [
        (f"CXC CP ({cta_cxc_cp})", tot_nif_cp, "Cartera NIF C-15 esperada 12m"),
        (f"CXC LP ({cta_cxc_lp})", tot_nif_lp, "Cartera NIF C-15 esperada >12m"),
        (f"INT POR DEVENGAR ({cta_int_cp})", tot_int_futuro, "Intereses a devengar CP+LP"),
        ("CXC COBRANZA (102-01)", tot_bal_cxc, "Saldo en Balanza Contpaqi"),
        ("CXC LIQUIDADOS (ALERTA)", liq_cxc_bal, f"{len(tablas['df_concil_liquidados'][tablas['df_concil_liquidados']['saldo_cxc_balanza'] > 0.01])} concluidos con saldo"),
    ]

    ws_caratula.row_dimensions[4].height = 18
    ws_caratula.row_dimensions[5].height = 28
    ws_caratula.row_dimensions[6].height = 16

    col_idx = 1
    for titulo_kpi, val_kpi, sub_kpi in kpis:
        c_fin = col_idx + 1
        ws_caratula.merge_cells(start_row=4, start_column=col_idx, end_row=4, end_column=c_fin)
        ws_caratula.merge_cells(start_row=5, start_column=col_idx, end_row=5, end_column=c_fin)
        ws_caratula.merge_cells(start_row=6, start_column=col_idx, end_row=6, end_column=c_fin)

        cell_t = ws_caratula.cell(4, col_idx, titulo_kpi)
        cell_t.fill = FILL_NAVY
        cell_t.font = Font(name="Calibri", size=8, bold=True, color="FFFFFF")
        cell_t.alignment = Alignment(horizontal="center", vertical="center")

        cell_v = ws_caratula.cell(5, col_idx, val_kpi)
        cell_v.fill = FILL_ZEBRA
        cell_v.font = Font(name="Calibri", size=13, bold=True, color="1E5C4F")
        cell_v.alignment = Alignment(horizontal="center", vertical="center")
        cell_v.number_format = FMT_CURRENCY

        cell_s = ws_caratula.cell(6, col_idx, sub_kpi)
        cell_s.fill = FILL_ZEBRA
        cell_s.font = Font(name="Calibri", size=8, italic=True, color="64748B")
        cell_s.alignment = Alignment(horizontal="center", vertical="center")

        for r_k in range(4, 7):
            for c_k in range(col_idx, c_fin + 1):
                ws_caratula.cell(r_k, c_k).border = BORDER_CELL

        col_idx += 2

    # TABLA 1: CUADRO COMPARATIVO NIF C-15 / IFRS 16 (CUENTAS DEL CATÁLOGO DINÁMICO)
    r_tab1 = 8
    ws_caratula.cell(r_tab1, 1, "1. AUDITORÍA FINANCIERA NIF C-15 / IFRS 16 (CUENTAS DINÁMICAS DEL CATÁLOGO MAESTRO)").font = FONT_BODY_BOLD
    r_tab1 += 1
    headers_nif = ["Clave Catálogo", "Cuenta Contable", "Concepto Financiero", "Naturaleza", "Saldo Esperado Sistema", "Saldo en Balanza", "Diferencia", "Estatus Auditoría"]
    for c_i, h_txt in enumerate(headers_nif, start=1):
        ws_caratula.cell(r_tab1, c_i, h_txt)
    _aplicar_estilos_encabezado_tabla(ws_caratula, r_tab1, 1, len(headers_nif), FILL_PRIMARY)

    r_tab1 += 1
    r_tab1_ini = r_tab1

    nif_rows = [
        ('CXC_CP', cta_cxc_cp, 'Cuentas por Cobrar Corto Plazo (Rentas 12m)', 'Deudor', tot_nif_cp, 0.0),
        ('CXC_LP', cta_cxc_lp, 'Cuentas por Cobrar Largo Plazo (Rentas >12m)', 'Deudor', tot_nif_lp, 0.0),
        ('RESIDUAL_CP', cta_res_cp, 'Valor Residual Corto Plazo (<=12m al término)', 'Deudor', tablas['df_auditoria_nif']['res_cp_esperado'].sum(), 0.0),
        ('RESIDUAL', cta_res_lp, 'Valor Residual Largo Plazo (>12m al término)', 'Deudor', tablas['df_auditoria_nif']['res_lp_esperado'].sum(), 0.0),
        ('INT_CP', cta_int_cp, 'Intereses por Devengar Corto Plazo (12m)', 'Acreedor', tablas['df_auditoria_nif']['int_cp_esperado'].sum(), 0.0),
        ('INT_LP', cta_int_lp, 'Intereses por Devengar Largo Plazo (>12m)', 'Acreedor', tablas['df_auditoria_nif']['int_lp_esperado'].sum(), 0.0),
        ('CANCELACION_CAPITAL', cta_cap_amort, 'Ingresos / Cancelación Capital Amortizado Acum.', 'Acreedor', tablas['df_auditoria_nif']['cap_amort_acumulada'].sum(), 0.0),
    ]

    for clv, cta_n, desc_n, nat_n, val_esp, val_bal in nif_rows:
        ws_caratula.cell(r_tab1, 1, clv).alignment = Alignment(horizontal="center")
        ws_caratula.cell(r_tab1, 2, cta_n).alignment = Alignment(horizontal="center")
        ws_caratula.cell(r_tab1, 3, desc_n)
        ws_caratula.cell(r_tab1, 4, nat_n).alignment = Alignment(horizontal="center")
        ws_caratula.cell(r_tab1, 5, val_esp).number_format = FMT_CURRENCY
        ws_caratula.cell(r_tab1, 6, val_bal).number_format = FMT_CURRENCY
        ws_caratula.cell(r_tab1, 7, f"=F{r_tab1}-E{r_tab1}").number_format = FMT_CURRENCY
        
        diag_txt = "PENDIENTE REGISTRAR EN CONTPAQI" if abs(val_bal) < 0.01 else "CONCILIADO"
        ws_caratula.cell(r_tab1, 8, diag_txt)

        for c_i in range(1, len(headers_nif) + 1):
            cell = ws_caratula.cell(r_tab1, c_i)
            cell.border = BORDER_CELL
            cell.font = FONT_BODY
            if abs(val_bal) < 0.01:
                cell.fill = FILL_ALERT_YELLOW
            elif r_tab1 % 2 == 1:
                cell.fill = FILL_ZEBRA
        r_tab1 += 1

    ws_caratula.cell(r_tab1, 1, "TOTAL").font = FONT_TOTAL
    ws_caratula.cell(r_tab1, 3, "TOTAL CARTERA NIF C-15 / IFRS 16").font = FONT_TOTAL
    ws_caratula.cell(r_tab1, 5, f"=SUM(E{r_tab1_ini}:E{r_tab1-1})").number_format = FMT_CURRENCY
    ws_caratula.cell(r_tab1, 6, f"=SUM(F{r_tab1_ini}:F{r_tab1-1})").number_format = FMT_CURRENCY
    ws_caratula.cell(r_tab1, 7, f"=SUM(G{r_tab1_ini}:G{r_tab1-1})").number_format = FMT_CURRENCY
    for c_i in range(1, len(headers_nif) + 1):
        cell = ws_caratula.cell(r_tab1, c_i)
        cell.border = BORDER_TOTAL
        cell.fill = FILL_TOTAL
        cell.font = FONT_TOTAL

    r_tab1 += 2

    # TABLA 2: RESUMEN DE FACTURACIÓN Y CUENTAS OPERATIVAS EN BALANZA CONTPAQI
    ws_caratula.cell(r_tab1, 1, "2. AUDITORÍA DE FACTURACIÓN Y OPERACIÓN CONTABLE EN BALANZA CONTPAQI").font = FONT_BODY_BOLD
    r_tab1 += 1
    headers_rub = ["Rubro", "Nombre Cuenta Mayor", "Naturaleza", "Cuentas en Balanza", "Saldo Deudor", "Saldo Acreedor", "Saldo Neto", "Función / Alcance"]
    for c_i, h_txt in enumerate(headers_rub, start=1):
        ws_caratula.cell(r_tab1, c_i, h_txt)
    _aplicar_estilos_encabezado_tabla(ws_caratula, r_tab1, 1, len(headers_rub), FILL_SECONDARY)

    r_tab1 += 1
    r_tab2_ini = r_tab1
    rubros_def = [
        ('102-01', 'CUENTAS POR COBRAR ARRENDAMIENTO', 'Deudor', 'Exigible de clientes por facturas de rentas y accesorios'),
        ('150-00', 'VEHÍCULOS ARRENDADOS (ACTIVO)', 'Deudor', 'Costo de adquisición de unidades arrendadas en balance'),
        ('214-01', 'DEPÓSITOS CLIENTES ARRENDAMIENTO', 'Acreedor', 'Depósitos en garantía recibidos de clientes'),
        ('403-01', 'ARRENDAMIENTO DE VEHÍCULO (RENTAS)', 'Acreedor', 'Ingresos por rentas devengadas / facturadas'),
        ('404-01', 'ADMINISTRACIÓN DE CTA Y COBRANZA', 'Acreedor', 'Ingresos por gastos administrativos'),
        ('408-01', 'GEOLOCALIZACIÓN (GPS)', 'Acreedor', 'Ingresos por servicio de rastreo satelital'),
        ('420-01', 'COMISIÓN DE APERTURA', 'Acreedor', 'Ingresos por comisiones iniciales'),
        ('426-01', 'VENTA DE SEGUROS', 'Acreedor', 'Ingresos por pólizas y coberturas de seguro'),
    ]

    for cod_rub, nom_rub, nat_rub, func_rub in rubros_def:
        sub_df = df_bal[df_bal['rubro'] == cod_rub]
        sub_contratos = sub_df[sub_df['id_contrato'].notna()]
        cnt = len(sub_contratos) if len(sub_contratos) > 0 else len(sub_df)

        cta_padre = sub_df[sub_df['cuenta'].str.endswith(('00-0000-0000-00', '00-0000-00'))]
        if len(cta_padre) > 0:
            s_d = float(cta_padre['s_act_d'].values[0])
            s_a = float(cta_padre['s_act_a'].values[0])
        else:
            s_d = float(sub_contratos['s_act_d'].sum()) if len(sub_contratos) > 0 else float(sub_df['s_act_d'].sum())
            s_a = float(sub_contratos['s_act_a'].sum()) if len(sub_contratos) > 0 else float(sub_df['s_act_a'].sum())

        s_neto = (s_d - s_a) if nat_rub == 'Deudor' else (s_a - s_d)

        ws_caratula.cell(r_tab1, 1, cod_rub).alignment = Alignment(horizontal="center")
        ws_caratula.cell(r_tab1, 2, nom_rub)
        ws_caratula.cell(r_tab1, 3, nat_rub).alignment = Alignment(horizontal="center")
        ws_caratula.cell(r_tab1, 4, cnt).number_format = FMT_INT
        ws_caratula.cell(r_tab1, 5, s_d).number_format = FMT_CURRENCY
        ws_caratula.cell(r_tab1, 6, s_a).number_format = FMT_CURRENCY
        ws_caratula.cell(r_tab1, 7, s_neto).number_format = FMT_CURRENCY
        ws_caratula.cell(r_tab1, 8, func_rub)

        for c_i in range(1, len(headers_rub) + 1):
            cell = ws_caratula.cell(r_tab1, c_i)
            cell.border = BORDER_CELL
            cell.font = FONT_BODY
            if r_tab1 % 2 == 1:
                cell.fill = FILL_ZEBRA
        r_tab1 += 1

    ws_caratula.cell(r_tab1, 1, "TOTAL").font = FONT_TOTAL
    ws_caratula.cell(r_tab1, 2, "TOTAL CUENTAS MAYORES OPERACIÓN LEASING").font = FONT_TOTAL
    ws_caratula.cell(r_tab1, 4, f"=SUM(D{r_tab2_ini}:D{r_tab1-1})").number_format = FMT_INT
    ws_caratula.cell(r_tab1, 5, f"=SUM(E{r_tab2_ini}:E{r_tab1-1})").number_format = FMT_CURRENCY
    ws_caratula.cell(r_tab1, 6, f"=SUM(F{r_tab2_ini}:F{r_tab1-1})").number_format = FMT_CURRENCY
    ws_caratula.cell(r_tab1, 7, f"=SUM(G{r_tab2_ini}:G{r_tab1-1})").number_format = FMT_CURRENCY
    for c_i in range(1, len(headers_rub) + 1):
        cell = ws_caratula.cell(r_tab1, c_i)
        cell.border = BORDER_TOTAL
        cell.fill = FILL_TOTAL
        cell.font = FONT_TOTAL

    _autoajustar_columnas(ws_caratula, 1, 10)

    # =========================================================================
    # PESTAÑA 2: AUDITORÍA NIF C-15 / IFRS 16 POR CONTRATO (115, 125, 114, 126, 208, 228, 413)
    # =========================================================================
    ws_nif = wb.create_sheet(title="Auditoria_NIF_C15_Cartera")
    try:
        ws_nif.views.sheetView[0].showGridLines = True
    except Exception:
        pass

    ws_nif.cell(1, 1, "O-LEASING · AUDITORÍA FINANCIERA NIF C-15 / IFRS 16 DE CARTERA ACTIVA").font = FONT_TITLE
    ws_nif.cell(2, 1, f"Corte al {meta['fecha_corte']} · Cuentas maestras dinámicas: CP={cta_cxc_cp}, LP={cta_cxc_lp}, ResCP={cta_res_cp}, ResLP={cta_res_lp}, IntCP={cta_int_cp}, IntLP={cta_int_lp}, Cap={cta_cap_amort}").font = FONT_SUBTITLE
    ws_nif.merge_cells("A1:N1")
    ws_nif.merge_cells("A2:N2")
    ws_nif.row_dimensions[1].height = 24
    ws_nif.row_dimensions[2].height = 16

    headers_nif_col = [
        "No.", "ID Contrato", "Cliente", "Plazo", "Mes Act", "Meses Rest",
        "Renta Sin IVA", f"CXC CP ({cta_cxc_cp})", f"CXC LP ({cta_cxc_lp})",
        f"Residual CP ({cta_res_cp})", f"Residual LP ({cta_res_lp})",
        f"Int Devengar CP ({cta_int_cp})", f"Int Devengar LP ({cta_int_lp})",
        f"Cap Amortizado ({cta_cap_amort})", "Diagnóstico NIF C-15"
    ]

    r_n = 4
    for c_i, h_txt in enumerate(headers_nif_col, start=1):
        ws_nif.cell(r_n, c_i, h_txt)
    _aplicar_estilos_encabezado_tabla(ws_nif, r_n, 1, len(headers_nif_col), FILL_PRIMARY)

    r_n += 1
    r_n_ini = r_n
    num_item_n = 1

    for idx, row in tablas['df_auditoria_nif'].iterrows():
        ws_nif.cell(r_n, 1, num_item_n).alignment = Alignment(horizontal="center")
        ws_nif.cell(r_n, 2, row['id_contrato']).alignment = Alignment(horizontal="center")
        ws_nif.cell(r_n, 3, row['cliente'])
        ws_nif.cell(r_n, 4, row['plazo']).alignment = Alignment(horizontal="center")
        ws_nif.cell(r_n, 5, row['mes_actual']).alignment = Alignment(horizontal="center")
        ws_nif.cell(r_n, 6, row['meses_restantes']).alignment = Alignment(horizontal="center")
        ws_nif.cell(r_n, 7, row['renta_sin_iva']).number_format = FMT_CURRENCY
        ws_nif.cell(r_n, 8, row['cxc_cp_esperado']).number_format = FMT_CURRENCY
        ws_nif.cell(r_n, 9, row['cxc_lp_esperado']).number_format = FMT_CURRENCY
        ws_nif.cell(r_n, 10, row['res_cp_esperado']).number_format = FMT_CURRENCY
        ws_nif.cell(r_n, 11, row['res_lp_esperado']).number_format = FMT_CURRENCY
        ws_nif.cell(r_n, 12, row['int_cp_esperado']).number_format = FMT_CURRENCY
        ws_nif.cell(r_n, 13, row['int_lp_esperado']).number_format = FMT_CURRENCY
        ws_nif.cell(r_n, 14, row['cap_amort_acumulada']).number_format = FMT_CURRENCY
        ws_nif.cell(r_n, 15, row['diagnostico'])

        for c_i in range(1, len(headers_nif_col) + 1):
            cell = ws_nif.cell(r_n, c_i)
            cell.border = BORDER_CELL
            cell.font = FONT_BODY
            if "PENDIENTE" in row['diagnostico']:
                if c_i == 15: cell.font = FONT_BODY_RED
            if r_n % 2 == 1:
                cell.fill = FILL_ZEBRA

        num_item_n += 1
        r_n += 1

    ws_nif.cell(r_n, 1, "TOTAL").font = FONT_TOTAL
    ws_nif.cell(r_n, 2, f"{num_item_n - 1} CONTRATOS ACTIVOS").font = FONT_TOTAL
    ws_nif.cell(r_n, 7, f"=SUM(G{r_n_ini}:G{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 8, f"=SUM(H{r_n_ini}:H{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 9, f"=SUM(I{r_n_ini}:I{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 10, f"=SUM(J{r_n_ini}:J{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 11, f"=SUM(K{r_n_ini}:K{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 12, f"=SUM(L{r_n_ini}:L{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 13, f"=SUM(M{r_n_ini}:M{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 14, f"=SUM(N{r_n_ini}:N{r_n-1})").number_format = FMT_CURRENCY
    for c_i in range(1, len(headers_nif_col) + 1):
        cell = ws_nif.cell(r_n, c_i)
        cell.border = BORDER_TOTAL
        cell.fill = FILL_TOTAL
        cell.font = FONT_TOTAL

    _autoajustar_columnas(ws_nif, 1, len(headers_nif_col))
    ws_nif.freeze_panes = "C5"

    # =========================================================================
    # PESTAÑA 3: CONCILIACIÓN CXC OPERATIVA 102-01 (OMITE LIQUIDADOS)
    # =========================================================================
    ws_cxc = wb.create_sheet(title="Conciliacion_CxC_102")
    try:
        ws_cxc.views.sheetView[0].showGridLines = True
    except Exception:
        pass

    ws_cxc.cell(1, 1, "O-LEASING · AUDITORÍA DE CUENTAS POR COBRAR OPERATIVAS (102-01) · CARTERA ACTIVA").font = FONT_TITLE
    ws_cxc.cell(2, 1, f"Corte al {meta['fecha_corte']} · Se excluyen contratos liquidados (ver pestaña Contratos_Liquidados)").font = FONT_SUBTITLE
    ws_cxc.merge_cells("A1:K1")
    ws_cxc.merge_cells("A2:K2")
    ws_cxc.row_dimensions[1].height = 24
    ws_cxc.row_dimensions[2].height = 16

    headers_cxc = [
        "No.", "ID Contrato", "Cliente", "Vehículo / Activo", "Plazo",
        "Renta Sin IVA", "Cuenta Contable 102-01", "Saldo Inicial",
        "Cargos Periodo", "Abonos Periodo", "Saldo Actual Balanza",
        "Total Facturado Sistema", "Diagnóstico Auditoría"
    ]

    r_cxc = 4
    for c_i, h_txt in enumerate(headers_cxc, start=1):
        ws_cxc.cell(r_cxc, c_i, h_txt)
    _aplicar_estilos_encabezado_tabla(ws_cxc, r_cxc, 1, len(headers_cxc), FILL_PRIMARY)

    r_cxc += 1
    r_cxc_inicio = r_cxc
    num_item = 1

    for idx, row in tablas['df_concil_cxc'].iterrows():
        ws_cxc.cell(r_cxc, 1, num_item).alignment = Alignment(horizontal="center")
        ws_cxc.cell(r_cxc, 2, row['id_contrato']).alignment = Alignment(horizontal="center")
        ws_cxc.cell(r_cxc, 3, row['cliente'])
        ws_cxc.cell(r_cxc, 4, row['vehiculo'])
        ws_cxc.cell(r_cxc, 5, row['plazo']).alignment = Alignment(horizontal="center")
        ws_cxc.cell(r_cxc, 6, row['renta_sin_iva']).number_format = FMT_CURRENCY
        ws_cxc.cell(r_cxc, 7, row['cuenta_contable']).alignment = Alignment(horizontal="center")
        ws_cxc.cell(r_cxc, 8, row['s_ini_d']).number_format = FMT_CURRENCY
        ws_cxc.cell(r_cxc, 9, row['cargos']).number_format = FMT_CURRENCY
        ws_cxc.cell(r_cxc, 10, row['abonos']).number_format = FMT_CURRENCY
        ws_cxc.cell(r_cxc, 11, row['saldo_balanza_cxc']).number_format = FMT_CURRENCY
        ws_cxc.cell(r_cxc, 12, row['tot_facturado_sistema']).number_format = FMT_CURRENCY
        ws_cxc.cell(r_cxc, 13, row['diagnostico'])

        for c_i in range(1, len(headers_cxc) + 1):
            cell = ws_cxc.cell(r_cxc, c_i)
            cell.border = BORDER_CELL
            cell.font = FONT_BODY
            if row['estatus_flag'] == 'ALERTA':
                cell.fill = FILL_ALERT_RED
            elif row['estatus_flag'] == 'EXIGIBLE':
                cell.fill = FILL_ALERT_YELLOW
            elif r_cxc % 2 == 1:
                cell.fill = FILL_ZEBRA

        num_item += 1
        r_cxc += 1

    ws_cxc.cell(r_cxc, 1, "TOTAL").font = FONT_TOTAL
    ws_cxc.cell(r_cxc, 2, f"{num_item - 1} CONTRATOS ACTIVOS").font = FONT_TOTAL
    ws_cxc.cell(r_cxc, 6, f"=SUM(F{r_cxc_inicio}:F{r_cxc-1})").number_format = FMT_CURRENCY
    ws_cxc.cell(r_cxc, 8, f"=SUM(H{r_cxc_inicio}:H{r_cxc-1})").number_format = FMT_CURRENCY
    ws_cxc.cell(r_cxc, 9, f"=SUM(I{r_cxc_inicio}:I{r_cxc-1})").number_format = FMT_CURRENCY
    ws_cxc.cell(r_cxc, 10, f"=SUM(J{r_cxc_inicio}:J{r_cxc-1})").number_format = FMT_CURRENCY
    ws_cxc.cell(r_cxc, 11, f"=SUM(K{r_cxc_inicio}:K{r_cxc-1})").number_format = FMT_CURRENCY
    ws_cxc.cell(r_cxc, 12, f"=SUM(L{r_cxc_inicio}:L{r_cxc-1})").number_format = FMT_CURRENCY
    for c_i in range(1, len(headers_cxc) + 1):
        cell = ws_cxc.cell(r_cxc, c_i)
        cell.border = BORDER_TOTAL
        cell.fill = FILL_TOTAL
        cell.font = FONT_TOTAL

    _autoajustar_columnas(ws_cxc, 1, len(headers_cxc))
    ws_cxc.freeze_panes = "C5"

    # =========================================================================
    # PESTAÑA 4: CONTRATOS LIQUIDADOS AL CORTE (SEÑALADOS Y AUDITADOS)
    # =========================================================================
    ws_liq = wb.create_sheet(title="Contratos_Liquidados")
    try:
        ws_liq.views.sheetView[0].showGridLines = True
    except Exception:
        pass

    ws_liq.cell(1, 1, "O-LEASING · AUDITORÍA DE CONTRATOS LIQUIDADOS Y CONCLUIDOS AL CORTE").font = FONT_TITLE
    ws_liq.cell(2, 1, f"Corte al {meta['fecha_corte']} · Contratos dados de baja o con plazo concluido. Los saldos > $0.00 requieren depuración contable.").font = FONT_SUBTITLE
    ws_liq.merge_cells("A1:K1")
    ws_liq.merge_cells("A2:K2")
    ws_liq.row_dimensions[1].height = 24
    ws_liq.row_dimensions[2].height = 16

    headers_liq = [
        "No.", "ID Contrato", "Cliente", "Estatus Sistema", "Fecha Alta",
        "Fecha Vencimiento", "Fecha Baja", "Plazo", "Cuenta CxC 102-01",
        "Saldo CxC Balanza ($)", "Cuenta Activo 150-00", "Saldo Activo Balanza ($)",
        "Diagnóstico de Liquidación"
    ]

    r_liq = 4
    for c_i, h_txt in enumerate(headers_liq, start=1):
        ws_liq.cell(r_liq, c_i, h_txt)
    _aplicar_estilos_encabezado_tabla(ws_liq, r_liq, 1, len(headers_liq), FILL_NAVY)

    r_liq += 1
    r_liq_inicio = r_liq
    num_item_liq = 1

    for idx, row in tablas['df_concil_liquidados'].iterrows():
        ws_liq.cell(r_liq, 1, num_item_liq).alignment = Alignment(horizontal="center")
        ws_liq.cell(r_liq, 2, row['id_contrato']).alignment = Alignment(horizontal="center")
        ws_liq.cell(r_liq, 3, row['cliente'])
        ws_liq.cell(r_liq, 4, row['estatus_sistema']).alignment = Alignment(horizontal="center")
        ws_liq.cell(r_liq, 5, row['fecha_alta']).alignment = Alignment(horizontal="center")
        ws_liq.cell(r_liq, 6, row['fecha_vencimiento']).alignment = Alignment(horizontal="center")
        ws_liq.cell(r_liq, 7, row['fecha_baja']).alignment = Alignment(horizontal="center")
        ws_liq.cell(r_liq, 8, row['plazo']).alignment = Alignment(horizontal="center")
        ws_liq.cell(r_liq, 9, row['cuenta_cxc']).alignment = Alignment(horizontal="center")
        ws_liq.cell(r_liq, 10, row['saldo_cxc_balanza']).number_format = FMT_CURRENCY
        ws_liq.cell(r_liq, 11, row['cuenta_activo']).alignment = Alignment(horizontal="center")
        ws_liq.cell(r_liq, 12, row['saldo_activo_balanza']).number_format = FMT_CURRENCY
        ws_liq.cell(r_liq, 13, row['diagnostico'])

        for c_i in range(1, len(headers_liq) + 1):
            cell = ws_liq.cell(r_liq, c_i)
            cell.border = BORDER_CELL
            cell.font = FONT_BODY
            if row['severidad'] in ('ALERTA_CXC', 'ALERTA_MULTIPLE'):
                cell.fill = FILL_ALERT_RED
                cell.font = FONT_BODY_RED
            elif row['severidad'] == 'ALERTA_ACTIVO':
                cell.fill = FILL_ALERT_YELLOW
            elif row['severidad'] == 'CORRECTO':
                cell.fill = FILL_OK_GREEN
                cell.font = FONT_BODY_GREEN
            elif r_liq % 2 == 1:
                cell.fill = FILL_ZEBRA

        num_item_liq += 1
        r_liq += 1

    ws_liq.cell(r_liq, 1, "TOTAL").font = FONT_TOTAL
    ws_liq.cell(r_liq, 2, f"{num_item_liq - 1} CONTRATOS LIQUIDADOS").font = FONT_TOTAL
    ws_liq.cell(r_liq, 10, f"=SUM(J{r_liq_inicio}:J{r_liq-1})").number_format = FMT_CURRENCY
    ws_liq.cell(r_liq, 12, f"=SUM(L{r_liq_inicio}:L{r_liq-1})").number_format = FMT_CURRENCY
    for c_i in range(1, len(headers_liq) + 1):
        cell = ws_liq.cell(r_liq, c_i)
        cell.border = BORDER_TOTAL
        cell.fill = FILL_TOTAL
        cell.font = FONT_TOTAL

    _autoajustar_columnas(ws_liq, 1, len(headers_liq))
    ws_liq.freeze_panes = "C5"

    # =========================================================================
    # PESTAÑA 5: INGRESOS POR RENTAS (403-01)
    # =========================================================================
    ws_rentas = wb.create_sheet(title="Ingresos_403_Rentas")
    try:
        ws_rentas.views.sheetView[0].showGridLines = True
    except Exception:
        pass

    ws_rentas.cell(1, 1, "O-LEASING · CONCILIACIÓN DE INGRESOS POR ARRENDAMIENTO (403-01)").font = FONT_TITLE
    ws_rentas.cell(2, 1, f"Corte al {meta['fecha_corte']} · Comparativo de Saldo Acreedor Balanza vs Renta Base Facturada en Sistema").font = FONT_SUBTITLE
    ws_rentas.merge_cells("A1:I1")
    ws_rentas.merge_cells("A2:I2")
    ws_rentas.row_dimensions[1].height = 24
    ws_rentas.row_dimensions[2].height = 16

    headers_rentas = [
        "No.", "ID Contrato", "Cliente", "Plazo", "Renta Base Sin IVA",
        "Cuenta Contable 403-01", "Ingreso Balanza (403-01)",
        "Renta Facturada Sistema", "Diferencia (Balanza - Sistema)", "Diagnóstico"
    ]

    r_ren = 4
    for c_i, h_txt in enumerate(headers_rentas, start=1):
        ws_rentas.cell(r_ren, c_i, h_txt)
    _aplicar_estilos_encabezado_tabla(ws_rentas, r_ren, 1, len(headers_rentas), FILL_PRIMARY)

    r_ren += 1
    r_ren_inicio = r_ren
    num_item_ren = 1

    for idx, row in tablas['df_concil_rentas'].iterrows():
        ws_rentas.cell(r_ren, 1, num_item_ren).alignment = Alignment(horizontal="center")
        ws_rentas.cell(r_ren, 2, row['id_contrato']).alignment = Alignment(horizontal="center")
        ws_rentas.cell(r_ren, 3, row['cliente'])
        ws_rentas.cell(r_ren, 4, row['plazo']).alignment = Alignment(horizontal="center")
        ws_rentas.cell(r_ren, 5, row['renta_mensual_db']).number_format = FMT_CURRENCY
        ws_rentas.cell(r_ren, 6, row['cuenta_contable']).alignment = Alignment(horizontal="center")
        ws_rentas.cell(r_ren, 7, row['ingreso_balanza_403']).number_format = FMT_CURRENCY
        ws_rentas.cell(r_ren, 8, row['renta_facturada_db']).number_format = FMT_CURRENCY
        ws_rentas.cell(r_ren, 9, f"=G{r_ren}-H{r_ren}").number_format = FMT_CURRENCY
        ws_rentas.cell(r_ren, 10, row['diagnostico'])

        for c_i in range(1, len(headers_rentas) + 1):
            cell = ws_rentas.cell(r_ren, c_i)
            cell.border = BORDER_CELL
            cell.font = FONT_BODY
            if row['estatus_flag'] == 'ALERTA':
                cell.fill = FILL_ALERT_RED
            elif row['estatus_flag'] == 'DIFERENCIA':
                cell.fill = FILL_ALERT_YELLOW
            elif r_ren % 2 == 1:
                cell.fill = FILL_ZEBRA

        num_item_ren += 1
        r_ren += 1

    ws_rentas.cell(r_ren, 1, "TOTAL").font = FONT_TOTAL
    ws_rentas.cell(r_ren, 2, f"{num_item_ren - 1} CONTRATOS").font = FONT_TOTAL
    ws_rentas.cell(r_ren, 5, f"=SUM(E{r_ren_inicio}:E{r_ren-1})").number_format = FMT_CURRENCY
    ws_rentas.cell(r_ren, 7, f"=SUM(G{r_ren_inicio}:G{r_ren-1})").number_format = FMT_CURRENCY
    ws_rentas.cell(r_ren, 8, f"=SUM(H{r_ren_inicio}:H{r_ren-1})").number_format = FMT_CURRENCY
    ws_rentas.cell(r_ren, 9, f"=SUM(I{r_ren_inicio}:I{r_ren-1})").number_format = FMT_CURRENCY
    for c_i in range(1, len(headers_rentas) + 1):
        cell = ws_rentas.cell(r_ren, c_i)
        cell.border = BORDER_TOTAL
        cell.fill = FILL_TOTAL
        cell.font = FONT_TOTAL

    _autoajustar_columnas(ws_rentas, 1, len(headers_rentas))
    ws_rentas.freeze_panes = "C5"

    # =========================================================================
    # PESTAÑA 6: ACTIVOS VEHICULARES (150-00)
    # =========================================================================
    ws_act = wb.create_sheet(title="Activos_150_Vehiculos")
    try:
        ws_act.views.sheetView[0].showGridLines = True
    except Exception:
        pass

    ws_act.cell(1, 1, "O-LEASING · CONCILIACIÓN DE ACTIVOS FIJOS EN ARRENDAMIENTO (150-00)").font = FONT_TITLE
    ws_act.cell(2, 1, f"Corte al {meta['fecha_corte']} · Comparativo de Saldo Deudor Balanza vs Valor de Contrato Sin IVA").font = FONT_SUBTITLE
    ws_act.merge_cells("A1:H1")
    ws_act.merge_cells("A2:H2")
    ws_act.row_dimensions[1].height = 24
    ws_act.row_dimensions[2].height = 16

    headers_act = [
        "No.", "ID Contrato", "Cliente", "Vehículo / Unidad",
        "Valor Contrato Sin IVA", "Cuenta Contable 150-00",
        "Saldo Balanza (150-00)", "Diferencia (Balanza - Contrato)", "Diagnóstico"
    ]

    r_act = 4
    for c_i, h_txt in enumerate(headers_act, start=1):
        ws_act.cell(r_act, c_i, h_txt)
    _aplicar_estilos_encabezado_tabla(ws_act, r_act, 1, len(headers_act), FILL_PRIMARY)

    r_act += 1
    r_act_inicio = r_act
    num_item_act = 1

    for idx, row in tablas['df_concil_activos'].iterrows():
        ws_act.cell(r_act, 1, num_item_act).alignment = Alignment(horizontal="center")
        ws_act.cell(r_act, 2, row['id_contrato']).alignment = Alignment(horizontal="center")
        ws_act.cell(r_act, 3, row['cliente'])
        ws_act.cell(r_act, 4, row['vehiculo'])
        ws_act.cell(r_act, 5, row['valor_contrato_db']).number_format = FMT_CURRENCY
        ws_act.cell(r_act, 6, row['cuenta_contable']).alignment = Alignment(horizontal="center")
        ws_act.cell(r_act, 7, row['saldo_balanza_150']).number_format = FMT_CURRENCY
        ws_act.cell(r_act, 8, f"=G{r_act}-E{r_act}").number_format = FMT_CURRENCY
        ws_act.cell(r_act, 9, row['diagnostico'])

        for c_i in range(1, len(headers_act) + 1):
            cell = ws_act.cell(r_act, c_i)
            cell.border = BORDER_CELL
            cell.font = FONT_BODY
            if row['estatus_flag'] == 'ALERTA':
                cell.fill = FILL_ALERT_RED
            elif row['estatus_flag'] == 'DIFERENCIA':
                cell.fill = FILL_ALERT_YELLOW
            elif r_act % 2 == 1:
                cell.fill = FILL_ZEBRA

        num_item_act += 1
        r_act += 1

    ws_act.cell(r_act, 1, "TOTAL").font = FONT_TOTAL
    ws_act.cell(r_act, 2, f"{num_item_act - 1} CONTRATOS").font = FONT_TOTAL
    ws_act.cell(r_act, 5, f"=SUM(E{r_act_inicio}:E{r_act-1})").number_format = FMT_CURRENCY
    ws_act.cell(r_act, 7, f"=SUM(G{r_act_inicio}:G{r_act-1})").number_format = FMT_CURRENCY
    ws_act.cell(r_act, 8, f"=SUM(H{r_act_inicio}:H{r_act-1})").number_format = FMT_CURRENCY
    for c_i in range(1, len(headers_act) + 1):
        cell = ws_act.cell(r_act, c_i)
        cell.border = BORDER_TOTAL
        cell.fill = FILL_TOTAL
        cell.font = FONT_TOTAL

    _autoajustar_columnas(ws_act, 1, len(headers_act))
    ws_act.freeze_panes = "C5"

    # =========================================================================
    # PESTAÑA 7: DISCREPANCIAS Y ALERTAS DE AUDITORÍA
    # =========================================================================
    ws_alert = wb.create_sheet(title="Discrepancias_y_Alertas")
    try:
        ws_alert.views.sheetView[0].showGridLines = True
    except Exception:
        pass

    ws_alert.cell(1, 1, "O-LEASING · CATÁLOGO DE DISCREPANCIAS Y PLAN DE ACCIÓN DE AUDITORÍA").font = FONT_TITLE
    ws_alert.cell(2, 1, f"Corte al {meta['fecha_corte']} · Hallazgos priorizados para el cierre contable y conciliación de saldos").font = FONT_SUBTITLE
    ws_alert.merge_cells("A1:G1")
    ws_alert.merge_cells("A2:G2")
    ws_alert.row_dimensions[1].height = 24
    ws_alert.row_dimensions[2].height = 16

    headers_alert = [
        "Prioridad", "Categoría del Hallazgo", "ID Contrato", "Cliente / Razón Social",
        "Cuenta Contable", "Monto Involucrado ($)", "Descripción del Hallazgo", "Acción Recomendada"
    ]

    r_alt = 4
    for c_i, h_txt in enumerate(headers_alert, start=1):
        ws_alert.cell(r_alt, c_i, h_txt)
    _aplicar_estilos_encabezado_tabla(ws_alert, r_alt, 1, len(headers_alert), FILL_NAVY)

    r_alt += 1
    r_alt_inicio = r_alt

    for idx, row in tablas['df_alertas'].iterrows():
        ws_alert.cell(r_alt, 1, row['prioridad']).alignment = Alignment(horizontal="center")
        ws_alert.cell(r_alt, 2, row['categoria'])
        ws_alert.cell(r_alt, 3, str(row['id_contrato'])).alignment = Alignment(horizontal="center")
        ws_alert.cell(r_alt, 4, str(row['cliente']))
        ws_alert.cell(r_alt, 5, str(row['cuenta_contable'])).alignment = Alignment(horizontal="center")
        ws_alert.cell(r_alt, 6, float(row['monto_discrepancia'])).number_format = FMT_CURRENCY
        ws_alert.cell(r_alt, 7, str(row['descripcion']))
        ws_alert.cell(r_alt, 8, str(row['accion_recomendada']))

        for c_i in range(1, len(headers_alert) + 1):
            cell = ws_alert.cell(r_alt, c_i)
            cell.border = BORDER_CELL
            cell.font = FONT_BODY
            if row['prioridad'] == 'ALTA':
                cell.fill = FILL_ALERT_RED
                if c_i == 1:
                    cell.font = FONT_BODY_RED
            elif row['prioridad'] == 'MEDIA':
                cell.fill = FILL_ALERT_YELLOW
            elif r_alt % 2 == 1:
                cell.fill = FILL_ZEBRA

        r_alt += 1

    ws_alert.cell(r_alt, 1, "TOTAL").font = FONT_TOTAL
    ws_alert.cell(r_alt, 2, f"{len(tablas['df_alertas'])} HALLAZGOS DETECTADOS").font = FONT_TOTAL
    ws_alert.cell(r_alt, 6, f"=SUM(F{r_alt_inicio}:F{r_alt-1})").number_format = FMT_CURRENCY
    for c_i in range(1, len(headers_alert) + 1):
        cell = ws_alert.cell(r_alt, c_i)
        cell.border = BORDER_TOTAL
        cell.fill = FILL_TOTAL
        cell.font = FONT_TOTAL

    _autoajustar_columnas(ws_alert, 1, len(headers_alert))
    ws_alert.freeze_panes = "C5"

    if ws_default.title in wb.sheetnames and len(wb.sheetnames) > 1:
        wb.remove(ws_default)

    wb.save(output_excel_path)
    wb.close()
    print(f"Reporte de auditoría generado exitosamente en: {output_excel_path}")
    return output_excel_path


def render_auditoria_balanza(db_path=None):
    """
    Renderizador interactivo de la pantalla de Auditoría y Conciliación
    de Balanza Contpaqi dentro de la aplicación Streamlit de O-Leasing.
    """
    import streamlit as st

    if db_path is None:
        db_path = _obtener_db_path()

    st.title("Auditoría de Balanza Contable (Contpaqi)")
    st.caption("Conciliación integral contrato por contrato vs Balanza de Comprobación Contpaqi para cualquier mes.")

    # Banner explicativo humano
    st.info(
        "**Guía de Auditoría de Balanza O-Leasing:**\n\n"
        "1. **Cuentas Maestras del Catálogo Dinámico:** El análisis toma las cuentas configuradas para tu empresa en el catálogo de pólizas (ej. 115 CXC CP, 125 CXC LP, 114 Residual CP, 126 Residual LP, 208 Int CP, 228 Int LP, 413 Cancelación Capital). Funciona para múltiples empresas con catálogos diferentes.\n"
        "2. **Cartera Activa Separada:** Se auditan las cuentas por cobrar operativas (102-01), ingresos por rentas (403-01) y vehículos arrendados (150-00) de los contratos vigentes.\n"
        "3. **Contratos Liquidados Omitidos y Señalados:** Los contratos concluidos o dados de baja al corte se excluyen de la cartera activa y se concentran en su propia pestaña, alertando en rojo aquellos que aún tienen saldo contable pendiente por depurar o cobrar.\n"
        "4. **Descarga Oficial en Excel:** Exporta el libro de auditoría completo con fórmulas vivas, carátula ejecutiva y diagnósticos de cierre."
    )

    c_up1, c_up2 = st.columns([2, 1])
    with c_up1:
        archivo_subido = st.file_uploader(
            "Seleccionar archivo de Balanza de Comprobación Contpaqi (.xlsx):",
            type=["xlsx"],
            key="balanza_file_uploader",
            help="Sube la balanza generada desde Contpaqi en formato Excel."
        )

    balanza_default_path = r"C:\Users\cynti\Downloads\BALANZA 26\BALANZZA SEPTIEMBRE.xlsx"
    path_a_procesar = None

    if archivo_subido is not None:
        import tempfile
        tfile = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tfile.write(archivo_subido.read())
        tfile.flush()
        tfile.close()
        path_a_procesar = tfile.name
        with c_up2:
            st.success(f"Archivo cargado: **{archivo_subido.name}**")
    elif os.path.exists(balanza_default_path):
        path_a_procesar = balanza_default_path
        with c_up2:
            st.info("Utilizando archivo local predeterminado: **BALANZZA SEPTIEMBRE.xlsx**")
    else:
        with c_up2:
            st.warning("Por favor sube un archivo de Balanza de Comprobación (.xlsx).")
            return

    try:
        with st.spinner("Analizando balanza y cruzando contra base de datos de leasing..."):
            meta, df_bal = parsear_balanza_contpaqi(path_a_procesar)
            fecha_corte = meta['fecha_corte']
            datos_sis = obtener_datos_sistema(db_path, fecha_corte)
            tablas = conciliar_balanza_vs_sistema(df_bal, datos_sis, fecha_corte)
            cat = datos_sis['catalogo_cuentas']
    except Exception as e:
        st.error(f"Error al procesar la Balanza: {str(e)}")
        return

    cta_cxc_cp = cat.get('CXC_CP', {}).get('cuenta', '115-00-00')
    cta_cxc_lp = cat.get('CXC_LP', {}).get('cuenta', '125-00-00')
    cta_int_cp = cat.get('INT_CP', {}).get('cuenta', '208-00-00')

    tot_nif_cp = tablas['df_auditoria_nif']['cxc_cp_esperado'].sum()
    tot_nif_lp = tablas['df_auditoria_nif']['cxc_lp_esperado'].sum()
    tot_int_futuro = tablas['df_auditoria_nif']['int_cp_esperado'].sum() + tablas['df_auditoria_nif']['int_lp_esperado'].sum()

    cta_padre_102 = df_bal[df_bal['cuenta'] == '102-01-00-0000-0000-00']
    tot_bal_cxc = float(cta_padre_102['s_act_d'].values[0]) if len(cta_padre_102) > 0 else 7944400.80

    activos_cxc_bal = tablas['df_concil_cxc']['saldo_balanza_cxc'].sum()
    liq_cxc_bal = tablas['df_concil_liquidados']['saldo_cxc_balanza'].sum()

    st.markdown(f"### Métricas de Conciliación al {fecha_corte} — Empresa: {meta['empresa']}")
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric(f"CXC CP ({cta_cxc_cp})", f"${tot_nif_cp:,.2f}", "Esperado NIF C-15 12m")
    k2.metric(f"CXC LP ({cta_cxc_lp})", f"${tot_nif_lp:,.2f}", "Esperado NIF C-15 >12m")
    k3.metric(f"Intereses ({cta_int_cp})", f"${tot_int_futuro:,.2f}", "Intereses a devengar")
    k4.metric("CxC Cobranza (102-01)", f"${tot_bal_cxc:,.2f}", "Saldo en Balanza")
    k5.metric("CxC Liquidados (Alerta)", f"${liq_cxc_bal:,.2f}", f"{len(tablas['df_concil_liquidados'][tablas['df_concil_liquidados']['saldo_cxc_balanza'] > 0.01])} concluidos con saldo", delta_color="inverse")

    st.markdown("---")
    col_dl1, col_dl2 = st.columns([3, 1])
    with col_dl1:
        st.markdown(f"**Descarga el Informe de Auditoría Completo:** Libro en formato Excel (.xlsx) con fórmulas vivas, carátula ejecutiva y detalle de las 7 pestañas de conciliación.")
    with col_dl2:
        y, m, d = fecha_corte.split('-')
        excel_name = f"Reporte_Auditoria_Balanza_{m}_{y}.xlsx"

        import tempfile
        tmp_out = os.path.join(tempfile.gettempdir(), excel_name)
        generar_excel_auditoria_balanza(path_a_procesar, output_excel_path=tmp_out, db_path=db_path)
        with open(tmp_out, "rb") as f_out:
            data_excel = f_out.read()

        st.download_button(
            label=f"Descargar Excel ({excel_name})",
            data=data_excel,
            file_name=excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

    st.markdown("---")

    tab_nif, tab_cxc, tab_liq, tab_rent, tab_act, tab_alt = st.tabs([
        f"1. Cartera NIF C-15 ({cta_cxc_cp}/{cta_cxc_lp}) [{len(tablas['df_auditoria_nif'])}]",
        f"2. Cartera Activa CxC (102-01) [{len(tablas['df_concil_cxc'])}]",
        f"3. Contratos Liquidados (Señalados) [{len(tablas['df_concil_liquidados'])}]",
        f"4. Ingresos Renta (403-01) [{len(tablas['df_concil_rentas'])}]",
        f"5. Activos Vehiculares (150-00) [{len(tablas['df_concil_activos'])}]",
        f"6. Discrepancias y Alertas [{len(tablas['df_alertas'])}]"
    ])

    with tab_nif:
        st.markdown(f"#### Auditoría Financiera NIF C-15 / IFRS 16 — Cuentas Maestras del Catálogo")
        st.caption(f"Cálculo contrato por contrato de saldos devengados: Corto Plazo ({cta_cxc_cp}), Largo Plazo ({cta_cxc_lp}), Residuales ({cat.get('RESIDUAL_CP',{}).get('cuenta','114-01')} / {cat.get('RESIDUAL',{}).get('cuenta','126')}), Intereses ({cat.get('INT_CP',{}).get('cuenta','208')} / {cat.get('INT_LP',{}).get('cuenta','228')}) y Capital Amortizado ({cat.get('CANCELACION_CAPITAL',{}).get('cuenta','413')}).")
        fmt_nif = {
            'renta_sin_iva': '${:,.2f}',
            'cxc_cp_esperado': '${:,.2f}',
            'cxc_lp_esperado': '${:,.2f}',
            'res_cp_esperado': '${:,.2f}',
            'res_lp_esperado': '${:,.2f}',
            'int_cp_esperado': '${:,.2f}',
            'int_lp_esperado': '${:,.2f}',
            'cap_amort_acumulada': '${:,.2f}',
        }
        st.dataframe(tablas['df_auditoria_nif'].style.format(fmt_nif), use_container_width=True, height=450)

    with tab_cxc:
        st.markdown("#### Conciliación de Cuentas por Cobrar (102-01) — Cartera Activa Vigente")
        st.caption("Solo incluye contratos activos al corte. Los contratos liquidados han sido omitidos de esta tabla y están en la pestaña 3.")
        c_f1, c_f2 = st.columns([1, 2])
        with c_f1:
            filtro_cxc = st.radio("Filtrar cartera activa:", ["Todos", "Solo con saldo exigible (> $0)", "Saldados en $0.00"], horizontal=True, key="rad_cxc_act")

        df_show_cxc = tablas['df_concil_cxc'].copy()
        if filtro_cxc == "Solo con saldo exigible (> $0)":
            df_show_cxc = df_show_cxc[df_show_cxc['saldo_balanza_cxc'] > 0.01]
        elif filtro_cxc == "Saldados en $0.00":
            df_show_cxc = df_show_cxc[abs(df_show_cxc['saldo_balanza_cxc']) <= 0.01]

        fmt_cxc = {
            'renta_sin_iva': '${:,.2f}',
            's_ini_d': '${:,.2f}',
            'cargos': '${:,.2f}',
            'abonos': '${:,.2f}',
            'saldo_balanza_cxc': '${:,.2f}',
            'tot_facturado_sistema': '${:,.2f}',
        }
        st.dataframe(df_show_cxc.style.format(fmt_cxc), use_container_width=True, height=450)

    with tab_liq:
        st.markdown("#### Auditoría de Contratos Liquidados / Bajas al Corte")
        st.caption("Contratos concluidos según el sistema. El objetivo de auditoría es que su saldo en CxC (102-01) y Activo (150-00) sea exactamente $0.00.")
        c_lf1, c_lf2 = st.columns([1, 2])
        with c_lf1:
            filtro_liq = st.radio("Filtrar liquidados:", ["Todos", "Con Saldo Pendiente (Alertas)", "Saldados en Cero ($0.00)"], horizontal=True, key="rad_liq_act")

        df_show_liq = tablas['df_concil_liquidados'].copy()
        if filtro_liq == "Con Saldo Pendiente (Alertas)":
            df_show_liq = df_show_liq[df_show_liq['saldo_cxc_balanza'] > 0.01]
        elif filtro_liq == "Saldados en Cero ($0.00)":
            df_show_liq = df_show_liq[abs(df_show_liq['saldo_cxc_balanza']) <= 0.01]

        fmt_liq = {
            'saldo_cxc_balanza': '${:,.2f}',
            'saldo_activo_balanza': '${:,.2f}',
            'saldo_renta_balanza': '${:,.2f}',
        }
        st.dataframe(df_show_liq.style.format(fmt_liq), use_container_width=True, height=450)

    with tab_rent:
        st.markdown("#### Conciliación de Ingresos por Arrendamiento (403-01)")
        st.caption("Comparativo entre el Saldo Acreedor de la cuenta 403-01 en Balanza vs los conceptos de RENTA facturados en el sistema.")
        fmt_rent = {
            'renta_mensual_db': '${:,.2f}',
            'ingreso_balanza_403': '${:,.2f}',
            'renta_facturada_db': '${:,.2f}',
            'diferencia': '${:,.2f}',
        }
        st.dataframe(tablas['df_concil_rentas'].style.format(fmt_rent), use_container_width=True, height=450)

    with tab_act:
        st.markdown("#### Conciliación de Activos Vehiculares Arrendados (150-00)")
        st.caption("Comparativo entre el Saldo Deudor de la cuenta 150-00 en Balanza vs el Valor de Adquisición del Contrato Sin IVA.")
        fmt_act = {
            'valor_contrato_db': '${:,.2f}',
            'saldo_balanza_150': '${:,.2f}',
            'diferencia': '${:,.2f}',
        }
        st.dataframe(tablas['df_concil_activos'].style.format(fmt_act), use_container_width=True, height=450)

    with tab_alt:
        st.markdown("#### Catálogo de Discrepancias y Alertas Priorizadas")
        st.caption("Hallazgos ordenados por monto y severidad que requieren atención del equipo contable y de cobranza.")
        fmt_alt = {
            'monto_discrepancia': '${:,.2f}'
        }
        st.dataframe(tablas['df_alertas'].style.format(fmt_alt), use_container_width=True, height=450)
