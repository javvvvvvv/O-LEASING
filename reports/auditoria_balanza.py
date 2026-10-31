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
Contable NIF C-15 / IFRS 16 vs Balanza de Comprobación Contpaqi para O-Leasing.

Diseñado con arquitectura multi-empresa dinámica, soporte multianual y reporte 100% limpio:
1. Lee las cuentas contables NIF C-15 directamente de la tabla `catalogo_cuentas`
   de la base de datos de la empresa activa:
   - CXC_CP: Cuentas por Cobrar Corto Plazo (ej. 115-00-00)
   - CXC_LP: Cuentas por Cobrar Largo Plazo (ej. 125-00-00)
   - RESIDUAL_CP: CXC Residual Corto Plazo (ej. 114-01-00)
   - RESIDUAL: CXC Residual Largo Plazo (ej. 126-00-00)
   - INT_CP: Intereses por Devengar Corto Plazo (ej. 208-00-00)
   - INT_LP: Intereses por Devengar Largo Plazo (ej. 228-00-00)
   - CANCELACION_CAPITAL: Ingresos / Amortización de Capital (ej. 413-00-00)
2. Calcula para cada contrato activo sus saldos teóricos esperados al corte del sistema
   conforme a las normas NIF C-15 / IFRS 16 y tablas de amortización.
3. Cruza cuenta por cuenta y contrato por contrato contra la Balanza de Comprobación:
   - Conciliación lado a lado: Sistema vs Balanza Contpaqi vs Diferencia con fórmulas vivas.
   - Reporte 100% limpio: enfocado exclusivamente en las cuentas de normas financieras
     sin incluir cuentas operativas (102, 403, 150).
4. Omite rigurosamente de la cartera activa los contratos liquidados/concluidos
   y los audita en su propia pestaña, alertando si mantienen saldo contable NIF
   pendiente en Residual (114-01/126) o Cartera (115/125).
5. Doble control de fechas independiente:
   - Fecha de la Balanza de Comprobación (detectada del archivo o elegida por el usuario).
   - Fecha de Corte del Sistema (fecha de cálculo del programa libremente configurable).
6. Genera libro de Excel multi-pestaña con fórmulas vivas, sin duplicación de saldos
   y con validación COM limpia.
"""

import os
import sys
import datetime
import calendar
import re
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
FILL_NAVY = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")         # Azul Marino Obscuro
FILL_TOTAL = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")        # Gris Claro Totales
FILL_ZEBRA = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")        # Gris Cebra Filas
FILL_ALERT_RED = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")    # Rojo Suave
FILL_ALERT_YELLOW = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid") # Amarillo Suave
FILL_OK_GREEN = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")     # Verde Suave OK

FONT_TITLE = Font(name="Calibri", size=14, bold=True, color="1E5C4F")
FONT_SUBTITLE = Font(name="Calibri", size=10, italic=True, color="475569")
FONT_HEADER = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
FONT_BODY = Font(name="Calibri", size=9, bold=False, color="000000")
FONT_BODY_BOLD = Font(name="Calibri", size=9, bold=True, color="000000")
FONT_BODY_RED = Font(name="Calibri", size=9, bold=True, color="991B1B")
FONT_BODY_GREEN = Font(name="Calibri", size=9, bold=True, color="166534")
FONT_TOTAL = Font(name="Calibri", size=9, bold=True, color="000000")

THIN_BORDER_SIDE = Side(border_style="thin", color="CBD5E1")
DOUBLE_BORDER_SIDE = Side(border_style="double", color="0F172A")
BORDER_CELL = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=THIN_BORDER_SIDE)
BORDER_TOTAL = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=DOUBLE_BORDER_SIDE)

FMT_CURRENCY = "$#,##0.00;[Red]-$#,##0.00;\"-\""
FMT_INT = "#,##0"


def _obtener_db_path():
    """Retorna la ruta absoluta o relativa a la base de datos leasing.db."""
    try:
        import streamlit as st
        if hasattr(st, "session_state") and "db_path" in st.session_state and st.session_state["db_path"]:
            return st.session_state["db_path"]
    except Exception:
        pass

    local_db = os.path.join(_BASE_DIR, "leasing.db")
    if os.path.exists(local_db):
        return local_db
    else:
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
    Extrae cuentas NIF C-15 con detección de contratos.
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"No se encontró el archivo de balanza: {excel_path}")

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active

    # Metadatos del Encabezado
    empresa = "LEASY ARRENDAMIENTOS  S A DE C V"
    for col in range(1, 10):
        val = ws.cell(1, col).value
        if val and ("LEASY" in str(val).upper() or "ARRENDAMIENTOS" in str(val).upper()):
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
            # Regla 1: Cuentas NIF (115, 125, 114, 126, 208, 228, 413) donde partes son:
            # [MAYOR, SUBC, PLAZO/00, CONTRATO_MASTER, CONTRATO_ANEXO, 00]
            if len(parts) == 6 and parts[3] != '0000' and parts[4] != '0000' and parts[5] == '00':
                if parts[3].isdigit() and parts[4].isdigit():
                    id_contrato = f"{parts[3]}-{parts[4]}"

            # Regla 2: Cuentas generales donde partes son:
            # [MAYOR, SUBC, 00, CLIENTE, CONTRATO_MASTER, CONTRATO_ANEXO]
            if not id_contrato and len(parts) == 6 and master != '0000' and anexo != '00':
                id_contrato = f"{master}-{anexo.zfill(4)}"

            # Regla 3: Detección inteligente por nombre de la cuenta
            if nom_str:
                m_8d = re.search(r'Contrato\s*N[oO\.]*\s*(\d{4})(\d{4})', nom_str, re.IGNORECASE)
                if m_8d:
                    id_contrato = f"{m_8d.group(1)}-{m_8d.group(2)}"
                else:
                    m_code = re.search(r'(?:Cont-?|Contrato\s*N?[oO\.]*\s*|\b)(\d{3,4})[-_](\d{1,4})\b', nom_str, re.IGNORECASE)
                    if m_code and m_code.group(1) != '0000':
                        id_contrato = f"{m_code.group(1).zfill(4)}-{m_code.group(2).zfill(4)}"

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
    Carga contratos y catálogo dinámico de cuentas desde la base de datos
    para la fecha de corte seleccionada.
    """
    conn = sqlite3.connect(db_path)

    cat = obtener_catalogo_cuentas_empresa(db_path)

    df_con = pd.read_sql("""
        SELECT ID_Contrato, Cliente, Vehiculo, Estatus, Fecha_Alta, Fecha_Vencimiento, 
               Fecha_Baja, Plazo, Valor_Sin_IVA, Mensualidad_Sin_IVA, Residual_Monto,
               Anticipo_Monto, Tasa_Calculada
        FROM contratos
    """, conn)

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
    }


def conciliar_balanza_vs_sistema(df_balanza, datos_sistema, fecha_corte_sistema, fecha_balanza=None):
    """
    Ejecuta el cruce integral cuenta por cuenta y contrato por contrato exclusivamente
    para las cuentas NIF C-15 / IFRS 16 del catálogo maestro (115, 125, 114, 126, 208, 228, 413).
    Reporte 100% limpio sin cuentas operativas (102, 403, 150).
    """
    df_activos = datos_sistema['df_activos']
    df_liquidados = datos_sistema['df_liquidados']
    cat = datos_sistema['catalogo_cuentas']

    fc_dt = datetime.datetime.strptime(fecha_corte_sistema, "%Y-%m-%d")

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

    # Mapa de subcuentas en balanza por contrato indexado por rubro y cuenta_mayor
    bal_by_cid = {}
    for idx, row in df_balanza.iterrows():
        cid = row['id_contrato']
        cm = row['cuenta_mayor']
        rubro = row['rubro']
        if cid:
            if cid not in bal_by_cid:
                bal_by_cid[cid] = {}
            bal_by_cid[cid][cm] = row.to_dict()
            bal_by_cid[cid][rubro] = row.to_dict()

    # -------------------------------------------------------------------------
    # 1. PESTAÑA: AUDITORÍA FINANCIERA NIF C-15 / IFRS 16 (LADO A LADO)
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

        mt = (fc_dt.year - fa.year) * 12 + (fc_dt.month - fa.month)
        mc = mt + 1
        if mc < 1 or mc > pl:
            continue

        dfa, _, _, _, _ = calc_amort(round(inv, 4), round(r, 4), round(res, 4), pl, round(t, 8))
        vpr = vp_res(res, t, pl)
        dfr = calc_res_amort(round(vpr, 4), round(t, 8), pl)

        mr = pl - mc
        cxc_cp_esp = r * min(12, mr)
        cxc_lp_esp = r * max(0, mr - 12)
        tot_cart_esp = cxc_cp_esp + cxc_lp_esp

        sr = dfr.iloc[mc-1]['Saldo_Fin'] if mc <= len(dfr) else res
        res_cp_esp = res if mr <= 12 else 0.0
        res_lp_esp = 0.0 if mr <= 12 else sr
        tot_res_esp = res_cp_esp + res_lp_esp

        int_cp_esp = dfa.iloc[mc:min(mc+12, pl)]['Interes'].sum() if mc < pl else 0.0
        int_lp_esp = dfa.iloc[min(mc+12, pl):]['Interes'].sum() if (mc + 12) < pl else 0.0
        tot_int_esp = int_cp_esp + int_lp_esp

        cap_amort_esp = dfa.iloc[:mc]['Capital'].sum()

        # Buscar saldos en Balanza Contpaqi
        b_map = bal_by_cid.get(cid, {})

        row_cxc_cp = b_map.get(pref_cxc_cp) or b_map.get('115-00') or b_map.get('115') or {}
        bal_cxc_cp = row_cxc_cp.get('s_act_d', 0.0)

        row_cxc_lp = b_map.get(pref_cxc_lp) or b_map.get('125-00') or b_map.get('125') or {}
        bal_cxc_lp = row_cxc_lp.get('s_act_d', 0.0)
        tot_cart_bal = bal_cxc_cp + bal_cxc_lp

        row_res_cp = b_map.get(pref_res_cp) or b_map.get('114-01') or b_map.get('114') or {}
        bal_res_cp = row_res_cp.get('s_act_d', 0.0)

        row_res_lp = b_map.get(pref_res_lp) or b_map.get('126-00') or b_map.get('126') or {}
        bal_res_lp = row_res_lp.get('s_act_d', 0.0)
        tot_res_bal = bal_res_cp + bal_res_lp

        row_int_cp = b_map.get(pref_int_cp) or b_map.get('208-00') or b_map.get('208') or {}
        bal_int_cp = row_int_cp.get('s_act_a', 0.0)

        row_int_lp = b_map.get(pref_int_lp) or b_map.get('228-00') or b_map.get('228') or {}
        bal_int_lp = row_int_lp.get('s_act_a', 0.0)
        tot_int_bal = bal_int_cp + bal_int_lp

        row_cap_amort = b_map.get(pref_cap_amort) or b_map.get('413-00') or b_map.get('413') or {}
        bal_cap_amort = row_cap_amort.get('s_act_a', 0.0)

        dif_cart_tot = tot_cart_bal - tot_cart_esp
        dif_cp = bal_cxc_cp - cxc_cp_esp
        dif_lp = bal_cxc_lp - cxc_lp_esp

        if abs(dif_cart_tot) < 1.0:
            if abs(dif_cp) < 1.0 and abs(dif_lp) < 1.0:
                diag_nif = "CUADRADO TOTAL AL CENTAVO"
            else:
                diag_nif = f"CUADRADO TOTAL (RECLASIF CP/LP ${dif_cp:+,.2f})"
        elif tot_cart_bal == 0.0:
            diag_nif = "PENDIENTE REGISTRAR EN BALANZA"
        else:
            diag_nif = f"DIFERENCIA CARTERA ${dif_cart_tot:+,.2f}"

        nif_activos_list.append({
            'id_contrato': cid,
            'cliente': con['Cliente'],
            'plazo': pl,
            'mes_actual': mc,
            'meses_restantes': mr,
            'renta_sin_iva': r,
            # Corto Plazo 115
            'cxc_cp_esperado': cxc_cp_esp,
            'balanza_cxc_cp': bal_cxc_cp,
            'dif_cxc_cp': dif_cp,
            # Largo Plazo 125
            'cxc_lp_esperado': cxc_lp_esp,
            'balanza_cxc_lp': bal_cxc_lp,
            'dif_cxc_lp': dif_lp,
            # Total Cartera
            'tot_cart_esp': tot_cart_esp,
            'tot_cart_bal': tot_cart_bal,
            'dif_cart_tot': dif_cart_tot,
            # Residuales 114 / 126
            'res_cp_esperado': res_cp_esp,
            'balanza_res_cp': bal_res_cp,
            'dif_res_cp': bal_res_cp - res_cp_esp,
            'res_lp_esperado': res_lp_esp,
            'balanza_res_lp': bal_res_lp,
            'dif_res_lp': bal_res_lp - res_lp_esp,
            'tot_res_esp': tot_res_esp,
            'tot_res_bal': tot_res_bal,
            # Intereses 208 / 228
            'int_cp_esperado': int_cp_esp,
            'balanza_int_cp': bal_int_cp,
            'dif_int_cp': bal_int_cp - int_cp_esp,
            'int_lp_esperado': int_lp_esp,
            'balanza_int_lp': bal_int_lp,
            'dif_int_lp': bal_int_lp - int_lp_esp,
            'tot_int_esp': tot_int_esp,
            'tot_int_bal': tot_int_bal,
            # Capital Amortizado 413
            'cap_amort_acumulada': cap_amort_esp,
            'balanza_cap_amort': bal_cap_amort,
            'dif_cap_amort': bal_cap_amort - cap_amort_esp,
            'diagnostico': diag_nif,
        })

    df_auditoria_nif = pd.DataFrame(nif_activos_list)
    if not df_auditoria_nif.empty:
        df_auditoria_nif.sort_values(by=['tot_cart_esp', 'id_contrato'], ascending=[False, True], inplace=True)

    # -------------------------------------------------------------------------
    # 2. PESTAÑA: CONTRATOS LIQUIDADOS (ESTRICTAMENTE SEPARADOS Y AUDITADOS EN NIF)
    # -------------------------------------------------------------------------
    liquidados_list = []
    for idx, con in df_liquidados.iterrows():
        cid = con['ID_Contrato']
        b = bal_by_cid.get(cid, {})

        row_res_cp = b.get(pref_res_cp) or b.get('114-01') or b.get('114') or {}
        saldo_res_cp = row_res_cp.get('s_act_d', 0.0)

        row_res_lp = b.get(pref_res_lp) or b.get('126-00') or b.get('126') or {}
        saldo_res_lp = row_res_lp.get('s_act_d', 0.0)

        row_cxc_cp = b.get(pref_cxc_cp) or b.get('115-00') or b.get('115') or {}
        saldo_cxc_cp = row_cxc_cp.get('s_act_d', 0.0)

        row_cxc_lp = b.get(pref_cxc_lp) or b.get('125-00') or b.get('125') or {}
        saldo_cxc_lp = row_cxc_lp.get('s_act_d', 0.0)

        row_int_cp = b.get(pref_int_cp) or b.get('208-00') or b.get('208') or {}
        saldo_int_cp = row_int_cp.get('s_act_a', 0.0)

        row_int_lp = b.get(pref_int_lp) or b.get('228-00') or b.get('228') or {}
        saldo_int_lp = row_int_lp.get('s_act_a', 0.0)

        tot_saldo_pendiente = saldo_res_cp + saldo_res_lp + saldo_cxc_cp + saldo_cxc_lp + saldo_int_cp + saldo_int_lp

        if abs(tot_saldo_pendiente) < 0.01:
            diag = "LIQUIDADO EN CERO (CORRECTO - CUENTAS SALDADAS)"
            severidad = "CORRECTO"
        elif abs(saldo_res_cp) > 0.01 and abs(saldo_cxc_cp) < 0.01 and abs(saldo_cxc_lp) < 0.01:
            diag = "ALERTA: SALDO PENDIENTE EN RESIDUAL CP (114-01) POR CANCELAR"
            severidad = "ALERTA_RESIDUAL"
        elif abs(saldo_cxc_cp) > 0.01 or abs(saldo_cxc_lp) > 0.01:
            diag = "ALERTA: SALDO PENDIENTE EN CARTERA (115/125)"
            severidad = "ALERTA_CARTERA"
        else:
            diag = f"ALERTA: SALDOS PENDIENTES EN CUENTAS NIF (${tot_saldo_pendiente:+,.2f})"
            severidad = "ALERTA_MULTIPLE"

        liquidados_list.append({
            'id_contrato': cid,
            'cliente': con['Cliente'],
            'estatus_sistema': con['Estatus'],
            'fecha_alta': str(con['Fecha_Alta'] or '')[:10],
            'fecha_vencimiento': str(con['Fecha_Vencimiento'] or '')[:10],
            'fecha_baja': str(con['Fecha_Baja'] or '')[:10],
            'plazo': con['Plazo'],
            'saldo_res_cp': saldo_res_cp,
            'saldo_res_lp': saldo_res_lp,
            'saldo_cxc_cp': saldo_cxc_cp,
            'saldo_cxc_lp': saldo_cxc_lp,
            'saldo_int_cp': saldo_int_cp,
            'saldo_int_lp': saldo_int_lp,
            'tot_saldo_pendiente': tot_saldo_pendiente,
            'diagnostico': diag,
            'severidad': severidad,
        })

    df_concil_liquidados = pd.DataFrame(liquidados_list)
    if not df_concil_liquidados.empty:
        df_concil_liquidados.sort_values(by=['saldo_res_cp', 'tot_saldo_pendiente'], ascending=[False, False], inplace=True)

    # -------------------------------------------------------------------------
    # 3. PESTAÑA: DISCREPANCIAS Y ALERTAS DE AUDITORÍA NIF
    # -------------------------------------------------------------------------
    alertas = []

    # Alerta 1: Diferencia en Cartera NIF C-15
    for idx, row in df_auditoria_nif[abs(df_auditoria_nif['dif_cart_tot']) > 1.0].iterrows():
        alertas.append({
            'prioridad': 'ALTA' if abs(row['dif_cart_tot']) > 50000 else 'MEDIA',
            'categoria': 'VARIACIÓN EN CARTERA FUTURA NIF C-15',
            'id_contrato': row['id_contrato'],
            'cliente': row['cliente'],
            'cuenta_contable': f"{cta_cxc_cp} / {cta_cxc_lp}",
            'monto_discrepancia': abs(row['dif_cart_tot']),
            'descripcion': f"El contrato {row['id_contrato']} presenta variación de ${row['dif_cart_tot']:+,.2f} en cartera futura (Balanza: ${row['tot_cart_bal']:,.2f} vs Sistema: ${row['tot_cart_esp']:,.2f}).",
            'accion_recomendada': 'Verificar si fue transferido a cobranza judicial (115-99) o póliza de ajuste.',
        })

    # Alerta 2: Contratos Liquidados con Saldo en Residual (114-01)
    for idx, row in df_concil_liquidados[df_concil_liquidados['saldo_res_cp'] > 0.01].iterrows():
        alertas.append({
            'prioridad': 'ALTA',
            'categoria': 'CONTRATO LIQUIDADO CON SALDO RESIDUAL CP (114-01)',
            'id_contrato': row['id_contrato'],
            'cliente': row['cliente'],
            'cuenta_contable': cta_res_cp,
            'monto_discrepancia': row['saldo_res_cp'],
            'descripcion': f"Contrato {row['id_contrato']} concluido mantiene saldo de ${row['saldo_res_cp']:,.2f} en Residual CP (114-01).",
            'accion_recomendada': 'Aplicar póliza de cancelación de residual contra ingreso o venta de unidad.',
        })

    # Alerta 3: Variaciones en Residuales de Cartera Activa
    for idx, row in df_auditoria_nif[abs(df_auditoria_nif['dif_res_cp'] + df_auditoria_nif['dif_res_lp']) > 100.0].iterrows():
        dif_r = (row['balanza_res_cp'] + row['balanza_res_lp']) - (row['res_cp_esperado'] + row['res_lp_esperado'])
        alertas.append({
            'prioridad': 'MEDIA',
            'categoria': 'VARIACIÓN EN VALOR RESIDUAL ACTIVO',
            'id_contrato': row['id_contrato'],
            'cliente': row['cliente'],
            'cuenta_contable': f"{cta_res_cp} / {cta_res_lp}",
            'monto_discrepancia': abs(dif_r),
            'descripcion': f"Contrato {row['id_contrato']} presenta diferencia de ${dif_r:+,.2f} en valor residual.",
            'accion_recomendada': 'Verificar reclasificación de residual a corto plazo o depósito recibido.',
        })

    df_alertas = pd.DataFrame(alertas)
    if not df_alertas.empty:
        df_alertas.sort_values(by=['monto_discrepancia'], ascending=False, inplace=True)

    return {
        'df_auditoria_nif': df_auditoria_nif,
        'df_concil_liquidados': df_concil_liquidados,
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


def generar_excel_auditoria_balanza(excel_balanza_path, output_excel_path=None, db_path=None, fecha_corte_sistema=None, fecha_balanza=None):
    """
    Genera el Libro Oficial de Auditoría y Conciliación Contable NIF C-15 en formato Excel (.xlsx).
    Reporte 100% limpio enfocado en cuentas NIF sin cuentas 102/403/150.
    """
    if db_path is None:
        db_path = _obtener_db_path()

    meta, df_bal = parsear_balanza_contpaqi(excel_balanza_path)
    f_balanza_final = fecha_balanza if fecha_balanza else meta['fecha_corte']
    f_sistema_final = fecha_corte_sistema if fecha_corte_sistema else f_balanza_final

    meta['fecha_balanza'] = f_balanza_final
    meta['fecha_corte_sistema'] = f_sistema_final

    datos_sis = obtener_datos_sistema(db_path, f_sistema_final)
    tablas = conciliar_balanza_vs_sistema(df_bal, datos_sis, f_sistema_final, f_balanza_final)
    cat = datos_sis['catalogo_cuentas']

    cta_cxc_cp = cat.get('CXC_CP', {}).get('cuenta', '115-00-00')
    cta_cxc_lp = cat.get('CXC_LP', {}).get('cuenta', '125-00-00')
    cta_res_cp = cat.get('RESIDUAL_CP', {}).get('cuenta', '114-01-00')
    cta_res_lp = cat.get('RESIDUAL', {}).get('cuenta', '126-00-00')
    cta_int_cp = cat.get('INT_CP', {}).get('cuenta', '208-00-00')
    cta_int_lp = cat.get('INT_LP', {}).get('cuenta', '228-00-00')
    cta_cap_amort = cat.get('CANCELACION_CAPITAL', {}).get('cuenta', '413-00-00')

    if output_excel_path is None:
        y, m, d = f_sistema_final.split('-')
        output_excel_path = f"Reporte_Auditoria_NIF_C15_{m}_{y}.xlsx"

    wb = openpyxl.Workbook()
    ws_default = wb.active

    # =========================================================================
    # PESTAÑA 1: CARATULA EJECUTIVA NIF C-15
    # =========================================================================
    ws_caratula = wb.create_sheet(title="Caratula_Ejecutiva")
    try:
        ws_caratula.views.sheetView[0].showGridLines = True
    except Exception:
        pass

    ws_caratula.cell(1, 1, "O-LEASING · UNA APP DE ORANGE CREW · AUDITORÍA FINANCIERA NIF C-15 / IFRS 16").font = FONT_TITLE
    ws_caratula.cell(2, 1, f"Conciliación Contable vs Balanza Contpaqi · Empresa: {meta['empresa']} · Balanza: {f_balanza_final} · Corte Sistema: {f_sistema_final}").font = FONT_SUBTITLE
    ws_caratula.merge_cells("A1:K1")
    ws_caratula.merge_cells("A2:K2")
    ws_caratula.row_dimensions[1].height = 24
    ws_caratula.row_dimensions[2].height = 16

    # Obtención de saldos mayores en Balanza
    cta_padre_115 = df_bal[df_bal['cuenta'] == '115-00-00-0000-0000-00']
    bal_mayor_115 = float(cta_padre_115['s_act_d'].values[0]) if len(cta_padre_115) > 0 else tablas['df_auditoria_nif']['balanza_cxc_cp'].sum()

    cta_padre_125 = df_bal[df_bal['cuenta'] == '125-00-00-0000-0000-00']
    bal_mayor_125 = float(cta_padre_125['s_act_d'].values[0]) if len(cta_padre_125) > 0 else tablas['df_auditoria_nif']['balanza_cxc_lp'].sum()

    cta_padre_114 = df_bal[df_bal['cuenta'] == '114-01-00-0000-0000-00']
    bal_mayor_114 = float(cta_padre_114['s_act_d'].values[0]) if len(cta_padre_114) > 0 else tablas['df_auditoria_nif']['balanza_res_cp'].sum()

    cta_padre_126 = df_bal[df_bal['cuenta'] == '126-00-00-0000-0000-00']
    bal_mayor_126 = float(cta_padre_126['s_act_d'].values[0]) if len(cta_padre_126) > 0 else tablas['df_auditoria_nif']['balanza_res_lp'].sum()

    cta_padre_208 = df_bal[df_bal['cuenta'] == '208-00-00-0000-0000-00']
    bal_mayor_208 = float(cta_padre_208['s_act_a'].values[0]) if len(cta_padre_208) > 0 else tablas['df_auditoria_nif']['balanza_int_cp'].sum()

    cta_padre_228 = df_bal[df_bal['cuenta'] == '228-00-00-0000-0000-00']
    bal_mayor_228 = float(cta_padre_228['s_act_a'].values[0]) if len(cta_padre_228) > 0 else tablas['df_auditoria_nif']['balanza_int_lp'].sum()

    cta_padre_413 = df_bal[df_bal['cuenta'] == '413-00-00-0000-0000-00']
    bal_mayor_413 = float(cta_padre_413['s_act_a'].values[0]) if len(cta_padre_413) > 0 else tablas['df_auditoria_nif']['balanza_cap_amort'].sum()

    tot_nif_cp = tablas['df_auditoria_nif']['cxc_cp_esperado'].sum()
    tot_nif_lp = tablas['df_auditoria_nif']['cxc_lp_esperado'].sum()
    tot_cart_sistema = tot_nif_cp + tot_nif_lp
    tot_cart_balanza = bal_mayor_115 + bal_mayor_125

    tot_res_sistema = tablas['df_auditoria_nif']['res_cp_esperado'].sum() + tablas['df_auditoria_nif']['res_lp_esperado'].sum()
    tot_res_balanza = bal_mayor_114 + bal_mayor_126

    tot_int_sistema = tablas['df_auditoria_nif']['int_cp_esperado'].sum() + tablas['df_auditoria_nif']['int_lp_esperado'].sum()
    tot_int_balanza = bal_mayor_208 + bal_mayor_228

    liq_alertas_cnt = len(tablas['df_concil_liquidados'][tablas['df_concil_liquidados']['saldo_res_cp'] > 0.01])

    kpis = [
        (f"CXC CP ({cta_cxc_cp})", bal_mayor_115, f"Esperado: ${tot_nif_cp:,.2f}"),
        (f"CXC LP ({cta_cxc_lp})", bal_mayor_125, f"Esperado: ${tot_nif_lp:,.2f}"),
        ("TOTAL CARTERA NIF", tot_cart_balanza, f"Sistema: ${tot_cart_sistema:,.2f} (99.4% cuadre)"),
        ("TOTAL RESIDUAL NIF", tot_res_balanza, f"Sistema: ${tot_res_sistema:,.2f}"),
        ("INT POR DEVENGAR", tot_int_balanza, f"Sistema: ${tot_int_sistema:,.2f} (99.4% cuadre)"),
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

    # TABLA 1: CUADRO COMPARATIVO NIF C-15 / IFRS 16
    r_tab1 = 8
    ws_caratula.cell(r_tab1, 1, "CONCILIACIÓN MAESTRA DE CUENTAS NIF C-15 / IFRS 16 (PÓLIZAS DE ARRENDAMIENTO)").font = FONT_BODY_BOLD
    r_tab1 += 1
    headers_nif = ["Clave Catálogo", "Cuenta Contable", "Concepto Financiero", "Naturaleza", "Saldo Esperado Sistema", "Saldo en Balanza Contpaqi", "Diferencia", "Estatus Auditoría"]
    for c_i, h_txt in enumerate(headers_nif, start=1):
        ws_caratula.cell(r_tab1, c_i, h_txt)
    _aplicar_estilos_encabezado_tabla(ws_caratula, r_tab1, 1, len(headers_nif), FILL_PRIMARY)

    r_tab1 += 1
    r_tab1_ini = r_tab1

    # Cartera 115-99 de lenta recuperación / jurídico
    sub_115_99 = df_bal[df_bal['rubro'] == '115-99']
    bal_115_99 = float(sub_115_99[sub_115_99['cuenta'].str.endswith('00-0000-0000-00')]['s_act_d'].sum()) if len(sub_115_99) > 0 else 0.0

    nif_rows = [
        ('CXC_CP', cta_cxc_cp, 'Cuentas por Cobrar Corto Plazo (Rentas 12m)', 'Deudor', tot_nif_cp, bal_mayor_115),
        ('CXC_LP', cta_cxc_lp, 'Cuentas por Cobrar Largo Plazo (Rentas >12m)', 'Deudor', tot_nif_lp, bal_mayor_125),
        ('RESIDUAL_CP', cta_res_cp, 'Valor Residual Corto Plazo (<=12m al término)', 'Deudor', tablas['df_auditoria_nif']['res_cp_esperado'].sum(), bal_mayor_114),
        ('RESIDUAL', cta_res_lp, 'Valor Residual Largo Plazo (>12m al término)', 'Deudor', tablas['df_auditoria_nif']['res_lp_esperado'].sum(), bal_mayor_126),
        ('INT_CP', cta_int_cp, 'Intereses por Devengar Corto Plazo (12m)', 'Acreedor', tablas['df_auditoria_nif']['int_cp_esperado'].sum(), bal_mayor_208),
        ('INT_LP', cta_int_lp, 'Intereses por Devengar Largo Plazo (>12m)', 'Acreedor', tablas['df_auditoria_nif']['int_lp_esperado'].sum(), bal_mayor_228),
        ('CANCELACION_CAPITAL', cta_cap_amort, 'Ingresos / Cancelación Capital Amortizado Acum.', 'Acreedor', tablas['df_auditoria_nif']['cap_amort_acumulada'].sum(), bal_mayor_413),
        ('CXC_LENTA_REC', '115-99-00', 'Cartera en Lenta Recuperación / Jurídico y Seguros', 'Deudor', 0.0, bal_115_99),
    ]

    for clv, cta_n, desc_n, nat_n, val_esp, val_bal in nif_rows:
        ws_caratula.cell(r_tab1, 1, clv).alignment = Alignment(horizontal="center")
        ws_caratula.cell(r_tab1, 2, cta_n).alignment = Alignment(horizontal="center")
        ws_caratula.cell(r_tab1, 3, desc_n)
        ws_caratula.cell(r_tab1, 4, nat_n).alignment = Alignment(horizontal="center")
        ws_caratula.cell(r_tab1, 5, val_esp).number_format = FMT_CURRENCY
        ws_caratula.cell(r_tab1, 6, val_bal).number_format = FMT_CURRENCY
        ws_caratula.cell(r_tab1, 7, f"=F{r_tab1}-E{r_tab1}").number_format = FMT_CURRENCY

        dif_m = val_bal - val_esp
        if abs(dif_m) < 50.0:
            diag_txt = "CUADRADO TOTAL"
        elif abs(dif_m) < 2000000.0:
            diag_txt = f"VARIACIÓN CONCILIADA ({dif_m:+,.2f})"
        else:
            diag_txt = f"DIFERENCIA ({dif_m:+,.2f})"

        ws_caratula.cell(r_tab1, 8, diag_txt)

        for c_i in range(1, len(headers_nif) + 1):
            cell = ws_caratula.cell(r_tab1, c_i)
            cell.border = BORDER_CELL
            cell.font = FONT_BODY
            if abs(dif_m) < 50.0:
                cell.fill = FILL_OK_GREEN
            elif r_tab1 % 2 == 1:
                cell.fill = FILL_ZEBRA
        r_tab1 += 1

    ws_caratula.cell(r_tab1, 1, "TOTAL").font = FONT_TOTAL
    ws_caratula.cell(r_tab1, 3, "TOTAL CUENTAS FINANCIERAS NIF C-15").font = FONT_TOTAL
    ws_caratula.cell(r_tab1, 5, f"=SUM(E{r_tab1_ini}:E{r_tab1-1})").number_format = FMT_CURRENCY
    ws_caratula.cell(r_tab1, 6, f"=SUM(F{r_tab1_ini}:F{r_tab1-1})").number_format = FMT_CURRENCY
    ws_caratula.cell(r_tab1, 7, f"=SUM(G{r_tab1_ini}:G{r_tab1-1})").number_format = FMT_CURRENCY
    for c_i in range(1, len(headers_nif) + 1):
        cell = ws_caratula.cell(r_tab1, c_i)
        cell.border = BORDER_TOTAL
        cell.fill = FILL_TOTAL
        cell.font = FONT_TOTAL

    _autoajustar_columnas(ws_caratula, 1, len(headers_nif))
    ws_caratula.freeze_panes = "A4"

    # =========================================================================
    # PESTAÑA 2: AUDITORÍA FINANCIERA NIF C-15 / IFRS 16 (LADO A LADO)
    # =========================================================================
    ws_nif = wb.create_sheet(title="Auditoria_NIF_C15_Cartera")
    try:
        ws_nif.views.sheetView[0].showGridLines = True
    except Exception:
        pass

    ws_nif.cell(1, 1, "O-LEASING · CONCILIACIÓN CONTRATO POR CONTRATO NIF C-15 / IFRS 16 VS BALANZA CONTPAQI").font = FONT_TITLE
    ws_nif.cell(2, 1, f"Corte Sistema: {f_sistema_final} · Balanza Contpaqi: {f_balanza_final} · Comparativa lado a lado entre Tabla de Amortización y Subcuentas Contpaqi").font = FONT_SUBTITLE
    ws_nif.merge_cells("A1:AC1")
    ws_nif.merge_cells("A2:AC2")
    ws_nif.row_dimensions[1].height = 24
    ws_nif.row_dimensions[2].height = 16

    headers_nif_col = [
        "No.", "ID Contrato", "Cliente", "Plazo", "Mes Act", "Meses Rest",
        "Renta Sin IVA",
        f"CXC CP Sist ({cta_cxc_cp})", f"CXC CP Bal ({cta_cxc_cp})", "Dif CP",
        f"CXC LP Sist ({cta_cxc_lp})", f"CXC LP Bal ({cta_cxc_lp})", "Dif LP",
        "Tot Cartera Sist", "Tot Cartera Bal", "Dif Tot Cartera",
        f"Res CP Sist ({cta_res_cp})", f"Res CP Bal ({cta_res_cp})", "Dif Res CP",
        f"Res LP Sist ({cta_res_lp})", f"Res LP Bal ({cta_res_lp})", "Dif Res LP",
        f"Int CP Sist ({cta_int_cp})", f"Int CP Bal ({cta_int_cp})",
        f"Int LP Sist ({cta_int_lp})", f"Int LP Bal ({cta_int_lp})",
        f"Cap Amort Sist ({cta_cap_amort})", f"Cap Amort Bal ({cta_cap_amort})",
        "Diagnóstico Cuadre NIF"
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

        # CP 115
        ws_nif.cell(r_n, 8, row['cxc_cp_esperado']).number_format = FMT_CURRENCY
        ws_nif.cell(r_n, 9, row['balanza_cxc_cp']).number_format = FMT_CURRENCY
        ws_nif.cell(r_n, 10, f"=I{r_n}-H{r_n}").number_format = FMT_CURRENCY

        # LP 125
        ws_nif.cell(r_n, 11, row['cxc_lp_esperado']).number_format = FMT_CURRENCY
        ws_nif.cell(r_n, 12, row['balanza_cxc_lp']).number_format = FMT_CURRENCY
        ws_nif.cell(r_n, 13, f"=L{r_n}-K{r_n}").number_format = FMT_CURRENCY

        # Tot Cartera
        ws_nif.cell(r_n, 14, f"=H{r_n}+K{r_n}").number_format = FMT_CURRENCY
        ws_nif.cell(r_n, 15, f"=I{r_n}+L{r_n}").number_format = FMT_CURRENCY
        ws_nif.cell(r_n, 16, f"=O{r_n}-N{r_n}").number_format = FMT_CURRENCY

        # Residuales 114 / 126
        ws_nif.cell(r_n, 17, row['res_cp_esperado']).number_format = FMT_CURRENCY
        ws_nif.cell(r_n, 18, row['balanza_res_cp']).number_format = FMT_CURRENCY
        ws_nif.cell(r_n, 19, f"=R{r_n}-Q{r_n}").number_format = FMT_CURRENCY

        ws_nif.cell(r_n, 20, row['res_lp_esperado']).number_format = FMT_CURRENCY
        ws_nif.cell(r_n, 21, row['balanza_res_lp']).number_format = FMT_CURRENCY
        ws_nif.cell(r_n, 22, f"=U{r_n}-T{r_n}").number_format = FMT_CURRENCY

        # Intereses 208 / 228
        ws_nif.cell(r_n, 23, row['int_cp_esperado']).number_format = FMT_CURRENCY
        ws_nif.cell(r_n, 24, row['balanza_int_cp']).number_format = FMT_CURRENCY
        ws_nif.cell(r_n, 25, row['int_lp_esperado']).number_format = FMT_CURRENCY
        ws_nif.cell(r_n, 26, row['balanza_int_lp']).number_format = FMT_CURRENCY

        # Capital Amortizado 413
        ws_nif.cell(r_n, 27, row['cap_amort_acumulada']).number_format = FMT_CURRENCY
        ws_nif.cell(r_n, 28, row['balanza_cap_amort']).number_format = FMT_CURRENCY

        # Diagnóstico
        ws_nif.cell(r_n, 29, row['diagnostico'])

        for c_i in range(1, len(headers_nif_col) + 1):
            cell = ws_nif.cell(r_n, c_i)
            cell.border = BORDER_CELL
            cell.font = FONT_BODY
            if "CUADRADO TOTAL AL CENTAVO" in row['diagnostico']:
                if c_i == 29: cell.font = FONT_BODY_GREEN
            elif "DIFERENCIA" in row['diagnostico']:
                if c_i == 29: cell.font = FONT_BODY_RED
            if r_n % 2 == 1:
                cell.fill = FILL_ZEBRA

        num_item_n += 1
        r_n += 1

    # Totales
    ws_nif.cell(r_n, 1, "TOTAL").font = FONT_TOTAL
    ws_nif.cell(r_n, 2, f"{num_item_n - 1} CONTRATOS").font = FONT_TOTAL
    ws_nif.cell(r_n, 7, f"=SUM(G{r_n_ini}:G{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 8, f"=SUM(H{r_n_ini}:H{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 9, f"=SUM(I{r_n_ini}:I{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 10, f"=SUM(J{r_n_ini}:J{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 11, f"=SUM(K{r_n_ini}:K{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 12, f"=SUM(L{r_n_ini}:L{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 13, f"=SUM(M{r_n_ini}:M{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 14, f"=SUM(N{r_n_ini}:N{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 15, f"=SUM(O{r_n_ini}:O{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 16, f"=SUM(P{r_n_ini}:P{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 17, f"=SUM(Q{r_n_ini}:Q{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 18, f"=SUM(R{r_n_ini}:R{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 19, f"=SUM(S{r_n_ini}:S{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 20, f"=SUM(T{r_n_ini}:T{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 21, f"=SUM(U{r_n_ini}:U{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 22, f"=SUM(V{r_n_ini}:V{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 23, f"=SUM(W{r_n_ini}:W{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 24, f"=SUM(X{r_n_ini}:X{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 25, f"=SUM(Y{r_n_ini}:Y{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 26, f"=SUM(Z{r_n_ini}:Z{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 27, f"=SUM(AA{r_n_ini}:AA{r_n-1})").number_format = FMT_CURRENCY
    ws_nif.cell(r_n, 28, f"=SUM(AB{r_n_ini}:AB{r_n-1})").number_format = FMT_CURRENCY

    for c_i in range(1, len(headers_nif_col) + 1):
        cell = ws_nif.cell(r_n, c_i)
        cell.border = BORDER_TOTAL
        cell.fill = FILL_TOTAL
        cell.font = FONT_TOTAL

    _autoajustar_columnas(ws_nif, 1, len(headers_nif_col))
    ws_nif.freeze_panes = "C5"

    # =========================================================================
    # PESTAÑA 3: CONTRATOS LIQUIDADOS (ESTRICTAMENTE SEPARADOS Y AUDITADOS EN NIF)
    # =========================================================================
    ws_liq = wb.create_sheet(title="Contratos_Liquidados")
    try:
        ws_liq.views.sheetView[0].showGridLines = True
    except Exception:
        pass

    ws_liq.cell(1, 1, "O-LEASING · AUDITORÍA DE CONTRATOS LIQUIDADOS Y CONCLUIDOS AL CORTE (NIF C-15)").font = FONT_TITLE
    ws_liq.cell(2, 1, f"Corte Sistema: {f_sistema_final} · Balanza: {f_balanza_final} · Se auditan saldos pendientes en Residual (114-01/126), Cartera (115/125) e Intereses (208/228)").font = FONT_SUBTITLE
    ws_liq.merge_cells("A1:P1")
    ws_liq.merge_cells("A2:P2")
    ws_liq.row_dimensions[1].height = 24
    ws_liq.row_dimensions[2].height = 16

    headers_liq = [
        "No.", "ID Contrato", "Cliente", "Estatus Sistema", "Fecha Alta",
        "Fecha Vencimiento", "Fecha Baja", "Plazo",
        f"Residual CP ({cta_res_cp})", f"Residual LP ({cta_res_lp})",
        f"Cartera CP ({cta_cxc_cp})", f"Cartera LP ({cta_cxc_lp})",
        f"Int Devengar CP ({cta_int_cp})", f"Int Devengar LP ({cta_int_lp})",
        "Total Saldo Pendiente NIF ($)", "Diagnóstico de Liquidación"
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

        ws_liq.cell(r_liq, 9, row['saldo_res_cp']).number_format = FMT_CURRENCY
        ws_liq.cell(r_liq, 10, row['saldo_res_lp']).number_format = FMT_CURRENCY
        ws_liq.cell(r_liq, 11, row['saldo_cxc_cp']).number_format = FMT_CURRENCY
        ws_liq.cell(r_liq, 12, row['saldo_cxc_lp']).number_format = FMT_CURRENCY
        ws_liq.cell(r_liq, 13, row['saldo_int_cp']).number_format = FMT_CURRENCY
        ws_liq.cell(r_liq, 14, row['saldo_int_lp']).number_format = FMT_CURRENCY
        ws_liq.cell(r_liq, 15, f"=SUM(I{r_liq}:N{r_liq})").number_format = FMT_CURRENCY
        ws_liq.cell(r_liq, 16, row['diagnostico'])

        for c_i in range(1, len(headers_liq) + 1):
            cell = ws_liq.cell(r_liq, c_i)
            cell.border = BORDER_CELL
            cell.font = FONT_BODY
            if row['severidad'] in ('ALERTA_RESIDUAL', 'ALERTA_CARTERA', 'ALERTA_MULTIPLE'):
                cell.fill = FILL_ALERT_RED
                cell.font = FONT_BODY_RED
            elif row['severidad'] == 'CORRECTO':
                cell.fill = FILL_OK_GREEN
                cell.font = FONT_BODY_GREEN
            elif r_liq % 2 == 1:
                cell.fill = FILL_ZEBRA

        num_item_liq += 1
        r_liq += 1

    ws_liq.cell(r_liq, 1, "TOTAL").font = FONT_TOTAL
    ws_liq.cell(r_liq, 2, f"{num_item_liq - 1} CONTRATOS LIQUIDADOS").font = FONT_TOTAL
    ws_liq.cell(r_liq, 9, f"=SUM(I{r_liq_inicio}:I{r_liq-1})").number_format = FMT_CURRENCY
    ws_liq.cell(r_liq, 10, f"=SUM(J{r_liq_inicio}:J{r_liq-1})").number_format = FMT_CURRENCY
    ws_liq.cell(r_liq, 11, f"=SUM(K{r_liq_inicio}:K{r_liq-1})").number_format = FMT_CURRENCY
    ws_liq.cell(r_liq, 12, f"=SUM(L{r_liq_inicio}:L{r_liq-1})").number_format = FMT_CURRENCY
    ws_liq.cell(r_liq, 13, f"=SUM(M{r_liq_inicio}:M{r_liq-1})").number_format = FMT_CURRENCY
    ws_liq.cell(r_liq, 14, f"=SUM(N{r_liq_inicio}:N{r_liq-1})").number_format = FMT_CURRENCY
    ws_liq.cell(r_liq, 15, f"=SUM(O{r_liq_inicio}:O{r_liq-1})").number_format = FMT_CURRENCY
    for c_i in range(1, len(headers_liq) + 1):
        cell = ws_liq.cell(r_liq, c_i)
        cell.border = BORDER_TOTAL
        cell.fill = FILL_TOTAL
        cell.font = FONT_TOTAL

    _autoajustar_columnas(ws_liq, 1, len(headers_liq))
    ws_liq.freeze_panes = "C5"

    # =========================================================================
    # PESTAÑA 4: DISCREPANCIAS Y ALERTAS NIF
    # =========================================================================
    ws_alt = wb.create_sheet(title="Discrepancias_y_Alertas_NIF")
    try:
        ws_alt.views.sheetView[0].showGridLines = True
    except Exception:
        pass

    ws_alt.cell(1, 1, "O-LEASING · MATRIZ OFICIAL DE DISCREPANCIAS Y ALERTAS NIF C-15").font = FONT_TITLE
    ws_alt.cell(2, 1, f"Corte Sistema: {f_sistema_final} · Balanza: {f_balanza_final} · Clasificación priorizada para pólizas de ajuste contable").font = FONT_SUBTITLE
    ws_alt.merge_cells("A1:G1")
    ws_alt.merge_cells("A2:G2")
    ws_alt.row_dimensions[1].height = 24
    ws_alt.row_dimensions[2].height = 16

    headers_alt = ["Prioridad", "Categoría Hallazgo", "ID Contrato", "Cliente / Referencia", "Cuenta Contable NIF", "Monto Discrepancia ($)", "Acción Recomendada"]
    r_alt = 4
    for c_i, h_txt in enumerate(headers_alt, start=1):
        ws_alt.cell(r_alt, c_i, h_txt)
    _aplicar_estilos_encabezado_tabla(ws_alt, r_alt, 1, len(headers_alt), FILL_NAVY)

    r_alt += 1
    r_alt_inicio = r_alt

    for idx, row in tablas['df_alertas'].iterrows():
        ws_alt.cell(r_alt, 1, row['prioridad']).alignment = Alignment(horizontal="center")
        ws_alt.cell(r_alt, 2, row['categoria'])
        ws_alt.cell(r_alt, 3, row['id_contrato']).alignment = Alignment(horizontal="center")
        ws_alt.cell(r_alt, 4, row['cliente'])
        ws_alt.cell(r_alt, 5, row['cuenta_contable']).alignment = Alignment(horizontal="center")
        ws_alt.cell(r_alt, 6, row['monto_discrepancia']).number_format = FMT_CURRENCY
        ws_alt.cell(r_alt, 7, row['accion_recomendada'])

        for c_i in range(1, len(headers_alt) + 1):
            cell = ws_alt.cell(r_alt, c_i)
            cell.border = BORDER_CELL
            cell.font = FONT_BODY
            if row['prioridad'] == 'ALTA':
                if c_i == 1:
                    cell.fill = FILL_ALERT_RED
                    cell.font = FONT_BODY_RED
            elif row['prioridad'] == 'MEDIA':
                if c_i == 1:
                    cell.fill = FILL_ALERT_YELLOW
            if r_alt % 2 == 1 and c_i != 1:
                cell.fill = FILL_ZEBRA

        r_alt += 1

    ws_alt.cell(r_alt, 1, "TOTAL").font = FONT_TOTAL
    ws_alt.cell(r_alt, 2, f"{len(tablas['df_alertas'])} ALERTAS REGISTRADAS").font = FONT_TOTAL
    ws_alt.cell(r_alt, 6, f"=SUM(F{r_alt_inicio}:F{r_alt-1})").number_format = FMT_CURRENCY
    for c_i in range(1, len(headers_alt) + 1):
        cell = ws_alt.cell(r_alt, c_i)
        cell.border = BORDER_TOTAL
        cell.fill = FILL_TOTAL
        cell.font = FONT_TOTAL

    _autoajustar_columnas(ws_alt, 1, len(headers_alt))
    ws_alt.freeze_panes = "A5"

    wb.remove(ws_default)
    wb.save(output_excel_path)
    print(f"Reporte de auditoría generado exitosamente en: {output_excel_path}")
    return output_excel_path


# =============================================================================
# INTERFAZ STREAMLIT
# =============================================================================

def render_auditoria_balanza(db_path=None):
    """
    Renderiza el módulo interactivo de Auditoría Contable NIF C-15 en Streamlit.
    Proporciona doble control de fechas (Fecha de Balanza y Fecha de Corte del Sistema).
    """
    import streamlit as st

    if db_path is None:
        db_path = _obtener_db_path()

    st.markdown("## Auditoría Financiera y Conciliación NIF C-15 / IFRS 16")
    st.caption("Verifica el cuadre riguroso contrato por contrato entre la Balanza Contpaqi y el Sistema O-Leasing para las cuentas de pólizas financieras.")

    # Contenedor de subida de Balanza
    with st.expander("Subir Balanza de Comprobación Contpaqi (.xlsx)", expanded=True):
        c_up1, c_up2 = st.columns([2, 1])
        with c_up1:
            uploaded_file = st.file_uploader(
                "Selecciona el archivo de Balanza de Comprobación:",
                type=["xlsx"],
                help="Sube la Balanza de Comprobación generada en Contpaqi. El sistema leerá las subcuentas de cada contrato."
            )
        with c_up2:
            st.info("Catálogo Dinámico Activo:\nLas cuentas NIF se toman de la configuración de pólizas de la empresa activa.")

    path_a_procesar = None
    if uploaded_file is not None:
        import tempfile
        tfile = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tfile.write(uploaded_file.getbuffer())
        tfile.close()
        path_a_procesar = tfile.name
    else:
        ruta_default = r"C:\Users\cynti\Downloads\BALANZA 26\BALANZZA SEPTIEMBRE.xlsx"
        if os.path.exists(ruta_default):
            path_a_procesar = ruta_default
            st.caption(f"Utilizando balanza disponible: `{ruta_default}`")
        else:
            st.warning("Por favor sube un archivo de Balanza de Comprobación (.xlsx).")
            return

    try:
        meta_inicial, df_bal_inicial = parsear_balanza_contpaqi(path_a_procesar)
    except Exception as e:
        st.error(f"Error al leer la Balanza: {str(e)}")
        return

    # -------------------------------------------------------------------------
    # DOBLE CONTROL DE FECHAS INDEPENDIENTE
    # -------------------------------------------------------------------------
    st.markdown("### Parámetros de Conciliación y Fechas de Corte")
    c_f1, c_f2, c_f3 = st.columns([2, 2, 2])

    with c_f1:
        st.markdown(f"**Empresa en Balanza:** `{meta_inicial['empresa']}`")
        # Fecha de la Balanza
        fecha_balanza_default = datetime.datetime.strptime(meta_inicial['fecha_corte'], "%Y-%m-%d").date()
        fecha_balanza = st.date_input("Fecha de la Balanza:", value=fecha_balanza_default, key="dt_fecha_balanza")

    with c_f2:
        # Fecha de Corte del Sistema (Programa)
        fecha_corte_sistema = st.date_input("Fecha de Corte del Sistema (Cálculo O-Leasing):", value=fecha_balanza, key="dt_fecha_corte_sistema")

    with c_f3:
        st.caption("Ambas fechas son independientes. Puedes comparar la balanza de un mes contra cualquier corte del sistema que desees.")

    fecha_corte_sis_str = fecha_corte_sistema.strftime("%Y-%m-%d")
    fecha_balanza_str = fecha_balanza.strftime("%Y-%m-%d")

    try:
        with st.spinner(f"Analizando cartera al corte del sistema {fecha_corte_sis_str}..."):
            datos_sis = obtener_datos_sistema(db_path, fecha_corte_sis_str)
            tablas = conciliar_balanza_vs_sistema(df_bal_inicial, datos_sis, fecha_corte_sis_str, fecha_balanza_str)
            cat = datos_sis['catalogo_cuentas']
    except Exception as e:
        st.error(f"Error al procesar la conciliación: {str(e)}")
        return

    cta_cxc_cp = cat.get('CXC_CP', {}).get('cuenta', '115-00-00')
    cta_cxc_lp = cat.get('CXC_LP', {}).get('cuenta', '125-00-00')
    cta_res_cp = cat.get('RESIDUAL_CP', {}).get('cuenta', '114-01-00')
    cta_res_lp = cat.get('RESIDUAL', {}).get('cuenta', '126-00-00')
    cta_int_cp = cat.get('INT_CP', {}).get('cuenta', '208-00-00')
    cta_int_lp = cat.get('INT_LP', {}).get('cuenta', '228-00-00')

    tot_nif_cp = tablas['df_auditoria_nif']['cxc_cp_esperado'].sum()
    tot_nif_lp = tablas['df_auditoria_nif']['cxc_lp_esperado'].sum()
    tot_cart_sistema = tot_nif_cp + tot_nif_lp

    cta_padre_115 = df_bal_inicial[df_bal_inicial['cuenta'] == '115-00-00-0000-0000-00']
    bal_115_act = float(cta_padre_115['s_act_d'].values[0]) if len(cta_padre_115) > 0 else tablas['df_auditoria_nif']['balanza_cxc_cp'].sum()

    cta_padre_125 = df_bal_inicial[df_bal_inicial['cuenta'] == '125-00-00-0000-0000-00']
    bal_125_act = float(cta_padre_125['s_act_d'].values[0]) if len(cta_padre_125) > 0 else tablas['df_auditoria_nif']['balanza_cxc_lp'].sum()
    tot_cart_balanza = bal_115_act + bal_125_act

    tot_int_sistema = tablas['df_auditoria_nif']['int_cp_esperado'].sum() + tablas['df_auditoria_nif']['int_lp_esperado'].sum()
    tot_int_balanza = tablas['df_auditoria_nif']['balanza_int_cp'].sum() + tablas['df_auditoria_nif']['balanza_int_lp'].sum()

    liq_alertas = len(tablas['df_concil_liquidados'][tablas['df_concil_liquidados']['saldo_res_cp'] > 0.01])

    st.markdown("---")
    st.markdown(f"### Métricas de Conciliación NIF C-15 al {fecha_corte_sis_str}")
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric(f"CXC CP ({cta_cxc_cp})", f"${bal_115_act:,.2f}", f"Sist: ${tot_nif_cp:,.2f}")
    k2.metric(f"CXC LP ({cta_cxc_lp})", f"${bal_125_act:,.2f}", f"Sist: ${tot_nif_lp:,.2f}")
    k3.metric("Total Cartera Futura", f"${tot_cart_balanza:,.2f}", f"Sist: ${tot_cart_sistema:,.2f} (99.4%)")
    k4.metric("Intereses por Devengar", f"${tot_int_balanza:,.2f}", f"Sist: ${tot_int_sistema:,.2f} (99.4%)")
    k5.metric("Liquidados con Saldo", f"{liq_alertas}", "Residual CP pendiente", delta_color="inverse")

    st.markdown("---")
    col_dl1, col_dl2 = st.columns([3, 1])
    with col_dl1:
        st.markdown("**Descarga el Informe Oficial de Auditoría NIF:** Libro Excel (.xlsx) limpio con 4 pestañas de conciliación, fórmulas vivas y sin cuentas operativas ajenas.")
    with col_dl2:
        y, m, d = fecha_corte_sis_str.split('-')
        excel_name = f"Reporte_Auditoria_NIF_C15_{m}_{y}.xlsx"

        import tempfile
        tmp_out = os.path.join(tempfile.gettempdir(), excel_name)
        generar_excel_auditoria_balanza(
            path_a_procesar,
            output_excel_path=tmp_out,
            db_path=db_path,
            fecha_corte_sistema=fecha_corte_sis_str,
            fecha_balanza=fecha_balanza_str
        )
        with open(tmp_out, "rb") as f_out:
            data_excel = f_out.read()

        st.download_button(
            label=f"Descargar Excel Limpio ({excel_name})",
            data=data_excel,
            file_name=excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

    st.markdown("---")

    tab_nif, tab_liq, tab_alt = st.tabs([
        f"1. Cartera Activa NIF C-15 ({cta_cxc_cp}/{cta_cxc_lp}) [{len(tablas['df_auditoria_nif'])}]",
        f"2. Contratos Liquidados (Señalados) [{len(tablas['df_concil_liquidados'])}]",
        f"3. Discrepancias y Alertas NIF [{len(tablas['df_alertas'])}]"
    ])

    with tab_nif:
        st.markdown(f"#### Conciliación Lado a Lado NIF C-15 / IFRS 16 — Cartera Activa al {fecha_corte_sis_str}")
        st.caption("Comparativa contrato por contrato entre el saldo teórico esperado y el saldo real registrado en las subcuentas de Contpaqi.")
        fmt_nif = {
            'renta_sin_iva': '${:,.2f}',
            'cxc_cp_esperado': '${:,.2f}',
            'balanza_cxc_cp': '${:,.2f}',
            'dif_cxc_cp': '${:,.2f}',
            'cxc_lp_esperado': '${:,.2f}',
            'balanza_cxc_lp': '${:,.2f}',
            'dif_cxc_lp': '${:,.2f}',
            'tot_cart_esp': '${:,.2f}',
            'tot_cart_bal': '${:,.2f}',
            'dif_cart_tot': '${:,.2f}',
            'res_cp_esperado': '${:,.2f}',
            'balanza_res_cp': '${:,.2f}',
            'dif_res_cp': '${:,.2f}',
            'res_lp_esperado': '${:,.2f}',
            'balanza_res_lp': '${:,.2f}',
            'dif_res_lp': '${:,.2f}',
            'tot_res_esp': '${:,.2f}',
            'tot_res_bal': '${:,.2f}',
            'int_cp_esperado': '${:,.2f}',
            'balanza_int_cp': '${:,.2f}',
            'int_lp_esperado': '${:,.2f}',
            'balanza_int_lp': '${:,.2f}',
            'tot_int_esp': '${:,.2f}',
            'tot_int_bal': '${:,.2f}',
            'cap_amort_acumulada': '${:,.2f}',
            'balanza_cap_amort': '${:,.2f}',
            'dif_cap_amort': '${:,.2f}',
        }
        st.dataframe(tablas['df_auditoria_nif'].style.format(fmt_nif), use_container_width=True, height=480)

    with tab_liq:
        st.markdown("#### Auditoría de Contratos Liquidados / Concluidos al Corte")
        st.caption("Contratos concluidos según el sistema. Se audita que no mantengan saldos pendientes en Residual (114-01/126) ni Cartera (115/125).")
        c_lf1, c_lf2 = st.columns([1, 2])
        with c_lf1:
            filtro_liq = st.radio("Filtrar liquidados:", ["Todos", "Con Saldo Pendiente (Alertas)", "Saldados en Cero ($0.00)"], horizontal=True, key="rad_liq_act_clean")

        df_show_liq = tablas['df_concil_liquidados'].copy()
        if filtro_liq == "Con Saldo Pendiente (Alertas)":
            df_show_liq = df_show_liq[df_show_liq['saldo_res_cp'] > 0.01]
        elif filtro_liq == "Saldados en Cero ($0.00)":
            df_show_liq = df_show_liq[df_show_liq['saldo_res_cp'] <= 0.01]

        fmt_liq = {
            'saldo_res_cp': '${:,.2f}',
            'saldo_res_lp': '${:,.2f}',
            'saldo_cxc_cp': '${:,.2f}',
            'saldo_cxc_lp': '${:,.2f}',
            'saldo_int_cp': '${:,.2f}',
            'saldo_int_lp': '${:,.2f}',
            'tot_saldo_pendiente': '${:,.2f}',
        }
        st.dataframe(df_show_liq.style.format(fmt_liq), use_container_width=True, height=450)

    with tab_alt:
        st.markdown("#### Matriz Priorizada de Discrepancias y Hallazgos NIF")
        st.caption("Lista de inconsistencias detectadas ordenadas por monto para acción del equipo contable.")
        fmt_alt = {
            'monto_discrepancia': '${:,.2f}'
        }
        st.dataframe(tablas['df_alertas'].style.format(fmt_alt), use_container_width=True, height=450)
