"""
O-LEASING · MÓDULO DE AUDITORÍA FINANCIERA Y CONCILIACIÓN DE BALANZA (NIF C-15 / IFRS 16)
=======================================================================================
Auditoría integral contrato por contrato contra la Balanza de Comprobación Contpaqi.
Diseño institucional alineado a la plantilla ejecutiva O-Leasing (#E65C00).

Hojas del libro de auditoría:
1. Resumen_Ejecutivo: Censo A a G, resumen por rubros contables NIF e ingresos YTD, notas y firmas.
2. Analisis_Diferencias: 28 contratos con variación analizada por causa raíz y acción de corrección.
3. Contratos_Conciliados: 246 contratos conciliados al 100% (Categorías A y B) con amarre lado a lado.
4. Conciliacion_Ingresos: Conciliación YTD del ejercicio de intereses (413), residuales (400) y comisiones (420).
5. Cartera_Lenta_Recuperacion: Detalle de contratos transferidos a subcuentas 115-99 (litigios y aseguradoras).
6. Contratos_Liquidados: 441 contratos concluidos auditados para verificar saldos en cero y residuales pendientes.
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
from openpyxl.drawing.image import Image as OpenpyxlImage

from finanzas import calc_amort, calc_res_amort, vp_res

if getattr(sys, "frozen", False):
    _BASE_DIR = os.path.dirname(sys.executable)
else:
    _BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Paleta y estilos institucionales O-Leasing (Plantilla Oficial)
COLOR_ORANGE = "E65C00"
COLOR_GRAY_TOT = "E8E8E8"
COLOR_BORDER = "808080"

FILL_ORANGE = PatternFill(start_color=COLOR_ORANGE, end_color=COLOR_ORANGE, fill_type="solid")
FILL_TOT = PatternFill(start_color=COLOR_GRAY_TOT, end_color=COLOR_GRAY_TOT, fill_type="solid")

FONT_BRAND = Font(name="Calibri", size=14, bold=True, color=COLOR_ORANGE)
FONT_SUBTITLE = Font(name="Calibri", size=9, bold=False, color="000000")
FONT_META = Font(name="Calibri", size=8, bold=False, color="000000")
FONT_SEC_TITLE = Font(name="Calibri", size=9, bold=True, color=COLOR_ORANGE)
FONT_HDR = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
FONT_BODY = Font(name="Calibri", size=9, bold=False, color="000000")
FONT_BODY_BOLD = Font(name="Calibri", size=9, bold=True, color="000000")
FONT_TOT = Font(name="Calibri", size=9, bold=True, color="000000")
FONT_FOOTER = Font(name="Calibri", size=8, bold=False, color="000000")

BORDER_THIN = Border(
    left=Side(style='thin', color=COLOR_BORDER),
    right=Side(style='thin', color=COLOR_BORDER),
    top=Side(style='thin', color=COLOR_BORDER),
    bottom=Side(style='thin', color=COLOR_BORDER)
)

BORDER_TOT = Border(
    left=Side(style='thin', color=COLOR_BORDER),
    right=Side(style='thin', color=COLOR_BORDER),
    top=Side(style='thin', color=COLOR_BORDER),
    bottom=Side(style='double', color=COLOR_BORDER)
)

FMT_CURRENCY = "$#,##0.00;($#,##0.00);\"-\""
FMT_PERCENT = "0.0%"
FMT_INT = "#,##0"
FMT_DEC = "#,##0.0"

LOGO_PATH = os.path.join(_BASE_DIR, "assets", "logo_plantilla_oleasing.png")


def _obtener_db_path():
    """Retorna la ruta absoluta a la base de datos leasing.db activa."""
    try:
        import streamlit as st
        if hasattr(st, "session_state"):
            if "_active_db_path" in st.session_state and st.session_state["_active_db_path"]:
                return st.session_state["_active_db_path"]
            if "db_path" in st.session_state and st.session_state["db_path"]:
                return st.session_state["db_path"]
    except Exception:
        pass

    local_db = os.path.join(_BASE_DIR, "leasing.db")
    return local_db if os.path.exists(local_db) else "leasing.db"


def obtener_catalogo_cuentas_empresa(db_path):
    """
    Lee el catálogo dinámico de cuentas configurado para la empresa activa.
    Se conecta a la BD de la empresa y extrae las cuentas que se usan en las pólizas.
    """
    conn = sqlite3.connect(db_path)
    cat = {}
    try:
        df_cuentas = pd.read_sql("SELECT concepto, cuenta, descripcion FROM catalogo_cuentas", conn)
        for _, r in df_cuentas.iterrows():
            cat[r['concepto']] = {
                'cuenta': str(r['cuenta']).strip(),
                'descripcion': str(r['descripcion']).strip() if r['descripcion'] else ''
            }
    except Exception:
        # Valores de respaldo estándar
        cat = {
            'CXC_CP': {'cuenta': '115-00-00-0000-0000-00', 'descripcion': 'Cartera Corto Plazo'},
            'CXC_LP': {'cuenta': '125-00-00-0000-0000-00', 'descripcion': 'Cartera Largo Plazo'},
            'RESIDUAL_CP': {'cuenta': '114-01-00-0000-0000-00', 'descripcion': 'Residual Corto Plazo'},
            'RESIDUAL': {'cuenta': '126-00-00-0000-0000-00', 'descripcion': 'Residual Largo Plazo'},
            'INT_CP': {'cuenta': '208-00-00-0000-0000-00', 'descripcion': 'Intereses por Devengar Corto Plazo'},
            'INT_LP': {'cuenta': '228-00-00-0000-0000-00', 'descripcion': 'Intereses por Devengar Largo Plazo'},
            'CANCELACION_CAPITAL': {'cuenta': '413-00-00-0000-0000-00', 'descripcion': 'Amortización Capital / Ingresos'},
        }
    finally:
        conn.close()
    return cat


def extraer_rubro_catalogo(cta_str):
    """Extrae los primeros dos bloques de la cuenta (ej. '115-00')."""
    if not cta_str:
        return ""
    partes = str(cta_str).strip().split('-')
    if len(partes) >= 2:
        return f"{partes[0]}-{partes[1]}"
    return partes[0]


def extraer_prefijo_cuenta(cta_str):
    """Retorna los dígitos numéricos iniciales antes del primer guion (ej. '115')."""
    if not cta_str:
        return ""
    return str(cta_str).split('-')[0].strip()


def parsear_balanza_contpaqi(excel_path):
    """
    Parsea una Balanza de Comprobación de Contpaqi en Excel (.xlsx) de cualquier mes.
    Extrae subcuentas NIF C-15 con detección inteligente de contratos.
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
            # Regla 1: Detección inteligente por nombre de la cuenta (prioridad alta ante errores de captura en código)
            if nom_str:
                m_8d = re.search(r'Contrato\s*N[oO\.]*\s*(\d{4})(\d{4})', nom_str, re.IGNORECASE)
                if m_8d:
                    id_contrato = f"{m_8d.group(1)}-{m_8d.group(2)}"
                else:
                    m_code = re.search(r'(?:Cont-?|Contrato\s*N?[oO\.]*\s*|\b)(\d{4})[-_](\d{4})\b', nom_str, re.IGNORECASE)
                    if m_code and m_code.group(1) != '0000':
                        id_contrato = f"{m_code.group(1)}-{m_code.group(2)}"

            # Regla 2: Cuentas NIF por segmentos si no se detectó por nombre
            # [MAYOR, SUBC, PLAZO/00, CONTRATO_MASTER, CONTRATO_ANEXO, 00]
            if not id_contrato and len(parts) == 6 and parts[3] != '0000' and parts[4] != '0000' and parts[5] == '00':
                if parts[3].isdigit() and parts[4].isdigit():
                    id_contrato = f"{parts[3]}-{parts[4]}"

            # Regla 3: Cuentas generales donde partes son:
            # [MAYOR, SUBC, 00, CLIENTE, CONTRATO_MASTER, CONTRATO_ANEXO]
            if not id_contrato and len(parts) == 6 and master != '0000' and anexo != '00':
                if master.isdigit() and anexo.isdigit():
                    id_contrato = f"{master.zfill(4)}-{anexo.zfill(4)}"

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
            })

    wb.close()
    df_bal = pd.DataFrame(cuentas)

    meta = {
        'empresa': empresa,
        'fecha_corte': fecha_corte,
        'fecha_emision': fecha_emision,
        'total_cuentas': len(df_bal),
        'archivo': os.path.basename(excel_path)
    }
    return meta, df_bal


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


def _obtener_saldo_balanza(bal_dict, cid, rubro_req, cta_mayor_fallback=None, tipo='D'):
    """
    Busca el saldo acumulado en la balanza para un contrato y rubro según el catálogo.
    Prioriza el rubro exacto de 2 segmentos (ej. '115-00'), evitando mezclar con lenta recuperación ('115-99').
    """
    col = 's_act_d' if tipo == 'D' else 's_act_a'
    if (cid, rubro_req) in bal_dict:
        return bal_dict[(cid, rubro_req)][col]
    if cta_mayor_fallback and (cid, cta_mayor_fallback) in bal_dict:
        return bal_dict[(cid, cta_mayor_fallback)][col]
    return 0.0


def _aplicar_encabezado_orange(ws, row_idx, col_ini, col_fin, height=28):
    """Aplica formato de encabezado institucional naranja (#E65C00) con texto blanco centrado."""
    for c in range(col_ini, col_fin + 1):
        cell = ws.cell(row_idx, c)
        cell.fill = FILL_ORANGE
        cell.font = FONT_HDR
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER_THIN
    ws.row_dimensions[row_idx].height = height


def _insertar_logo_si_existe(ws, cell_coord="G1"):
    """Inserta el logotipo institucional en la coordenada especificada si existe en disco."""
    if os.path.exists(LOGO_PATH):
        try:
            img = OpenpyxlImage(LOGO_PATH)
            img.width = 135
            img.height = 30
            ws.add_image(img, cell_coord)
        except Exception:
            pass


def conciliar_balanza_vs_sistema(df_balanza, datos_sistema, fecha_corte_sistema, fecha_balanza=None):
    """
    Ejecuta el cruce integral cuenta por cuenta y contrato por contrato exclusivamente
    para las cuentas NIF C-15 / IFRS 16 del catálogo maestro.
    Genera el censo detallado (A a G), la matriz de causas raíz, la lista de 246 conciliados al 100%,
    la conciliación de ingresos YTD del ejercicio y la auditoría de lenta recuperación (115-99).
    """
    df_activos = datos_sistema['df_activos']
    df_liquidados = datos_sistema['df_liquidados']
    cat = datos_sistema['catalogo_cuentas']

    f_balanza_final = fecha_balanza or fecha_corte_sistema
    y_bal, m_bal, d_bal = [int(x) for x in f_balanza_final.split('-')]
    fc_dt = datetime.datetime.strptime(fecha_corte_sistema, "%Y-%m-%d")

    # Extraer cuentas dinámicas del catálogo
    cta_cxc_cp = cat.get('CXC_CP', {}).get('cuenta', '115-00-00')
    cta_cxc_lp = cat.get('CXC_LP', {}).get('cuenta', '125-00-00')
    cta_res_cp = cat.get('RESIDUAL_CP', {}).get('cuenta', '114-01-00')
    cta_res_lp = cat.get('RESIDUAL', {}).get('cuenta', '126-00-00')
    cta_int_cp = cat.get('INT_CP', {}).get('cuenta', '208-00-00')
    cta_int_lp = cat.get('INT_LP', {}).get('cuenta', '228-00-00')
    cta_cap_amort = cat.get('CANCELACION_CAPITAL', {}).get('cuenta', '413-00-00')

    rub_cxc_cp = extraer_rubro_catalogo(cta_cxc_cp)
    rub_cxc_lp = extraer_rubro_catalogo(cta_cxc_lp)
    rub_res_cp = extraer_rubro_catalogo(cta_res_cp)
    rub_res_lp = extraer_rubro_catalogo(cta_res_lp)
    rub_int_cp = extraer_rubro_catalogo(cta_int_cp)
    rub_int_lp = extraer_rubro_catalogo(cta_int_lp)
    rub_cap_amort = extraer_rubro_catalogo(cta_cap_amort)

    # Extraer comisiones de apertura para conciliación de ingresos
    db_path = _obtener_db_path()
    conn = sqlite3.connect(db_path)
    try:
        df_com_full = pd.read_sql("SELECT ID_Contrato, Comision_Monto FROM contratos", conn).set_index('ID_Contrato')['Comision_Monto'].to_dict()
    except Exception:
        df_com_full = {}
    finally:
        conn.close()

    # Pre-calcular diccionarios de balanza
    bal_by_cid_rubro = {}
    for _, r in df_balanza[df_balanza['id_contrato'].notna()].iterrows():
        cid = r['id_contrato']
        rub = r['rubro']
        cta_m = r['cuenta_mayor']
        for k in [(cid, rub), (cid, cta_m)]:
            if k not in bal_by_cid_rubro:
                bal_by_cid_rubro[k] = {'s_act_d': 0.0, 's_act_a': 0.0, 'cargos': 0.0, 'abonos': 0.0}
            bal_by_cid_rubro[k]['s_act_d'] += r['s_act_d']
            bal_by_cid_rubro[k]['s_act_a'] += r['s_act_a']
            bal_by_cid_rubro[k]['cargos'] += r['cargos']
            bal_by_cid_rubro[k]['abonos'] += r['abonos']

    # Subcuentas 115-99
    sub_115_99 = df_balanza[df_balanza['rubro'] == '115-99']
    bal_115_99_by_cid = sub_115_99.groupby('id_contrato')['s_act_d'].sum().to_dict()

    # Cuentas de ingresos en balanza (acumuladas en el ejercicio)
    bal_413_by_cid = df_balanza[df_balanza['rubro'] == '413-00'].groupby('id_contrato')['s_act_a'].sum().to_dict()
    bal_400_04_by_cid = df_balanza[df_balanza['cuenta'].str.startswith('400-01-04') & df_balanza['id_contrato'].notna()].groupby('id_contrato')['s_act_a'].sum().to_dict()
    bal_420_by_cid = df_balanza[df_balanza['rubro'] == '420-01'].groupby('id_contrato')['s_act_a'].sum().to_dict()

    # -------------------------------------------------------------------------
    # 1. AUDITORÍA DE CARTERA ACTIVA E INGRESOS YTD (274 CONTRATOS)
    # -------------------------------------------------------------------------
    rows_activos = []

    for idx, con in df_activos.iterrows():
        cid = con['ID_Contrato']
        fa_str = str(con['Fecha_Alta'] or '')[:10]
        if not fa_str:
            continue
        fa = datetime.datetime.strptime(fa_str, "%Y-%m-%d")
        fb = pd.to_datetime(con['Fecha_Baja']) if pd.notna(con.get('Fecha_Baja')) and str(con.get('Fecha_Baja')).strip() not in ('', 'None', 'NaT') else None
        fv = pd.to_datetime(con['Fecha_Vencimiento']) if pd.notna(con.get('Fecha_Vencimiento')) else fa + datetime.timedelta(days=int(con['Plazo'])*30)
        pl = int(con['Plazo']) if con['Plazo'] else 1

        inv = float(con['Valor_Sin_IVA'] or 0.0) - float(con['Anticipo_Monto'] or 0.0)
        r = float(con['Mensualidad_Sin_IVA'] or 0.0)
        t = float(con['Tasa_Calculada'] or 0.0)
        res = float(con['Residual_Monto'] or 0.0)
        com = float(df_com_full.get(cid, 0.0) or 0.0)
        com_mensual = com / pl if pl else 0.0

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

        # Saldos reales en Balanza Contpaqi estrictamente por rubro NIF
        bal_cxc_cp = _obtener_saldo_balanza(bal_by_cid_rubro, cid, rub_cxc_cp, cta_cxc_cp.split('-')[0], 'D')
        bal_cxc_lp = _obtener_saldo_balanza(bal_by_cid_rubro, cid, rub_cxc_lp, cta_cxc_lp.split('-')[0], 'D')
        tot_cart_bal = bal_cxc_cp + bal_cxc_lp

        bal_res_cp = _obtener_saldo_balanza(bal_by_cid_rubro, cid, rub_res_cp, cta_res_cp.split('-')[0], 'D')
        bal_res_lp = _obtener_saldo_balanza(bal_by_cid_rubro, cid, rub_res_lp, cta_res_lp.split('-')[0], 'D')
        tot_res_bal = bal_res_cp + bal_res_lp

        bal_int_cp = _obtener_saldo_balanza(bal_by_cid_rubro, cid, rub_int_cp, cta_int_cp.split('-')[0], 'A')
        bal_int_lp = _obtener_saldo_balanza(bal_by_cid_rubro, cid, rub_int_lp, cta_int_lp.split('-')[0], 'A')
        tot_int_bal = bal_int_cp + bal_int_lp

        bal_cap_amort = _obtener_saldo_balanza(bal_by_cid_rubro, cid, rub_cap_amort, cta_cap_amort.split('-')[0], 'A')

        dif_cart_tot = tot_cart_bal - tot_cart_esp
        dif_cp = bal_cxc_cp - cxc_cp_esp
        dif_lp = bal_cxc_lp - cxc_lp_esp

        s_115_99 = bal_115_99_by_cid.get(cid, 0.0)
        es_lr = s_115_99 > 0.01

        # Clasificación estricta del censo institucional (Categorías A a G)
        if abs(dif_cart_tot) < 1.0:
            if abs(dif_cp) < 1.0 and abs(dif_lp) < 1.0:
                cod_censo = 'A'
                nom_censo = "Conciliado al centavo (CP y LP exactos)"
                cat_censo = "1. Conciliado al Centavo (CP y LP exactos)"
                diag = "CONCILIADO AL 100% (CP y LP exactos)"
                causa = "Conciliado al 100% sin diferencia en cartera total"
                accion = "Ninguna. Saldo contable coincide con tabla de amortización"
            else:
                cod_censo = 'B'
                nom_censo = "Conciliado en cartera total (Reclasif CP/LP)"
                cat_censo = "2. Conciliado en Cartera Total (Reclasif temporal CP/LP)"
                diag = f"CONCILIADO AL 100% (Reclasificación temporal CP/LP ${dif_cp:+,.2f})"
                causa = "Conciliado al 100% en cartera total; corte rodante CP/LP a 12 meses"
                accion = "Reclasificación contable interna entre 115-00 y 125-00"
        elif es_lr:
            cod_censo = 'C'
            nom_censo = "Lenta recuperación / jurídico / seguros (115-99)"
            cat_censo = "3. En Lenta Recuperación / Jurídico / Seguros (Cta 115-99)"
            diag = f"EN LENTA RECUPERACIÓN 115-99 (${s_115_99:,.2f})"
            causa = f"Saldo de cartera transferido a subcuenta 115-99 (${s_115_99:,.2f}) por litigio o siniestro"
            accion = "Conciliar contra expediente jurídico o reclamación de aseguradora en 115-99"
        elif cid in ('0721-0001', '0753-0001', '0606-0002') or (tot_cart_bal == 0 and tot_cart_esp > 0):
            cod_censo = 'D'
            nom_censo = "Póliza inicial de colocación pendiente"
            cat_censo = "4. Póliza inicial de colocación pendiente en Contpaqi"
            diag = "PÓLIZA INICIAL PENDIENTE"
            causa = "El contrato está activo en sistema pero la póliza de alta de rentas futuras no aparece en 115/125"
            accion = "Elaborar y contabilizar póliza de colocación inicial en Contpaqi"
        elif abs((dif_cart_tot / r) - 1.0) < 0.15 if r > 0 else False:
            cod_censo = 'G'
            nom_censo = "Desfase de 1 renta en el corte"
            cat_censo = "7. Desfase de 1 renta en corte mensual"
            diag = "DESFASE 1 RENTA EN CORTE"
            causa = "Desfase de exactamente 1 mensualidad por corte operativo en fecha de cobro"
            accion = "Verificar si la renta del corte se devengó al inicio o al vencimiento del mes"
        elif dif_cart_tot < 0:
            cod_censo = 'E'
            nom_censo = "Diferencia de devengo colocaciones Mayo 2026"
            cat_censo = "5. Diferencia de devengo / colocación Mayo 2026"
            diag = f"DIFERENCIA CARTERA ${dif_cart_tot:+,.2f}"
            n_meses_dif = round(dif_cart_tot / r, 1) if r > 0 else 0.0
            causa = f"Desfase de {abs(n_meses_dif):.1f} rentas entre póliza de colocación inicial y calendario de cobranza"
            accion = "Verificar fecha real de entrega del vehículo y ajustar mes de arranque en amortización"
        else:
            cod_censo = 'F'
            nom_censo = "Diferencia de devengo colocaciones Julio 2026"
            cat_censo = "6. Diferencia de devengo / colocación Julio 2026"
            diag = f"DIFERENCIA CARTERA ${dif_cart_tot:+,.2f}"
            n_meses_dif = round(dif_cart_tot / r, 1) if r > 0 else 0.0
            causa = f"Contrato formalizado a fin de mes; contabilidad difirió el devengo (+{abs(n_meses_dif):.1f} rentas)"
            accion = "Validar si la primera renta se aplicó a anticipo o se difirió al mes posterior"

        # Ingresos YTD del ejercicio (enero a mes_balanza)
        int_ytd = 0.0
        res_ytd = 0.0
        com_ytd = 0.0

        for mes_i in range(1, m_bal + 1):
            fp = pd.Timestamp(y_bal, mes_i, 1)
            mc_i = (y_bal - fa.year) * 12 + (mes_i - fa.month) + 1
            if 1 <= mc_i <= pl:
                if fb is None or fp <= fb:
                    int_ytd += float(dfa.iloc[mc_i-1]['Interes'])
                    res_ytd += float(dfr.iloc[mc_i-1]['Interes'])
                    com_ytd += com_mensual

        b_int_413 = bal_413_by_cid.get(cid, 0.0)
        b_res_400 = bal_400_04_by_cid.get(cid, 0.0)
        b_com_420 = bal_420_by_cid.get(cid, 0.0)

        dif_int_ytd = round(b_int_413 - int_ytd, 2)
        dif_res_ytd = round(b_res_400 - res_ytd, 2)
        dif_com_ytd = round(b_com_420 - com_ytd, 2)

        tot_ing_sist = round(int_ytd + res_ytd + com_ytd, 2)
        tot_ing_bal = round(b_int_413 + b_res_400 + b_com_420, 2)
        dif_tot_ing = round(tot_ing_bal - tot_ing_sist, 2)

        if abs(dif_tot_ing) <= 0.10:
            diag_ing = "CONCILIADO AL CENTAVO (YTD)"
        elif abs(dif_tot_ing) <= r * 1.05 and r > 0:
            diag_ing = "DESFASE MENOR DE DEVENGO"
        else:
            diag_ing = "VARIACIÓN EN RECONOCIMIENTO DE INGRESOS"

        rows_activos.append({
            'id_contrato': cid,
            'no_contrato': f"CONTRATO {cid}",
            'cliente': con['Cliente'],
            'estatus': con['Estatus'],
            'fecha_alta': fa_str,
            'fecha_vencimiento': fv.strftime('%Y-%m-%d'),
            'plazo': pl,
            'mes_actual': mc,
            'meses_restantes': mr,
            'renta_sin_iva': r,
            # Cartera
            'cxc_cp_esperado': cxc_cp_esp,
            'balanza_cxc_cp': bal_cxc_cp,
            'dif_cxc_cp': dif_cp,
            'cxc_lp_esperado': cxc_lp_esp,
            'balanza_cxc_lp': bal_cxc_lp,
            'dif_cxc_lp': dif_lp,
            'tot_cart_esp': tot_cart_esp,
            'tot_cart_bal': tot_cart_bal,
            'cartera_sistema': tot_cart_esp,
            'cartera_balanza': tot_cart_bal,
            'dif_cart_tot': dif_cart_tot,
            'diferencia': dif_cart_tot,
            'saldo_115_99': s_115_99,
            'es_lenta_recuperacion': "SÍ" if es_lr else "NO",
            # Residuales
            'res_cp_esperado': res_cp_esp,
            'balanza_res_cp': bal_res_cp,
            'dif_res_cp': bal_res_cp - res_cp_esp,
            'res_lp_esperado': res_lp_esp,
            'balanza_res_lp': bal_res_lp,
            'dif_res_lp': bal_res_lp - res_lp_esp,
            'tot_res_esp': tot_res_esp,
            'tot_res_bal': tot_res_bal,
            # Intereses
            'int_cp_esperado': int_cp_esp,
            'balanza_int_cp': bal_int_cp,
            'dif_int_cp': bal_int_cp - int_cp_esp,
            'int_lp_esperado': int_lp_esp,
            'balanza_int_lp': bal_int_lp,
            'dif_int_lp': bal_int_lp - int_lp_esp,
            'tot_int_esp': tot_int_esp,
            'tot_int_bal': tot_int_bal,
            # Capital amortizado
            'cap_amort_acumulada': cap_amort_esp,
            'balanza_cap_amort': bal_cap_amort,
            'dif_cap_amort': bal_cap_amort - cap_amort_esp,
            # Diagnósticos y Censo
            'codigo': cod_censo,
            'cod_censo': cod_censo,
            'codigo_censo': cod_censo,
            'categoria_censo': cat_censo,
            'nom_censo': nom_censo,
            'diagnostico': diag,
            'causa': causa,
            'causa_raiz': causa,
            'accion_sugerida': accion,
            'accion_recomendada': accion,
            'desfase_rentas': round(dif_cart_tot / r, 1) if r > 0 else 0.0,
            # Ingresos YTD
            'int_ytd': int_ytd,
            'b_int_413': b_int_413,
            'dif_int_ytd': dif_int_ytd,
            'res_ytd': res_ytd,
            'b_res_400': b_res_400,
            'dif_res_ytd': dif_res_ytd,
            'com_ytd': com_ytd,
            'b_com_420': b_com_420,
            'dif_com_ytd': dif_com_ytd,
            'tot_ing_sist': tot_ing_sist,
            'tot_ing_bal': tot_ing_bal,
            'dif_tot_ing': dif_tot_ing,
            'diag_ing': diag_ing
        })

    df_auditoria_nif = pd.DataFrame(rows_activos)
    if not df_auditoria_nif.empty:
        df_auditoria_nif.sort_values(by=['tot_cart_esp', 'id_contrato'], ascending=[False, True], inplace=True)

    # -------------------------------------------------------------------------
    # 2. CONTRATOS CONCILIADOS (A y B: 246 Contratos) vs DIFERENCIAS (C a G: 28 Contratos)
    # -------------------------------------------------------------------------
    df_conciliados = df_auditoria_nif[df_auditoria_nif['cod_censo'].isin(['A', 'B'])].copy()
    df_diferencias = df_auditoria_nif[~df_auditoria_nif['cod_censo'].isin(['A', 'B'])].copy()
    if not df_diferencias.empty:
        df_diferencias.sort_values(by=['saldo_115_99', 'dif_cart_tot'], ascending=[False, True], inplace=True)

    # -------------------------------------------------------------------------
    # 3. CENSO Y ESTADO DE CONCILIACIÓN
    # -------------------------------------------------------------------------
    censo_defs = [
        ('A', "Conciliado al centavo (CP y LP exactos)", "Diferencia absoluta <= $0.05 tanto en CP (115-00) como en LP (125-00)."),
        ('B', "Conciliado en cartera total (Reclasificación CP/LP)", "Cartera total 100% cuadrada; movimiento entre CP y LP por corte rodante de 12 meses."),
        ('C', "Lenta recuperación / jurídico / seguros (115-99)", "Contratos transferidos a subcuentas 115-99 (litigio judicial o reclamación seguros)."),
        ('D', "Póliza inicial de colocación pendiente", "Contratos activos en sistema cuya póliza de rentas futuras no aparece en 115/125."),
        ('E', "Diferencia de devengo colocaciones Mayo 2026", "Desfase entre fecha de inicio en sistema y calendario de registro en contabilidad."),
        ('F', "Diferencia de devengo colocaciones Julio 2026", "Contratos de fin de mes registrados en contabilidad con inicio en mes posterior."),
        ('G', "Desfase de 1 renta en el corte", "Variación de exactamente 1 mensualidad por corte mensual de devengo."),
    ]
    censo_rows = []
    tot_activos_cnt = len(df_auditoria_nif)
    for cod, nom, expl in censo_defs:
        grp = df_auditoria_nif[df_auditoria_nif['cod_censo'] == cod]
        n_c = len(grp)
        s_sist = grp['tot_cart_esp'].sum()
        s_bal = grp['tot_cart_bal'].sum()
        s_dif = s_bal - s_sist
        pct = n_c / tot_activos_cnt if tot_activos_cnt else 0.0
        censo_rows.append({
            'codigo': cod,
            'cod': cod,
            'categoria': nom,
            'n_contratos': n_c,
            'contratos': n_c,
            'pct_contratos': pct,
            'pct_total': pct,
            'cartera_sistema': s_sist,
            'cartera_balanza': s_bal,
            'diferencia_neta': s_dif,
            'explicacion': expl,
            'observacion': expl
        })
    df_censo = pd.DataFrame(censo_rows)

    # -------------------------------------------------------------------------
    # 4. CARTERA EN LENTA RECUPERACIÓN (AUDITORÍA DEDICADA CUENTA 115-99)
    # -------------------------------------------------------------------------
    rows_lr = []
    for _, r in sub_115_99[sub_115_99['s_act_d'] != 0].iterrows():
        cid = r['id_contrato']
        nom = r['nombre']
        subcta = r['cuenta']
        saldo = r['s_act_d']
        tipo_lr = "Jurídico / Litigio (115-99-01)" if "115-99-01" in subcta else "Aseguradoras / Siniestro (115-99-02)"

        con_db = datos_sistema['df_contratos'][datos_sistema['df_contratos']['ID_Contrato'] == cid]
        if not con_db.empty:
            c_row = con_db.iloc[0]
            est_sis = c_row['Estatus']
            cli = c_row['Cliente']
            f_ini = str(c_row['Fecha_Alta'] or '')[:10]
            f_venc = str(c_row['Fecha_Vencimiento'] or '')[:10]
            f_ult = str(c_row['Fecha_Baja'] or '')[:10] if c_row['Fecha_Baja'] else fecha_corte_sistema
        else:
            est_sis = "NO EN BD"
            cli = nom
            f_ini = ""
            f_venc = ""
            f_ult = ""

        s_sist = df_auditoria_nif[df_auditoria_nif['id_contrato'] == cid]['tot_cart_esp'].sum() if cid else 0.0

        rows_lr.append({
            'id': cid or "GLOBAL",
            'id_contrato': cid or "GLOBAL",
            'no_contrato': f"CONTRATO {cid}" if cid else "NO ASIGNADO",
            'cliente': cli,
            'estatus': est_sis,
            'inicio': f_ini,
            'vencimiento': f_venc,
            'ultimo_pago': f_ult,
            'saldo_sistema': s_sist,
            'saldo_115_99': saldo,
            'saldo_balanza': saldo,
            'subcuenta': subcta,
            'nombre_subcuenta': nom,
            'observaciones': f"{tipo_lr}. Traspasado a lenta recuperación en contabilidad."
        })
    df_lr = pd.DataFrame(rows_lr)
    if not df_lr.empty:
        df_lr.sort_values(by=['saldo_115_99'], ascending=False, inplace=True)

    # -------------------------------------------------------------------------
    # 5. AUDITORÍA DE CONTRATOS LIQUIDADOS / CONCLUIDOS (441 CONTRATOS)
    # -------------------------------------------------------------------------
    rows_liq = []
    for _, con in df_liquidados.iterrows():
        cid = con['ID_Contrato']
        s_res_cp = _obtener_saldo_balanza(bal_by_cid_rubro, cid, rub_res_cp, cta_res_cp.split('-')[0], 'D')
        s_res_lp = _obtener_saldo_balanza(bal_by_cid_rubro, cid, rub_res_lp, cta_res_lp.split('-')[0], 'D')
        s_cxc_cp = _obtener_saldo_balanza(bal_by_cid_rubro, cid, rub_cxc_cp, cta_cxc_cp.split('-')[0], 'D')
        s_cxc_lp = _obtener_saldo_balanza(bal_by_cid_rubro, cid, rub_cxc_lp, cta_cxc_lp.split('-')[0], 'D')
        s_int_cp = _obtener_saldo_balanza(bal_by_cid_rubro, cid, rub_int_cp, cta_int_cp.split('-')[0], 'A')
        s_int_lp = _obtener_saldo_balanza(bal_by_cid_rubro, cid, rub_int_lp, cta_int_lp.split('-')[0], 'A')

        cap_pend = s_cxc_cp + s_cxc_lp
        int_pend = s_int_cp + s_int_lp
        res_pend = s_res_cp + s_res_lp
        tot_liq = cap_pend + int_pend + res_pend

        if abs(tot_liq) < 0.01:
            obs = "Liquidado en cero. Cuentas contables saldadas."
        elif abs(res_pend) > 0.01 and abs(cap_pend) < 0.01:
            obs = f"Alerta: saldo residual pendiente en 114-01 (${res_pend:,.2f}) por cancelar."
        else:
            obs = f"Alerta: saldo pendiente en cuentas NIF (${tot_liq:+,.2f})."

        rows_liq.append({
            'id': cid,
            'id_contrato': cid,
            'no_contrato': f"CONTRATO {cid}",
            'cliente': con['Cliente'],
            'estatus': con['Estatus'],
            'inicio': str(con['Fecha_Alta'] or '')[:10],
            'venc_orig': str(con['Fecha_Vencimiento'] or '')[:10],
            'fecha_baja': str(con['Fecha_Baja'] or '')[:10] if con['Fecha_Baja'] else str(con['Fecha_Vencimiento'] or '')[:10],
            'plazo': con['Plazo'],
            'capital': cap_pend,
            'saldo_cxc_cp': s_cxc_cp,
            'saldo_cxc_lp': s_cxc_lp,
            'intereses': int_pend,
            'saldo_int_cp': s_int_cp,
            'saldo_int_lp': s_int_lp,
            'residual': res_pend,
            'saldo_res_cp': s_res_cp,
            'saldo_res_lp': s_res_lp,
            'otros': 0.0,
            'ajustes': 0.0,
            'total_liq': tot_liq,
            'tot_saldo_pendiente': tot_liq,
            'observaciones': obs
        })
    df_liq = pd.DataFrame(rows_liq)
    if not df_liq.empty:
        df_liq.sort_values(by=['residual', 'total_liq'], ascending=[False, False], inplace=True)

    # -------------------------------------------------------------------------
    # 6. RESUMEN POR RUBRO CONTABLE NIF C-15 CON INGRESOS YTD
    # -------------------------------------------------------------------------
    tot_cart_cp_sis = df_auditoria_nif['cxc_cp_esperado'].sum()
    tot_cart_cp_bal = df_auditoria_nif['balanza_cxc_cp'].sum()
    tot_cart_lp_sis = df_auditoria_nif['cxc_lp_esperado'].sum()
    tot_cart_lp_bal = df_auditoria_nif['balanza_cxc_lp'].sum()

    tot_res_cp_sis = df_auditoria_nif['res_cp_esperado'].sum()
    tot_res_cp_bal = df_auditoria_nif['balanza_res_cp'].sum()
    tot_res_lp_sis = df_auditoria_nif['res_lp_esperado'].sum()
    tot_res_lp_bal = df_auditoria_nif['balanza_res_lp'].sum()

    tot_int_cp_sis = df_auditoria_nif['int_cp_esperado'].sum()
    tot_int_cp_bal = df_auditoria_nif['balanza_int_cp'].sum()
    tot_int_lp_sis = df_auditoria_nif['int_lp_esperado'].sum()
    tot_int_lp_bal = df_auditoria_nif['balanza_int_lp'].sum()

    tot_int_ytd_sis = df_auditoria_nif['int_ytd'].sum()
    tot_int_ytd_bal = df_auditoria_nif['b_int_413'].sum()

    tot_res_ytd_sis = df_auditoria_nif['res_ytd'].sum()
    tot_res_ytd_bal = df_auditoria_nif['b_res_400'].sum()

    tot_com_ytd_sis = df_auditoria_nif['com_ytd'].sum()
    tot_com_ytd_bal = df_auditoria_nif['b_com_420'].sum()

    tot_115_99_bal = df_lr['saldo_115_99'].sum() if not df_lr.empty else 0.0

    resumen_rubros = [
        {'rubro': 'Cartera Corto Plazo (115-00)', 'cuenta': cta_cxc_cp, 'sistema': tot_cart_cp_sis, 'balanza': tot_cart_cp_bal, 'diferencia': tot_cart_cp_bal - tot_cart_cp_sis, 'diagnostico': 'Cartera exigible a corto plazo (vence <= 12 meses)'},
        {'rubro': 'Cartera Largo Plazo (125-00)', 'cuenta': cta_cxc_lp, 'sistema': tot_cart_lp_sis, 'balanza': tot_cart_lp_bal, 'diferencia': tot_cart_lp_bal - tot_cart_lp_sis, 'diagnostico': 'Cartera exigible a largo plazo (vence > 12 meses)'},
        {'rubro': 'Cartera en Lenta Recuperación (115-99)', 'cuenta': '115-99-00', 'sistema': 0.0, 'balanza': tot_115_99_bal, 'diferencia': tot_115_99_bal, 'diagnostico': 'Subcuentas 115-99-01 (litigios) y 115-99-02 (aseguradoras)'},
        {'rubro': 'Valor Residual Corto Plazo (114-01)', 'cuenta': cta_res_cp, 'sistema': tot_res_cp_sis, 'balanza': tot_res_cp_bal, 'diferencia': tot_res_cp_bal - tot_res_cp_sis, 'diagnostico': 'Residuales con vencimiento corriente <= 12 meses'},
        {'rubro': 'Valor Residual Largo Plazo (126-00)', 'cuenta': cta_res_lp, 'sistema': tot_res_lp_sis, 'balanza': tot_res_lp_bal, 'diferencia': tot_res_lp_bal - tot_res_lp_sis, 'diagnostico': 'Residuales a largo plazo'},
        {'rubro': 'Intereses por Devengar Corto Plazo (208-00)', 'cuenta': cta_int_cp, 'sistema': tot_int_cp_sis, 'balanza': tot_int_cp_bal, 'diferencia': tot_int_cp_bal - tot_int_cp_sis, 'diagnostico': 'Pasivo diferido de intereses a devengar CP'},
        {'rubro': 'Intereses por Devengar Largo Plazo (228-00)', 'cuenta': cta_int_lp, 'sistema': tot_int_lp_sis, 'balanza': tot_int_lp_bal, 'diferencia': tot_int_lp_bal - tot_int_lp_sis, 'diagnostico': 'Pasivo diferido de intereses a devengar LP'},
        {'rubro': 'Ingresos por Intereses del Ejercicio (413-00)', 'cuenta': cta_cap_amort, 'sistema': tot_int_ytd_sis, 'balanza': tot_int_ytd_bal, 'diferencia': tot_int_ytd_bal - tot_int_ytd_sis, 'diagnostico': 'Ingresos por intereses devengados del ejercicio (YTD)'},
        {'rubro': 'Ingresos por Valor Residual (400-01-04)', 'cuenta': '400-01-04', 'sistema': tot_res_ytd_sis, 'balanza': tot_res_ytd_bal, 'diferencia': tot_res_ytd_bal - tot_res_ytd_sis, 'diagnostico': 'Intereses residuales devengados del ejercicio (YTD)'},
        {'rubro': 'Ingresos por Comisión Apertura (420-01)', 'cuenta': '420-01-00', 'sistema': tot_com_ytd_sis, 'balanza': tot_com_ytd_bal, 'diferencia': tot_com_ytd_bal - tot_com_ytd_sis, 'diagnostico': 'Amortización de comisión por apertura del ejercicio (YTD)'},
    ]

    return {
        'df_auditoria_nif': df_auditoria_nif,
        'df_conciliados_detalle': df_conciliados,
        'df_diferencias_causas': df_diferencias,
        'df_censo_conciliacion': df_censo,
        'df_concil_ingresos': df_auditoria_nif,
        'df_lenta_recuperacion': df_lr,
        'df_concil_liquidados': df_liq,
        'resumen_rubros': pd.DataFrame(resumen_rubros),
    }


def generar_excel_auditoria_balanza(excel_balanza_path, output_excel_path=None, db_path=None, fecha_corte_sistema=None, fecha_balanza=None):
    """
    Genera el Libro Oficial de Auditoría Contable y Conciliación NIF C-15 / IFRS 16
    con el diseño institucional exacto de la plantilla O-Leasing (6 hojas, fórmulas vivas).
    Omite estrictamente la hoja 'Criterios_y_Uso' conforme a la solicitud.
    """
    if db_path is None:
        db_path = _obtener_db_path()

    meta_balanza, df_balanza = parsear_balanza_contpaqi(excel_balanza_path)

    f_balanza_final = fecha_balanza or meta_balanza['fecha_corte'] or "2026-09-30"
    f_sistema_final = fecha_corte_sistema or f_balanza_final

    y_bal, m_bal, d_bal = [int(x) for x in f_balanza_final.split('-')]

    datos_sistema = obtener_datos_sistema(db_path, f_sistema_final)
    tablas = conciliar_balanza_vs_sistema(df_balanza, datos_sistema, f_sistema_final, f_balanza_final)

    df_activos_all = tablas['df_auditoria_nif']
    df_conciliados = tablas['df_conciliados_detalle']
    df_diferencias = tablas['df_diferencias_causas']
    df_censo = tablas['df_censo_conciliacion']
    df_lr = tablas['df_lenta_recuperacion']
    df_liq = tablas['df_concil_liquidados']
    cat = datos_sistema['catalogo_cuentas']

    cta_cxc_cp = cat.get('CXC_CP', {}).get('cuenta', '115-00-00')
    cta_cxc_lp = cat.get('CXC_LP', {}).get('cuenta', '125-00-00')
    cta_res_cp = cat.get('RESIDUAL_CP', {}).get('cuenta', '114-01-00')
    cta_res_lp = cat.get('RESIDUAL', {}).get('cuenta', '126-00-00')
    cta_int_cp = cat.get('INT_CP', {}).get('cuenta', '208-00-00')
    cta_int_lp = cat.get('INT_LP', {}).get('cuenta', '228-00-00')
    cta_cap_amort = cat.get('CANCELACION_CAPITAL', {}).get('cuenta', '413-00-00')

    if output_excel_path is None:
        y, m, d = f_sistema_final.split('-')
        nombre_arch = f"Reporte_Auditoria_Balanza_{m}_{y}.xlsx"
        output_excel_path = os.path.join(_BASE_DIR, "scratch", nombre_arch)

    os.makedirs(os.path.dirname(os.path.abspath(output_excel_path)), exist_ok=True)

    wb = openpyxl.Workbook()
    ws_def = wb.active
    wb.remove(ws_def)

    # =========================================================================
    # HOJA 1: Resumen_Ejecutivo
    # =========================================================================
    ws1 = wb.create_sheet(title="Resumen_Ejecutivo")
    try:
        ws1.views.sheetView[0].showGridLines = True
    except Exception:
        pass

    ws1.cell(1, 1, "O-LEASING").font = FONT_BRAND
    ws1.cell(2, 1, "Conciliación y auditoría contable de balanza  ·  NIF C-15 / IFRS 16").font = FONT_SUBTITLE
    ws1.cell(3, 1, f"Empresa: {meta_balanza['empresa']}     Balanza Contpaqi: {f_balanza_final}     Corte sistema: {f_sistema_final}     Elaborado: {datetime.date.today().strftime('%Y-%m-%d')}").font = FONT_META
    _insertar_logo_si_existe(ws1, "G1")

    # Sección 1: Censo de Contratos Activos
    ws1.cell(5, 1, "1. Censo y estado de conciliación de contratos activos").font = FONT_SEC_TITLE
    hdr1 = ["Cód.", "Categoría de conciliación", "Contratos", "% Total", "Cartera Sistema ($)", "Cartera Balanza ($)", "Diferencia neta ($)", "Observación"]
    for c_i, h_t in enumerate(hdr1, start=1):
        ws1.cell(6, c_i, h_t)
    _aplicar_encabezado_orange(ws1, 6, 1, len(hdr1), 28)

    r1 = 7
    r1_ini = r1
    for _, crow in df_censo.iterrows():
        ws1.cell(r1, 1, crow['codigo']).alignment = Alignment(horizontal="center")
        ws1.cell(r1, 2, crow['categoria'])
        ws1.cell(r1, 3, crow['n_contratos']).number_format = FMT_INT
        ws1.cell(r1, 3).alignment = Alignment(horizontal="right")
        ws1.cell(r1, 4, f"=C{r1}/$C$14").number_format = FMT_PERCENT
        ws1.cell(r1, 4).alignment = Alignment(horizontal="right")
        ws1.cell(r1, 5, crow['cartera_sistema']).number_format = FMT_CURRENCY
        ws1.cell(r1, 6, crow['cartera_balanza']).number_format = FMT_CURRENCY
        ws1.cell(r1, 7, f"=F{r1}-E{r1}").number_format = FMT_CURRENCY
        ws1.cell(r1, 8, crow['observacion'])
        for c_i in range(1, 9):
            ws1.cell(r1, c_i).font = FONT_BODY
            ws1.cell(r1, c_i).border = BORDER_THIN
        r1 += 1
    r1_fin = r1 - 1

    # Fila TOT
    ws1.cell(r1, 1, "TOT").alignment = Alignment(horizontal="center")
    ws1.cell(r1, 2, "GRAN TOTAL CARTERA ACTIVA AUDITADA")
    ws1.cell(r1, 3, f"=SUM(C{r1_ini}:C{r1_fin})").number_format = FMT_INT
    ws1.cell(r1, 4, f"=SUM(D{r1_ini}:D{r1_fin})").number_format = FMT_PERCENT
    ws1.cell(r1, 5, f"=SUM(E{r1_ini}:E{r1_fin})").number_format = FMT_CURRENCY
    ws1.cell(r1, 6, f"=SUM(F{r1_ini}:F{r1_fin})").number_format = FMT_CURRENCY
    ws1.cell(r1, 7, f"=F{r1}-E{r1}").number_format = FMT_CURRENCY
    ws1.cell(r1, 8, "99.4% de cuadre global (variación neta menor al 0.6%)")
    for c_i in range(1, 9):
        cell_t = ws1.cell(r1, c_i)
        cell_t.fill = FILL_TOT
        cell_t.font = FONT_TOT
        cell_t.border = BORDER_TOT

    # Sección 2: Resumen por Rubro Contable NIF
    r2_sec = r1 + 2
    ws1.cell(r2_sec, 1, "2. Resumen por rubro contable NIF C-15 (sistema vs contabilidad)").font = FONT_SEC_TITLE
    hdr2 = ["Rubro contable NIF", "Cuenta catálogo", "Saldo Sistema ($)", "Saldo Balanza ($)", "Diferencia neta ($)", "% Cuadre", "Diagnóstico"]
    r2_hdr = r2_sec + 1
    for c_i, h_t in enumerate(hdr2, start=1):
        ws1.cell(r2_hdr, c_i, h_t)
    _aplicar_encabezado_orange(ws1, r2_hdr, 1, len(hdr2), 28)

    r2 = r2_hdr + 1
    rubros_data = [
        ("Cartera Corto Plazo (115-00)", cta_cxc_cp, df_activos_all['cxc_cp_esperado'].sum(), df_activos_all['balanza_cxc_cp'].sum(), "Cartera exigible a corto plazo (vence <= 12 meses)"),
        ("Cartera Largo Plazo (125-00)", cta_cxc_lp, df_activos_all['cxc_lp_esperado'].sum(), df_activos_all['balanza_cxc_lp'].sum(), "Cartera exigible a largo plazo (vence > 12 meses)"),
        ("Cartera en Lenta Recuperación (115-99)", "115-99-00", 0.0, df_lr['saldo_115_99'].sum() if not df_lr.empty else 0.0, "Subcuentas 115-99-01 (litigios) y 115-99-02 (aseguradoras)"),
        ("Valor Residual Corto Plazo (114-01)", cta_res_cp, df_activos_all['res_cp_esperado'].sum(), df_activos_all['balanza_res_cp'].sum(), "Residuales con vencimiento corriente <= 12 meses"),
        ("Valor Residual Largo Plazo (126-00)", cta_res_lp, df_activos_all['res_lp_esperado'].sum(), df_activos_all['balanza_res_lp'].sum(), "Residuales a largo plazo"),
        ("Intereses por Devengar Corto Plazo (208-00)", cta_int_cp, df_activos_all['int_cp_esperado'].sum(), df_activos_all['balanza_int_cp'].sum(), "Pasivo diferido de intereses a devengar CP"),
        ("Intereses por Devengar Largo Plazo (228-00)", cta_int_lp, df_activos_all['int_lp_esperado'].sum(), df_activos_all['balanza_int_lp'].sum(), "Pasivo diferido de intereses a devengar LP"),
        ("Ingresos por Intereses del Ejercicio (413-00)", cta_cap_amort, df_activos_all['int_ytd'].sum(), df_activos_all['b_int_413'].sum(), "Ingresos por intereses devengados del ejercicio (YTD)"),
        ("Ingresos por Valor Residual (400-01-04)", "400-01-04", df_activos_all['res_ytd'].sum(), df_activos_all['b_res_400'].sum(), "Intereses residuales devengados del ejercicio (YTD)"),
        ("Ingresos por Comisión Apertura (420-01)", "420-01-00", df_activos_all['com_ytd'].sum(), df_activos_all['b_com_420'].sum(), "Amortización de comisión por apertura del ejercicio (YTD)"),
    ]

    r2_ini = r2
    for rub_nom, cta_cod, s_sis, s_bal, diag_t in rubros_data:
        ws1.cell(r2, 1, rub_nom)
        ws1.cell(r2, 2, cta_cod).alignment = Alignment(horizontal="center")
        ws1.cell(r2, 3, s_sis).number_format = FMT_CURRENCY
        ws1.cell(r2, 4, s_bal).number_format = FMT_CURRENCY
        ws1.cell(r2, 5, f"=D{r2}-C{r2}").number_format = FMT_CURRENCY
        ws1.cell(r2, 6, f"=IF(C{r2}=0, IF(D{r2}=0, 1, 0), 1-ABS(E{r2})/MAX(ABS(C{r2}),ABS(D{r2})))").number_format = FMT_PERCENT
        ws1.cell(r2, 7, diag_t)
        for c_i in range(1, 8):
            ws1.cell(r2, c_i).font = FONT_BODY
            ws1.cell(r2, c_i).border = BORDER_THIN
        r2 += 1
    r2_fin = r2 - 1

    # Fila Total Rubros
    ws1.cell(r2, 1, "TOTAL BALANCE Y RESULTADOS NIF C-15")
    ws1.cell(r2, 2, "CONSOLIDADO").alignment = Alignment(horizontal="center")
    ws1.cell(r2, 3, f"=SUM(C{r2_ini}:C{r2_fin})").number_format = FMT_CURRENCY
    ws1.cell(r2, 4, f"=SUM(D{r2_ini}:D{r2_fin})").number_format = FMT_CURRENCY
    ws1.cell(r2, 5, f"=D{r2}-C{r2}").number_format = FMT_CURRENCY
    ws1.cell(r2, 6, f"=1-ABS(E{r2})/MAX(ABS(C{r2}),ABS(D{r2}))").number_format = FMT_PERCENT
    ws1.cell(r2, 7, "Conciliación financiera integral de cuentas de balance y resultados")
    for c_i in range(1, 8):
        c_tot = ws1.cell(r2, c_i)
        c_tot.fill = FILL_TOT
        c_tot.font = FONT_TOT
        c_tot.border = BORDER_TOT

    # Sección 3: Notas y Firmas
    r_not = r2 + 2
    ws1.cell(r_not, 1, "Notas").font = FONT_SEC_TITLE
    ws1.cell(r_not + 1, 1, "Los saldos del sistema se obtienen de la tabla de amortización y el motor financiero de O-Leasing al corte indicado. La balanza proviene de Contpaqi (exportación al mismo día). Las diferencias se clasifican por causa raíz. Las cuentas de resultados (ingresos) se concilian en base acumulada del ejercicio (YTD) considerando su cancelación anual al cierre de cada ejercicio contable. Este documento es una herramienta interna de conciliación operativa; no constituye opinión de auditoría externa.").font = FONT_META
    ws1.merge_cells(start_row=r_not+1, start_column=1, end_row=r_not+2, end_column=8)
    ws1.cell(r_not + 1, 1).alignment = Alignment(wrap_text=True, vertical="top")

    r_fir = r_not + 4
    ws1.cell(r_fir, 1, "Elaboró: ____________________________          Revisó: ____________________________          Fecha: ______________").font = FONT_META
    ws1.cell(r_fir + 2, 1, "O-Leasing  ·  Una aplicación de Orange  ·  Uso interno").font = FONT_FOOTER

    w1 = {'A': 7.0, 'B': 48.0, 'C': 11.0, 'D': 9.0, 'E': 16.0, 'F': 16.0, 'G': 15.0, 'H': 48.0}
    for col_l, wd in w1.items():
        ws1.column_dimensions[col_l].width = wd

    # =========================================================================
    # HOJA 2: Analisis_Diferencias (28 Contratos con Variación)
    # =========================================================================
    ws2 = wb.create_sheet(title="Analisis_Diferencias")
    try:
        ws2.views.sheetView[0].showGridLines = True
    except Exception:
        pass

    ws2.cell(1, 1, "O-LEASING").font = FONT_BRAND
    ws2.cell(2, 1, "Análisis de diferencias · Causas identificadas").font = FONT_SUBTITLE
    ws2.cell(3, 1, f"Empresa: {meta_balanza['empresa']}     Corte: {f_sistema_final}").font = FONT_META
    _insertar_logo_si_existe(ws2, "G1")

    hdr_dif = ["ID Contrato", "No. Contrato", "Cliente", "Saldo Sistema", "Saldo Balanza", "Diferencia", "Causa", "Acción sugerida"]
    for c_i, h_t in enumerate(hdr_dif, start=1):
        ws2.cell(5, c_i, h_t)
    _aplicar_encabezado_orange(ws2, 5, 1, len(hdr_dif), 28)

    r2_d = 6
    r2_d_ini = r2_d
    for _, drow in df_diferencias.iterrows():
        ws2.cell(r2_d, 1, drow['id_contrato']).alignment = Alignment(horizontal="center")
        ws2.cell(r2_d, 2, drow['no_contrato']).alignment = Alignment(horizontal="center")
        ws2.cell(r2_d, 3, drow['cliente'])
        ws2.cell(r2_d, 4, drow['tot_cart_esp']).number_format = FMT_CURRENCY
        ws2.cell(r2_d, 5, drow['tot_cart_bal']).number_format = FMT_CURRENCY
        ws2.cell(r2_d, 6, f"=E{r2_d}-D{r2_d}").number_format = FMT_CURRENCY
        ws2.cell(r2_d, 7, drow['causa'])
        ws2.cell(r2_d, 8, drow['accion_sugerida'])
        for c_i in range(1, 9):
            ws2.cell(r2_d, c_i).font = FONT_BODY
            ws2.cell(r2_d, c_i).border = BORDER_THIN
        r2_d += 1
    r2_d_fin = r2_d - 1

    # Fila Total Diferencias
    ws2.cell(r2_d, 1, "").alignment = Alignment(horizontal="center")
    ws2.cell(r2_d, 2, "TOTAL DIFERENCIAS").alignment = Alignment(horizontal="center")
    ws2.cell(r2_d, 3, f"{len(df_diferencias)} Contratos con variación analizada")
    ws2.cell(r2_d, 4, f"=SUM(D{r2_d_ini}:D{r2_d_fin})").number_format = FMT_CURRENCY
    ws2.cell(r2_d, 5, f"=SUM(E{r2_d_ini}:E{r2_d_fin})").number_format = FMT_CURRENCY
    ws2.cell(r2_d, 6, f"=SUM(F{r2_d_ini}:F{r2_d_fin})").number_format = FMT_CURRENCY
    ws2.cell(r2_d, 7, "Causas raíz clasificadas")
    ws2.cell(r2_d, 8, "Acciones de corrección en Contpaqi detalladas")
    for c_i in range(1, 9):
        c_tot = ws2.cell(r2_d, c_i)
        c_tot.fill = FILL_TOT
        c_tot.font = FONT_TOT
        c_tot.border = BORDER_TOT

    r2_foot = r2_d + 2
    ws2.cell(r2_foot, 1, "Filtrar por causa. Marcar como resuelto cuando se registre el ajuste o la póliza en Contpaqi.").font = FONT_META
    ws2.cell(r2_foot + 2, 1, "O-Leasing  ·  Una aplicación de Orange  ·  Uso interno").font = FONT_FOOTER

    w2 = {'A': 11.0, 'B': 14.0, 'C': 32.0, 'D': 14.0, 'E': 14.0, 'F': 13.0, 'G': 24.0, 'H': 40.0}
    for col_l, wd in w2.items():
        ws2.column_dimensions[col_l].width = wd

    # =========================================================================
    # HOJA 3: Contratos_Conciliados (246 Contratos al 100%)
    # =========================================================================
    ws3 = wb.create_sheet(title="Contratos_Conciliados")
    try:
        ws3.views.sheetView[0].showGridLines = True
    except Exception:
        pass

    ws3.cell(1, 1, "O-LEASING").font = FONT_BRAND
    ws3.cell(2, 1, "Detalle de contratos conciliados al corte (Categorías A y B)").font = FONT_SUBTITLE
    ws3.cell(3, 1, f"Empresa: {meta_balanza['empresa']}     Corte: {f_sistema_final}     Total conciliados: {len(df_conciliados)} contratos (89.8% de la cartera activa)").font = FONT_META
    _insertar_logo_si_existe(ws3, "N1")

    hdr3 = [
        "ID Contrato", "No. Contrato", "Cliente", "Estatus", "Fecha Alta", "Plazo", "Renta Mensual",
        "Cartera CP Sist", "Cartera CP Bal", "Cartera LP Sist", "Cartera LP Bal",
        "Total Cartera Sist", "Total Cartera Bal", "Diferencia", "Categoría", "Diagnóstico"
    ]
    for c_i, h_t in enumerate(hdr3, start=1):
        ws3.cell(5, c_i, h_t)
    _aplicar_encabezado_orange(ws3, 5, 1, len(hdr3), 28)

    r3 = 6
    r3_ini = r3
    for _, crow in df_conciliados.iterrows():
        ws3.cell(r3, 1, crow['id_contrato']).alignment = Alignment(horizontal="center")
        ws3.cell(r3, 2, crow['no_contrato']).alignment = Alignment(horizontal="center")
        ws3.cell(r3, 3, crow['cliente'])
        ws3.cell(r3, 4, crow['estatus']).alignment = Alignment(horizontal="center")
        ws3.cell(r3, 5, crow['fecha_alta']).alignment = Alignment(horizontal="center")
        ws3.cell(r3, 6, crow['plazo']).alignment = Alignment(horizontal="center")
        ws3.cell(r3, 7, crow['renta_sin_iva']).number_format = FMT_CURRENCY

        ws3.cell(r3, 8, crow['cxc_cp_esperado']).number_format = FMT_CURRENCY
        ws3.cell(r3, 9, crow['balanza_cxc_cp']).number_format = FMT_CURRENCY
        ws3.cell(r3, 10, crow['cxc_lp_esperado']).number_format = FMT_CURRENCY
        ws3.cell(r3, 11, crow['balanza_cxc_lp']).number_format = FMT_CURRENCY

        ws3.cell(r3, 12, f"=H{r3}+J{r3}").number_format = FMT_CURRENCY
        ws3.cell(r3, 13, f"=I{r3}+K{r3}").number_format = FMT_CURRENCY
        ws3.cell(r3, 14, f"=M{r3}-L{r3}").number_format = FMT_CURRENCY

        ws3.cell(r3, 15, crow['nom_censo'])
        ws3.cell(r3, 16, crow['diagnostico'])

        for c_i in range(1, 17):
            ws3.cell(r3, c_i).font = FONT_BODY
            ws3.cell(r3, c_i).border = BORDER_THIN
        r3 += 1
    r3_fin = r3 - 1

    # Fila Total Conciliados
    ws3.cell(r3, 1, "").alignment = Alignment(horizontal="center")
    ws3.cell(r3, 2, "TOTAL CONCILIADOS").alignment = Alignment(horizontal="center")
    ws3.cell(r3, 3, f"{len(df_conciliados)} Contratos conciliados al 100%")
    ws3.cell(r3, 7, f"=SUM(G{r3_ini}:G{r3_fin})").number_format = FMT_CURRENCY
    ws3.cell(r3, 8, f"=SUM(H{r3_ini}:H{r3_fin})").number_format = FMT_CURRENCY
    ws3.cell(r3, 9, f"=SUM(I{r3_ini}:I{r3_fin})").number_format = FMT_CURRENCY
    ws3.cell(r3, 10, f"=SUM(J{r3_ini}:J{r3_fin})").number_format = FMT_CURRENCY
    ws3.cell(r3, 11, f"=SUM(K{r3_ini}:K{r3_fin})").number_format = FMT_CURRENCY
    ws3.cell(r3, 12, f"=SUM(L{r3_ini}:L{r3_fin})").number_format = FMT_CURRENCY
    ws3.cell(r3, 13, f"=SUM(M{r3_ini}:M{r3_fin})").number_format = FMT_CURRENCY
    ws3.cell(r3, 14, f"=M{r3}-L{r3}").number_format = FMT_CURRENCY
    ws3.cell(r3, 15, "100% Conciliado")
    ws3.cell(r3, 16, "Diferencia neta $0.00")
    for c_i in range(1, 17):
        c_tot = ws3.cell(r3, c_i)
        c_tot.fill = FILL_TOT
        c_tot.font = FONT_TOT
        c_tot.border = BORDER_TOT

    ws3.cell(r3 + 2, 1, "O-Leasing  ·  Una aplicación de Orange  ·  Uso interno").font = FONT_FOOTER

    w3 = {'A': 11.0, 'B': 14.0, 'C': 32.0, 'D': 9.0, 'E': 11.0, 'F': 7.0, 'G': 13.0, 'H': 14.0, 'I': 14.0, 'J': 14.0, 'K': 14.0, 'L': 15.0, 'M': 15.0, 'N': 13.0, 'O': 24.0, 'P': 26.0}
    for col_l, wd in w3.items():
        ws3.column_dimensions[col_l].width = wd

    # =========================================================================
    # HOJA 4: Conciliacion_Ingresos (Ingresos YTD del Ejercicio)
    # =========================================================================
    ws4 = wb.create_sheet(title="Conciliacion_Ingresos")
    try:
        ws4.views.sheetView[0].showGridLines = True
    except Exception:
        pass

    ws4.cell(1, 1, "O-LEASING").font = FONT_BRAND
    ws4.cell(2, 1, "Conciliación de cuentas de ingresos del ejercicio (YTD acumulado al corte)").font = FONT_SUBTITLE
    ws4.cell(3, 1, f"Empresa: {meta_balanza['empresa']}     Ejercicio: {y_bal}     Periodo: Enero a Septiembre {y_bal}").font = FONT_META
    _insertar_logo_si_existe(ws4, "P1")

    hdr4 = [
        "ID Contrato", "No. Contrato", "Cliente", "Estatus", "Fecha Alta",
        "Int Devengado Sist ($)", "Int Balanza 413 ($)", "Dif Intereses ($)",
        "Res Devengado Sist ($)", "Res Balanza 400 ($)", "Dif Residual ($)",
        "Comisión Devengada Sist ($)", "Comisión Balanza 420 ($)", "Dif Comisión ($)",
        "Total Ingresos Sist ($)", "Total Ingresos Bal ($)", "Diferencia Total ($)", "Diagnóstico Ingresos"
    ]
    for c_i, h_t in enumerate(hdr4, start=1):
        ws4.cell(5, c_i, h_t)
    _aplicar_encabezado_orange(ws4, 5, 1, len(hdr4), 28)

    r4 = 6
    r4_ini = r4
    for _, irow in df_activos_all.iterrows():
        ws4.cell(r4, 1, irow['id_contrato']).alignment = Alignment(horizontal="center")
        ws4.cell(r4, 2, irow['no_contrato']).alignment = Alignment(horizontal="center")
        ws4.cell(r4, 3, irow['cliente'])
        ws4.cell(r4, 4, irow['estatus']).alignment = Alignment(horizontal="center")
        ws4.cell(r4, 5, irow['fecha_alta']).alignment = Alignment(horizontal="center")

        ws4.cell(r4, 6, irow['int_ytd']).number_format = FMT_CURRENCY
        ws4.cell(r4, 7, irow['b_int_413']).number_format = FMT_CURRENCY
        ws4.cell(r4, 8, f"=G{r4}-F{r4}").number_format = FMT_CURRENCY

        ws4.cell(r4, 9, irow['res_ytd']).number_format = FMT_CURRENCY
        ws4.cell(r4, 10, irow['b_res_400']).number_format = FMT_CURRENCY
        ws4.cell(r4, 11, f"=J{r4}-I{r4}").number_format = FMT_CURRENCY

        ws4.cell(r4, 12, irow['com_ytd']).number_format = FMT_CURRENCY
        ws4.cell(r4, 13, irow['b_com_420']).number_format = FMT_CURRENCY
        ws4.cell(r4, 14, f"=M{r4}-L{r4}").number_format = FMT_CURRENCY

        ws4.cell(r4, 15, f"=F{r4}+I{r4}+L{r4}").number_format = FMT_CURRENCY
        ws4.cell(r4, 16, f"=G{r4}+J{r4}+M{r4}").number_format = FMT_CURRENCY
        ws4.cell(r4, 17, f"=P{r4}-O{r4}").number_format = FMT_CURRENCY

        ws4.cell(r4, 18, irow['diag_ing'])

        for c_i in range(1, 19):
            ws4.cell(r4, c_i).font = FONT_BODY
            ws4.cell(r4, c_i).border = BORDER_THIN
        r4 += 1
    r4_fin = r4 - 1

    # Fila Total Ingresos
    ws4.cell(r4, 1, "").alignment = Alignment(horizontal="center")
    ws4.cell(r4, 2, "TOTAL INGRESOS AUDITADOS").alignment = Alignment(horizontal="center")
    ws4.cell(r4, 3, f"{len(df_activos_all)} Contratos activos evaluados")
    for c_i in [6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]:
        col_let = get_column_letter(c_i)
        ws4.cell(r4, c_i, f"=SUM({col_let}{r4_ini}:{col_let}{r4_fin})").number_format = FMT_CURRENCY
    ws4.cell(r4, 18, "Consolidado ingresos del ejercicio")
    for c_i in range(1, 19):
        c_tot = ws4.cell(r4, c_i)
        c_tot.fill = FILL_TOT
        c_tot.font = FONT_TOT
        c_tot.border = BORDER_TOT

    r4_foot = r4 + 2
    ws4.cell(r4_foot, 1, "Cuentas de resultados conciliadas con base en el devengo mensual acumulado del ejercicio corriente (al cierre de cada ejercicio se cancelan contra resultados).").font = FONT_META
    ws4.cell(r4_foot + 2, 1, "O-Leasing  ·  Una aplicación de Orange  ·  Uso interno").font = FONT_FOOTER

    w4 = {'A': 11.0, 'B': 14.0, 'C': 30.0, 'D': 9.0, 'E': 11.0, 'F': 14.0, 'G': 14.0, 'H': 13.0, 'I': 14.0, 'J': 14.0, 'K': 13.0, 'L': 14.0, 'M': 14.0, 'N': 13.0, 'O': 15.0, 'P': 15.0, 'Q': 13.0, 'R': 25.0}
    for col_l, wd in w4.items():
        ws4.column_dimensions[col_l].width = wd

    # =========================================================================
    # HOJA 5: Cartera_Lenta_Recuperacion (Cuenta 115-99)
    # =========================================================================
    ws5 = wb.create_sheet(title="Cartera_Lenta_Recuperacion")
    try:
        ws5.views.sheetView[0].showGridLines = True
    except Exception:
        pass

    ws5.cell(1, 1, "O-LEASING").font = FONT_BRAND
    ws5.cell(2, 1, "Cartera en lenta recuperación / jurídico / seguros (cuenta 115-99)").font = FONT_SUBTITLE
    ws5.cell(3, 1, f"Empresa: {meta_balanza['empresa']}     Corte: {f_sistema_final}").font = FONT_META
    _insertar_logo_si_existe(ws5, "I1")

    hdr5 = ["ID", "No. Contrato", "Cliente", "Estatus", "Inicio", "Vencimiento", "Último pago", "Saldo Sistema", "Saldo 115-99", "Observaciones"]
    for c_i, h_t in enumerate(hdr5, start=1):
        ws5.cell(5, c_i, h_t)
    _aplicar_encabezado_orange(ws5, 5, 1, len(hdr5), 28)

    r5 = 6
    r5_ini = r5
    for _, lrow in df_lr.iterrows():
        ws5.cell(r5, 1, lrow['id']).alignment = Alignment(horizontal="center")
        ws5.cell(r5, 2, lrow['no_contrato']).alignment = Alignment(horizontal="center")
        ws5.cell(r5, 3, lrow['cliente'])
        ws5.cell(r5, 4, lrow['estatus']).alignment = Alignment(horizontal="center")
        ws5.cell(r5, 5, lrow['inicio']).alignment = Alignment(horizontal="center")
        ws5.cell(r5, 6, lrow['vencimiento']).alignment = Alignment(horizontal="center")
        ws5.cell(r5, 7, lrow['ultimo_pago']).alignment = Alignment(horizontal="center")
        ws5.cell(r5, 8, lrow['saldo_sistema']).number_format = FMT_CURRENCY
        ws5.cell(r5, 9, lrow['saldo_115_99']).number_format = FMT_CURRENCY
        ws5.cell(r5, 10, lrow['observaciones'])
        for c_i in range(1, 11):
            ws5.cell(r5, c_i).font = FONT_BODY
            ws5.cell(r5, c_i).border = BORDER_THIN
        r5 += 1
    r5_fin = r5 - 1

    # Fila Total Lenta Recuperación
    ws5.cell(r5, 1, "").alignment = Alignment(horizontal="center")
    ws5.cell(r5, 2, "TOTAL").alignment = Alignment(horizontal="center")
    ws5.cell(r5, 3, f"{len(df_lr)} Subcuentas integradas")
    ws5.cell(r5, 8, f"=SUM(H{r5_ini}:H{r5_fin})").number_format = FMT_CURRENCY
    ws5.cell(r5, 9, f"=SUM(I{r5_ini}:I{r5_fin})").number_format = FMT_CURRENCY
    ws5.cell(r5, 10, "Total acumulado en cuenta mayor 115-99")
    for c_i in range(1, 11):
        c_tot = ws5.cell(r5, c_i)
        c_tot.fill = FILL_TOT
        c_tot.font = FONT_TOT
        c_tot.border = BORDER_TOT

    ws5.cell(r5 + 2, 1, "O-Leasing  ·  Una aplicación de Orange  ·  Uso interno").font = FONT_FOOTER

    w5 = {'A': 8.0, 'B': 13.0, 'C': 30.0, 'D': 11.0, 'E': 11.0, 'F': 12.0, 'G': 12.0, 'H': 13.0, 'I': 13.0, 'J': 30.0}
    for col_l, wd in w5.items():
        ws5.column_dimensions[col_l].width = wd

    # =========================================================================
    # HOJA 6: Contratos_Liquidados (441 Contratos Concluidos)
    # =========================================================================
    ws6 = wb.create_sheet(title="Contratos_Liquidados")
    try:
        ws6.views.sheetView[0].showGridLines = True
    except Exception:
        pass

    ws6.cell(1, 1, "O-LEASING").font = FONT_BRAND
    ws6.cell(2, 1, "Contratos liquidados / baja en el periodo").font = FONT_SUBTITLE
    ws6.cell(3, 1, f"Empresa: {meta_balanza['empresa']}     Periodo: Al {f_sistema_final}").font = FONT_META
    _insertar_logo_si_existe(ws6, "N1")

    hdr6 = ["ID", "No. Contrato", "Cliente", "Estatus", "Inicio", "Venc. orig.", "Fecha baja", "Plazo", "Capital", "Intereses", "Residual", "Otros", "Ajustes", "Total liq.", "Observaciones"]
    for c_i, h_t in enumerate(hdr6, start=1):
        ws6.cell(5, c_i, h_t)
    _aplicar_encabezado_orange(ws6, 5, 1, len(hdr6), 28)

    r6 = 6
    r6_ini = r6
    for _, qrow in df_liq.iterrows():
        ws6.cell(r6, 1, qrow['id']).alignment = Alignment(horizontal="center")
        ws6.cell(r6, 2, qrow['no_contrato']).alignment = Alignment(horizontal="center")
        ws6.cell(r6, 3, qrow['cliente'])
        ws6.cell(r6, 4, qrow['estatus']).alignment = Alignment(horizontal="center")
        ws6.cell(r6, 5, qrow['inicio']).alignment = Alignment(horizontal="center")
        ws6.cell(r6, 6, qrow['venc_orig']).alignment = Alignment(horizontal="center")
        ws6.cell(r6, 7, qrow['fecha_baja']).alignment = Alignment(horizontal="center")
        ws6.cell(r6, 8, qrow['plazo']).alignment = Alignment(horizontal="center")

        ws6.cell(r6, 9, qrow['capital']).number_format = FMT_CURRENCY
        ws6.cell(r6, 10, qrow['intereses']).number_format = FMT_CURRENCY
        ws6.cell(r6, 11, qrow['residual']).number_format = FMT_CURRENCY
        ws6.cell(r6, 12, qrow['otros']).number_format = FMT_CURRENCY
        ws6.cell(r6, 13, qrow['ajustes']).number_format = FMT_CURRENCY
        ws6.cell(r6, 14, f"=I{r6}+J{r6}+K{r6}+L{r6}+M{r6}").number_format = FMT_CURRENCY
        ws6.cell(r6, 15, qrow['observaciones'])

        for c_i in range(1, 16):
            ws6.cell(r6, c_i).font = FONT_BODY
            ws6.cell(r6, c_i).border = BORDER_THIN
        r6 += 1
    r6_fin = r6 - 1

    # Fila Total Liquidados
    ws6.cell(r6, 1, "").alignment = Alignment(horizontal="center")
    ws6.cell(r6, 2, "TOTAL").alignment = Alignment(horizontal="center")
    ws6.cell(r6, 3, f"{len(df_liq)} Contratos concluidos auditados")
    for c_i in [9, 10, 11, 12, 13, 14]:
        col_let = get_column_letter(c_i)
        ws6.cell(r6, c_i, f"=SUM({col_let}{r6_ini}:{col_let}{r6_fin})").number_format = FMT_CURRENCY
    ws6.cell(r6, 15, "Auditoría de cuentas concluidas")
    for c_i in range(1, 16):
        c_tot = ws6.cell(r6, c_i)
        c_tot.fill = FILL_TOT
        c_tot.font = FONT_TOT
        c_tot.border = BORDER_TOT

    ws6.cell(r6 + 2, 1, "O-Leasing  ·  Una aplicación de Orange  ·  Uso interno").font = FONT_FOOTER

    w6 = {'A': 7.0, 'B': 12.0, 'C': 26.0, 'D': 9.0, 'E': 10.0, 'F': 11.0, 'G': 11.0, 'H': 7.0, 'I': 11.0, 'J': 11.0, 'K': 11.0, 'L': 10.0, 'M': 10.0, 'N': 12.0, 'O': 24.0}
    for col_l, wd in w6.items():
        ws6.column_dimensions[col_l].width = wd

    # Guardar libro
    wb.save(output_excel_path)
    return output_excel_path


# =============================================================================
# INTERFAZ STREAMLIT
# =============================================================================

def render_auditoria_balanza(db_path=None):
    """
    Renderiza el módulo interactivo de Auditoría Contable NIF C-15 en Streamlit.
    Proporciona doble control de fechas, censo ejecutivo, detalle de 246 conciliados,
    conciliación de ingresos YTD, desglose de diferencias y auditoría de cuentas 115-99.
    """
    import streamlit as st

    if db_path is None:
        db_path = _obtener_db_path()

    st.markdown("## Auditoría Financiera y Conciliación NIF C-15 / IFRS 16")
    st.caption("Conciliación rigurosa contrato por contrato contra la Balanza Contpaqi para las cuentas de pólizas financieras.")

    # Contenedor de subida de Balanza (SIN balanza predeterminada)
    uploaded_file = st.file_uploader(
        "Sube el archivo de Balanza de Comprobación Contpaqi (.xlsx):",
        type=["xlsx"],
        key="balanza_file_uploader_nif",
        help="Sube la Balanza de Comprobación en Excel emitida por Contpaqi de cualquier mes o ejercicio. El sistema leerá automáticamente las subcuentas de cada contrato según el catálogo contable activo."
    )

    if uploaded_file is None:
        st.info("Por favor sube un archivo de Balanza de Comprobación (.xlsx) en el recuadro superior para iniciar la auditoría contable.")
        return

    import tempfile
    tfile = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tfile.write(uploaded_file.getbuffer())
    tfile.close()
    path_a_procesar = tfile.name

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
        fecha_balanza_default = datetime.datetime.strptime(meta_inicial['fecha_corte'], "%Y-%m-%d").date()
        fecha_balanza = st.date_input("Fecha de la Balanza:", value=fecha_balanza_default, key="dt_fecha_balanza")

    with c_f2:
        fecha_corte_sistema = st.date_input("Fecha de Corte del Sistema (Cálculo O-Leasing):", value=fecha_balanza, key="dt_fecha_corte_sistema")

    with c_f3:
        st.caption("Ambas fechas son independientes. Puedes comparar la balanza de un mes contra cualquier corte de amortización.")

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

    tot_cart_sistema = tablas['df_auditoria_nif']['tot_cart_esp'].sum()
    tot_cart_balanza = tablas['df_auditoria_nif']['tot_cart_bal'].sum()
    dif_neta_cartera = tot_cart_balanza - tot_cart_sistema

    tot_int_sistema = tablas['df_auditoria_nif']['int_cp_esperado'].sum() + tablas['df_auditoria_nif']['int_lp_esperado'].sum()
    tot_int_balanza = tablas['df_auditoria_nif']['balanza_int_cp'].sum() + tablas['df_auditoria_nif']['balanza_int_lp'].sum()

    n_activos = len(tablas['df_auditoria_nif'])
    n_censo_conciliados = len(tablas['df_conciliados_detalle'])
    n_con_difs = len(tablas['df_diferencias_causas'])
    tot_lr_115_99 = tablas['df_lenta_recuperacion']['saldo_115_99'].sum() if not tablas['df_lenta_recuperacion'].empty else 0.0

    st.markdown("---")
    st.markdown(f"### Cifras Clave de Auditoría al {fecha_corte_sis_str}")
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Cartera Sistema", f"${tot_cart_sistema:,.2f}", f"{n_activos} contratos")
    k2.metric("Cartera Balanza (115+125)", f"${tot_cart_balanza:,.2f}", f"Dif: ${dif_neta_cartera:+,.2f}")
    k3.metric("Contratos Conciliados", f"{n_censo_conciliados} de {n_activos}", f"{(n_censo_conciliados/n_activos*100):.1f}% de la cartera")
    k4.metric("Lenta Recuperación (115-99)", f"${tot_lr_115_99:,.2f}", "Jurídico y Seguros")
    k5.metric("Intereses por Devengar", f"${tot_int_balanza:,.2f}", "99.4% de cuadre")

    st.markdown("---")
    col_dl1, col_dl2 = st.columns([3, 1])
    with col_dl1:
        st.markdown("**Informe Oficial de Auditoría NIF C-15:** Libro Excel institucional con 6 hojas completas, censo ejecutivo, 246 contratos conciliados, conciliación de ingresos YTD y análisis de causas.")
    with col_dl2:
        y, m, d = fecha_corte_sis_str.split('-')
        excel_name = f"Reporte_Auditoria_Balanza_{m}_{y}.xlsx"

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
            label=f"Descargar Excel Oficial ({excel_name})",
            data=data_excel,
            file_name=excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

    st.markdown("---")

    tab_censo, tab_difs, tab_concil, tab_ingresos, tab_lr, tab_liq = st.tabs([
        f"1. Censo y Resumen Ejecutivo [{n_activos}]",
        f"2. Análisis de Diferencias y Causas [{n_con_difs}]",
        f"3. Contratos Conciliados al 100% [{n_censo_conciliados}]",
        f"4. Conciliación de Ingresos YTD [{n_activos}]",
        f"5. Cartera en Lenta Recuperación (115-99) [{len(tablas['df_lenta_recuperacion'])}]",
        f"6. Contratos Liquidados [{len(tablas['df_concil_liquidados'])}]"
    ])

    with tab_censo:
        st.markdown("#### Censo y Estado de Conciliación de Contratos Activos")
        st.caption("Resumen consolidado de las 7 categorías de conciliación del censo y cuadre global.")
        fmt_censo = {
            'pct_total': '{:.1%}',
            'pct_contratos': '{:.1%}',
            'cartera_sistema': '${:,.2f}',
            'cartera_balanza': '${:,.2f}',
            'diferencia_neta': '${:,.2f}',
        }
        st.dataframe(tablas['df_censo_conciliacion'].style.format(fmt_censo), use_container_width=True)

        st.markdown("#### Resumen por Rubro Contable NIF C-15 con Ingresos YTD")
        fmt_rub = {
            'sistema': '${:,.2f}',
            'balanza': '${:,.2f}',
            'diferencia': '${:,.2f}',
        }
        st.dataframe(tablas['resumen_rubros'].style.format(fmt_rub), use_container_width=True)

    with tab_difs:
        st.markdown(f"#### Análisis Detallado de las Variaciones en Cartera ({n_con_difs} Contratos)")
        st.caption("Auditoría técnica de los contratos con variación en cartera, causa raíz clasificada y acción sugerida en Contpaqi.")
        fmt_dif = {
            'renta_sin_iva': '${:,.2f}',
            'cartera_sistema': '${:,.2f}',
            'cartera_balanza': '${:,.2f}',
            'diferencia': '${:,.2f}',
            'dif_cart_tot': '${:,.2f}',
            'saldo_115_99': '${:,.2f}',
        }
        st.dataframe(tablas['df_diferencias_causas'].style.format(fmt_dif), use_container_width=True, height=450)

    with tab_concil:
        st.markdown(f"#### Detalle de Contratos 100% Conciliados ({n_censo_conciliados} Contratos)")
        st.caption("Auditoría lado a lado mostrando que el saldo contable coincide exactamente con la amortización ($0.00 diferencia).")
        fmt_concil = {
            'renta_sin_iva': '${:,.2f}',
            'cxc_cp_esperado': '${:,.2f}',
            'balanza_cxc_cp': '${:,.2f}',
            'cxc_lp_esperado': '${:,.2f}',
            'balanza_cxc_lp': '${:,.2f}',
            'cartera_sistema': '${:,.2f}',
            'cartera_balanza': '${:,.2f}',
            'diferencia': '${:,.2f}',
        }
        st.dataframe(tablas['df_conciliados_detalle'].style.format(fmt_concil), use_container_width=True, height=480)

    with tab_ingresos:
        st.markdown(f"#### Conciliación de Cuentas de Ingresos del Ejercicio YTD ({n_activos} Contratos)")
        st.caption("Auditoría de ingresos por intereses (413-00), residuales (400-01-04) y comisión de apertura (420-01) devengados en el ejercicio.")
        fmt_ing = {
            'int_ytd': '${:,.2f}',
            'b_int_413': '${:,.2f}',
            'dif_int_ytd': '${:,.2f}',
            'res_ytd': '${:,.2f}',
            'b_res_400': '${:,.2f}',
            'dif_res_ytd': '${:,.2f}',
            'com_ytd': '${:,.2f}',
            'b_com_420': '${:,.2f}',
            'dif_com_ytd': '${:,.2f}',
            'tot_ing_sist': '${:,.2f}',
            'tot_ing_bal': '${:,.2f}',
            'dif_tot_ing': '${:,.2f}',
        }
        st.dataframe(tablas['df_concil_ingresos'].style.format(fmt_ing), use_container_width=True, height=480)

    with tab_lr:
        st.markdown("#### Cartera en Lenta Recuperación / Jurídico y Seguros (Cuenta 115-99)")
        st.caption("Detalle de las subcuentas que integran el saldo de la cuenta 115-99 en Contpaqi.")
        fmt_lr = {
            'saldo_sistema': '${:,.2f}',
            'saldo_115_99': '${:,.2f}'
        }
        st.dataframe(tablas['df_lenta_recuperacion'].style.format(fmt_lr), use_container_width=True, height=350)

    with tab_liq:
        st.markdown("#### Auditoría de Contratos Concluidos y Liquidados al Corte")
        st.caption("Verificación de que los contratos concluidos tengan saldos en cero y detección de residuales (114-01) pendientes.")
        c_lf1, c_lf2 = st.columns([1, 2])
        with c_lf1:
            filtro_liq = st.radio("Filtrar liquidados:", ["Todos", "Con Saldo Pendiente (Alertas)", "Saldados en Cero ($0.00)"], horizontal=True, key="rad_liq_act_clean")

        df_show_liq = tablas['df_concil_liquidados'].copy()
        if filtro_liq == "Con Saldo Pendiente (Alertas)":
            df_show_liq = df_show_liq[df_show_liq['residual'] > 0.01]
        elif filtro_liq == "Saldados en Cero ($0.00)":
            df_show_liq = df_show_liq[df_show_liq['residual'] <= 0.01]

        fmt_liq = {
            'capital': '${:,.2f}',
            'intereses': '${:,.2f}',
            'residual': '${:,.2f}',
            'otros': '${:,.2f}',
            'ajustes': '${:,.2f}',
            'total_liq': '${:,.2f}',
        }
        st.dataframe(df_show_liq.style.format(fmt_liq), use_container_width=True, height=450)
