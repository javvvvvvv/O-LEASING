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
reports/auditoria_excel.py — Generador oficial de libro de Excel para O-Leasing
desarrollado por Orange Crew con fórmulas matemáticas vivas de auditoría financiera,
desglose mensual de devengos y matrices completas de saldos insolutos, saldos de intereses
y saldos de residual.

Diseñado específicamente para auditorías financieras y fiscales (NIF C-15 / IFRS 16):
1. Identidad institucional O-Leasing · Una app de Orange Crew con logotipos oficiales.
2. Contratos en vertical (filas) con unidad/vehículo adquirida, porcentajes iniciales
   (Anticipo %, Comisión %, Residual %) y fórmulas nativas de Excel (=RATE, =PV, =SUM, =AVERAGE).
3. Todas las cifras limitadas a 2 decimales ($#,##0.00 y 0.00%), excepto tasas completas
   con precisión total a 8 decimales (0.00000000%).
4. Hoja dinámica 'Matriz_Consolidada' con lista desplegable para seleccionar contrato,
   actualización automática mediante fórmulas =VLOOKUP y corte al mes exacto del plazo.
5. Matrices cronológicas horizontales con fórmulas vivas de derivación mensual:
   - Intereses_Leasing_Mensual
   - Intereses_Residual_Mensual
   - Saldo_Residual_Activo
   - Amortizacion_Capital_Mensual
   - Saldos_Capital_Mensual
   - Saldos_Intereses_Mensual
   - Saldos_Cartera_Total
   - Saldos_Exigible_Global
   - Matriz_Consolidada_Cartera
   - Matriz_Consolidada_Saldos
"""

import io
import os
import sys
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.drawing.image import Image as OpenpyxlImage

if getattr(sys, "frozen", False):
    _BASE_DIR = os.path.dirname(sys.executable)
else:
    _BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

LOGO_OLEASING_PNG_DEFAULT = os.path.join(_BASE_DIR, "assets", "o-leasing-logo.png")
LOGO_ORANGE_PNG_DEFAULT = os.path.join(_BASE_DIR, "assets", "orange-crew-logo.png")

# Paleta corporativa sobria de auditoria O-Leasing / Orange Crew
FILL_HEADER = PatternFill(start_color="1E5C4F", end_color="1E5C4F", fill_type="solid")
FILL_SUBHEADER = PatternFill(start_color="2D7D6F", end_color="2D7D6F", fill_type="solid")
FILL_ZEBRA = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
FILL_TOTAL = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
FILL_HIGHLIGHT = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

FONT_HEADER = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
FONT_SUBHEADER = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
FONT_BODY = Font(name="Calibri", size=9, color="1E293B")
FONT_BODY_BOLD = Font(name="Calibri", size=9, bold=True, color="1E293B")
FONT_TOTAL = Font(name="Calibri", size=9, bold=True, color="1E5C4F")
FONT_TITLE = Font(name="Calibri", size=13, bold=True, color="1E5C4F")
FONT_SUBTITLE = Font(name="Calibri", size=8, italic=True, color="64748B")

THIN_BORDER_SIDE = Side(style="thin", color="CBD5E1")
DOUBLE_BOTTOM_SIDE = Side(style="double", color="1E5C4F")
BORDER_CELL = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=THIN_BORDER_SIDE)
BORDER_TOTAL = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=DOUBLE_BOTTOM_SIDE)

# Formatos numericos normativos
FMT_CURRENCY = "$#,##0.00;[Red]-$#,##0.00;\"-\""
FMT_PCT_RATE = "0.00000000%"          # Tasas completas con 8 decimales
FMT_PCT_ANNUAL = "0.00000000%"        # Tasas anuales completas
FMT_PCT_PARAM = "0.00%"               # Porcentajes de captura a 2 decimales
FMT_INT = "#,##0"


def _obtener_nombre_empresa():
    try:
        from models.configuracion import get_cfg
        cfg_nom = get_cfg("nombre_empresa")
        if cfg_nom and str(cfg_nom).strip() and str(cfg_nom).strip().lower() != "mi empresa":
            return str(cfg_nom).strip()
    except Exception:
        pass
    return "O-LEASING"


def _preparar_encabezado_hoja(ws, titulo_hoja, subtitulo, total_cols, titulo_empresa=None, logo_path=None):
    ws.cell(row=1, column=1, value=f"O-LEASING · UNA APP DE ORANGE CREW · {titulo_hoja}").font = FONT_TITLE
    ws.cell(row=2, column=1, value=f"{subtitulo} · Auditoría Financiera NIF C-15 / IFRS 16 · Orange Crew").font = FONT_SUBTITLE
    if total_cols > 1:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=min(total_cols, 15))
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=min(total_cols, 15))
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 16
    try:
        ws.views.sheetView[0].showGridLines = True
    except Exception:
        pass


def _construir_caratula_ejecutiva(ws, titulo_empresa, fila_tot_resumen, anio_filtro, total_contratos):
    try:
        ws.views.sheetView[0].showGridLines = True
    except Exception:
        pass

    ws.row_dimensions[1].height = 10
    ws.row_dimensions[2].height = 44
    ws.row_dimensions[3].height = 10
    ws.row_dimensions[4].height = 24
    ws.row_dimensions[5].height = 22
    ws.row_dimensions[6].height = 18
    ws.row_dimensions[7].height = 12

    # Insercion de Logotipos Institucionales
    try:
        if os.path.exists(LOGO_OLEASING_PNG_DEFAULT):
            img_oleasing = OpenpyxlImage(LOGO_OLEASING_PNG_DEFAULT)
            img_oleasing.width = 175
            img_oleasing.height = 39
            ws.add_image(img_oleasing, "B2")
        if os.path.exists(LOGO_ORANGE_PNG_DEFAULT):
            img_orange = OpenpyxlImage(LOGO_ORANGE_PNG_DEFAULT)
            img_orange.width = 135
            img_orange.height = 39
            ws.add_image(img_orange, "E2")
    except Exception:
        pass

    ws.cell(row=4, column=2, value="O-LEASING · UNA APP DE ORANGE CREW").font = Font(name="Calibri", size=14, bold=True, color="1E5C4F")
    ws.cell(row=5, column=2, value="REPORTE INTEGRAL DE AUDITORÍA Y SALDOS DE CARTERA").font = Font(name="Calibri", size=16, bold=True, color="1E293B")

    alcance_txt = f"Ejercicio Fiscal {anio_filtro}" if anio_filtro else "Histórico Multianual Consolidado (2022 - 2030)"
    ws.cell(row=6, column=2, value=f"Normatividad NIF C-15 / IFRS 16 · Fórmulas Financieras Vivas de Excel · Desarrollado por Orange Crew · {alcance_txt}").font = FONT_SUBTITLE

    # Bloque de KPIs Ejecutivos (Filas 8 a 11)
    kpis_config = [
        (9, 2, "CONTRATOS AUDITADOS", f"='Resumen_Cartera'!G{fila_tot_resumen}", FMT_INT),
        (9, 4, "INVERSIÓN NETA TOTAL", f"='Resumen_Cartera'!K{fila_tot_resumen}", FMT_CURRENCY),
        (9, 6, "RENTA MENSUAL TOTAL", f"='Resumen_Cartera'!N{fila_tot_resumen}", FMT_CURRENCY),
        (9, 8, "RESIDUAL TOTAL PACTADO", f"='Resumen_Cartera'!P{fila_tot_resumen}", FMT_CURRENCY),
        (11, 2, "TASA MENSUAL PROMEDIO", f"='Resumen_Cartera'!R{fila_tot_resumen}", FMT_PCT_RATE),
        (11, 4, "TOTAL INTERÉS LEASING", f"='Resumen_Cartera'!V{fila_tot_resumen}", FMT_CURRENCY),
        (11, 6, "TOTAL INTERÉS RESIDUAL", f"='Resumen_Cartera'!W{fila_tot_resumen}", FMT_CURRENCY),
        (11, 8, "INGRESO FINANCIERO TOTAL", f"='Resumen_Cartera'!X{fila_tot_resumen}", FMT_CURRENCY),
    ]

    for r_kpi, c_kpi, label, form_val, fmt in kpis_config:
        c_lbl = ws.cell(row=r_kpi - 1, column=c_kpi, value=label)
        c_lbl.font = Font(name="Calibri", size=8, bold=True, color="64748B")
        c_lbl.alignment = Alignment(horizontal="left", vertical="center")

        c_v = ws.cell(row=r_kpi, column=c_kpi, value=form_val)
        c_v.font = Font(name="Calibri", size=13, bold=True, color="1E5C4F")
        c_v.number_format = fmt
        c_v.alignment = Alignment(horizontal="left", vertical="center")
        c_v.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        c_v.border = Border(
            left=Side(style="thin", color="CBD5E1"),
            right=Side(style="thin", color="CBD5E1"),
            top=Side(style="thin", color="CBD5E1"),
            bottom=Side(style="medium", color="1E5C4F")
        )

    # Indice de Hojas de Auditoria
    ws.cell(row=13, column=2, value="ÍNDICE DE HOJAS Y METODOLOGÍA DEL LIBRO OFICIAL (O-LEASING · ORANGE CREW)").font = Font(name="Calibri", size=11, bold=True, color="1E5C4F")

    headers_ind = [
        ("No.", 6, Alignment(horizontal="center")),
        ("Hoja de Cálculo (Clic para ir)", 28, Alignment(horizontal="left")),
        ("Propósito y Contenido Contable", 42, Alignment(horizontal="left")),
        ("Fórmulas Matemáticas Clave", 42, Alignment(horizontal="left")),
    ]

    r_hdr_ind = 14
    ws.row_dimensions[r_hdr_ind].height = 24
    for idx_hi, (nom_h, anc_h, al_h) in enumerate(headers_ind, start=2):
        c_h = ws.cell(row=r_hdr_ind, column=idx_hi, value=nom_h)
        c_h.fill = FILL_HEADER
        c_h.font = FONT_HEADER
        c_h.alignment = Alignment(horizontal="center", vertical="center")
        c_h.border = BORDER_CELL
        ws.column_dimensions[get_column_letter(idx_hi)].width = anc_h

    indice_filas = [
        ("1", "Resumen_Cartera", "Padrón completo de contratos en vertical con unidad vehicular, porcentajes y rentabilidad.", "=RATE, =PV, =SUM, =AVERAGE"),
        ("2", "Calculo_Tasa_Implicita", "Metodología NIF C-15 / IFRS 16 de cálculo de tasa mensual y comprobación a diferencia cero.", "=RATE, =PV, =PV-Inversion, TIR"),
        ("3", "Matriz_Consolidada", "Matriz interactiva con selector de contrato (desplegable) y corte automático exacto al plazo.", "=VLOOKUP, =ROUND(Saldo*Tasa, 2), =IF"),
        ("4", "Intereses_Leasing_Mensual", "Matriz horizontal de intereses de leasing devengados mes a mes por contrato.", "=IF(Saldo_Ant>0, ROUND(Saldo_Ant*Tasa, 2), '')"),
        ("5", "Intereses_Residual_Mensual", "Matriz horizontal de intereses sobre el valor presente del residual (NIF C-15).", "=IF(Saldo_Res_Ant>0, ROUND(Saldo_Res*Tasa, 2), '')"),
        ("6", "Ingreso_Financiero_Total", "Matriz de ingreso financiero consolidado (Leasing + Residual) mes a mes.", "=Intereses_Leasing_Mensual! + Residual!"),
        ("7", "Amortizacion_Capital_Mensual", "Matriz horizontal de amortización de principal mes a mes.", "=IF(Leasing>0, Renta - Leasing, '')"),
        ("8", "Saldos_Capital_Mensual", "Saldo insoluto de principal mes a mes al cierre de cada periodo contable.", "=IF(Cap>0, Saldo_Ant - Cap, '')"),
        ("9", "Saldos_Intereses_Mensual", "Saldo de intereses remanentes por devengar/cobrar desde cada corte mensual.", "=IF(Int>0, Saldo_Ant - Int, '')"),
        ("10", "Saldos_Cartera_Total", "Saldo total exigible de mensualidades restantes (Capital + Intereses).", "=Saldos_Capital! + Saldos_Intereses!"),
        ("11", "Saldo_Residual_Activo", "Evolución mensual del saldo activo del valor residual pactado.", "=IF(Int_Res>0, Saldo_Ant + Int_Res, '')"),
        ("12", "Saldos_Exigible_Global", "Deuda integral total de cartera más valor residual activo.", "=Saldos_Cartera_Total! + Saldo_Residual!"),
        ("13", "Matriz_Consolidada_Cartera", "Bloques cronológicos multi-concepto con leasing, residual, ingreso y capital de cartera.", "Fórmulas inter-hoja vivas"),
        ("14", "Matriz_Consolidada_Saldos", "Bloques cronológicos multi-concepto con los 5 rubros de saldos lado a lado.", "Fórmulas inter-hoja vivas"),
    ]

    for idx_row_ind, (num_i, hoja_i, desc_i, form_i) in enumerate(indice_filas, start=15):
        ws.row_dimensions[idx_row_ind].height = 20
        c_num = ws.cell(row=idx_row_ind, column=2, value=int(num_i))
        c_num.alignment = Alignment(horizontal="center", vertical="center")

        c_link = ws.cell(row=idx_row_ind, column=3, value=f"=HYPERLINK(\"#'{hoja_i}'!A1\", \"{hoja_i}\")")
        c_link.font = Font(name="Calibri", size=9, bold=True, color="1E5C4F", underline="single")
        c_link.alignment = Alignment(horizontal="left", vertical="center")

        c_desc = ws.cell(row=idx_row_ind, column=4, value=desc_i)
        c_desc.font = FONT_BODY
        c_desc.alignment = Alignment(horizontal="left", vertical="center")

        c_form = ws.cell(row=idx_row_ind, column=5, value=form_i)
        c_form.font = Font(name="Consolas", size=8, color="047857")
        c_form.alignment = Alignment(horizontal="left", vertical="center")

        es_par_ind = (idx_row_ind % 2 == 0)
        for col_c in range(2, 6):
            cell_box = ws.cell(row=idx_row_ind, column=col_c)
            cell_box.border = BORDER_CELL
            if es_par_ind:
                cell_box.fill = FILL_ZEBRA

    r_pie = len(indice_filas) + 16
    ws.cell(row=r_pie, column=2, value="O-LEASING es una aplicación financiera desarrollada y mantenida por ORANGE CREW. Todos los derechos reservados conforme a la LFDA.").font = FONT_SUBTITLE

    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 6
    ws.column_dimensions["C"].width = 28
    ws.column_dimensions["D"].width = 44
    ws.column_dimensions["E"].width = 44


def _obtener_db_path(db_path=None):
    if db_path and os.path.exists(db_path):
        return db_path
    try:
        import streamlit as st
        p = st.session_state.get("_active_db_path")
        if p and os.path.exists(p):
            return p
    except Exception:
        pass
    cand = os.path.join(_BASE_DIR, "leasing.db")
    if os.path.exists(cand):
        return cand
    cand_data = os.path.join(_BASE_DIR, "data", "leasing.db")
    if os.path.exists(cand_data):
        return cand_data
    return cand


def generar_excel_auditoria(
    anio_filtro=None,
    filtro_estatus="ACTIVO",
    titulo_empresa=None,
    logo_path=None,
    db_path=None,
    solo_con_movimiento=True,
    *args,
    **kwargs
) -> io.BytesIO:
    """Genera el libro completo de auditoria O-Leasing desarrollado por Orange Crew.

    Garantiza:
    1. Cifras a 2 decimales ($#,##0.00 y 0.00%), excepto tasas completas a 8 decimales (0.00000000%).
    2. Formulas vivas en Intereses_Leasing_Mensual, Intereses_Residual_Mensual y Saldo_Residual_Activo.
    3. Hoja Matriz_Consolidada interactiva con lista desplegable y corte exacto al plazo.
    4. Logotipos de O-Leasing y Orange Crew con membretes institucionales.
    """
    import sqlite3
    from finanzas import calc_amort, calc_res_amort, vp_res

    db_path = _obtener_db_path(db_path)
    conn = sqlite3.connect(db_path)
    df_c = pd.read_sql("SELECT * FROM contratos ORDER BY ID_Contrato", conn)
    conn.close()

    if df_c.empty:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sin_Datos"
        ws["A1"] = "No se encontraron contratos registrados."
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    # Filtrar por estatus si se especifica
    if filtro_estatus and str(filtro_estatus).strip().upper() not in ("TODOS", "ALL", "HISTORICO"):
        est_req = str(filtro_estatus).strip().upper()
        df_c = df_c[df_c["Estatus"].astype(str).str.strip().str.upper() == est_req]

    # Precomputo de amortizaciones, devengos y saldos por contrato
    contratos_pre = []
    min_ym = None
    max_ym = None

    for _, r in df_c.iterrows():
        id_c = str(r["ID_Contrato"]).strip()
        cli = str(r.get("Cliente") or "").strip()
        veh = str(r.get("Vehiculo") or "").strip()
        est = str(r.get("Estatus") or "").strip()
        fa = pd.to_datetime(r.get("Fecha_Alta"))
        fv = pd.to_datetime(r.get("Fecha_Vencimiento"))
        fb = pd.to_datetime(r.get("Fecha_Baja"))

        if pd.isna(fa) or pd.isna(fv):
            continue
        try:
            pl = int(r.get("Plazo") or 0)
        except Exception:
            continue
        if pl <= 0:
            continue

        fa_y = int(fa.year)
        fa_m = int(fa.month)
        fv_y = int(fv.year)
        fv_m = int(fv.month)

        es_baja = est.upper() == "BAJA"
        fb_y, fb_m = 9999, 99
        if es_baja and pd.notna(fb):
            fb_y = int(fb.year)
            fb_m = int(fb.month)

        if anio_filtro is not None and solo_con_movimiento:
            a_f = int(anio_filtro)
            if (fa_y > a_f) or (fv_y < a_f) or (es_baja and fb_y < a_f):
                continue

        ym_start = (fa_y, fa_m)
        ym_end = (fv_y, fv_m)
        if min_ym is None or ym_start < min_ym:
            min_ym = ym_start
        if max_ym is None or ym_end > max_ym:
            max_ym = ym_end

        v_sin_iva = round(float(r.get("Valor_Sin_IVA") or 0.0), 2)
        anticipo = round(float(r.get("Anticipo_Monto") or 0.0), 2)
        inv = round(v_sin_iva - anticipo, 2)
        renta = round(float(r.get("Mensualidad_Sin_IVA") or 0.0), 2)
        res_monto = round(float(r.get("Residual_Monto") or 0.0), 2)
        com_monto = round(float(r.get("Comision_Monto") or 0.0), 2)
        tasa_reg = round(float(r.get("Tasa_Calculada") or 0.0), 8)

        # Porcentajes de captura (en base decimal para Excel: 20% = 0.20)
        ant_pct_raw = float(r.get("Anticipo_Pct") or 0.0)
        if ant_pct_raw > 0:
            ant_pct = round((ant_pct_raw / 100.0 if ant_pct_raw > 1.0 else ant_pct_raw), 4)
        else:
            ant_pct = round(((anticipo / v_sin_iva) if v_sin_iva > 0 else 0.0), 4)

        com_pct_raw = float(r.get("Comision_Apertura_Pct") or 0.0)
        if com_pct_raw > 0:
            com_pct = round((com_pct_raw / 100.0 if com_pct_raw > 1.0 else com_pct_raw), 4)
        else:
            com_pct = round(((com_monto / v_sin_iva) if v_sin_iva > 0 else 0.0), 4)

        res_pct_raw = float(r.get("Residual_Pct") or 0.0)
        if res_pct_raw > 0:
            res_pct = round((res_pct_raw / 100.0 if res_pct_raw > 1.0 else res_pct_raw), 4)
        else:
            res_pct = round(((res_monto / v_sin_iva) if v_sin_iva > 0 else 0.0), 4)

        vpr = float(r.get("VP_Residual") or 0.0)
        if not vpr or vpr <= 0:
            vpr = float(vp_res(res_monto, tasa_reg, pl))
        vpr = round(vpr, 2)

        dfa, _, _, _, _ = calc_amort(round(inv, 4), round(renta, 4), round(res_monto, 4), pl, tasa_reg)
        dfr = calc_res_amort(round(vpr, 4), tasa_reg, pl)

        il_list = [round(float(x), 2) for x in dfa["Interes"].tolist()]
        cap_list = [round(float(x), 2) for x in dfa["Capital"].tolist()]
        sal_list = [round(float(x), 2) for x in dfa["Saldo"].tolist()]
        ir_list = [round(float(x), 2) for x in dfr["Interes"].tolist()]
        sr_list = [round(float(x), 2) for x in dfr["Saldo_Fin"].tolist()]

        sint_list = [round(sum(il_list[idx + 1:]), 2) for idx in range(pl)]
        tot_interes = round(sum(il_list), 2)

        contratos_pre.append({
            "id": id_c,
            "cliente": cli,
            "vehiculo": veh,
            "estatus": est,
            "fa_str": fa.strftime("%Y-%m-%d"),
            "fv_str": fv.strftime("%Y-%m-%d"),
            "fa_y": fa_y,
            "fa_m": fa_m,
            "fv_y": fv_y,
            "fv_m": fv_m,
            "fb_y": fb_y,
            "fb_m": fb_m,
            "pl": pl,
            "valor": v_sin_iva,
            "anticipo_pct": ant_pct,
            "anticipo": anticipo,
            "inversion": inv,
            "comision_pct": com_pct,
            "comision": com_monto,
            "renta": renta,
            "residual_pct": res_pct,
            "residual": res_monto,
            "tasa_reg": tasa_reg,
            "vpr": vpr,
            "il": il_list,
            "cap": cap_list,
            "sal": sal_list,
            "sint": sint_list,
            "tot_interes": tot_interes,
            "ir": ir_list,
            "sr": sr_list,
        })

    # Determinar meses del horizonte
    if anio_filtro is not None:
        anio_filtro = int(anio_filtro)
        lista_meses = [(anio_filtro, m) for m in range(1, 13)]
    else:
        if min_ym is None:
            min_ym = (2022, 1)
        if max_ym is None:
            max_ym = (2030, 12)

        lista_meses = []
        cur_y, cur_m = min_ym
        end_y, end_m = max_ym
        while (cur_y < end_y) or (cur_y == end_y and cur_m <= end_m):
            lista_meses.append((cur_y, cur_m))
            cur_m += 1
            if cur_m > 12:
                cur_m = 1
                cur_y += 1

    wb = openpyxl.Workbook()
    ws_caratula = wb.active
    ws_caratula.title = "Caratula_Ejecutiva"

    # =========================================================================
    # 1. HOJA: Resumen_Cartera (Contratos en vertical + Formulas Financieras)
    # =========================================================================
    ws_resumen = wb.create_sheet(title="Resumen_Cartera")
    cols_resumen = [
        ("ID_Contrato", 14, Alignment(horizontal="center")),              # A (1)
        ("Cliente", 34, Alignment(horizontal="left")),                    # B (2)
        ("Vehiculo", 26, Alignment(horizontal="left")),                   # C (3)
        ("Estatus", 12, Alignment(horizontal="center")),                  # D (4)
        ("Fecha_Alta", 13, Alignment(horizontal="center")),               # E (5)
        ("Fecha_Vencimiento", 15, Alignment(horizontal="center")),        # F (6)
        ("Plazo_Meses", 12, Alignment(horizontal="right")),               # G (7)
        ("Valor_Sin_IVA", 16, Alignment(horizontal="right")),             # H (8)
        ("Anticipo_Pct", 13, Alignment(horizontal="right")),              # I (9)
        ("Anticipo_Monto", 16, Alignment(horizontal="right")),            # J (10) Formula =ROUND(H*I, 2)
        ("Inversion_Neta", 16, Alignment(horizontal="right")),            # K (11) Formula =ROUND(H-J, 2)
        ("Comision_Pct", 13, Alignment(horizontal="right")),              # L (12)
        ("Comision_Apertura", 16, Alignment(horizontal="right")),         # M (13) Formula =ROUND(H*L, 2)
        ("Renta_Sin_IVA", 16, Alignment(horizontal="right")),             # N (14)
        ("Residual_Pct", 13, Alignment(horizontal="right")),              # O (15)
        ("Residual_Pactado", 16, Alignment(horizontal="right")),          # P (16) Formula =ROUND(H*O, 2)
        ("Tasa_Mensual_Reg", 18, Alignment(horizontal="right")),          # Q (17) 8 decimales
        ("Tasa_Mensual_RATE", 18, Alignment(horizontal="right")),         # R (18) Formula RATE (8 dec)
        ("Tasa_Nominal_Anual", 18, Alignment(horizontal="right")),        # S (19) Formula =Q*12 (8 dec)
        ("TIR_Anual_Efectiva", 18, Alignment(horizontal="right")),        # T (20) Formula ((1+Q)^12)-1 (8 dec)
        ("VP_Residual", 16, Alignment(horizontal="right")),               # U (21) Formula VP (2 dec)
        ("Total_Interes_Leasing", 18, Alignment(horizontal="right")),     # V (22) Formula =(N*G)-K+P
        ("Total_Interes_Residual", 18, Alignment(horizontal="right")),    # W (23) Formula =P-U
        ("Ingreso_Financiero_Total", 20, Alignment(horizontal="right")),  # X (24) Formula =V+W
        ("Ganancia_Neta", 18, Alignment(horizontal="right")),             # Y (25) Formula =(N*G+P+M)-K
        ("Margen_Sobre_Inv", 16, Alignment(horizontal="right")),          # Z (26) Formula =Y/K
    ]

    _preparar_encabezado_hoja(
        ws_resumen,
        "Resumen de Cartera y Parámetros Financieros",
        "Parámetros contractuales, unidad adquirida, porcentajes iniciales y fórmulas matemáticas de rentabilidad",
        len(cols_resumen),
        titulo_empresa=titulo_empresa,
        logo_path=logo_path
    )

    fila_header_res = 4
    ws_resumen.row_dimensions[fila_header_res].height = 26
    for idx_col, (nombre_col, ancho, align) in enumerate(cols_resumen, start=1):
        c = ws_resumen.cell(row=fila_header_res, column=idx_col, value=nombre_col)
        c.fill = FILL_HEADER
        c.font = FONT_HEADER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER_CELL
        letra = get_column_letter(idx_col)
        ws_resumen.column_dimensions[letra].width = ancho

    fila_inicio_datos = 5
    fila_actual = fila_inicio_datos

    for c_item in contratos_pre:
        r_idx = fila_actual
        ws_resumen.cell(row=r_idx, column=1, value=c_item["id"]).alignment = cols_resumen[0][2]
        ws_resumen.cell(row=r_idx, column=2, value=c_item["cliente"]).alignment = cols_resumen[1][2]
        ws_resumen.cell(row=r_idx, column=3, value=c_item["vehiculo"]).alignment = cols_resumen[2][2]
        ws_resumen.cell(row=r_idx, column=4, value=c_item["estatus"]).alignment = cols_resumen[3][2]
        ws_resumen.cell(row=r_idx, column=5, value=c_item["fa_str"]).alignment = cols_resumen[4][2]
        ws_resumen.cell(row=r_idx, column=6, value=c_item["fv_str"]).alignment = cols_resumen[5][2]

        c_pl = ws_resumen.cell(row=r_idx, column=7, value=c_item["pl"])
        c_pl.number_format = FMT_INT
        c_pl.alignment = cols_resumen[6][2]

        c_val = ws_resumen.cell(row=r_idx, column=8, value=c_item["valor"])
        c_val.number_format = FMT_CURRENCY
        c_val.alignment = cols_resumen[7][2]

        c_apct = ws_resumen.cell(row=r_idx, column=9, value=c_item["anticipo_pct"])
        c_apct.number_format = FMT_PCT_PARAM
        c_apct.alignment = cols_resumen[8][2]

        # Anticipo Monto = ROUND(H * I, 2)
        c_ant = ws_resumen.cell(row=r_idx, column=10, value=f"=ROUND(H{r_idx}*I{r_idx}, 2)")
        c_ant.number_format = FMT_CURRENCY
        c_ant.alignment = cols_resumen[9][2]

        # Inversion Neta = ROUND(H - J, 2)
        c_inv = ws_resumen.cell(row=r_idx, column=11, value=f"=ROUND(H{r_idx}-J{r_idx}, 2)")
        c_inv.number_format = FMT_CURRENCY
        c_inv.alignment = cols_resumen[10][2]

        c_cpct = ws_resumen.cell(row=r_idx, column=12, value=c_item["comision_pct"])
        c_cpct.number_format = FMT_PCT_PARAM
        c_cpct.alignment = cols_resumen[11][2]

        # Comision Apertura = ROUND(H * L, 2)
        c_com = ws_resumen.cell(row=r_idx, column=13, value=f"=ROUND(H{r_idx}*L{r_idx}, 2)")
        c_com.number_format = FMT_CURRENCY
        c_com.alignment = cols_resumen[12][2]

        c_ren = ws_resumen.cell(row=r_idx, column=14, value=c_item["renta"])
        c_ren.number_format = FMT_CURRENCY
        c_ren.alignment = cols_resumen[13][2]

        c_rpct = ws_resumen.cell(row=r_idx, column=15, value=c_item["residual_pct"])
        c_rpct.number_format = FMT_PCT_PARAM
        c_rpct.alignment = cols_resumen[14][2]

        # Residual Pactado = ROUND(H * O, 2)
        c_res = ws_resumen.cell(row=r_idx, column=16, value=f"=ROUND(H{r_idx}*O{r_idx}, 2)")
        c_res.number_format = FMT_CURRENCY
        c_res.alignment = cols_resumen[15][2]

        # Tasas Completas a 8 decimales
        c_treg = ws_resumen.cell(row=r_idx, column=17, value=c_item["tasa_reg"])
        c_treg.number_format = FMT_PCT_RATE
        c_treg.alignment = cols_resumen[16][2]

        c_trate = ws_resumen.cell(row=r_idx, column=18, value=f"=IFERROR(RATE(G{r_idx}, -N{r_idx}, K{r_idx}, -P{r_idx}), Q{r_idx})")
        c_trate.number_format = FMT_PCT_RATE
        c_trate.alignment = cols_resumen[17][2]

        c_tnom = ws_resumen.cell(row=r_idx, column=19, value=f"=Q{r_idx}*12")
        c_tnom.number_format = FMT_PCT_ANNUAL
        c_tnom.alignment = cols_resumen[18][2]

        c_tir = ws_resumen.cell(row=r_idx, column=20, value=f"=((1+Q{r_idx})^12)-1")
        c_tir.number_format = FMT_PCT_ANNUAL
        c_tir.alignment = cols_resumen[19][2]

        # VP Residual = IF(Q>0, ROUND(P/((1+Q)^G), 2), P)
        c_vpr = ws_resumen.cell(row=r_idx, column=21, value=f"=IF(Q{r_idx}>0, ROUND(P{r_idx}/((1+Q{r_idx})^G{r_idx}), 2), P{r_idx})")
        c_vpr.number_format = FMT_CURRENCY
        c_vpr.alignment = cols_resumen[20][2]

        # Total Interes Leasing = ROUND((N * G) - K + P, 2)
        c_til = ws_resumen.cell(row=r_idx, column=22, value=f"=ROUND((N{r_idx}*G{r_idx})-K{r_idx}+P{r_idx}, 2)")
        c_til.number_format = FMT_CURRENCY
        c_til.alignment = cols_resumen[21][2]

        # Total Interes Residual = ROUND(P - U, 2)
        c_tir_res = ws_resumen.cell(row=r_idx, column=23, value=f"=ROUND(P{r_idx}-U{r_idx}, 2)")
        c_tir_res.number_format = FMT_CURRENCY
        c_tir_res.alignment = cols_resumen[22][2]

        # Ingreso Financiero Total = ROUND(V + W, 2)
        c_ing = ws_resumen.cell(row=r_idx, column=24, value=f"=ROUND(V{r_idx}+W{r_idx}, 2)")
        c_ing.number_format = FMT_CURRENCY
        c_ing.alignment = cols_resumen[23][2]

        # Ganancia Neta = ROUND((N * G + P + M) - K, 2)
        c_gan = ws_resumen.cell(row=r_idx, column=25, value=f"=ROUND((N{r_idx}*G{r_idx}+P{r_idx}+M{r_idx})-K{r_idx}, 2)")
        c_gan.number_format = FMT_CURRENCY
        c_gan.alignment = cols_resumen[24][2]

        # Margen = IF(K>0, Y/K, 0)
        c_mar = ws_resumen.cell(row=r_idx, column=26, value=f"=IF(K{r_idx}>0, Y{r_idx}/K{r_idx}, 0)")
        c_mar.number_format = FMT_PCT_PARAM
        c_mar.alignment = cols_resumen[25][2]

        es_par = (fila_actual - fila_inicio_datos) % 2 == 1
        for col_i in range(1, len(cols_resumen) + 1):
            cell_obj = ws_resumen.cell(row=r_idx, column=col_i)
            cell_obj.border = BORDER_CELL
            cell_obj.font = FONT_BODY
            if es_par:
                cell_obj.fill = FILL_ZEBRA

        fila_actual += 1

    # Fila de Totales de Cartera
    fila_fin_datos = fila_actual - 1
    ws_resumen.cell(row=fila_actual, column=1, value="TOTAL CARTERA").alignment = Alignment(horizontal="center")
    ws_resumen.cell(row=fila_actual, column=2, value=f"{len(contratos_pre)} Contratos Registrados").alignment = Alignment(horizontal="left")

    for c_idx in [8, 10, 11, 13, 14, 16, 21, 22, 23, 24, 25]:
        letra_col = get_column_letter(c_idx)
        cell_t = ws_resumen.cell(row=fila_actual, column=c_idx, value=f"=SUM({letra_col}{fila_inicio_datos}:{letra_col}{fila_fin_datos})")
        cell_t.number_format = FMT_CURRENCY
        cell_t.alignment = Alignment(horizontal="right")

    for c_idx in [9, 12, 15, 17, 18, 19, 20]:
        letra_col = get_column_letter(c_idx)
        fmt_t = FMT_PCT_RATE if c_idx in [17, 18, 19, 20] else FMT_PCT_PARAM
        cell_t = ws_resumen.cell(row=fila_actual, column=c_idx, value=f"=AVERAGE({letra_col}{fila_inicio_datos}:{letra_col}{fila_fin_datos})")
        cell_t.number_format = fmt_t
        cell_t.alignment = Alignment(horizontal="right")

    cell_mar_tot = ws_resumen.cell(row=fila_actual, column=26, value=f"=IF(K{fila_actual}>0, Y{fila_actual}/K{fila_actual}, 0)")
    cell_mar_tot.number_format = FMT_PCT_PARAM
    cell_mar_tot.alignment = Alignment(horizontal="right")

    for col_i in range(1, len(cols_resumen) + 1):
        c_tot = ws_resumen.cell(row=fila_actual, column=col_i)
        c_tot.fill = FILL_TOTAL
        c_tot.font = FONT_TOTAL
        c_tot.border = BORDER_TOTAL

    ws_resumen.freeze_panes = "D5"
    ws_resumen.auto_filter.ref = f"A{fila_header_res}:{get_column_letter(len(cols_resumen))}{fila_fin_datos}"

    # Construir contenido ejecutivo e indice en la caratula
    _construir_caratula_ejecutiva(ws_caratula, titulo_empresa, fila_actual, anio_filtro, len(contratos_pre))

    # =========================================================================
    # 2. HOJA: Calculo_Tasa_Implicita (Explicacion Matematica y Comprobacion PV)
    # =========================================================================
    ws_tasa = wb.create_sheet(title="Calculo_Tasa_Implicita")
    cols_tasa = [
        ("ID_Contrato", 14, Alignment(horizontal="center")),              # A (1)
        ("Cliente", 34, Alignment(horizontal="left")),                    # B (2)
        ("Vehiculo", 26, Alignment(horizontal="left")),                   # C (3)
        ("Estatus", 12, Alignment(horizontal="center")),                  # D (4)
        ("Plazo_Meses", 12, Alignment(horizontal="right")),               # E (5)
        ("Valor_Sin_IVA", 16, Alignment(horizontal="right")),             # F (6)
        ("Anticipo_Pct", 13, Alignment(horizontal="right")),              # G (7)
        ("Anticipo_Monto", 16, Alignment(horizontal="right")),            # H (8) Formula =ROUND(F*G, 2)
        ("Inversion_Neta", 16, Alignment(horizontal="right")),            # I (9) Formula =ROUND(F-H, 2)
        ("Renta_Sin_IVA", 16, Alignment(horizontal="right")),             # J (10)
        ("Residual_Pct", 13, Alignment(horizontal="right")),              # K (11)
        ("Residual_Pactado", 16, Alignment(horizontal="right")),          # L (12) Formula =ROUND(F*K, 2)
        ("Tasa_Mensual_Registrada", 18, Alignment(horizontal="right")),   # M (13) 8 dec
        ("Tasa_Calculada_RATE", 18, Alignment(horizontal="right")),       # N (14) Formula RATE (8 dec)
        ("Comprobacion_PV", 18, Alignment(horizontal="right")),           # O (15) Formula PV (2 dec)
        ("Diferencia_Cuadre", 16, Alignment(horizontal="right")),         # P (16) Formula =ROUND(O-I, 2)
        ("Tasa_Nominal_Anual", 18, Alignment(horizontal="right")),        # Q (17) Formula =N*12 (8 dec)
        ("TIR_Anual_Efectiva", 18, Alignment(horizontal="right")),        # R (18) Formula ((1+N)^12)-1 (8 dec)
        ("Diferencia_vs_Registrada", 18, Alignment(horizontal="right")),  # S (19) Formula =N-M (8 dec)
        ("Estatus_Verificacion", 16, Alignment(horizontal="center")),     # T (20) Formula IF
    ]

    _preparar_encabezado_hoja(
        ws_tasa,
        "Cálculo de Tasa Implícita y Comprobación de Flujos (NIF C-15 / IFRS 16)",
        "Determinación de tasa mensual mediante RATE(nper, -pmt, pv, -fv), unidad adquirida y comprobación PV con diferencia cero",
        len(cols_tasa),
        titulo_empresa=titulo_empresa,
        logo_path=logo_path
    )

    fila_header_tasa = 4
    ws_tasa.row_dimensions[fila_header_tasa].height = 28
    for idx, (nom_col, ancho, alig) in enumerate(cols_tasa, start=1):
        c = ws_tasa.cell(row=fila_header_tasa, column=idx, value=nom_col)
        c.fill = FILL_HEADER
        c.font = FONT_HEADER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER_CELL
        ws_tasa.column_dimensions[get_column_letter(idx)].width = ancho

    fila_actual_tasa = fila_inicio_datos
    for c_item in contratos_pre:
        r_t = fila_actual_tasa
        ws_tasa.cell(row=r_t, column=1, value=c_item["id"]).alignment = cols_tasa[0][2]
        ws_tasa.cell(row=r_t, column=2, value=c_item["cliente"]).alignment = cols_tasa[1][2]
        ws_tasa.cell(row=r_t, column=3, value=c_item["vehiculo"]).alignment = cols_tasa[2][2]
        ws_tasa.cell(row=r_t, column=4, value=c_item["estatus"]).alignment = cols_tasa[3][2]

        c_pl = ws_tasa.cell(row=r_t, column=5, value=c_item["pl"])
        c_pl.number_format = FMT_INT
        c_pl.alignment = cols_tasa[4][2]

        c_val = ws_tasa.cell(row=r_t, column=6, value=c_item["valor"])
        c_val.number_format = FMT_CURRENCY
        c_val.alignment = cols_tasa[5][2]

        c_apct = ws_tasa.cell(row=r_t, column=7, value=c_item["anticipo_pct"])
        c_apct.number_format = FMT_PCT_PARAM
        c_apct.alignment = cols_tasa[6][2]

        # Anticipo Monto = ROUND(F * G, 2)
        c_ant = ws_tasa.cell(row=r_t, column=8, value=f"=ROUND(F{r_t}*G{r_t}, 2)")
        c_ant.number_format = FMT_CURRENCY
        c_ant.alignment = cols_tasa[7][2]

        # Inversion Neta = ROUND(F - H, 2)
        c_inv = ws_tasa.cell(row=r_t, column=9, value=f"=ROUND(F{r_t}-H{r_t}, 2)")
        c_inv.number_format = FMT_CURRENCY
        c_inv.alignment = cols_tasa[8][2]

        c_ren = ws_tasa.cell(row=r_t, column=10, value=c_item["renta"])
        c_ren.number_format = FMT_CURRENCY
        c_ren.alignment = cols_tasa[9][2]

        c_rpct = ws_tasa.cell(row=r_t, column=11, value=c_item["residual_pct"])
        c_rpct.number_format = FMT_PCT_PARAM
        c_rpct.alignment = cols_tasa[10][2]

        # Residual Pactado = ROUND(F * K, 2)
        c_res = ws_tasa.cell(row=r_t, column=12, value=f"=ROUND(F{r_t}*K{r_t}, 2)")
        c_res.number_format = FMT_CURRENCY
        c_res.alignment = cols_tasa[11][2]

        # Tasas completas con 8 decimales
        c_treg = ws_tasa.cell(row=r_t, column=13, value=c_item["tasa_reg"])
        c_treg.number_format = FMT_PCT_RATE
        c_treg.alignment = cols_tasa[12][2]

        c_trate = ws_tasa.cell(row=r_t, column=14, value=f"=IFERROR(RATE(E{r_t}, -J{r_t}, I{r_t}, -L{r_t}), M{r_t})")
        c_trate.number_format = FMT_PCT_RATE
        c_trate.alignment = cols_tasa[13][2]

        # Comprobacion PV = IF(N>0, ROUND(PV(N, E, -J, -L), 2), I)
        c_pv = ws_tasa.cell(row=r_t, column=15, value=f"=IF(N{r_t}>0, ROUND(PV(N{r_t}, E{r_t}, -J{r_t}, -L{r_t}), 2), I{r_t})")
        c_pv.number_format = FMT_CURRENCY
        c_pv.alignment = cols_tasa[14][2]

        # Diferencia Cuadre = ROUND(O - I, 2) (debe dar $0.00 exacto)
        c_dif = ws_tasa.cell(row=r_t, column=16, value=f"=ROUND(O{r_t}-I{r_t}, 2)")
        c_dif.number_format = FMT_CURRENCY
        c_dif.alignment = cols_tasa[15][2]

        c_tnom = ws_tasa.cell(row=r_t, column=17, value=f"=N{r_t}*12")
        c_tnom.number_format = FMT_PCT_ANNUAL
        c_tnom.alignment = cols_tasa[16][2]

        c_tir = ws_tasa.cell(row=r_t, column=18, value=f"=((1+N{r_t})^12)-1")
        c_tir.number_format = FMT_PCT_ANNUAL
        c_tir.alignment = cols_tasa[17][2]

        c_dift = ws_tasa.cell(row=r_t, column=19, value=f"=N{r_t}-M{r_t}")
        c_dift.number_format = FMT_PCT_RATE
        c_dift.alignment = cols_tasa[18][2]

        c_ver = ws_tasa.cell(row=r_t, column=20, value=f'=IF(ABS(P{r_t})<0.01, "EXACTO", "REVISAR")')
        c_ver.alignment = cols_tasa[19][2]

        es_par_t = (fila_actual_tasa - fila_inicio_datos) % 2 == 1
        for col_i in range(1, len(cols_tasa) + 1):
            cell_obj = ws_tasa.cell(row=r_t, column=col_i)
            cell_obj.border = BORDER_CELL
            cell_obj.font = FONT_BODY
            if es_par_t:
                cell_obj.fill = FILL_ZEBRA

        fila_actual_tasa += 1

    fila_fin_tasa = fila_actual_tasa - 1
    ws_tasa.cell(row=fila_actual_tasa, column=1, value="TOTAL CARTERA").alignment = Alignment(horizontal="center")
    ws_tasa.cell(row=fila_actual_tasa, column=2, value=f"{len(contratos_pre)} Contratos").alignment = Alignment(horizontal="left")

    for c_idx in [6, 8, 9, 10, 12, 15, 16]:
        letra_col = get_column_letter(c_idx)
        cell_t = ws_tasa.cell(row=fila_actual_tasa, column=c_idx, value=f"=SUM({letra_col}{fila_inicio_datos}:{letra_col}{fila_fin_tasa})")
        cell_t.number_format = FMT_CURRENCY
        cell_t.alignment = Alignment(horizontal="right")

    for c_idx in [7, 11, 13, 14, 17, 18, 19]:
        letra_col = get_column_letter(c_idx)
        fmt_t = FMT_PCT_RATE if c_idx in [13, 14, 17, 18, 19] else FMT_PCT_PARAM
        cell_t = ws_tasa.cell(row=fila_actual_tasa, column=c_idx, value=f"=AVERAGE({letra_col}{fila_inicio_datos}:{letra_col}{fila_fin_tasa})")
        cell_t.number_format = fmt_t
        cell_t.alignment = Alignment(horizontal="right")

    ws_tasa.cell(row=fila_actual_tasa, column=20, value="AUDITADO").alignment = Alignment(horizontal="center")

    for col_i in range(1, len(cols_tasa) + 1):
        c_tot = ws_tasa.cell(row=fila_actual_tasa, column=col_i)
        c_tot.fill = FILL_TOTAL
        c_tot.font = FONT_TOTAL
        c_tot.border = BORDER_TOTAL

    ws_tasa.freeze_panes = "D5"
    ws_tasa.auto_filter.ref = f"A{fila_header_tasa}:{get_column_letter(len(cols_tasa))}{fila_fin_tasa}"

    # =========================================================================
    # 3. HOJA: Matriz_Consolidada (INTERACTIVA: Selector de Contrato y Corte al Plazo)
    # =========================================================================
    ws_cons = wb.create_sheet(title="Matriz_Consolidada")
    _preparar_encabezado_hoja(
        ws_cons,
        "Matriz Consolidada Dinámica por Contrato",
        "Seleccione cualquier contrato en la lista desplegable; la tabla actualiza todos los datos y se corta al mes exacto de terminación",
        15,
        titulo_empresa=titulo_empresa,
        logo_path=logo_path
    )

    # Insercion de Logotipo en Matriz Consolidada
    try:
        if os.path.exists(LOGO_OLEASING_PNG_DEFAULT):
            img_oleasing_m = OpenpyxlImage(LOGO_OLEASING_PNG_DEFAULT)
            img_oleasing_m.width = 150
            img_oleasing_m.height = 33
            ws_cons.add_image(img_oleasing_m, "K2")
    except Exception:
        pass

    # Bloque de Control y Seleccion de Contrato
    ws_cons.row_dimensions[4].height = 24
    ws_cons.cell(row=4, column=2, value="SELECCIONAR CONTRATO:").font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    ws_cons.cell(row=4, column=2).fill = FILL_HEADER
    ws_cons.cell(row=4, column=2).alignment = Alignment(horizontal="center", vertical="center")
    ws_cons.cell(row=4, column=2).border = BORDER_CELL

    primer_id = contratos_pre[0]["id"] if contratos_pre else ""
    c_sel = ws_cons.cell(row=4, column=3, value=primer_id)
    c_sel.font = Font(name="Calibri", size=12, bold=True, color="1E5C4F")
    c_sel.fill = FILL_HIGHLIGHT
    c_sel.alignment = Alignment(horizontal="center", vertical="center")
    c_sel.border = Border(
        left=Side(style="medium", color="1E5C4F"),
        right=Side(style="medium", color="1E5C4F"),
        top=Side(style="medium", color="1E5C4F"),
        bottom=Side(style="medium", color="1E5C4F")
    )

    # Validacion de datos con lista desplegable de todos los contratos de Resumen_Cartera
    dv = DataValidation(type="list", formula1=f"='Resumen_Cartera'!$A$5:$A${fila_fin_datos}", allow_blank=False)
    ws_cons.add_data_validation(dv)
    dv.add(ws_cons["C4"])

    ws_cons.cell(row=4, column=4, value="<-- Seleccione de la lista desplegable o escriba el ID del contrato para actualizar toda la tabla").font = Font(name="Calibri", size=9, italic=True, color="64748B")
    ws_cons.merge_cells("D4:K4")

    # Bloque de Parametros Contractuales vinculados con =VLOOKUP (Filas 5 a 8)
    meta_campos = [
        # (Fila, Col_Lbl, Label, Col_Val, Form_VLOOKUP, FMT)
        (5, 2, "Cliente:", 3, f"=IFERROR(VLOOKUP($C$4, 'Resumen_Cartera'!$A$5:$Z${fila_fin_datos}, 2, FALSE), \"\")", "@"),
        (5, 6, "Unidad / Vehículo:", 7, f"=IFERROR(VLOOKUP($C$4, 'Resumen_Cartera'!$A$5:$Z${fila_fin_datos}, 3, FALSE), \"\")", "@"),
        (5, 10, "Estatus:", 11, f"=IFERROR(VLOOKUP($C$4, 'Resumen_Cartera'!$A$5:$Z${fila_fin_datos}, 4, FALSE), \"\")", "@"),

        (6, 2, "Fecha Alta:", 3, f"=IFERROR(VLOOKUP($C$4, 'Resumen_Cartera'!$A$5:$Z${fila_fin_datos}, 5, FALSE), \"\")", "@"),
        (6, 4, "Fecha Vencimiento:", 5, f"=IFERROR(VLOOKUP($C$4, 'Resumen_Cartera'!$A$5:$Z${fila_fin_datos}, 6, FALSE), \"\")", "@"),
        (6, 7, "Plazo (Meses):", 8, f"=IFERROR(VLOOKUP($C$4, 'Resumen_Cartera'!$A$5:$Z${fila_fin_datos}, 7, FALSE), 0)", FMT_INT),
        (6, 10, "Valor Sin IVA:", 11, f"=IFERROR(VLOOKUP($C$4, 'Resumen_Cartera'!$A$5:$Z${fila_fin_datos}, 8, FALSE), 0)", FMT_CURRENCY),

        (7, 2, "Anticipo Monto:", 3, f"=IFERROR(VLOOKUP($C$4, 'Resumen_Cartera'!$A$5:$Z${fila_fin_datos}, 10, FALSE), 0)", FMT_CURRENCY),
        (7, 4, "Inversión Neta:", 5, f"=IFERROR(VLOOKUP($C$4, 'Resumen_Cartera'!$A$5:$Z${fila_fin_datos}, 11, FALSE), 0)", FMT_CURRENCY),
        (7, 7, "Renta Sin IVA:", 8, f"=IFERROR(VLOOKUP($C$4, 'Resumen_Cartera'!$A$5:$Z${fila_fin_datos}, 14, FALSE), 0)", FMT_CURRENCY),
        (7, 10, "Residual Pactado:", 11, f"=IFERROR(VLOOKUP($C$4, 'Resumen_Cartera'!$A$5:$Z${fila_fin_datos}, 16, FALSE), 0)", FMT_CURRENCY),

        (8, 2, "Tasa Mensual:", 3, f"=IFERROR(VLOOKUP($C$4, 'Resumen_Cartera'!$A$5:$Z${fila_fin_datos}, 17, FALSE), 0)", FMT_PCT_RATE),
        (8, 4, "Tasa Nominal Anual:", 5, f"=IFERROR(VLOOKUP($C$4, 'Resumen_Cartera'!$A$5:$Z${fila_fin_datos}, 19, FALSE), 0)", FMT_PCT_ANNUAL),
        (8, 7, "TIR Anual Efectiva:", 8, f"=IFERROR(VLOOKUP($C$4, 'Resumen_Cartera'!$A$5:$Z${fila_fin_datos}, 20, FALSE), 0)", FMT_PCT_ANNUAL),
        (8, 10, "VP Residual Inicial:", 11, f"=IFERROR(VLOOKUP($C$4, 'Resumen_Cartera'!$A$5:$Z${fila_fin_datos}, 21, FALSE), 0)", FMT_CURRENCY),
    ]

    for r_f, c_lbl, lbl_txt, c_val, form_str, num_f in meta_campos:
        ws_cons.cell(row=r_f, column=c_lbl, value=lbl_txt).font = FONT_BODY_BOLD
        ws_cons.cell(row=r_f, column=c_lbl).alignment = Alignment(horizontal="right", vertical="center")
        cell_v = ws_cons.cell(row=r_f, column=c_val, value=form_str)
        cell_v.font = FONT_BODY
        cell_v.number_format = num_f
        cell_v.alignment = Alignment(horizontal="left", vertical="center")
        cell_v.fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        cell_v.border = BORDER_CELL

    ws_cons.cell(row=9, column=1, value="TABLA DINÁMICA DE AMORTIZACIÓN Y SALDOS: La tabla se corta automáticamente al mes exacto de terminación del plazo contractual ($H$6 meses).").font = FONT_SUBTITLE
    ws_cons.merge_cells("A9:O9")

    cols_cons_dinamica = [
        ("Mes", 8, Alignment(horizontal="center")),                          # A (1)
        ("Fecha_Periodo", 14, Alignment(horizontal="center")),               # B (2) Formula EDATE
        ("Saldo_Inicial_Capital", 18, Alignment(horizontal="right")),        # C (3) Formula Saldo
        ("Tasa_Mensual", 18, Alignment(horizontal="right")),                 # D (4) Formula Tasa (8 dec)
        ("Interes_Leasing", 16, Alignment(horizontal="right")),              # E (5) Formula ROUND(C*D, 2)
        ("Renta_Sin_IVA", 16, Alignment(horizontal="right")),                # F (6) Formula Renta
        ("Capital_Amortizado", 16, Alignment(horizontal="right")),           # G (7) Formula F-E
        ("Saldo_Insoluto_Capital", 18, Alignment(horizontal="right")),       # H (8) Formula C-G
        ("Saldo_Inicial_VP_Residual", 20, Alignment(horizontal="right")),    # I (9) Formula VP Inicial
        ("Interes_Residual", 16, Alignment(horizontal="right")),             # J (10) Formula ROUND(I*D, 2)
        ("Saldo_Final_VP_Residual", 20, Alignment(horizontal="right")),      # K (11) Formula I+J
        ("Ingreso_Financiero_Total", 20, Alignment(horizontal="right")),     # L (12) Formula E+J
        ("Saldo_Intereses_Por_Devengar", 22, Alignment(horizontal="right")), # M (13) Formula SUM(E_sig)
        ("Saldo_Cartera_Total", 18, Alignment(horizontal="right")),          # N (14) Formula H+M
        ("Saldo_Global_Exigible", 18, Alignment(horizontal="right")),        # O (15) Formula N+K
    ]

    fila_hdr_cons = 10
    ws_cons.row_dimensions[fila_hdr_cons].height = 24
    for idx_col, (nom_col, anc, al) in enumerate(cols_cons_dinamica, start=1):
        c = ws_cons.cell(row=fila_hdr_cons, column=idx_col, value=nom_col)
        c.fill = FILL_HEADER
        c.font = FONT_HEADER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER_CELL
        ws_cons.column_dimensions[get_column_letter(idx_col)].width = anc

    fila_cons_inicio = 11
    max_plazo_meses = 72  # Permite contratos de hasta 72 meses

    for m_i in range(1, max_plazo_meses + 1):
        r_c = fila_cons_inicio + m_i - 1

        # Col A: Mes con corte exacto al plazo ($H$6)
        ws_cons.cell(row=r_c, column=1, value=f'=IF(ROW()-10<=$H$6, ROW()-10, "")').alignment = Alignment(horizontal="center")

        # Col B: Fecha con corte exacto
        ws_cons.cell(row=r_c, column=2, value=f'=IF($A{r_c}<="", "", TEXT(EDATE($C$6, $A{r_c}-1), "YYYY-MM"))').alignment = Alignment(horizontal="center")

        # Col C: Saldo Inicial Capital
        if m_i == 1:
            ws_cons.cell(row=r_c, column=3, value=f'=IF($A11<="", "", $E$7)')
        else:
            ws_cons.cell(row=r_c, column=3, value=f'=IF($A{r_c}<="", "", H{r_c - 1})')
        ws_cons.cell(row=r_c, column=3).number_format = FMT_CURRENCY
        ws_cons.cell(row=r_c, column=3).alignment = Alignment(horizontal="right")

        # Col D: Tasa Mensual (completa con 8 decimales)
        ws_cons.cell(row=r_c, column=4, value=f'=IF($A{r_c}<="", "", $C$8)')
        ws_cons.cell(row=r_c, column=4).number_format = FMT_PCT_RATE
        ws_cons.cell(row=r_c, column=4).alignment = Alignment(horizontal="right")

        # Col E: Interes Leasing
        ws_cons.cell(row=r_c, column=5, value=f'=IF($A{r_c}<="", "", ROUND(C{r_c}*D{r_c}, 2))')
        ws_cons.cell(row=r_c, column=5).number_format = FMT_CURRENCY
        ws_cons.cell(row=r_c, column=5).alignment = Alignment(horizontal="right")

        # Col F: Renta Mensual Sin IVA
        ws_cons.cell(row=r_c, column=6, value=f'=IF($A{r_c}<="", "", $H$7)')
        ws_cons.cell(row=r_c, column=6).number_format = FMT_CURRENCY
        ws_cons.cell(row=r_c, column=6).alignment = Alignment(horizontal="right")

        # Col G: Capital Amortizado = Renta - Interes Leasing
        ws_cons.cell(row=r_c, column=7, value=f'=IF($A{r_c}<="", "", F{r_c}-E{r_c})')
        ws_cons.cell(row=r_c, column=7).number_format = FMT_CURRENCY
        ws_cons.cell(row=r_c, column=7).alignment = Alignment(horizontal="right")

        # Col H: Saldo Insoluto Capital = Saldo Inicial - Capital Amortizado
        ws_cons.cell(row=r_c, column=8, value=f'=IF($A{r_c}<="", "", C{r_c}-G{r_c})')
        ws_cons.cell(row=r_c, column=8).number_format = FMT_CURRENCY
        ws_cons.cell(row=r_c, column=8).alignment = Alignment(horizontal="right")

        # Col I: Saldo Inicial VP Residual
        if m_i == 1:
            ws_cons.cell(row=r_c, column=9, value=f'=IF($A11<="", "", $K$8)')
        else:
            ws_cons.cell(row=r_c, column=9, value=f'=IF($A{r_c}<="", "", K{r_c - 1})')
        ws_cons.cell(row=r_c, column=9).number_format = FMT_CURRENCY
        ws_cons.cell(row=r_c, column=9).alignment = Alignment(horizontal="right")

        # Col J: Interes Residual = ROUND(VP_Inicial * Tasa, 2)
        ws_cons.cell(row=r_c, column=10, value=f'=IF($A{r_c}<="", "", ROUND(I{r_c}*D{r_c}, 2))')
        ws_cons.cell(row=r_c, column=10).number_format = FMT_CURRENCY
        ws_cons.cell(row=r_c, column=10).alignment = Alignment(horizontal="right")

        # Col K: Saldo Final VP Residual = VP_Inicial + Interes Residual
        ws_cons.cell(row=r_c, column=11, value=f'=IF($A{r_c}<="", "", I{r_c}+J{r_c})')
        ws_cons.cell(row=r_c, column=11).number_format = FMT_CURRENCY
        ws_cons.cell(row=r_c, column=11).alignment = Alignment(horizontal="right")

        # Col L: Ingreso Financiero Total = Interes Leasing + Interes Residual
        ws_cons.cell(row=r_c, column=12, value=f'=IF($A{r_c}<="", "", E{r_c}+J{r_c})')
        ws_cons.cell(row=r_c, column=12).number_format = FMT_CURRENCY
        ws_cons.cell(row=r_c, column=12).alignment = Alignment(horizontal="right")

        # Col M: Saldo Intereses Remanentes por Devengar
        if m_i < max_plazo_meses:
            ws_cons.cell(row=r_c, column=13, value=f'=IF($A{r_c}<="", "", SUM(E{r_c + 1}:$E${fila_cons_inicio + max_plazo_meses - 1}))')
        else:
            ws_cons.cell(row=r_c, column=13, value=f'=IF($A{r_c}<="", "", 0)')
        ws_cons.cell(row=r_c, column=13).number_format = FMT_CURRENCY
        ws_cons.cell(row=r_c, column=13).alignment = Alignment(horizontal="right")

        # Col N: Saldo Cartera Total = Saldo Insoluto Capital (H) + Saldo Intereses (M)
        ws_cons.cell(row=r_c, column=14, value=f'=IF($A{r_c}<="", "", H{r_c}+M{r_c})')
        ws_cons.cell(row=r_c, column=14).number_format = FMT_CURRENCY
        ws_cons.cell(row=r_c, column=14).alignment = Alignment(horizontal="right")

        # Col O: Saldo Global Exigible = Saldo Cartera Total (N) + Saldo Final VP Residual (K)
        ws_cons.cell(row=r_c, column=15, value=f'=IF($A{r_c}<="", "", N{r_c}+K{r_c})')
        ws_cons.cell(row=r_c, column=15).number_format = FMT_CURRENCY
        ws_cons.cell(row=r_c, column=15).alignment = Alignment(horizontal="right")

        es_par_c = (m_i % 2 == 0)
        for c_k in range(1, 16):
            cell_k = ws_cons.cell(row=r_c, column=c_k)
            cell_k.border = BORDER_CELL
            cell_k.font = FONT_BODY
            if es_par_c:
                cell_k.fill = FILL_ZEBRA

    fila_tot_cons = fila_cons_inicio + max_plazo_meses
    ws_cons.cell(row=fila_tot_cons, column=1, value="TOTALES DEL CONTRATO").alignment = Alignment(horizontal="center")
    ws_cons.merge_cells(f"A{fila_tot_cons}:D{fila_tot_cons}")

    for col_m in [5, 6, 7, 10, 12]:
        letra_m = get_column_letter(col_m)
        c_m_t = ws_cons.cell(row=fila_tot_cons, column=col_m, value=f"=SUM({letra_m}{fila_cons_inicio}:{letra_m}{fila_tot_cons - 1})")
        c_m_t.number_format = FMT_CURRENCY
        c_m_t.alignment = Alignment(horizontal="right")

    for col_k in range(1, 16):
        c_tot_k = ws_cons.cell(row=fila_tot_cons, column=col_k)
        c_tot_k.fill = FILL_TOTAL
        c_tot_k.font = FONT_TOTAL
        c_tot_k.border = BORDER_TOTAL

    ws_cons.freeze_panes = "C11"

    # =========================================================================
    # 4. DEFINICION DE MATRICES CRONOLOGICAS HORIZONTALES (DEVENGOS Y SALDOS)
    # =========================================================================
    matrices_config = [
        # --- BLOQUE 1: DEVENGOS Y FLUJOS MENSUALES ---
        {
            "sheet_name": "Intereses_Leasing_Mensual",
            "titulo": "Intereses de Leasing Devengados Mes a Mes",
            "subtitulo": "Devengo mensual de intereses de leasing con fórmulas vivas de saldo de capital por tasa",
            "campo": "il_formula",
            "fixed_cols": [
                ("ID_Contrato", 14, Alignment(horizontal="center")),
                ("Cliente", 34, Alignment(horizontal="left")),
                ("Vehiculo", 26, Alignment(horizontal="left")),
                ("Estatus", 12, Alignment(horizontal="center")),
                ("Fecha_Alta", 13, Alignment(horizontal="center")),
                ("Fecha_Vencimiento", 15, Alignment(horizontal="center")),
                ("Plazo", 10, Alignment(horizontal="right")),
                ("Inversion_Neta", 16, Alignment(horizontal="right")),
            ],
            "col_total_nombre": "Total_Devengado_Leasing"
        },
        {
            "sheet_name": "Intereses_Residual_Mensual",
            "titulo": "Intereses de Valor Residual Devengados (NIF C-15)",
            "subtitulo": "Acumulación y devengo mensual con fórmulas vivas sobre el saldo activo del residual",
            "campo": "ir_formula",
            "fixed_cols": [
                ("ID_Contrato", 14, Alignment(horizontal="center")),
                ("Cliente", 34, Alignment(horizontal="left")),
                ("Vehiculo", 26, Alignment(horizontal="left")),
                ("Estatus", 12, Alignment(horizontal="center")),
                ("Fecha_Alta", 13, Alignment(horizontal="center")),
                ("Fecha_Vencimiento", 15, Alignment(horizontal="center")),
                ("Plazo", 10, Alignment(horizontal="right")),
                ("Residual_Pactado", 16, Alignment(horizontal="right")),
            ],
            "col_total_nombre": "Total_Devengado_Residual"
        },
        {
            "sheet_name": "Ingreso_Financiero_Total",
            "titulo": "Ingreso Financiero Total Devengado (Leasing + Residual)",
            "subtitulo": "Fórmula inter-hoja sumando intereses de leasing y de residual en cada período",
            "campo": "ingreso_total_formula",
            "fixed_cols": [
                ("ID_Contrato", 14, Alignment(horizontal="center")),
                ("Cliente", 34, Alignment(horizontal="left")),
                ("Vehiculo", 26, Alignment(horizontal="left")),
                ("Estatus", 12, Alignment(horizontal="center")),
                ("Plazo", 10, Alignment(horizontal="right")),
            ],
            "col_total_nombre": "Total_Ingreso_Financiero"
        },
        {
            "sheet_name": "Amortizacion_Capital_Mensual",
            "titulo": "Amortización de Capital de Cartera Mes a Mes",
            "subtitulo": "Recuperación mensual de capital principal mediante fórmula: Capital = Renta - Intereses",
            "campo": "cap_formula",
            "fixed_cols": [
                ("ID_Contrato", 14, Alignment(horizontal="center")),
                ("Cliente", 34, Alignment(horizontal="left")),
                ("Vehiculo", 26, Alignment(horizontal="left")),
                ("Estatus", 12, Alignment(horizontal="center")),
                ("Fecha_Alta", 13, Alignment(horizontal="center")),
                ("Fecha_Vencimiento", 15, Alignment(horizontal="center")),
                ("Plazo", 10, Alignment(horizontal="right")),
                ("Inversion_Neta", 16, Alignment(horizontal="right")),
                ("Renta_Sin_IVA", 16, Alignment(horizontal="right")),
            ],
            "col_total_nombre": "Total_Capital_Recuperado"
        },
        # --- BLOQUE 2: SALDOS INSOLUTOS, SALDOS DE INTERESES Y SALDOS DE CARTERA ---
        {
            "sheet_name": "Saldos_Capital_Mensual",
            "titulo": "Saldos Insolutos de Capital al Cierre de Cada Mes",
            "subtitulo": "Saldo insoluto de principal con fórmulas vivas: Saldo_Anterior - Capital_Amortizado",
            "campo": "sal_formula",
            "fixed_cols": [
                ("ID_Contrato", 14, Alignment(horizontal="center")),
                ("Cliente", 34, Alignment(horizontal="left")),
                ("Vehiculo", 26, Alignment(horizontal="left")),
                ("Estatus", 12, Alignment(horizontal="center")),
                ("Fecha_Alta", 13, Alignment(horizontal="center")),
                ("Fecha_Vencimiento", 15, Alignment(horizontal="center")),
                ("Plazo", 10, Alignment(horizontal="right")),
                ("Inversion_Neta", 16, Alignment(horizontal="right")),
            ],
            "col_total_nombre": "Saldo_Capital_Cierre"
        },
        {
            "sheet_name": "Saldos_Intereses_Mensual",
            "titulo": "Saldos de Intereses por Devengar al Cierre de Cada Mes",
            "subtitulo": "Intereses remanentes con fórmulas vivas: Saldo_Anterior - Intereses_Leasing_Devengados",
            "campo": "sint_formula",
            "fixed_cols": [
                ("ID_Contrato", 14, Alignment(horizontal="center")),
                ("Cliente", 34, Alignment(horizontal="left")),
                ("Vehiculo", 26, Alignment(horizontal="left")),
                ("Estatus", 12, Alignment(horizontal="center")),
                ("Fecha_Alta", 13, Alignment(horizontal="center")),
                ("Fecha_Vencimiento", 15, Alignment(horizontal="center")),
                ("Plazo", 10, Alignment(horizontal="right")),
                ("Total_Interes_Pactado", 16, Alignment(horizontal="right")),
            ],
            "col_total_nombre": "Saldo_Interes_Cierre"
        },
        {
            "sheet_name": "Saldos_Cartera_Total",
            "titulo": "Saldo Total de Cartera al Cierre de Cada Mes (Capital + Intereses)",
            "subtitulo": "Fórmula inter-hoja viva sumando saldo insoluto de capital más saldo de intereses por devengar",
            "campo": "saldo_total_cartera_formula",
            "fixed_cols": [
                ("ID_Contrato", 14, Alignment(horizontal="center")),
                ("Cliente", 34, Alignment(horizontal="left")),
                ("Vehiculo", 26, Alignment(horizontal="left")),
                ("Estatus", 12, Alignment(horizontal="center")),
                ("Plazo", 10, Alignment(horizontal="right")),
            ],
            "col_total_nombre": "Saldo_Cartera_Cierre"
        },
        {
            "sheet_name": "Saldo_Residual_Activo",
            "titulo": "Saldo del Valor Residual Activo al Cierre de Cada Mes",
            "subtitulo": "Evolución mensual con fórmulas vivas: Saldo_Residual_Anterior + Interés_Residual_Mes",
            "campo": "sr_formula",
            "fixed_cols": [
                ("ID_Contrato", 14, Alignment(horizontal="center")),
                ("Cliente", 34, Alignment(horizontal="left")),
                ("Vehiculo", 26, Alignment(horizontal="left")),
                ("Estatus", 12, Alignment(horizontal="center")),
                ("Fecha_Alta", 13, Alignment(horizontal="center")),
                ("Fecha_Vencimiento", 15, Alignment(horizontal="center")),
                ("Plazo", 10, Alignment(horizontal="right")),
                ("Residual_Pactado", 16, Alignment(horizontal="right")),
            ],
            "col_total_nombre": "Saldo_Residual_Cierre"
        },
        {
            "sheet_name": "Saldos_Exigible_Global",
            "titulo": "Saldo Total Exigible Global de Cartera (Capital + Intereses + Residual)",
            "subtitulo": "Fórmula inter-hoja viva sumando saldo total de cartera más saldo de valor residual activo",
            "campo": "saldo_exigible_global_formula",
            "fixed_cols": [
                ("ID_Contrato", 14, Alignment(horizontal="center")),
                ("Cliente", 34, Alignment(horizontal="left")),
                ("Vehiculo", 26, Alignment(horizontal="left")),
                ("Estatus", 12, Alignment(horizontal="center")),
                ("Plazo", 10, Alignment(horizontal="right")),
            ],
            "col_total_nombre": "Saldo_Global_Cierre"
        },
    ]

    for m_cfg in matrices_config:
        s_name = m_cfg["sheet_name"]
        ws_m = wb.create_sheet(title=s_name)
        f_cols = m_cfg["fixed_cols"]
        n_fixed = len(f_cols)
        n_meses = len(lista_meses)
        total_cols_m = n_fixed + n_meses + 1

        _preparar_encabezado_hoja(
            ws_m,
            m_cfg["titulo"],
            m_cfg["subtitulo"],
            total_cols_m,
            titulo_empresa=titulo_empresa,
            logo_path=logo_path
        )

        ws_m.row_dimensions[fila_header_res].height = 26
        for idx_c, (nom_c, anc, al) in enumerate(f_cols, start=1):
            c_cell = ws_m.cell(row=fila_header_res, column=idx_c, value=nom_c)
            c_cell.fill = FILL_HEADER
            c_cell.font = FONT_HEADER
            c_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c_cell.border = BORDER_CELL
            ws_m.column_dimensions[get_column_letter(idx_c)].width = anc

        start_month_col_idx = n_fixed + 1
        for idx_m, (y, m) in enumerate(lista_meses):
            col_pos = start_month_col_idx + idx_m
            mes_str = f"{y:04d}-{m:02d}"
            c_mes = ws_m.cell(row=fila_header_res, column=col_pos, value=mes_str)
            c_mes.fill = FILL_HEADER
            c_mes.font = FONT_HEADER
            c_mes.alignment = Alignment(horizontal="center", vertical="center")
            c_mes.border = BORDER_CELL
            ws_m.column_dimensions[get_column_letter(col_pos)].width = 14

        end_month_col_idx = start_month_col_idx + n_meses - 1
        col_total_row_idx = end_month_col_idx + 1
        c_tot_hdr = ws_m.cell(row=fila_header_res, column=col_total_row_idx, value=m_cfg["col_total_nombre"])
        c_tot_hdr.fill = FILL_SUBHEADER
        c_tot_hdr.font = FONT_HEADER
        c_tot_hdr.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c_tot_hdr.border = BORDER_CELL
        ws_m.column_dimensions[get_column_letter(col_total_row_idx)].width = 20

        campo_m = m_cfg["campo"]
        r_actual = fila_inicio_datos

        for c_item in contratos_pre:
            fa_y = c_item["fa_y"]
            fa_m = c_item["fa_m"]
            fv_y = c_item["fv_y"]
            fv_m = c_item["fv_m"]
            fb_y = c_item["fb_y"]
            fb_m = c_item["fb_m"]
            pl = c_item["pl"]

            ws_m.cell(row=r_actual, column=1, value=c_item["id"]).alignment = f_cols[0][2]
            ws_m.cell(row=r_actual, column=2, value=c_item["cliente"]).alignment = f_cols[1][2]
            ws_m.cell(row=r_actual, column=3, value=c_item["vehiculo"]).alignment = f_cols[2][2]
            ws_m.cell(row=r_actual, column=4, value=c_item["estatus"]).alignment = f_cols[3][2]

            if n_fixed == 5:
                # Ingreso_Financiero_Total, Saldos_Cartera_Total o Saldos_Exigible_Global
                c_p = ws_m.cell(row=r_actual, column=5, value=pl)
                c_p.number_format = FMT_INT
                c_p.alignment = f_cols[4][2]
            elif n_fixed >= 8:
                ws_m.cell(row=r_actual, column=5, value=c_item["fa_str"]).alignment = f_cols[4][2]
                ws_m.cell(row=r_actual, column=6, value=c_item["fv_str"]).alignment = f_cols[5][2]
                c_p = ws_m.cell(row=r_actual, column=7, value=pl)
                c_p.number_format = FMT_INT
                c_p.alignment = f_cols[6][2]

                if "Residual" in f_cols[7][0]:
                    val_col8 = c_item["residual"]
                elif "Interes" in f_cols[7][0]:
                    val_col8 = c_item["tot_interes"]
                else:
                    val_col8 = c_item["inversion"]

                c_monto_fijo = ws_m.cell(row=r_actual, column=8, value=val_col8)
                c_monto_fijo.number_format = FMT_CURRENCY
                c_monto_fijo.alignment = f_cols[7][2]

                if n_fixed == 9:
                    # Amortizacion_Capital_Mensual tiene Renta_Sin_IVA en col 9 (Col I)
                    c_ren_m = ws_m.cell(row=r_actual, column=9, value=c_item["renta"])
                    c_ren_m.number_format = FMT_CURRENCY
                    c_ren_m.alignment = f_cols[8][2]

            for idx_m, (y, m) in enumerate(lista_meses):
                c_pos = start_month_col_idx + idx_m
                col_act_letra = get_column_letter(c_pos)

                esta_activo_mes = True
                if (y < fa_y) or (y == fa_y and m < fa_m):
                    esta_activo_mes = False
                elif (y > fv_y) or (y == fv_y and m > fv_m):
                    esta_activo_mes = False
                elif (y > fb_y) or (y == fb_y and m > fb_m):
                    esta_activo_mes = False

                mt = (y - fa_y) * 12 + (m - fa_m)
                if not (0 <= mt < pl):
                    esta_activo_mes = False

                if not esta_activo_mes:
                    cell_val_obj = ws_m.cell(row=r_actual, column=c_pos, value="")
                    cell_val_obj.number_format = FMT_CURRENCY
                    cell_val_obj.alignment = Alignment(horizontal="right")
                    continue

                # --- 1. Intereses Leasing Mensual (Fórmulas Vivas de Saldo * Tasa) ---
                if campo_m == "il_formula":
                    if mt == 0:
                        f_val = f"=ROUND('Resumen_Cartera'!$K{r_actual} * 'Resumen_Cartera'!$Q{r_actual}, 2)"
                    else:
                        if idx_m > 0:
                            col_ant_letra = get_column_letter(c_pos - 1)
                            f_val = f'=IF(Saldos_Capital_Mensual!{col_ant_letra}{r_actual}>0, ROUND(Saldos_Capital_Mensual!{col_ant_letra}{r_actual} * \'Resumen_Cartera\'!$Q{r_actual}, 2), "")'
                        else:
                            s_prev = round(c_item["sal"][mt - 1], 2)
                            f_val = f"=ROUND({s_prev} * 'Resumen_Cartera'!$Q{r_actual}, 2)"
                    cell_val_obj = ws_m.cell(row=r_actual, column=c_pos, value=f_val)
                    cell_val_obj.number_format = FMT_CURRENCY
                    cell_val_obj.alignment = Alignment(horizontal="right")

                # --- 2. Intereses Residual Mensual (Fórmulas Vivas de Saldo Residual * Tasa) ---
                elif campo_m == "ir_formula":
                    if mt == 0:
                        f_val = f"=ROUND('Resumen_Cartera'!$U{r_actual} * 'Resumen_Cartera'!$Q{r_actual}, 2)"
                    else:
                        if idx_m > 0:
                            col_ant_letra = get_column_letter(c_pos - 1)
                            f_val = f'=IF(Saldo_Residual_Activo!{col_ant_letra}{r_actual}>0, ROUND(Saldo_Residual_Activo!{col_ant_letra}{r_actual} * \'Resumen_Cartera\'!$Q{r_actual}, 2), "")'
                        else:
                            sr_prev = round(c_item["sr"][mt - 1], 2)
                            f_val = f"=ROUND({sr_prev} * 'Resumen_Cartera'!$Q{r_actual}, 2)"
                    cell_val_obj = ws_m.cell(row=r_actual, column=c_pos, value=f_val)
                    cell_val_obj.number_format = FMT_CURRENCY
                    cell_val_obj.alignment = Alignment(horizontal="right")

                # --- 3. Ingreso Financiero Total = Leasing + Residual ---
                elif campo_m == "ingreso_total_formula":
                    col_letra_origen = get_column_letter(8 + 1 + idx_m)
                    f_val = f"=ROUND(Intereses_Leasing_Mensual!{col_letra_origen}{r_actual} + Intereses_Residual_Mensual!{col_letra_origen}{r_actual}, 2)"
                    cell_val_obj = ws_m.cell(row=r_actual, column=c_pos, value=f_val)
                    cell_val_obj.number_format = FMT_CURRENCY
                    cell_val_obj.alignment = Alignment(horizontal="right")

                # --- 4. Amortización de Capital = Renta - Interes Leasing ---
                elif campo_m == "cap_formula":
                    col_letra_origen = get_column_letter(8 + 1 + idx_m)
                    f_cap = f'=IF(Intereses_Leasing_Mensual!{col_letra_origen}{r_actual}>0, ROUND($I{r_actual}-Intereses_Leasing_Mensual!{col_letra_origen}{r_actual}, 2), "")'
                    cell_val_obj = ws_m.cell(row=r_actual, column=c_pos, value=f_cap)
                    cell_val_obj.number_format = FMT_CURRENCY
                    cell_val_obj.alignment = Alignment(horizontal="right")

                # --- 5. Saldos Insolutos de Capital = Saldo_Anterior - Capital ---
                elif campo_m == "sal_formula":
                    col_cap_letra = get_column_letter(9 + 1 + idx_m)
                    if mt == 0:
                        f_val = f'=IF(Amortizacion_Capital_Mensual!{col_cap_letra}{r_actual}>0, ROUND($H{r_actual} - Amortizacion_Capital_Mensual!{col_cap_letra}{r_actual}, 2), "")'
                    else:
                        if idx_m > 0:
                            col_ant_letra = get_column_letter(c_pos - 1)
                            f_val = f'=IF(Amortizacion_Capital_Mensual!{col_cap_letra}{r_actual}>0, ROUND({col_ant_letra}{r_actual} - Amortizacion_Capital_Mensual!{col_cap_letra}{r_actual}, 2), "")'
                        else:
                            s_prev = round(c_item["sal"][mt - 1], 2)
                            f_val = f'=IF(Amortizacion_Capital_Mensual!{col_cap_letra}{r_actual}>0, ROUND({s_prev} - Amortizacion_Capital_Mensual!{col_cap_letra}{r_actual}, 2), "")'
                    cell_val_obj = ws_m.cell(row=r_actual, column=c_pos, value=f_val)
                    cell_val_obj.number_format = FMT_CURRENCY
                    cell_val_obj.alignment = Alignment(horizontal="right")

                # --- 6. Saldos Intereses por Devengar = Saldo_Anterior - Intereses_Leasing ---
                elif campo_m == "sint_formula":
                    col_il_letra = get_column_letter(8 + 1 + idx_m)
                    if mt == 0:
                        f_val = f'=IF(Intereses_Leasing_Mensual!{col_il_letra}{r_actual}>0, ROUND($H{r_actual} - Intereses_Leasing_Mensual!{col_il_letra}{r_actual}, 2), "")'
                    else:
                        if idx_m > 0:
                            col_ant_letra = get_column_letter(c_pos - 1)
                            f_val = f'=IF(Intereses_Leasing_Mensual!{col_il_letra}{r_actual}>0, ROUND({col_ant_letra}{r_actual} - Intereses_Leasing_Mensual!{col_il_letra}{r_actual}, 2), "")'
                        else:
                            sint_prev = round(c_item["sint"][mt - 1], 2)
                            f_val = f'=IF(Intereses_Leasing_Mensual!{col_il_letra}{r_actual}>0, ROUND({sint_prev} - Intereses_Leasing_Mensual!{col_il_letra}{r_actual}, 2), "")'
                    cell_val_obj = ws_m.cell(row=r_actual, column=c_pos, value=f_val)
                    cell_val_obj.number_format = FMT_CURRENCY
                    cell_val_obj.alignment = Alignment(horizontal="right")

                # --- 7. Saldo Total Cartera = Saldo Capital + Saldo Intereses ---
                elif campo_m == "saldo_total_cartera_formula":
                    col_letra_origen = get_column_letter(8 + 1 + idx_m)
                    f_val = f"=ROUND(Saldos_Capital_Mensual!{col_letra_origen}{r_actual} + Saldos_Intereses_Mensual!{col_letra_origen}{r_actual}, 2)"
                    cell_val_obj = ws_m.cell(row=r_actual, column=c_pos, value=f_val)
                    cell_val_obj.number_format = FMT_CURRENCY
                    cell_val_obj.alignment = Alignment(horizontal="right")

                # --- 8. Saldo Residual Activo = Saldo_Anterior + Interes_Residual ---
                elif campo_m == "sr_formula":
                    col_res_dev_letra = get_column_letter(8 + 1 + idx_m)
                    if mt == 0:
                        f_val = f"=ROUND('Resumen_Cartera'!$U{r_actual} + Intereses_Residual_Mensual!{col_res_dev_letra}{r_actual}, 2)"
                    else:
                        if idx_m > 0:
                            col_ant_letra = get_column_letter(c_pos - 1)
                            f_val = f'=IF(Intereses_Residual_Mensual!{col_res_dev_letra}{r_actual}>0, ROUND({col_ant_letra}{r_actual} + Intereses_Residual_Mensual!{col_res_dev_letra}{r_actual}, 2), "")'
                        else:
                            sr_prev = round(c_item["sr"][mt - 1], 2)
                            f_val = f'=IF(Intereses_Residual_Mensual!{col_res_dev_letra}{r_actual}>0, ROUND({sr_prev} + Intereses_Residual_Mensual!{col_res_dev_letra}{r_actual}, 2), "")'
                    cell_val_obj = ws_m.cell(row=r_actual, column=c_pos, value=f_val)
                    cell_val_obj.number_format = FMT_CURRENCY
                    cell_val_obj.alignment = Alignment(horizontal="right")

                # --- 9. Saldo Global Exigible = Saldo Cartera + Saldo Residual Activo ---
                elif campo_m == "saldo_exigible_global_formula":
                    col_tot_letra = get_column_letter(5 + 1 + idx_m)
                    col_res_letra = get_column_letter(8 + 1 + idx_m)
                    f_val = f"=ROUND(Saldos_Cartera_Total!{col_tot_letra}{r_actual} + Saldo_Residual_Activo!{col_res_letra}{r_actual}, 2)"
                    cell_val_obj = ws_m.cell(row=r_actual, column=c_pos, value=f_val)
                    cell_val_obj.number_format = FMT_CURRENCY
                    cell_val_obj.alignment = Alignment(horizontal="right")

            letra_ini_mes = get_column_letter(start_month_col_idx)
            letra_fin_mes = get_column_letter(end_month_col_idx)
            if "saldo" in campo_m or campo_m in ["sal_formula", "sint_formula", "sr_formula"]:
                cell_row_tot = ws_m.cell(row=r_actual, column=col_total_row_idx, value=f"=MAX({letra_ini_mes}{r_actual}:{letra_fin_mes}{r_actual})")
            else:
                cell_row_tot = ws_m.cell(row=r_actual, column=col_total_row_idx, value=f"=SUM({letra_ini_mes}{r_actual}:{letra_fin_mes}{r_actual})")
            cell_row_tot.number_format = FMT_CURRENCY
            cell_row_tot.alignment = Alignment(horizontal="right")

            es_par_m = (r_actual - fila_inicio_datos) % 2 == 1
            for c_i in range(1, total_cols_m + 1):
                c_elem = ws_m.cell(row=r_actual, column=c_i)
                c_elem.border = BORDER_CELL
                c_elem.font = FONT_BODY
                if es_par_m:
                    c_elem.fill = FILL_ZEBRA

            r_actual += 1

        f_fin_m = r_actual - 1
        ws_m.cell(row=r_actual, column=1, value="TOTAL CARTERA").alignment = Alignment(horizontal="center")
        ws_m.cell(row=r_actual, column=2, value=f"{len(contratos_pre)} Contratos").alignment = Alignment(horizontal="left")

        for idx_m in range(n_meses):
            c_pos = start_month_col_idx + idx_m
            letra_col_m = get_column_letter(c_pos)
            cell_m_tot = ws_m.cell(row=r_actual, column=c_pos, value=f"=SUM({letra_col_m}{fila_inicio_datos}:{letra_col_m}{f_fin_m})")
            cell_m_tot.number_format = FMT_CURRENCY
            cell_m_tot.alignment = Alignment(horizontal="right")

        letra_col_tot = get_column_letter(col_total_row_idx)
        cell_gran_tot = ws_m.cell(row=r_actual, column=col_total_row_idx, value=f"=SUM({letra_col_tot}{fila_inicio_datos}:{letra_col_tot}{f_fin_m})")
        cell_gran_tot.number_format = FMT_CURRENCY
        cell_gran_tot.alignment = Alignment(horizontal="right")

        for col_i in range(1, total_cols_m + 1):
            c_tot_m = ws_m.cell(row=r_actual, column=col_i)
            c_tot_m.fill = FILL_TOTAL
            c_tot_m.font = FONT_TOTAL
            c_tot_m.border = BORDER_TOTAL

        ws_m.freeze_panes = "D5"
        ws_m.auto_filter.ref = f"A{fila_header_res}:{get_column_letter(total_cols_m)}{f_fin_m}"

    # Hoja de compatibilidad previa Saldos_Insolutos_Mensual (copia referenciada a Saldos_Capital_Mensual)
    if "Saldos_Capital_Mensual" in wb.sheetnames:
        ws_insol = wb.create_sheet(title="Saldos_Insolutos_Mensual")
        ws_cap_ref = wb["Saldos_Capital_Mensual"]
        _preparar_encabezado_hoja(
            ws_insol,
            "Saldos Insolutos de Capital al Cierre de Cada Mes",
            "Saldo insoluto de cartera activa al término de cada período contable",
            ws_cap_ref.max_column,
            titulo_empresa=titulo_empresa,
            logo_path=logo_path
        )
        for r_k in range(fila_header_res, ws_cap_ref.max_row + 1):
            for c_k in range(1, ws_cap_ref.max_column + 1):
                val_orig = ws_cap_ref.cell(row=r_k, column=c_k).value
                cell_dest = ws_insol.cell(row=r_k, column=c_k)
                if r_k == fila_header_res:
                    cell_dest.value = val_orig
                    cell_dest.fill = FILL_HEADER
                    cell_dest.font = FONT_HEADER
                    cell_dest.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    ws_insol.column_dimensions[get_column_letter(c_k)].width = ws_cap_ref.column_dimensions[get_column_letter(c_k)].width
                elif r_k == ws_cap_ref.max_row:
                    letra_c_k = get_column_letter(c_k)
                    if c_k == 1:
                        cell_dest.value = "TOTAL CARTERA"
                    elif c_k == 2:
                        cell_dest.value = f"{len(contratos_pre)} Contratos"
                    elif c_k >= 9:
                        cell_dest.value = f"=SUM({letra_c_k}5:{letra_c_k}{r_k - 1})"
                        cell_dest.number_format = FMT_CURRENCY
                    cell_dest.fill = FILL_TOTAL
                    cell_dest.font = FONT_TOTAL
                    cell_dest.border = BORDER_TOTAL
                else:
                    if c_k <= 8:
                        cell_dest.value = f"=Saldos_Capital_Mensual!{get_column_letter(c_k)}{r_k}"
                        if c_k in [7]:
                            cell_dest.number_format = FMT_INT
                        elif c_k in [8]:
                            cell_dest.number_format = FMT_CURRENCY
                    else:
                        letra_c_k = get_column_letter(c_k)
                        cell_dest.value = f"=Saldos_Capital_Mensual!{letra_c_k}{r_k}"
                        cell_dest.number_format = FMT_CURRENCY
                    cell_dest.border = BORDER_CELL
                    cell_dest.font = FONT_BODY

    # =========================================================================
    # 5. HOJA: Matriz_Consolidada_Cartera (Desglose Multi-Concepto de Flujos por Mes)
    # =========================================================================
    meses_consolidada = lista_meses if len(lista_meses) <= 24 else lista_meses[-24:]
    ws_cons_cart = wb.create_sheet(title="Matriz_Consolidada_Cartera")

    cols_fijas_cons = [
        ("ID_Contrato", 14, Alignment(horizontal="center")),
        ("Cliente", 34, Alignment(horizontal="left")),
        ("Vehiculo", 26, Alignment(horizontal="left")),
        ("Plazo", 10, Alignment(horizontal="right")),
    ]

    conceptos_mes = [
        ("Int_Leasing", 14),
        ("Capital", 14),
        ("Saldo_Cap", 15),
        ("Int_Residual", 14),
        ("Total_Ingreso", 15),
    ]

    total_cols_cons = len(cols_fijas_cons) + len(meses_consolidada) * len(conceptos_mes)

    _preparar_encabezado_hoja(
        ws_cons_cart,
        "Matriz Cronológica Consolidada Multi-Concepto de Cartera (Flujos y Devengos)",
        "Interés leasing, amortización de capital, saldo insoluto e interés de residual mes por mes con fórmulas vivas",
        total_cols_cons,
        titulo_empresa=titulo_empresa,
        logo_path=logo_path
    )

    ws_cons_cart.row_dimensions[3].height = 20
    ws_cons_cart.row_dimensions[4].height = 22

    for idx_f, (n_f, anc_f, al_f) in enumerate(cols_fijas_cons, start=1):
        ws_cons_cart.merge_cells(start_row=3, start_column=idx_f, end_row=4, end_column=idx_f)
        c_f = ws_cons_cart.cell(row=3, column=idx_f, value=n_f)
        c_f.fill = FILL_HEADER
        c_f.font = FONT_HEADER
        c_f.alignment = Alignment(horizontal="center", vertical="center")
        c_f.border = BORDER_CELL
        ws_cons_cart.column_dimensions[get_column_letter(idx_f)].width = anc_f

    curr_col_cons = len(cols_fijas_cons) + 1
    for y_c, m_c in meses_consolidada:
        mes_etiqueta = f"{y_c:04d}-{m_c:02d}"
        col_ini_bloque = curr_col_cons
        col_fin_bloque = curr_col_cons + len(conceptos_mes) - 1

        ws_cons_cart.merge_cells(start_row=3, start_column=col_ini_bloque, end_row=3, end_column=col_fin_bloque)
        c_bloque = ws_cons_cart.cell(row=3, column=col_ini_bloque, value=mes_etiqueta)
        c_bloque.fill = FILL_HEADER
        c_bloque.font = FONT_HEADER
        c_bloque.alignment = Alignment(horizontal="center", vertical="center")
        c_bloque.border = BORDER_CELL

        for c_nombre, c_ancho in conceptos_mes:
            c_sub = ws_cons_cart.cell(row=4, column=curr_col_cons, value=c_nombre)
            c_sub.fill = FILL_SUBHEADER
            c_sub.font = FONT_SUBHEADER
            c_sub.alignment = Alignment(horizontal="center", vertical="center")
            c_sub.border = BORDER_CELL
            ws_cons_cart.column_dimensions[get_column_letter(curr_col_cons)].width = c_ancho
            curr_col_cons += 1

    r_act_c = 5
    for c_item in contratos_pre:
        ws_cons_cart.cell(row=r_act_c, column=1, value=c_item["id"]).alignment = cols_fijas_cons[0][2]
        ws_cons_cart.cell(row=r_act_c, column=2, value=c_item["cliente"]).alignment = cols_fijas_cons[1][2]
        ws_cons_cart.cell(row=r_act_c, column=3, value=c_item["vehiculo"]).alignment = cols_fijas_cons[2][2]
        c_pl_c = ws_cons_cart.cell(row=r_act_c, column=4, value=c_item["pl"])
        c_pl_c.number_format = FMT_INT
        c_pl_c.alignment = cols_fijas_cons[3][2]

        c_col_cursor = 5
        for y_c, m_c in meses_consolidada:
            idx_mes_global = lista_meses.index((y_c, m_c))
            col_origen_letra_8 = get_column_letter(8 + 1 + idx_mes_global)
            col_origen_letra_9 = get_column_letter(9 + 1 + idx_mes_global)

            # 1. Int Leasing
            f_il = f"=Intereses_Leasing_Mensual!{col_origen_letra_8}{r_act_c}"
            c1 = ws_cons_cart.cell(row=r_act_c, column=c_col_cursor, value=f_il)
            c1.number_format = FMT_CURRENCY
            c1.alignment = Alignment(horizontal="right")

            # 2. Capital
            f_cap = f"=Amortizacion_Capital_Mensual!{col_origen_letra_9}{r_act_c}"
            c2 = ws_cons_cart.cell(row=r_act_c, column=c_col_cursor + 1, value=f_cap)
            c2.number_format = FMT_CURRENCY
            c2.alignment = Alignment(horizontal="right")

            # 3. Saldo Cap
            f_sc = f"=Saldos_Capital_Mensual!{col_origen_letra_8}{r_act_c}"
            c3 = ws_cons_cart.cell(row=r_act_c, column=c_col_cursor + 2, value=f_sc)
            c3.number_format = FMT_CURRENCY
            c3.alignment = Alignment(horizontal="right")

            # 4. Int Residual
            f_ir = f"=Intereses_Residual_Mensual!{col_origen_letra_8}{r_act_c}"
            c4 = ws_cons_cart.cell(row=r_act_c, column=c_col_cursor + 3, value=f_ir)
            c4.number_format = FMT_CURRENCY
            c4.alignment = Alignment(horizontal="right")

            # 5. Total Ingreso Financiero
            letra_il = get_column_letter(c_col_cursor)
            letra_ir = get_column_letter(c_col_cursor + 3)
            f_tot_ing = f"=ROUND({letra_il}{r_act_c} + {letra_ir}{r_act_c}, 2)"
            c5 = ws_cons_cart.cell(row=r_act_c, column=c_col_cursor + 4, value=f_tot_ing)
            c5.number_format = FMT_CURRENCY
            c5.alignment = Alignment(horizontal="right")

            c_col_cursor += 5

        es_par_cons = (r_act_c - 5) % 2 == 1
        for col_idx in range(1, total_cols_cons + 1):
            cell_item = ws_cons_cart.cell(row=r_act_c, column=col_idx)
            cell_item.border = BORDER_CELL
            cell_item.font = FONT_BODY
            if es_par_cons:
                cell_item.fill = FILL_ZEBRA

        r_act_c += 1

    fin_datos_cons = r_act_c - 1
    ws_cons_cart.cell(row=r_act_c, column=1, value="TOTAL CARTERA").alignment = Alignment(horizontal="center")
    ws_cons_cart.cell(row=r_act_c, column=2, value=f"{len(contratos_pre)} Contratos").alignment = Alignment(horizontal="left")

    for col_k in range(5, total_cols_cons + 1):
        letra_col_k = get_column_letter(col_k)
        c_k_tot = ws_cons_cart.cell(row=r_act_c, column=col_k, value=f"=SUM({letra_col_k}5:{letra_col_k}{fin_datos_cons})")
        c_k_tot.number_format = FMT_CURRENCY
        c_k_tot.alignment = Alignment(horizontal="right")

    for col_k in range(1, total_cols_cons + 1):
        c_tot_obj = ws_cons_cart.cell(row=r_act_c, column=col_k)
        c_tot_obj.fill = FILL_TOTAL
        c_tot_obj.font = FONT_TOTAL
        c_tot_obj.border = BORDER_TOTAL

    ws_cons_cart.freeze_panes = "E5"
    ws_cons_cart.auto_filter.ref = f"A4:{get_column_letter(total_cols_cons)}{fin_datos_cons}"

    # =========================================================================
    # 6. HOJA: Matriz_Consolidada_Saldos (Desglose Multi-Concepto de Saldos por Mes)
    # =========================================================================
    ws_cons_s = wb.create_sheet(title="Matriz_Consolidada_Saldos")

    conceptos_saldos_mes = [
        ("Saldo_Capital", 15),
        ("Saldo_Intereses", 15),
        ("Saldo_Cartera", 16),
        ("Saldo_Residual", 15),
        ("Saldo_Global", 16),
    ]

    total_cols_cons_s = len(cols_fijas_cons) + len(meses_consolidada) * len(conceptos_saldos_mes)

    _preparar_encabezado_hoja(
        ws_cons_s,
        "Matriz Cronológica Consolidada de Saldos de Cartera",
        "Saldos de capital, saldos de intereses por devengar, saldos totales de cartera y saldos de residual mes por mes",
        total_cols_cons_s,
        titulo_empresa=titulo_empresa,
        logo_path=logo_path
    )

    ws_cons_s.row_dimensions[3].height = 20
    ws_cons_s.row_dimensions[4].height = 22

    for idx_f, (n_f, anc_f, al_f) in enumerate(cols_fijas_cons, start=1):
        ws_cons_s.merge_cells(start_row=3, start_column=idx_f, end_row=4, end_column=idx_f)
        c_f = ws_cons_s.cell(row=3, column=idx_f, value=n_f)
        c_f.fill = FILL_HEADER
        c_f.font = FONT_HEADER
        c_f.alignment = Alignment(horizontal="center", vertical="center")
        c_f.border = BORDER_CELL
        ws_cons_s.column_dimensions[get_column_letter(idx_f)].width = anc_f

    curr_col_s = len(cols_fijas_cons) + 1
    for y_c, m_c in meses_consolidada:
        mes_etiqueta = f"{y_c:04d}-{m_c:02d}"
        col_ini_bloque = curr_col_s
        col_fin_bloque = curr_col_s + len(conceptos_saldos_mes) - 1

        ws_cons_s.merge_cells(start_row=3, start_column=col_ini_bloque, end_row=3, end_column=col_fin_bloque)
        c_bloque = ws_cons_s.cell(row=3, column=col_ini_bloque, value=mes_etiqueta)
        c_bloque.fill = FILL_HEADER
        c_bloque.font = FONT_HEADER
        c_bloque.alignment = Alignment(horizontal="center", vertical="center")
        c_bloque.border = BORDER_CELL

        for c_nombre, c_ancho in conceptos_saldos_mes:
            c_sub = ws_cons_s.cell(row=4, column=curr_col_s, value=c_nombre)
            c_sub.fill = FILL_SUBHEADER
            c_sub.font = FONT_SUBHEADER
            c_sub.alignment = Alignment(horizontal="center", vertical="center")
            c_sub.border = BORDER_CELL
            ws_cons_s.column_dimensions[get_column_letter(curr_col_s)].width = c_ancho
            curr_col_s += 1

    r_act_cs = 5
    for c_item in contratos_pre:
        ws_cons_s.cell(row=r_act_cs, column=1, value=c_item["id"]).alignment = cols_fijas_cons[0][2]
        ws_cons_s.cell(row=r_act_cs, column=2, value=c_item["cliente"]).alignment = cols_fijas_cons[1][2]
        ws_cons_s.cell(row=r_act_cs, column=3, value=c_item["vehiculo"]).alignment = cols_fijas_cons[2][2]
        c_pl_s = ws_cons_s.cell(row=r_act_cs, column=4, value=c_item["pl"])
        c_pl_s.number_format = FMT_INT
        c_pl_s.alignment = cols_fijas_cons[3][2]

        c_col_cursor_s = 5
        for y_c, m_c in meses_consolidada:
            idx_mes_global = lista_meses.index((y_c, m_c))
            col_origen_letra_8 = get_column_letter(8 + 1 + idx_mes_global)

            # 1. Saldo Capital
            f_scap = f"=Saldos_Capital_Mensual!{col_origen_letra_8}{r_act_cs}"
            cs1 = ws_cons_s.cell(row=r_act_cs, column=c_col_cursor_s, value=f_scap)
            cs1.number_format = FMT_CURRENCY
            cs1.alignment = Alignment(horizontal="right")

            # 2. Saldo Intereses por Devengar
            f_sint = f"=Saldos_Intereses_Mensual!{col_origen_letra_8}{r_act_cs}"
            cs2 = ws_cons_s.cell(row=r_act_cs, column=c_col_cursor_s + 1, value=f_sint)
            cs2.number_format = FMT_CURRENCY
            cs2.alignment = Alignment(horizontal="right")

            # 3. Saldo Cartera Total = Saldo Cap + Saldo Int
            letra_scap = get_column_letter(c_col_cursor_s)
            letra_sint = get_column_letter(c_col_cursor_s + 1)
            f_scart = f"=ROUND({letra_scap}{r_act_cs} + {letra_sint}{r_act_cs}, 2)"
            cs3 = ws_cons_s.cell(row=r_act_cs, column=c_col_cursor_s + 2, value=f_scart)
            cs3.number_format = FMT_CURRENCY
            cs3.alignment = Alignment(horizontal="right")

            # 4. Saldo Residual Activo
            f_sres = f"=Saldo_Residual_Activo!{col_origen_letra_8}{r_act_cs}"
            cs4 = ws_cons_s.cell(row=r_act_cs, column=c_col_cursor_s + 3, value=f_sres)
            cs4.number_format = FMT_CURRENCY
            cs4.alignment = Alignment(horizontal="right")

            # 5. Saldo Global Integral = Saldo Cartera + Saldo Residual
            letra_scart = get_column_letter(c_col_cursor_s + 2)
            letra_sres = get_column_letter(c_col_cursor_s + 3)
            f_sglob = f"=ROUND({letra_scart}{r_act_cs} + {letra_sres}{r_act_cs}, 2)"
            cs5 = ws_cons_s.cell(row=r_act_cs, column=c_col_cursor_s + 4, value=f_sglob)
            cs5.number_format = FMT_CURRENCY
            cs5.alignment = Alignment(horizontal="right")

            c_col_cursor_s += 5

        es_par_cons_s = (r_act_cs - 5) % 2 == 1
        for col_idx in range(1, total_cols_cons_s + 1):
            cell_item = ws_cons_s.cell(row=r_act_cs, column=col_idx)
            cell_item.border = BORDER_CELL
            cell_item.font = FONT_BODY
            if es_par_cons_s:
                cell_item.fill = FILL_ZEBRA

        r_act_cs += 1

    fin_datos_cons_s = r_act_cs - 1
    ws_cons_s.cell(row=r_act_cs, column=1, value="TOTAL CARTERA").alignment = Alignment(horizontal="center")
    ws_cons_s.cell(row=r_act_cs, column=2, value=f"{len(contratos_pre)} Contratos").alignment = Alignment(horizontal="left")

    for col_k in range(5, total_cols_cons_s + 1):
        letra_col_k = get_column_letter(col_k)
        c_k_tot = ws_cons_s.cell(row=r_act_cs, column=col_k, value=f"=SUM({letra_col_k}5:{letra_col_k}{fin_datos_cons_s})")
        c_k_tot.number_format = FMT_CURRENCY
        c_k_tot.alignment = Alignment(horizontal="right")

    for col_k in range(1, total_cols_cons_s + 1):
        c_tot_obj = ws_cons_s.cell(row=r_act_cs, column=col_k)
        c_tot_obj.fill = FILL_TOTAL
        c_tot_obj.font = FONT_TOTAL
        c_tot_obj.border = BORDER_TOTAL

    ws_cons_s.freeze_panes = "E5"
    ws_cons_s.auto_filter.ref = f"A4:{get_column_letter(total_cols_cons_s)}{fin_datos_cons_s}"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
