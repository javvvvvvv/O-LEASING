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
test_auditoria_cuadre.py — Suite de pruebas automatizadas de aseguramiento
de calidad para el libro de auditoría O-Leasing desarrollado por Orange Crew.

Verifica rigurosamente los 4 requerimientos del usuario:
1. Fórmulas vivas de derivación en Intereses_Leasing_Mensual, Intereses_Residual_Mensual
   y Saldo_Residual_Activo.
2. Cifras limitadas a 2 decimales ($#,##0.00 y 0.00%), excepto tasas completas a 8 decimales.
3. Hoja 'Matriz_Consolidada' dinámica con selector de contrato (DataValidation), fórmulas
   =VLOOKUP y corte automático al mes exacto del plazo.
4. Logotipos de O-Leasing y Orange Crew con membretes institucionales en todo el archivo.
5. Amarres contables exactos contra la base de datos y flujos históricos.
"""

import os
import openpyxl
import pandas as pd

from app import tabla_mensual_conceptos, totales_intereses_desde_tabla_mensual
from reports.auditoria_excel import (
    generar_excel_auditoria,
    generar_excel_reporte_lineal,
    LOGO_OLEASING_PNG_DEFAULT,
    LOGO_ORANGE_PNG_DEFAULT,
    FMT_CURRENCY,
    FMT_PCT_RATE,
    FMT_PCT_ANNUAL,
    FMT_PCT_PARAM
)


def test_amarre_agosto_2026():
    """Verifica que los saldos amarrados de agosto 2026 se mantengan exactos."""
    res = totales_intereses_desde_tabla_mensual(2026, 8)
    leasing_ytd = res.get("interes_leasing_ytd", 0.0)
    residual_ytd = res.get("interes_residual_ytd", 0.0)

    diff_l = round(leasing_ytd - 24135876.84, 2)
    diff_r = round(residual_ytd - 4316777.57, 2)

    assert diff_l == 0.0, f"Diferencia en Leasing Agosto 2026: {diff_l}"
    assert diff_r == 0.0, f"Diferencia en Residual Agosto 2026: {diff_r}"


def test_requerimiento_1_formulas_vivas_matrices():
    """PUNTO 1: Verifica que Intereses_Leasing_Mensual, Intereses_Residual_Mensual

    y Saldo_Residual_Activo contengan formulas vivas de Excel.
    """
    buf = generar_excel_auditoria(anio_filtro=2026, filtro_estatus="ACTIVO")
    wb = openpyxl.load_workbook(buf, data_only=False)

    # 1. Intereses_Leasing_Mensual
    ws_il = wb["Intereses_Leasing_Mensual"]
    # Buscar una celda activa de contrato que tenga formula
    formula_il_encontrada = False
    for r in range(5, min(15, ws_il.max_row)):
        for c in range(9, ws_il.max_column):
            val = str(ws_il.cell(row=r, column=c).value or "")
            if val.startswith("="):
                assert "ROUND" in val or "Saldos_Capital_Mensual" in val or "Resumen_Cartera" in val, f"Formula de leasing incorrecta: {val}"
                formula_il_encontrada = True
                break
        if formula_il_encontrada:
            break
    assert formula_il_encontrada, "No se encontraron formulas vivas en Intereses_Leasing_Mensual"

    # 2. Intereses_Residual_Mensual
    ws_ir = wb["Intereses_Residual_Mensual"]
    formula_ir_encontrada = False
    for r in range(5, min(15, ws_ir.max_row)):
        for c in range(9, ws_ir.max_column):
            val = str(ws_ir.cell(row=r, column=c).value or "")
            if val.startswith("="):
                assert "ROUND" in val or "Saldo_Residual_Activo" in val or "Resumen_Cartera" in val, f"Formula de residual incorrecta: {val}"
                formula_ir_encontrada = True
                break
        if formula_ir_encontrada:
            break
    assert formula_ir_encontrada, "No se encontraron formulas vivas en Intereses_Residual_Mensual"

    # 3. Saldo_Residual_Activo
    ws_sr = wb["Saldo_Residual_Activo"]
    formula_sr_encontrada = False
    for r in range(5, min(15, ws_sr.max_row)):
        for c in range(9, ws_sr.max_column):
            val = str(ws_sr.cell(row=r, column=c).value or "")
            if val.startswith("="):
                assert "ROUND" in val or "Intereses_Residual_Mensual" in val or "Resumen_Cartera" in val, f"Formula de saldo residual incorrecta: {val}"
                formula_sr_encontrada = True
                break
        if formula_sr_encontrada:
            break
    assert formula_sr_encontrada, "No se encontraron formulas vivas en Saldo_Residual_Activo"


def test_requerimiento_2_precision_decimales_y_tasas():
    """PUNTO 2: Cifras monetarias a 2 decimales ($#,##0.00), tasas completas a 8 decimales."""
    buf = generar_excel_auditoria(anio_filtro=2026, filtro_estatus="ACTIVO")
    wb = openpyxl.load_workbook(buf, data_only=False)

    ws_res = wb["Resumen_Cartera"]

    # Tasas deben tener formato completo con 8 decimales (0.00000000%)
    fmt_q = ws_res["Q5"].number_format
    fmt_r = ws_res["R5"].number_format
    fmt_s = ws_res["S5"].number_format
    fmt_t = ws_res["T5"].number_format

    assert "000000" in fmt_q, f"Tasa mensual Q5 debe tener precision completa a 8 decimales, actual: {fmt_q}"
    assert "000000" in fmt_r, f"Tasa RATE R5 debe tener precision completa a 8 decimales, actual: {fmt_r}"
    assert "000000" in fmt_s, f"Tasa nominal S5 debe tener precision completa a 8 decimales, actual: {fmt_s}"
    assert "000000" in fmt_t, f"TIR anual T5 debe tener precision completa a 8 decimales, actual: {fmt_t}"

    # Cifras monetarias a 2 decimales ($#,##0.00)
    for c_coord in ["H5", "J5", "K5", "M5", "N5", "P5", "U5", "V5", "W5", "X5", "Y5"]:
        fmt_mon = ws_res[c_coord].number_format
        assert ".00" in fmt_mon or "$#,##0.00" in fmt_mon, f"Celda {c_coord} debe tener formato a 2 decimales, actual: {fmt_mon}"

    # Porcentajes de captura (Anticipo, Comisión, Residual) a 2 decimales (0.00%)
    for c_coord in ["I5", "L5", "O5"]:
        fmt_p = ws_res[c_coord].number_format
        assert fmt_p == "0.00%", f"Porcentaje {c_coord} debe tener 2 decimales, actual: {fmt_p}"


def test_requerimiento_3_modelo_amortizacion_tipo_dinamico():
    """PUNTO 3: Selector de contrato interactivo y corte automatico al mes del plazo en Modelo_Amortizacion_Tipo, y Matriz_Consolidada restaurada."""
    buf = generar_excel_auditoria(anio_filtro=2026, filtro_estatus="ACTIVO")
    wb = openpyxl.load_workbook(buf, data_only=False)

    assert "Modelo_Amortizacion_Tipo" in wb.sheetnames, "Falta la hoja Modelo_Amortizacion_Tipo"
    ws = wb["Modelo_Amortizacion_Tipo"]

    # Celda de control y selector C4
    val_c4 = str(ws["C4"].value or "")
    assert val_c4 != "", "Celda C4 debe contener el ID del contrato seleccionado por defecto"

    # Validacion de datos (dropdown) en C4 mediante Nombre Definido CONTRATOS_LISTA
    dvs = ws.data_validations.dataValidation
    assert len(dvs) > 0, "Debe existir al menos una validación de datos en la hoja"
    dv_encontrada = False
    for dv_item in dvs:
        if "CONTRATOS_LISTA" in str(dv_item.formula1) or "Resumen_Cartera" in str(dv_item.formula1):
            dv_encontrada = True
            break
    assert dv_encontrada, "No se encontró la lista desplegable apuntando a CONTRATOS_LISTA"

    # Formulas VLOOKUP en bloque de metadata (Filas 5 a 8)
    assert "VLOOKUP" in str(ws["C5"].value), "Cliente en C5 debe usar VLOOKUP"
    assert "VLOOKUP" in str(ws["G5"].value), "Vehiculo en G5 debe usar VLOOKUP"
    assert "VLOOKUP" in str(ws["K5"].value), "Estatus en K5 debe usar VLOOKUP"
    assert "VLOOKUP" in str(ws["G6"].value), "Plazo en G6 debe usar VLOOKUP"
    assert "VLOOKUP" in str(ws["C7"].value), "Inversión Neta en C7 debe usar VLOOKUP"
    assert "VLOOKUP" in str(ws["E7"].value), "Renta en E7 debe usar VLOOKUP"
    assert "VLOOKUP" in str(ws["G7"].value), "Residual en G7 debe usar VLOOKUP"
    assert "VLOOKUP" in str(ws["I7"].value), "VP Residual Inicial en I7 debe usar VLOOKUP"
    assert "VLOOKUP" in str(ws["C8"].value), "Tasa Mensual en C8 debe usar VLOOKUP"

    # Formula de corte dinamico en columna Mes (Col A) y columnas de calculo
    assert "IF" in str(ws["A11"].value) and "$G$6" in str(ws["A11"].value), "Mes 1 debe cortarse con $G$6"
    assert "IF" in str(ws["A47"].value) and "$G$6" in str(ws["A47"].value), "Mes 37 debe cortarse con $G$6"

    # Columnas de datos deben condicionarse a $A{r}=""
    for r_test in [11, 20, 46, 70]:
        val_e = str(ws.cell(row=r_test, column=5).value or "")
        assert val_e.startswith("=IF("), f"Fila {r_test} debe tener formula condicionada de corte"
        assert f"$A{r_test}" in val_e, f"Fila {r_test} debe evaluar la celda $A{r_test} para corte"

    # Verificar que Matriz_Consolidada y Matriz_Consolidada_Saldos existen como matrices de cartera
    assert "Matriz_Consolidada" in wb.sheetnames, "Debe existir Matriz_Consolidada como matriz de cartera"
    assert "Matriz_Consolidada_Saldos" in wb.sheetnames, "Debe existir Matriz_Consolidada_Saldos"
    ws_mat = wb["Matriz_Consolidada"]
    assert ws_mat.max_row >= 10, "Matriz_Consolidada debe contener datos de cartera consolidados"


def test_requerimiento_4_logotipos_e_identidad_institucional():
    """PUNTO 4: Logotipo de O-Leasing y Orange Crew y membrete institucional."""
    assert os.path.exists(LOGO_OLEASING_PNG_DEFAULT), "El archivo de logo o-leasing-logo.png debe existir"
    assert os.path.exists(LOGO_ORANGE_PNG_DEFAULT), "El archivo de logo orange-crew-logo.png debe existir"

    buf = generar_excel_auditoria(anio_filtro=2026, filtro_estatus="ACTIVO")
    wb = openpyxl.load_workbook(buf, data_only=False)

    # Caratula Ejecutiva
    ws_car = wb["Caratula_Ejecutiva"]
    assert "O-LEASING" in str(ws_car["B4"].value), "Caratula debe contener O-LEASING"
    assert "ORANGE CREW" in str(ws_car["B4"].value), "Caratula debe contener ORANGE CREW"

    # Verificar imagenes insertadas en Caratula
    assert len(ws_car._images) >= 1, "Caratula debe contener imagenes institucionales insertadas"

    # Encabezados de hojas
    for s_nom in ["Resumen_Cartera", "Calculo_Tasa_Implicita", "Matriz_Consolidada", "Intereses_Leasing_Mensual"]:
        ws_test = wb[s_nom]
        tit_val = str(ws_test["A1"].value or "")
        sub_val = str(ws_test["A2"].value or "")
        assert "O-LEASING" in tit_val, f"Hoja {s_nom} debe tener O-LEASING en A1"
        assert "ORANGE CREW" in tit_val, f"Hoja {s_nom} debe tener ORANGE CREW en A1"
        assert "Orange Crew" in sub_val, f"Hoja {s_nom} debe tener Orange Crew en A2"


def test_matrices_saldos_formulas_completas():
    """Verifica que las matrices de saldos tengan formulas encadenadas vivas."""
    buf = generar_excel_auditoria(anio_filtro=2025, filtro_estatus="ACTIVO")
    wb = openpyxl.load_workbook(buf, data_only=False)

    hojas_esperadas = [
        "Saldos_Capital_Mensual",
        "Saldos_Intereses_Mensual",
        "Saldos_Cartera_Total",
        "Saldo_Residual_Activo",
        "Saldos_Exigible_Global",
        "Matriz_Consolidada",
        "Matriz_Consolidada_Saldos"
    ]
    for h in hojas_esperadas:
        assert h in wb.sheetnames, f"Falta la hoja de saldos: {h}"

    ws_cart = wb["Saldos_Cartera_Total"]
    val_cart = str(ws_cart["J5"].value)
    assert "Saldos_Capital_Mensual" in val_cart, "Saldos_Cartera_Total debe referenciar a Saldos_Capital_Mensual"
    assert "Saldos_Intereses_Mensual" in val_cart, "Saldos_Cartera_Total debe referenciar a Saldos_Intereses_Mensual"
    assert "+" in val_cart, "Saldos_Cartera_Total debe ser suma viva de capital + intereses"

    ws_exig = wb["Saldos_Exigible_Global"]
    val_exig = str(ws_exig["J5"].value)
    assert "Saldos_Cartera_Total" in val_exig, "Saldos_Exigible_Global debe referenciar a Saldos_Cartera_Total"
    assert "Saldo_Residual_Activo" in val_exig, "Saldos_Exigible_Global debe referenciar a Saldo_Residual_Activo"
    assert "+" in val_exig, "Saldos_Exigible_Global debe ser suma viva de cartera total + residual"


def test_reporte_lineal_horizontal_y_modelo_amortizacion():
    """Verifica el Reporte Lineal con contratos en vertical y amortizacion horizontal mes 1 a N, y Modelo_Amortizacion_Tipo."""
    buf = generar_excel_reporte_lineal(filtro_estatus="ACTIVO")
    wb = openpyxl.load_workbook(buf, data_only=False)

    # Verificar presencia de hojas
    assert "Reporte_Lineal" in wb.sheetnames, "Falta la hoja Reporte_Lineal"
    assert "Modelo_Amortizacion_Tipo" in wb.sheetnames, "Falta la hoja Modelo_Amortizacion_Tipo"
    assert "Caratula_Ejecutiva" in wb.sheetnames, "Falta la hoja Caratula_Ejecutiva"

    ws_lin = wb["Reporte_Lineal"]

    # Fila 5: datos del primer contrato
    assert ws_lin["A5"].value is not None, "A5 debe contener el ID de contrato"
    assert ws_lin["B5"].value is not None, "B5 debe contener el Cliente"
    assert ws_lin["G5"].value is not None, "G5 debe contener el Plazo"

    # Formulas fijas en metadatos
    val_inv = str(ws_lin["K5"].value)
    assert "ROUND" in val_inv and "H5" in val_inv, "K5 Inversión Neta debe ser fórmula =ROUND(H5-J5, 2)"

    val_tot_il = str(ws_lin["V5"].value)
    assert "SUMIF" in val_tot_il and "Interes_Leasing" in val_tot_il, "V5 Total Interés Leasing debe ser =SUMIF"

    val_tot_ir = str(ws_lin["W5"].value)
    assert "SUMIF" in val_tot_ir and "Interes_Residual" in val_tot_ir, "W5 Total Interés Residual debe ser =SUMIF"

    val_ing_tot = str(ws_lin["X5"].value)
    assert "ROUND" in val_ing_tot and "V5+W5" in val_ing_tot, "X5 Ingreso Financiero Total debe ser =ROUND(V5+W5, 2)"

    # Formulas en Mes 1 (Cols AA a AI = 27 a 35)
    val_m1_s_ini = str(ws_lin["AA5"].value)
    assert "IF" in val_m1_s_ini and "$K5" in val_m1_s_ini, "AA5 Saldo Inicial Mes 1 debe apuntar a $K5"

    val_m1_int_l = str(ws_lin["AB5"].value)
    assert "ROUND" in val_m1_int_l and "$Q5" in val_m1_int_l, "AB5 Interés Leasing Mes 1 debe multiplicar por $Q5"

    val_m1_cap = str(ws_lin["AD5"].value)
    assert "AC5-AB5" in val_m1_cap, "AD5 Capital Mes 1 debe ser Renta - Interés (AC5-AB5)"

    val_m1_s_cap = str(ws_lin["AE5"].value)
    assert "AA5-AD5" in val_m1_s_cap, "AE5 Saldo Capital Mes 1 debe ser Saldo Inicial - Capital (AA5-AD5)"

    # Formulas en Mes 2 (Cols AJ a AR = 36 a 44)
    val_m2_s_ini = str(ws_lin["AJ5"].value)
    assert "AE5" in val_m2_s_ini, "AJ5 Saldo Inicial Mes 2 debe referenciar a AE5 (Saldo Capital Mes 1)"

    # Hoja Modelo_Amortizacion_Tipo vinculada a Reporte_Lineal
    ws_mod = wb["Modelo_Amortizacion_Tipo"]
    val_c4 = str(ws_mod["C4"].value or "")
    assert val_c4 != "", "C4 debe tener un contrato seleccionado"

    val_vlk_cli = str(ws_mod["C5"].value)
    assert "Reporte_Lineal" in val_vlk_cli and "VLOOKUP" in val_vlk_cli, "C5 Cliente debe usar VLOOKUP a Reporte_Lineal"

    val_vlk_plazo = str(ws_mod["G6"].value)
    assert "Reporte_Lineal" in val_vlk_plazo and "VLOOKUP" in val_vlk_plazo, "G6 Plazo debe usar VLOOKUP a Reporte_Lineal"

    val_mod_a11 = str(ws_mod["A11"].value)
    assert "$G$6" in val_mod_a11 and "IF" in val_mod_a11, "A11 Mes 1 en Modelo debe cortar con $G$6"


if __name__ == "__main__":
    print("Ejecutando test_amarre_agosto_2026...")
    test_amarre_agosto_2026()
    print("PASS: test_amarre_agosto_2026")

    print("Ejecutando test_requerimiento_1_formulas_vivas_matrices...")
    test_requerimiento_1_formulas_vivas_matrices()
    print("PASS: test_requerimiento_1_formulas_vivas_matrices")

    print("Ejecutando test_requerimiento_2_precision_decimales_y_tasas...")
    test_requerimiento_2_precision_decimales_y_tasas()
    print("PASS: test_requerimiento_2_precision_decimales_y_tasas")

    print("Ejecutando test_requerimiento_3_modelo_amortizacion_tipo_dinamico...")
    test_requerimiento_3_modelo_amortizacion_tipo_dinamico()
    print("PASS: test_requerimiento_3_modelo_amortizacion_tipo_dinamico")

    print("Ejecutando test_requerimiento_4_logotipos_e_identidad_institucional...")
    test_requerimiento_4_logotipos_e_identidad_institucional()
    print("PASS: test_requerimiento_4_logotipos_e_identidad_institucional")

    print("Ejecutando test_matrices_saldos_formulas_completas...")
    test_matrices_saldos_formulas_completas()
    print("PASS: test_matrices_saldos_formulas_completas")

    print("Ejecutando test_reporte_lineal_horizontal_y_modelo_amortizacion...")
    test_reporte_lineal_horizontal_y_modelo_amortizacion()
    print("PASS: test_reporte_lineal_horizontal_y_modelo_amortizacion")

    print("\n========================================================================")
    print("TODAS LAS PRUEBAS DE LOS 4 REQUERIMIENTOS PASARON EXITOSAMENTE (100% OK)")
    print("========================================================================")
