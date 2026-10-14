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
test_auditoria_cuadre.py — Suite de pruebas automatizadas para garantizar
que el libro de auditoría en Excel cuadre exactamente con la fuente de
verdad contable en el pasado, presente y con el paso del tiempo.
"""

import openpyxl
import pandas as pd

from app import tabla_mensual_conceptos, totales_intereses_desde_tabla_mensual
from reports.auditoria_excel import generar_excel_auditoria

MN = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
]


def test_amarre_agosto_2026():
    """Verifica que los saldos amarrados de agosto 2026 se mantengan exactos."""
    res = totales_intereses_desde_tabla_mensual(2026, 8)
    leasing_ytd = res.get("interes_leasing_ytd", 0.0)
    residual_ytd = res.get("interes_residual_ytd", 0.0)

    diff_l = round(leasing_ytd - 24135876.84, 2)
    diff_r = round(residual_ytd - 4316777.57, 2)

    assert diff_l == 0.0, f"Diferencia en Leasing Agosto 2026: {diff_l}"
    assert diff_r == 0.0, f"Diferencia en Residual Agosto 2026: {diff_r}"


def test_amarre_multianual_todos_los_meses():
    """Verifica que cada uno de los 105 meses (2022-2030) cuadre exactamente al centavo."""
    buf = generar_excel_auditoria(anio_filtro=None, filtro_estatus="TODOS")
    wb = openpyxl.load_workbook(buf, data_only=True)

    ws_il = wb["Intereses_Leasing_Mensual"]
    ws_ir = wb["Intereses_Residual_Mensual"]

    # Extraer columnas de meses dinamicamente
    col_meses = {}
    for c in range(1, ws_il.max_column):
        val = ws_il.cell(row=4, column=c).value
        if val and "-" in str(val):
            parts = str(val).split("-")
            if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                y, m = int(parts[0]), int(parts[1])
                col_meses[(y, m)] = c

    errores = []
    for (y, m), col_idx in col_meses.items():
        sum_il_excel = 0.0
        sum_ir_excel = 0.0
        for r in range(5, ws_il.max_row):
            v1 = ws_il.cell(row=r, column=col_idx).value
            v2 = ws_ir.cell(row=r, column=col_idx).value
            if v1:
                sum_il_excel += float(v1)
            if v2:
                sum_ir_excel += float(v2)
        sum_il_excel = round(sum_il_excel, 2)
        sum_ir_excel = round(sum_ir_excel, 2)

        di, dr, _, _, _ = tabla_mensual_conceptos(y)
        nombre_mes = MN[m - 1]
        sum_il_sis = round(float(di[nombre_mes].sum(skipna=True)), 2) if (di is not None and nombre_mes in di.columns) else 0.0
        sum_ir_sis = round(float(dr[nombre_mes].sum(skipna=True)), 2) if (dr is not None and nombre_mes in dr.columns) else 0.0

        diff_l = round(sum_il_excel - sum_il_sis, 2)
        diff_r = round(sum_ir_excel - sum_ir_sis, 2)

        if diff_l != 0.0 or diff_r != 0.0:
            errores.append(f"{y}-{m:02d}: Leasing diff={diff_l}, Residual diff={diff_r}")

    assert len(errores) == 0, f"Discrepancias encontradas: {errores}"


def test_formulas_vivas_excel():
    """Verifica que el libro contenga formulas vivas y no valores planos."""
    buf = generar_excel_auditoria(anio_filtro=2025, filtro_estatus="ACTIVO")
    wb = openpyxl.load_workbook(buf, data_only=False)

    ws_res = wb["Resumen_Cartera"]
    assert str(ws_res["J5"].value).startswith("="), "Anticipo_Monto debe ser formula =H*I"
    assert str(ws_res["K5"].value).startswith("="), "Inversion_Neta debe ser formula =H-J"
    assert str(ws_res["M5"].value).startswith("="), "Comision_Apertura debe ser formula =H*L"
    assert str(ws_res["P5"].value).startswith("="), "Residual_Pactado debe ser formula =H*O"
    assert "RATE" in str(ws_res["R5"].value), "Tasa mensual debe usar formula RATE"
    assert "^12" in str(ws_res["T5"].value), "TIR debe ser formula anualizada"
    assert "*" in str(ws_res["V5"].value) or "SUM" in str(ws_res["V5"].value), "Interes leasing debe ser formula"

    ws_tot = wb["Ingreso_Financiero_Total"]
    assert "Intereses_Leasing_Mensual" in str(ws_tot["F5"].value), "Debe tener formula inter-hoja de leasing"
    assert "Intereses_Residual_Mensual" in str(ws_tot["F5"].value), "Debe tener formula inter-hoja de residual"

    ws_cap = wb["Amortizacion_Capital_Mensual"]
    assert "Intereses_Leasing_Mensual" in str(ws_cap["J5"].value), "Debe tener formula de capital = Renta - Intereses"

    ws_mod = wb["Modelo_Amortizacion_Tipo"]
    assert "=" in str(ws_mod["D9"].value), "Modelo didactico debe tener formulas de interes"
    assert "=" in str(ws_mod["F9"].value), "Modelo didactico debe tener formulas de capital"


def test_hoja_calculo_tasa_implicita():
    """Verifica que la hoja Calculo_Tasa_Implicita contenga formulas de RATE, PV y comprobacion."""
    buf = generar_excel_auditoria(anio_filtro=2025, filtro_estatus="ACTIVO")
    wb = openpyxl.load_workbook(buf, data_only=False)

    assert "Calculo_Tasa_Implicita" in wb.sheetnames, "Falta la hoja Calculo_Tasa_Implicita"
    ws = wb["Calculo_Tasa_Implicita"]
    assert str(ws["H5"].value).startswith("="), "Anticipo Monto debe ser formula =F*G"
    assert str(ws["I5"].value).startswith("="), "Inversion Neta debe ser formula =F-H"
    assert str(ws["L5"].value).startswith("="), "Residual Pactado debe ser formula =F*K"
    assert "RATE" in str(ws["N5"].value), "Tasa mensual debe usar RATE"
    assert "PV" in str(ws["O5"].value), "Comprobacion debe usar formula PV"
    assert str(ws["P5"].value).startswith("="), "Diferencia de cuadre debe ser formula"
    assert "^12" in str(ws["R5"].value), "TIR anual debe ser formula"


if __name__ == "__main__":
    print("Ejecutando test_amarre_agosto_2026...")
    test_amarre_agosto_2026()
    print("PASS: test_amarre_agosto_2026")

    print("Ejecutando test_formulas_vivas_excel...")
    test_formulas_vivas_excel()
    print("PASS: test_formulas_vivas_excel")

    print("Ejecutando test_hoja_calculo_tasa_implicita...")
    test_hoja_calculo_tasa_implicita()
    print("PASS: test_hoja_calculo_tasa_implicita")

    print("Ejecutando test_amarre_multianual_todos_los_meses...")
    test_amarre_multianual_todos_los_meses()
    print("PASS: test_amarre_multianual_todos_los_meses")

    print("\nTODAS LAS PRUEBAS DE AMARRE Y CONTINUIDAD EN EL TIEMPO PASARON EXITOSAMENTE.")
