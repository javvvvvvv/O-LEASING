import re

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

func_export = """
def exportar_maestro_completo(anio, di, dcap, drenta, dr, dc, ds, total_cap, total_int, total_renta):
    import io
    import pandas as pd
    from reports.excel import excel_con_formato
    
    # Crear un dataframe para el resumen ejecutivo
    MN = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
    df_resumen = pd.DataFrame({
        'Mes': MN,
        'Capital Amortizado': total_cap.values,
        'Intereses Devengados': total_int.values,
        'Renta Neta (Flujo Total)': total_renta.values
    })
    
    # Preparamos las hojas completas
    # Formateamos un poco los de data quitando index si lo tienen
    skip_currency = ['ID_Contrato', 'Cliente', 'Vehiculo', 'Estatus', 'Fecha_Alta', 'Plazo', 'Fecha_Baja', 'Tasa_Anual_%']
    
    def pre(df):
        return df.reset_index().rename(columns={'index': 'ID_Contrato'})
        
    d1 = pre(dcap)
    d2 = pre(di)
    d3 = pre(drenta)
    d4 = pre(dr)
    d5 = pre(dc)
    d6 = pre(ds)
    
    curr_cols_1 = [c for c in d1.columns if c not in skip_currency]
    
    hojas = {
        'Resumen Ejecutivo': df_resumen,
        'Capital Leasing': d1,
        'Intereses Leasing': d2,
        'Renta Neta': d3,
        'Int. Residual': d4,
        'Amort. Comision': d5,
        'Saldo Residual': d6
    }
    
    # Diccionarios de formato
    c_cols = {
        'Resumen Ejecutivo': ['Capital Amortizado', 'Intereses Devengados', 'Renta Neta (Flujo Total)'],
        'Capital Leasing': curr_cols_1,
        'Intereses Leasing': curr_cols_1,
        'Renta Neta': curr_cols_1,
        'Int. Residual': curr_cols_1,
        'Amort. Comision': curr_cols_1,
        'Saldo Residual': curr_cols_1
    }
    
    p_cols = {
        'Capital Leasing': ['Tasa_Anual_%'],
        'Intereses Leasing': ['Tasa_Anual_%'],
        'Renta Neta': ['Tasa_Anual_%'],
        'Int. Residual': ['Tasa_Anual_%'],
        'Amort. Comision': ['Tasa_Anual_%'],
        'Saldo Residual': ['Tasa_Anual_%']
    }
    
    # Llamar al generador base (que ya le pone logo y colores corporativos a la tabla)
    buf = excel_con_formato(hojas, currency_cols=c_cols, pct_cols=p_cols)
    
    # AHORA VAMOS A ABRIR ESE BUFFER CON OPENPYXL PARA METER LA GRÁFICA
    import openpyxl
    from openpyxl.chart import BarChart, LineChart, Reference, Series
    from openpyxl.chart.axis import DateAxis
    
    buf.seek(0)
    wb = openpyxl.load_workbook(buf)
    
    # Obtener hoja de resumen
    ws = wb['Resumen Ejecutivo']
    
    # Crear gráfica apilada (Capital + Intereses)
    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 10
    chart1.grouping = "stacked"
    chart1.overlap = 100
    chart1.title = "Composición Mensual (Capital vs Intereses)"
    chart1.y_axis.title = "Monto ($)"
    chart1.x_axis.title = "Mes"
    
    # Crear gráfica de linea (Flujo Total)
    chart2 = LineChart()
    
    # Datos para chart1 (Columnas B y C -> Categ/X es A)
    # StartRow depende de si excel_con_formato metió el título en fila 1. Sí lo mete.
    # Los datos empiezan en A3 si hay titulo. Busquemos la celda A3 o A4 para ver dónde empieza la tabla.
    # En excel_con_formato usa startrow=2 si _titulo_emp, lo que significa fila 3 en excel (1-indexed).
    data1 = Reference(ws, min_col=2, min_row=3, max_col=3, max_row=15) # Fila 3 es cabecera, hasta la 15
    cats = Reference(ws, min_col=1, min_row=4, max_row=15)
    chart1.add_data(data1, titles_from_data=True)
    chart1.set_categories(cats)
    
    # Datos para chart2 (Columna D)
    data2 = Reference(ws, min_col=4, min_row=3, max_row=15)
    chart2.add_data(data2, titles_from_data=True)
    
    # Combinar graficas
    chart1 += chart2
    
    # Posicionar la gráfica
    ws.add_chart(chart1, "F4")
    chart1.width = 18
    chart1.height = 10
    
    # Guardar en un nuevo buffer
    out_buf = io.BytesIO()
    wb.save(out_buf)
    out_buf.seek(0)
    return out_buf
"""

# Insert function before `def reporte_maestro_mensual(anio):`
if 'exportar_maestro_completo(' not in content:
    content = content.replace('def reporte_maestro_mensual(anio):', func_export + '\ndef reporte_maestro_mensual(anio):')

# Now patch the UI to have the download button
ui_patch_target = r"st\.plotly_chart\(fig_exec, width='stretch', key=\"pc_maestro_exec\"\)"
ui_patch_replacement = """st.plotly_chart(fig_exec, width='stretch', key="pc_maestro_exec")
                
                # Generar Excel completo con Graficas
                buf_completo = exportar_maestro_completo(anio_m, di, dcap, drenta, dr, dc, ds, total_cap, total_int, total_renta)
                st.download_button("📥 Descargar Todo en un Solo Excel (con Gráficas)", buf_completo, f"reporte_maestro_completo_{anio_m}.xlsx", type="primary", use_container_width=True)"""

content = re.sub(ui_patch_target, ui_patch_replacement, content)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)
print("Funcionalidad de exportación a un solo Excel lista.")
