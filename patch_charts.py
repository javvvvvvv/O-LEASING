import re

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Match the chart creation in exportar_saldos_completo
saldos_target = r"chart = AreaChart\(\)\s*chart\.title = \"Evolución de Saldos de la Cartera\".*?chart\.height = 10"

saldos_replacement = """chart = AreaChart()
    chart.title = "Análisis de Abatimiento de Deuda (Saldos)"
    chart.style = 42
    chart.grouping = "stacked"
    chart.x_axis.title = "Mes"
    chart.y_axis.title = "Saldo Pendiente ($)"
    
    data_area = Reference(ws, min_col=2, min_row=3, max_col=3, max_row=15)
    cats = Reference(ws, min_col=1, min_row=4, max_row=15)
    chart.add_data(data_area, titles_from_data=True)
    chart.set_categories(cats)
    
    from openpyxl.chart import LineChart
    line = LineChart()
    data_line = Reference(ws, min_col=4, min_row=3, max_col=4, max_row=15)
    line.add_data(data_line, titles_from_data=True)
    
    chart += line
    
    ws.add_chart(chart, "F4")
    chart.width = 24
    chart.height = 14"""

content = re.sub(saldos_target, saldos_replacement, content, flags=re.DOTALL)

# Match the chart creation in exportar_maestro_completo
maestro_target = r"chart1 = BarChart\(\)\s*chart1\.type = \"col\".*?chart1\.height = 10"

maestro_replacement = """chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 2
    chart1.grouping = "stacked"
    chart1.overlap = 100
    chart1.title = "Proyección de Ingresos: Capital vs Intereses"
    chart1.y_axis.title = "Flujo de Efectivo ($)"
    chart1.x_axis.title = "Mes"
    
    chart2 = LineChart()
    
    data1 = Reference(ws, min_col=2, min_row=3, max_col=3, max_row=15)
    cats = Reference(ws, min_col=1, min_row=4, max_row=15)
    chart1.add_data(data1, titles_from_data=True)
    chart1.set_categories(cats)
    
    data2 = Reference(ws, min_col=4, min_row=3, max_col=4, max_row=15)
    chart2.add_data(data2, titles_from_data=True)
    
    chart1 += chart2
    
    ws.add_chart(chart1, "F4")
    chart1.width = 24
    chart1.height = 14"""

content = re.sub(maestro_target, maestro_replacement, content, flags=re.DOTALL)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)
print("Gráficas actualizadas!")
