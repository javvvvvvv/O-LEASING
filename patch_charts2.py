import re

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

saldos_target = r"chart = AreaChart\(\).*?chart\.height = 14"
saldos_replacement = """from openpyxl.chart import BarChart, Reference
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Saldos Pendientes por Mes"
    chart.y_axis.title = "Monto ($)"
    chart.x_axis.title = "Mes"
    chart.grouping = "clustered"
    chart.gapWidth = 150
    chart.overlap = 0
    
    data = Reference(ws, min_col=2, min_row=3, max_col=4, max_row=15)
    cats = Reference(ws, min_col=1, min_row=4, max_row=15)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    
    # Opcional: Para mostrar las etiquetas de datos encima de las barras de Total (serie 3)
    # En openpyxl es complejo, as que dejaremos la grfica super limpia.
    
    ws.add_chart(chart, "F4")
    chart.width = 25
    chart.height = 14"""

maestro_target = r"chart1 = BarChart\(\)\s*chart1\.type = \"col\".*?chart1\.height = 14"
maestro_replacement = """from openpyxl.chart import BarChart, Reference
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Proyección de Flujo por Mes"
    chart.y_axis.title = "Ingresos ($)"
    chart.x_axis.title = "Mes"
    chart.grouping = "clustered"
    chart.gapWidth = 150
    chart.overlap = 0
    
    data = Reference(ws, min_col=2, min_row=3, max_col=4, max_row=15)
    cats = Reference(ws, min_col=1, min_row=4, max_row=15)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    
    ws.add_chart(chart, "F4")
    chart.width = 25
    chart.height = 14"""

content = re.sub(saldos_target, saldos_replacement, content, flags=re.DOTALL)
content = re.sub(maestro_target, maestro_replacement, content, flags=re.DOTALL)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)
print("Gráficas actualizadas a columnas agrupadas.")
