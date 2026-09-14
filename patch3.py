with open('ui/theme.py', 'r', encoding='utf-8') as f:
    content = f.read()

old_css = """/* Tabular Nums para datos financieros */
[data-testid="stMetricValue"], [data-testid="stMetricDelta"], .stDataFrame, table, .tabular-nums { 
    font-variant-numeric: tabular-nums; 
}"""

new_css = """/* Tabular Nums para datos financieros y Ajuste de Tamaños */
[data-testid="stMetricValue"], [data-testid="stMetricDelta"], .stDataFrame, table, .tabular-nums { 
    font-variant-numeric: tabular-nums; 
}
[data-testid="stMetricValue"] {
    font-size: 1.4rem !important; /* Reducir el tamaño para que no se corten */
    word-wrap: break-word !important;
}
[data-testid="stMetricLabel"] {
    font-size: 0.85rem !important;
    white-space: normal !important; /* Permitir que los títulos largos usen 2 líneas */
    overflow: visible !important;
}"""

content = content.replace(old_css, new_css)
with open('ui/theme.py', 'w', encoding='utf-8') as f:
    f.write(content)
