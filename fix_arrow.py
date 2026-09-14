import re

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Fix 1: line ~2707
content = content.replace(
    "'Mes':          f['mes_contrato'] if f['mes_contrato'] is not None else '-',",
    "'Mes':          str(int(f['mes_contrato'])) if f['mes_contrato'] is not None else '-',"
)

# Fix 2: line ~3290
# df_rep['Mes'] = df_rep['Mes'].fillna('-')
content = re.sub(
    r"df_rep\['Mes'\] = df_rep\['Mes'\].fillna\('-'\)",
    "df_rep['Mes'] = df_rep['Mes'].astype(str).replace(['nan', 'None', '<NA>', 'NaN'], '-').fillna('-')",
    content
)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)
print("ArrowInvalid arreglado")
