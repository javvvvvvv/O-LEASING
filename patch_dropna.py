import re

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

old_drop1 = "df_temp = df_temp.dropna(subset=[c for c in df_temp.columns if 'CONTRATO' in str(c).upper()])"
new_drop1 = """col_contrato = next((c for c in df_temp.columns if str(c).upper().strip() == 'CONTRATO'), None)
                            if col_contrato:
                                df_temp = df_temp.dropna(subset=[col_contrato])"""

old_drop2 = "dfu = dfu.dropna(subset=[c for c in dfu.columns if 'CONTRATO' in str(c).upper()])"
new_drop2 = """col_c_dfu = next((c for c in dfu.columns if str(c).upper().strip() == 'CONTRATO'), None)
                        if col_c_dfu:
                            dfu = dfu.dropna(subset=[col_c_dfu])"""

content = content.replace(old_drop1, new_drop1)
content = content.replace(old_drop2, new_drop2)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)
