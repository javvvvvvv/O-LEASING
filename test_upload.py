import pandas as pd
from dateutil.relativedelta import relativedelta
from datetime import datetime
import unicodedata
import warnings
warnings.filterwarnings('ignore')

arch = 'C:\\Cartera Ago-26.xlsx'
xls = pd.ExcelFile(arch)
hojas_sel = ['Cartera']

dfs = []
for h in hojas_sel:
    df_raw = pd.read_excel(xls, sheet_name=h, header=None)
    header_idx = 0
    for r_idx in range(min(10, len(df_raw))):
        row_vals = [str(x).upper() for x in df_raw.iloc[r_idx].values]
        if 'CONTRATO' in row_vals or 'CLIENTE' in row_vals:
            header_idx = r_idx
            break
    print(f'Header found at row {header_idx}')
    df_temp = pd.read_excel(xls, sheet_name=h, header=header_idx)
    def clean_col(c):
        c = str(c).strip().upper()
        c = unicodedata.normalize('NFKD', c).encode('ASCII', 'ignore').decode('utf-8')
        return c
    df_temp.columns = [clean_col(c) for c in df_temp.columns]
    
    col_contrato = next((c for c in df_temp.columns if str(c).strip() == 'CONTRATO'), None)
    if col_contrato:
        df_temp = df_temp.dropna(subset=[col_contrato])
    dfs.append(df_temp)

dfu = pd.concat(dfs, ignore_index=True)
print(f'Total rows to process: {len(dfu)}')

nuevos=0; act=0
for i, r in dfu.iterrows():
    try:
        r = r.fillna('')
        c_val = str(r.get('CONTRATO','')).replace('.0','').strip()
        if not c_val: continue
        a_val = str(r.get('ANEXO','')).replace('.0','').strip()
        
        if '-' in c_val:
            c_parts = c_val.split('-')
            id_c = f"{c_parts[0].zfill(4)}-{c_parts[1].zfill(4)}"
        else:
            if a_val:
                id_c = f"{c_val.zfill(4)}-{a_val.zfill(4)}"
            else:
                id_c = c_val.zfill(4) + "-0001"
        
        cliente = str(r.get('CLIENTE','')).strip()
        if 'UNIDAD' in r and str(r['UNIDAD']).strip() != '':
            vehiculo = str(r['UNIDAD']).strip()
        else:
            vehiculo = f"{str(r.get('MARCA','')).strip()} {str(r.get('VERSION','')).strip()} {str(r.get('MODELO','')).strip()}".strip()
        if not vehiculo: vehiculo = 'VEHICULO NO ESPECIFICADO'
        
        if 'FECHA DE APERTURA' in r and str(r['FECHA DE APERTURA']).strip() != '':
            fa = pd.to_datetime(r['FECHA DE APERTURA'])
        else:
            fa = datetime.today()
            
        pl = int(float(r.get('PLAZO', 1)))
        if pl < 1: pl = 1
        
        val_raw = r.get('VALOR COTIZACION', r.get('MOI', 0))
        v_con_iva = float(str(val_raw).replace('$','').replace(',','').strip() or 0)
        v_sin_iva = v_con_iva / 1.16
        if v_sin_iva <= 0: raise ValueError(f"Valor Cotizacion invalido o 0: {val_raw}")
        
        renta_raw = r.get('RENTA', r.get('RENTAS', 0))
        rn_con_iva = float(str(renta_raw).replace('$','').replace(',','').strip() or 0)
        if 'RENTAS' in r and 'RENTA' not in r:
            rn_con_iva = rn_con_iva / pl if pl > 0 else rn_con_iva
        rn_sin_iva = rn_con_iva / 1.16
        
        def safe_pct(cols_posibles):
            for col in cols_posibles:
                if col in r and str(r[col]).strip() != '':
                    val = str(r[col]).replace('%','').strip()
                    try:
                        v_f = float(val)
                        return v_f*100 if v_f<=1 and v_f>0 else v_f
                    except: pass
            return 0.0
        
        pc = safe_pct(['% COMISION', 'COMISION %', 'COMISION'])
        pa = safe_pct(['% ANTICIPO', 'ANTICIPO %', 'ANTICIPO'])
        
        pr = safe_pct(['VALOR RESIDUAL (%)', 'RESIDUAL %', 'RESIDUAL', '%'])
        if pr == 0.0 and 'RESIDUAL SIN IVA' in r and str(r['RESIDUAL SIN IVA']).strip() != '':
            val_res = float(str(r['RESIDUAL SIN IVA']).replace('$','').replace(',','').strip())
            pr = (val_res / v_sin_iva) * 100 if v_sin_iva > 0 else 0.0
            
        est = str(r.get('STATUS','ACTIVO')).strip().upper()
        
    except Exception as e:
        print(f"Fila {i} Error: {e}")
print('Done!')
