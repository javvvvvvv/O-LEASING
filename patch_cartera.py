import re

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Replace plantilla()
old_plantilla = r"def plantilla\(\):.*?incluir_titulo=False,\n    \)"
new_plantilla = """def plantilla():
    df_plantilla = pd.DataFrame({
        'Contrato': ['0554'],
        'Anexo': ['0002'],
        'Cliente': ['Ejemplo S.A. de C.V.'],
        'Fecha De Apertura': [datetime.today().strftime('%Y-%m-%d')],
        'Plazo': [48],
        'Marca': ['RAM'],
        'Versión': ['LIMITED'],
        'Modelo': ['1500'],
        'Valor Cotización': [1600800],
        'Renta': [44483],
        '% Anticipo': [20.0],
        '% Comisión': [2.0],
        'Valor Residual (%)': [10.0],
        'Agencia': ['MOTORMEXA'],
        'Status': ['ACTIVO']
    })
    return excel_con_formato(
        {'Carga': df_plantilla},
        currency_cols=['Valor Cotización','Renta'],
        pct_cols=['% Anticipo','% Comisión','Valor Residual (%)'],
        incluir_titulo=False,
    )"""

content = re.sub(old_plantilla, new_plantilla, content, flags=re.DOTALL)

# Find the dfu processing block safely
idx_start = content.find("dfu=pd.read_csv(arch) if arch.name.endswith('.csv') else pd.read_excel(arch)")
if idx_start != -1:
    idx_end = content.find("bar.progress(1.0)", idx_start)
    if idx_end != -1:
        old_parser_block = content[idx_start:idx_end + len("bar.progress(1.0)")]
        new_parser_block = """# Leer el archivo crudo
                    if arch.name.endswith('.xlsx'):
                        if not hojas_sel: hojas_sel = [xls.sheet_names[0]]
                        dfs = []
                        for h in hojas_sel:
                            # Buscar dinamicamente la fila de headers (hasta fila 10)
                            df_raw = pd.read_excel(xls, sheet_name=h, header=None)
                            header_idx = 0
                            for r_idx in range(min(10, len(df_raw))):
                                row_vals = [str(x).upper() for x in df_raw.iloc[r_idx].values]
                                if 'CONTRATO' in row_vals or 'CLIENTE' in row_vals:
                                    header_idx = r_idx
                                    break
                            df_temp = pd.read_excel(xls, sheet_name=h, header=header_idx)
                            # Limpiar filas vacias
                            df_temp = df_temp.dropna(subset=[c for c in df_temp.columns if 'CONTRATO' in str(c).upper()])
                            dfs.append(df_temp)
                        dfu = pd.concat(dfs, ignore_index=True)
                    else:
                        dfu = pd.read_csv(arch)
                        dfu = dfu.dropna(subset=[c for c in dfu.columns if 'CONTRATO' in str(c).upper()])
                        
                    # Forzar nombres a mayúsculas, sin tildes y sin espacios extra para mapeo flexible
                    import unicodedata
                    def clean_col(c):
                        c = str(c).strip().upper()
                        c = unicodedata.normalize('NFKD', c).encode('ASCII', 'ignore').decode('utf-8')
                        return c
                    dfu.columns = [clean_col(c) for c in dfu.columns]
                    
                    nuevos=act=0; bar=st.progress(0); tot_f=len(dfu)
                    for i,r in dfu.iterrows():
                        try:
                            r = r.fillna('')
                            
                            # Contrato y Anexo
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
                            
                            # Vehiculo
                            if 'UNIDAD' in r and str(r['UNIDAD']).strip() != '':
                                vehiculo = str(r['UNIDAD']).strip()
                            else:
                                vehiculo = f"{str(r.get('MARCA','')).strip()} {str(r.get('VERSION','')).strip()} {str(r.get('MODELO','')).strip()}".strip()
                            if not vehiculo: vehiculo = "VEHICULO NO ESPECIFICADO"
                            
                            # Fecha Apertura
                            if 'FECHA DE APERTURA' in r and str(r['FECHA DE APERTURA']).strip() != '':
                                fa = pd.to_datetime(r['FECHA DE APERTURA'])
                            else:
                                fa = datetime.today()
                                
                            pl = int(float(r.get('PLAZO', 1)))
                            if pl < 1: pl = 1
                            
                            # Valores financieros (quitar IVA segun regla de negocio)
                            val_raw = r.get('VALOR COTIZACION', r.get('MOI', 0))
                            v_con_iva = float(str(val_raw).replace('$','').replace(',','').strip() or 0)
                            v_sin_iva = v_con_iva / 1.16
                            if v_sin_iva <= 0: raise ValueError("Valor Cotizacion invalido o 0.")
                            
                            renta_raw = r.get('RENTA', r.get('RENTAS', 0))
                            rn_con_iva = float(str(renta_raw).replace('$','').replace(',','').strip() or 0)
                            # Si la columna se llama RENTAS, asumimos que es el total
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
                            
                            # Residual
                            pr = safe_pct(['VALOR RESIDUAL (%)', 'RESIDUAL %', 'RESIDUAL', '%'])
                            if pr == 0.0 and 'RESIDUAL SIN IVA' in r and str(r['RESIDUAL SIN IVA']).strip() != '':
                                val_res = float(str(r['RESIDUAL SIN IVA']).replace('$','').replace(',','').strip())
                                pr = (val_res / v_sin_iva) * 100 if v_sin_iva > 0 else 0.0
                                
                            est = str(r.get('STATUS','ACTIVO')).strip().upper()
                            fb = pd.to_datetime(r['FECHA DE BAJA']) if 'FECHA DE BAJA' in r and str(r['FECHA DE BAJA']).strip() != '' else None
                            if est == 'BAJA' and fb is None: fb = datetime.today()
                            
                            existe = get_db().execute("SELECT 1 FROM contratos WHERE ID_Contrato=?",(id_c,)).fetchone()
                            
                            anotaciones = ""
                            for extra_col in ['SERIE', 'AGENCIA', 'EMISOR', 'ORIGEN']:
                                if extra_col in r and str(r[extra_col]).strip() != '':
                                    anotaciones += f"{extra_col.capitalize()}: {r[extra_col]}\\n"
                            
                            ct = dict(
                                ID_Contrato=id_c, Cliente=cliente, Vehiculo=vehiculo,
                                Fecha_Alta=fa, Fecha_Vencimiento=fa+relativedelta(months=pl),
                                Valor_Sin_IVA=v_sin_iva, Mensualidad_Sin_IVA=rn_sin_iva, Plazo=pl,
                                Comision_Apertura_Pct=pc, Comision_Monto=v_sin_iva*pc/100,
                                Anticipo_Pct=pa, Anticipo_Monto=v_sin_iva*pa/100,
                                Residual_Pct=pr, Residual_Monto=v_sin_iva*pr/100,
                                Tasa_Calculada=0.0, Estatus=est, Fecha_Baja=fb,
                                residual_transferred=0, Nivel_Morosidad=0,
                                Anotaciones=anotaciones.strip()
                            )
                            guardar(ct)
                            if existe: act+=1
                            else: nuevos+=1
                        except Exception as e: st.error(f"Fila {i+1} (Contrato {r.get('CONTRATO','')}): {e}")
                    bar.progress(1.0)"""
        content = content.replace(old_parser_block, new_parser_block)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)
print("Adaptado a Cartera Ago-26.xlsx exitosamente.")
