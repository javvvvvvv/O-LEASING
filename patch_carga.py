import re

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Replace plantilla()
old_plantilla = r"def plantilla\(\):.*?incluir_titulo=False,\n    \)"
new_plantilla = """def plantilla():
    df_plantilla = pd.DataFrame({
        'CONTRATO': ['551-01'],
        'CLIENTE': ['Ejemplo S.A. de C.V.'],
        'MARCA': ['RAM'],
        'VERSION': ['LIMITED'],
        'MODELO': ['1500'],
        'SERIE': ['1C6SRFHT2NN305890'],
        'PLAZO': [36],
        'VALOR COTIZACION': [1600800],
        'RENTA': [44483],
        'RESIDUAL %': [10.0],
        'ANTICIPO %': [20.0],
        'COMISION %': [2.0],
        'FECHA DE APERTURA': [datetime.today().strftime('%Y-%m-%d')],
        'STATUS': ['ACTIVO'],
        'EMISOR': ['MOTORMEXA']
    })
    return excel_con_formato(
        {'Carga': df_plantilla},
        currency_cols=['VALOR COTIZACION','RENTA'],
        pct_cols=['RESIDUAL %','ANTICIPO %','COMISION %'],
        incluir_titulo=False,
    )"""
content = re.sub(old_plantilla, new_plantilla, content, flags=re.DOTALL)

# Find the dfu processing block safely
idx_start = content.find("dfu=pd.read_csv(arch) if arch.name.endswith('.csv') else pd.read_excel(arch)")
if idx_start != -1:
    idx_end = content.find("bar.progress((i+1)/tot_f)", idx_start)
    if idx_end != -1:
        old_parser_block = content[idx_start:idx_end + len("bar.progress((i+1)/tot_f)")]
        new_parser_block = """dfu=pd.read_csv(arch) if arch.name.endswith('.csv') else pd.read_excel(arch)
                    # Forzar nombres a mayúsculas y quitar espacios extra
                    dfu.columns = [str(c).strip().upper() for c in dfu.columns]
                    
                    req=['CONTRATO', 'CLIENTE', 'MARCA', 'VERSION', 'MODELO', 'PLAZO', 'VALOR COTIZACION', 'RENTA']
                    miss=[c for c in req if c not in dfu.columns]
                    if miss: st.error(f"Faltan columnas obligatorias: {miss}")
                    else:
                        nuevos=act=0; bar=st.progress(0); tot_f=len(dfu)
                        for i,r in dfu.iterrows():
                            try:
                                r = r.fillna('')
                                c_raw = str(r['CONTRATO']).replace('.0','').strip()
                                c_parts = c_raw.split('-')
                                if len(c_parts) >= 2:
                                    id_c = f"{c_parts[0].zfill(4)}-{c_parts[1].zfill(4)}"
                                else:
                                    id_c = c_raw.zfill(4) + "-0001"
                                    
                                vehiculo = f"{str(r['MARCA']).strip()} {str(r['VERSION']).strip()} {str(r['MODELO']).strip()}"
                                
                                if 'FECHA DE APERTURA' in r and str(r['FECHA DE APERTURA']).strip() != '':
                                    fa = pd.to_datetime(r['FECHA DE APERTURA'])
                                else:
                                    fa = datetime.today()
                                    
                                pl = int(float(r['PLAZO']))
                                if pl < 1: raise ValueError(f"Plazo debe ser de al menos 1 mes.")
                                    
                                v_con_iva = float(str(r['VALOR COTIZACION']).replace('$','').replace(',','').strip())
                                v_sin_iva = v_con_iva / 1.16
                                if v_sin_iva <= 0: raise ValueError("Valor Cotizacion invalido.")
                                
                                rn_con_iva = float(str(r['RENTA']).replace('$','').replace(',','').strip())
                                rn_sin_iva = rn_con_iva / 1.16
                                
                                def safe_pct(col):
                                    if col not in r or str(r[col]).strip() == '': return 0.0
                                    val = str(r[col]).replace('%','').strip()
                                    try:
                                        v_f = float(val)
                                    except:
                                        return 0.0
                                    return v_f*100 if v_f<=1 and v_f>0 else v_f
                                
                                pc = safe_pct('COMISION %')
                                pa = safe_pct('ANTICIPO %')
                                pr = safe_pct('RESIDUAL %')
                                
                                est = str(r.get('STATUS','ACTIVO')).strip().upper()
                                fb = pd.to_datetime(r['FECHA DE BAJA']) if 'FECHA DE BAJA' in r and str(r['FECHA DE BAJA']).strip() != '' else None
                                if est == 'BAJA' and fb is None: fb = datetime.today()
                                
                                existe = get_db().execute("SELECT 1 FROM contratos WHERE ID_Contrato=?",(id_c,)).fetchone()
                                
                                anotaciones = ""
                                if 'SERIE' in r and str(r['SERIE']).strip() != '': anotaciones += f"Serie: {r['SERIE']}\\n"
                                if 'EMISOR' in r and str(r['EMISOR']).strip() != '': anotaciones += f"Emisor: {r['EMISOR']}\\n"
                                
                                ct = dict(
                                    ID_Contrato=id_c, Cliente=str(r['CLIENTE']).strip(), Vehiculo=vehiculo.strip(),
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
                        bar.progress((i+1)/tot_f)"""
        content = content.replace(old_parser_block, new_parser_block)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)
print("Carga masiva adaptada al nuevo formato Excel.")
