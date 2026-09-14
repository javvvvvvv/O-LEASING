import re

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Buscamos el inicio de "with tab_m:"
start_str = '        with tab_m:\n            st.subheader("Importar desde Excel")'
idx_start = content.find(start_str)

if idx_start != -1:
    # Buscar el inicio de tab_man (la siguiente tab)
    idx_end = content.find('        with tab_man:', idx_start)
    if idx_end != -1:
        old_block = content[idx_start:idx_end]
        
        new_block = """        with tab_m:
            st.subheader("Importar desde Excel")
            if st.button("Descargar Plantilla"):
                st.download_button("Guardar plantilla",plantilla(),"plantilla.xlsx")
            arch = st.file_uploader("Sube .xlsx o .csv",type=['xlsx','csv'])
            
            hojas_sel = []
            xls = None
            if arch and arch.name.endswith('.xlsx'):
                xls = pd.ExcelFile(arch)
                opciones = xls.sheet_names
                # Preseleccionar julio y agosto 2026 si existen, por requerimiento directo
                defs = [h for h in opciones if "JULIO 2026" in h.upper() or "AGOSTO 2026" in h.upper()]
                hojas_sel = st.multiselect("Selecciona las hojas a procesar (dejalo vacio para procesar la primera):", opciones, default=defs)
            
            if arch and st.button("Procesar y Cargar"):
                try:
                    if arch.name.endswith('.xlsx'):
                        if not hojas_sel: hojas_sel = [xls.sheet_names[0]]
                        dfs = []
                        for h in hojas_sel:
                            df_temp = pd.read_excel(xls, sheet_name=h, header=0)
                            if not any('CONTRATO' in str(c).upper() for c in df_temp.columns):
                                df_temp = pd.read_excel(xls, sheet_name=h, header=1) # Usar fila 2 como headers (formato legacy)
                            # Limpiar filas vacias (donde contrato sea NaN)
                            df_temp = df_temp.dropna(subset=[c for c in df_temp.columns if 'CONTRATO' in str(c).upper()])
                            dfs.append(df_temp)
                        dfu = pd.concat(dfs, ignore_index=True)
                    else:
                        dfu = pd.read_csv(arch)
                        dfu = dfu.dropna(subset=[c for c in dfu.columns if 'CONTRATO' in str(c).upper()])
                        
                    # Forzar nombres a mayúsculas y quitar espacios
                    dfu.columns = [str(c).strip().upper() for c in dfu.columns]
                    
                    nuevos=act=0; bar=st.progress(0); tot_f=len(dfu)
                    for i,r in dfu.iterrows():
                        try:
                            r = r.fillna('')
                            c_raw = str(r.get('CONTRATO','')).replace('.0','').strip()
                            if not c_raw: continue
                            c_parts = c_raw.split('-')
                            if len(c_parts) >= 2:
                                id_c = f"{c_parts[0].zfill(4)}-{c_parts[1].zfill(4)}"
                            else:
                                id_c = c_raw.zfill(4) + "-0001"
                                
                            cliente = str(r.get('CLIENTE','')).strip()
                            
                            # Adaptacion Hibrida (Plantilla nueva vs Archivo legacy)
                            if 'UNIDAD' in r:
                                vehiculo = str(r['UNIDAD']).strip()
                            else:
                                vehiculo = f"{str(r.get('MARCA','')).strip()} {str(r.get('VERSION','')).strip()} {str(r.get('MODELO','')).strip()}"
                            
                            if 'FECHA DE APERTURA' in r and str(r['FECHA DE APERTURA']).strip() != '':
                                fa = pd.to_datetime(r['FECHA DE APERTURA'])
                            else:
                                fa = datetime.today()
                                
                            pl = int(float(r.get('PLAZO', 1)))
                            if pl < 1: pl = 1
                            
                            # Valor Cotizacion (o MOI)
                            val_raw = r.get('VALOR COTIZACION', r.get('MOI', 0))
                            v_con_iva = float(str(val_raw).replace('$','').replace(',','').strip() or 0)
                            v_sin_iva = v_con_iva / 1.16
                            if v_sin_iva <= 0: raise ValueError("Valor Cotizacion/MOI invalido o 0.")
                            
                            # Renta Mensual (o RENTAS totales)
                            renta_raw = r.get('RENTA', r.get('RENTAS', 0))
                            rn_con_iva = float(str(renta_raw).replace('$','').replace(',','').strip() or 0)
                            # Si es el archivo legacy, RENTAS es el total, debemos dividir entre plazo
                            if 'RENTAS' in r and 'RENTA' not in r:
                                rn_con_iva = rn_con_iva / pl
                            rn_sin_iva = rn_con_iva / 1.16
                            
                            def safe_pct(col):
                                if col not in r or str(r[col]).strip() == '': return 0.0
                                val = str(r[col]).replace('%','').strip()
                                try:
                                    v_f = float(val)
                                except: return 0.0
                                return v_f*100 if v_f<=1 and v_f>0 else v_f
                            
                            pc = safe_pct('COMISION %')
                            pa = safe_pct('ANTICIPO %')
                            # Manejar Residual: Si existe RESIDUAL %, usarlo. Si no, calcular desde RESIDUAL SIN IVA
                            if 'RESIDUAL %' in r:
                                pr = safe_pct('RESIDUAL %')
                            elif 'RESIDUAL SIN IVA' in r and str(r['RESIDUAL SIN IVA']).strip() != '':
                                val_res = float(str(r['RESIDUAL SIN IVA']).replace('$','').replace(',','').strip())
                                pr = (val_res / v_sin_iva) * 100 if v_sin_iva > 0 else 0.0
                            else:
                                pr = safe_pct('%') # Fallback al % que viene en Julio/Agosto
                            
                            est = str(r.get('STATUS','ACTIVO')).strip().upper()
                            fb = pd.to_datetime(r['FECHA DE BAJA']) if 'FECHA DE BAJA' in r and str(r['FECHA DE BAJA']).strip() != '' else None
                            if est == 'BAJA' and fb is None: fb = datetime.today()
                            
                            existe = get_db().execute("SELECT 1 FROM contratos WHERE ID_Contrato=?",(id_c,)).fetchone()
                            
                            anotaciones = ""
                            if 'SERIE' in r and str(r['SERIE']).strip() != '': anotaciones += f"Serie: {r['SERIE']}\\n"
                            if 'EMISOR' in r and str(r['EMISOR']).strip() != '': anotaciones += f"Emisor: {r['EMISOR']}\\n"
                            
                            ct = dict(
                                ID_Contrato=id_c, Cliente=cliente, Vehiculo=vehiculo.strip(),
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
                    bar.progress(1.0)
                    st.success(f"Procesamiento listo: {nuevos} contratos nuevos, {act} actualizados.")
                except Exception as e:
                    st.error(f"Error procesando el archivo: {e}")
                    
"""
        
        content = content.replace(old_block, new_block)
        with open('app.py', 'w', encoding='utf-8') as f:
            f.write(content)
        print("Carga masiva adaptada exitosamente con soporte multi-hoja y legacy.")
    else:
        print("No se encontro el fin del bloque tab_m")
else:
    print("No se encontro el inicio del bloque tab_m")
