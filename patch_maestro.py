import re

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

func_pattern = re.compile(r'def reporte_maestro_mensual\(anio\):.*?return di, dcap, drenta, dr, dc, ds, None', re.DOTALL)

new_func = """def reporte_maestro_mensual(anio):
    df=obtener()
    if df.empty: return None,None,None,None,None,None,"Sin contratos registrados"
    MN=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
    ids=df['ID_Contrato'].tolist()
    di=pd.DataFrame(index=ids,columns=range(1,13),dtype=float)
    dr=di.copy(); dc=di.copy(); ds=di.copy(); dcap=di.copy(); drenta=di.copy()
    bar=st.progress(0,"Calculando Reporte Maestro...")
    tot=len(df)
    for i,(_,row) in enumerate(df.iterrows()):
        id_c=row['ID_Contrato']; fa=row['Fecha_Alta']; pl=int(row['Plazo'])
        t=round(row['Tasa_Calculada'],8); inv=row['Valor_Sin_IVA']-row['Anticipo_Monto']
        try:
            dfa,_,_,_,_=calc_amort(round(inv,4),round(row['Mensualidad_Sin_IVA'],4),round(row['Residual_Monto'],4),pl,t)
            dfr=calc_res_amort(round(float(row['VP_Residual']),4),t,pl)
        except: continue
        es_baja = str(row.get('Estatus','')).upper() == 'BAJA'
        fecha_baja = pd.to_datetime(row['Fecha_Baja']) if es_baja and pd.notna(row.get('Fecha_Baja')) else None
        for mes in range(1,13):
            fm=pd.Timestamp(anio,mes,1)
            if fm<pd.Timestamp(fa.year,fa.month,1) or fm>row['Fecha_Vencimiento']: continue
            if fecha_baja is not None and fm>pd.Timestamp(fecha_baja.year,fecha_baja.month,1): continue
            mc=(anio-fa.year)*12+(mes-fa.month)+1
            if mc<1 or mc>pl: continue
            
            interes_leasing = round(dfa.iloc[mc-1]['Interes'],2)
            capital_leasing = round(dfa.iloc[mc-1]['Capital'],2)
            
            di.at[id_c,mes] = interes_leasing
            dcap.at[id_c,mes] = capital_leasing
            drenta.at[id_c,mes] = interes_leasing + capital_leasing
            dr.at[id_c,mes] = round(dfr.iloc[mc-1]['Interes'],2)
            dc.at[id_c,mes] = round(row['Comision_Monto']/pl,2)
            ds.at[id_c,mes] = round(dfr.iloc[mc-1]['Saldo_Fin'],2)
        bar.progress((i+1)/tot,text=f"Procesando {i+1}/{tot}...")
    bar.empty()
    
    # Agregar metadatos descriptivos a cada dataframe
    df_meta = df.set_index('ID_Contrato')[['Cliente', 'Vehiculo', 'Estatus', 'Fecha_Alta', 'Plazo', 'Fecha_Baja', 'Valor_Sin_IVA', 'Tasa_Calculada']].copy()
    # Formatear algunos metadatos para que se vean bien en excel
    df_meta['Fecha_Alta'] = df_meta['Fecha_Alta'].dt.strftime('%Y-%m-%d')
    df_meta['Fecha_Baja'] = df_meta['Fecha_Baja'].dt.strftime('%Y-%m-%d').fillna('')
    df_meta['Tasa_Anual_%'] = (df_meta['Tasa_Calculada'] * 1200).round(2)
    df_meta.drop(columns=['Tasa_Calculada'], inplace=True)
    df_meta['Valor_Sin_IVA'] = df_meta['Valor_Sin_IVA'].round(2)
    
    out = []
    for d in [di,dcap,drenta,dr,dc,ds]: 
        d.columns=MN
        d.dropna(how='all',inplace=True)
        # Unir metadatos y reordenar para que queden al principio
        d_merged = df_meta.join(d, how='right')
        # Limpiar NaNs de los meses por ceros (opcional, o dejar nulo)
        out.append(d_merged)
        
    return out[0], out[1], out[2], out[3], out[4], out[5], None"""

content = func_pattern.sub(new_func, content)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)
print("Reporte Maestro mejorado con metadatos")
