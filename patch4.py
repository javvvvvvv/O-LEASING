import re

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Restore original tabla_mensual_conceptos
original_func = """def tabla_mensual_conceptos(anio):
    df=obtener()
    if df.empty: return None,None,None,None,"Sin contratos registrados"
    MN=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
    ids=df['ID_Contrato'].tolist()
    di=pd.DataFrame(index=ids,columns=range(1,13),dtype=float); dr=di.copy(); dc=di.copy(); ds=di.copy()
    bar=st.progress(0,"Calculando...")
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
            di.at[id_c,mes]=round(dfa.iloc[mc-1]['Interes'],2)
            dr.at[id_c,mes]=round(dfr.iloc[mc-1]['Interes'],2)
            dc.at[id_c,mes]=round(row['Comision_Monto']/pl,2)
            ds.at[id_c,mes]=round(dfr.iloc[mc-1]['Saldo_Fin'],2)
        bar.progress((i+1)/tot,text=f"Procesando {i+1}/{tot}...")
    bar.empty()
    for d in [di,dr,dc,ds]: 
        d.columns=MN
        d.dropna(how='all',inplace=True)
    return di,dr,dc,ds,None"""

new_master_func = """
def reporte_maestro_mensual(anio):
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
    for d in [di,dcap,drenta,dr,dc,ds]: 
        d.columns=MN
        d.dropna(how='all',inplace=True)
    return di, dcap, drenta, dr, dc, ds, None
"""

content = re.sub(r'def tabla_mensual_conceptos\(anio\):.*?return di, dcap, drenta, dr, dc, ds, None', original_func + new_master_func, content, flags=re.DOTALL)

# Add menu item "Reporte Maestro" in "Finanzas & Contabilidad"
if '"Tabla Mensual por Contrato"' in content and '"Reporte Maestro"' not in content:
    content = content.replace('"Tabla Mensual por Contrato",\n    ],', '"Tabla Mensual por Contrato",\n        "Reporte Maestro",\n    ],')
    content = content.replace('"Tabla Mensual por Contrato": "Consulta el detalle mes a mes de un contrato en particular.",', '"Tabla Mensual por Contrato": "Consulta las tablas base contables (Interés, Residual, Comisión).",\n    "Reporte Maestro": "Reporte gerencial completo con Capital y Renta neta por contrato.",')

# UI Block
old_ui_block = r'elif menu=="Tabla Mensual por Contrato":.*?st\.download_button\(f"\{titulo\[:25\]\}",buf,fname,key=f"dl_\{fname\}"\)'

original_ui_plus_new = """elif menu=="Tabla Mensual por Contrato":
        st.title("Tabla Mensual de Conceptos Financieros")
        st.markdown("Genera 4 tablas: **Intereses leasing** • **Intereses residual** • **Amort. comisión** • **Saldo residual activo**")
        anio_t=st.number_input("Año",min_value=2020,max_value=2050,value=datetime.now().year,step=1,key="yt")
        if st.button("Generar Tablas"):
            di,dr,dc,ds,err=tabla_mensual_conceptos(anio_t)
            if err: st.error(err)
            else:
                for titulo,dft,fname in [
                    ("1. Intereses Leasing — cta 208",di,f"int_208_{anio_t}.xlsx"),
                    ("2. Intereses Residual — cta 126/114",dr,f"int_residual_{anio_t}.xlsx"),
                    ("3. Amortización Comisión por Apertura",dc,f"amort_comision_{anio_t}.xlsx"),
                    ("4. Saldo del Valor Residual Activo",ds,f"saldo_residual_{anio_t}.xlsx"),
                ]:
                    st.subheader(titulo)
                    if dft is None or dft.empty: st.info("Sin datos.")
                    else:
                        st.dataframe(dft.style.format("{:,.2f}").highlight_null("lightgray"),width='stretch', key=f"df_026_{fname}")
                        _dft_export = dft.reset_index().rename(columns={'index': 'ID_Contrato'})
                        buf = excel_con_formato({titulo[:31]: _dft_export},
                            currency_cols=[c for c in _dft_export.columns if c != 'ID_Contrato'])
                        st.download_button(f"{titulo[:25]}",buf,fname,key=f"dl_{fname}")

    elif menu=="Reporte Maestro":
        st.title("Reporte Maestro: Capital e Intereses")
        st.markdown("Genera las tablas completas de amortización de toda la cartera activa: **Capital**, **Intereses**, **Renta Neta** y anexos (Residual/Comisiones).")
        anio_m=st.number_input("Año del reporte maestro",min_value=2020,max_value=2050,value=datetime.now().year,step=1,key="ym")
        if st.button("Generar Reporte Maestro", kind="primary"):
            di, dcap, drenta, dr, dc, ds, err = reporte_maestro_mensual(anio_m)
            if err: st.error(err)
            else:
                for titulo,dft,fname in [
                    ("1. Capital Leasing (Amortización Principal)", dcap, f"capital_leasing_{anio_m}.xlsx"),
                    ("2. Intereses Leasing (Devengados)", di, f"intereses_leasing_{anio_m}.xlsx"),
                    ("3. Renta Neta (Capital + Interés)", drenta, f"renta_neta_{anio_m}.xlsx"),
                    ("4. Intereses Residual (Acumulación)", dr, f"int_residual_{anio_m}.xlsx"),
                    ("5. Amortización Comisión por Apertura", dc, f"amort_comision_{anio_m}.xlsx"),
                    ("6. Saldo del Valor Residual Activo", ds, f"saldo_residual_{anio_m}.xlsx"),
                ]:
                    st.subheader(titulo)
                    if dft is None or dft.empty: st.info("Sin datos.")
                    else:
                        st.dataframe(dft.style.format("{:,.2f}").highlight_null("lightgray"),width='stretch', key=f"df_026_m_{fname}")
                        _dft_export = dft.reset_index().rename(columns={'index': 'ID_Contrato'})
                        buf = excel_con_formato({titulo[:31]: _dft_export},
                            currency_cols=[c for c in _dft_export.columns if c != 'ID_Contrato'])
                        st.download_button(f"{titulo[:25]}",buf,fname,key=f"dl_m_{fname}")"""

content = re.sub(old_ui_block, original_ui_plus_new, content, flags=re.DOTALL)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)
print("Reversión y separación exitosa.")
