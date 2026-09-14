import re

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Update the function
func_pattern = re.compile(r'def tabla_mensual_conceptos\(anio\):.*?return di,dr,dc,ds,None', re.DOTALL)
new_func = """def tabla_mensual_conceptos(anio):
    df=obtener()
    if df.empty: return None,None,None,None,None,None,"Sin contratos registrados"
    MN=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
    ids=df['ID_Contrato'].tolist()
    di=pd.DataFrame(index=ids,columns=range(1,13),dtype=float)
    dr=di.copy(); dc=di.copy(); ds=di.copy(); dcap=di.copy(); drenta=di.copy()
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
    return di, dcap, drenta, dr, dc, ds, None"""

content = func_pattern.sub(new_func, content)

# 2. Update the UI
ui_pattern = re.compile(r'elif menu=="Tabla Mensual por Contrato":.*?\]:', re.DOTALL)
new_ui = """elif menu=="Tabla Mensual por Contrato":
        st.title("Reporte Global: Capital e Intereses")
        st.markdown("Genera las tablas completas de amortización de toda la cartera activa: **Capital**, **Intereses**, **Renta Neta** y anexos (Residual/Comisiones).")
        anio_t=st.number_input("Año del reporte",min_value=2020,max_value=2050,value=datetime.now().year,step=1,key="yt")
        if st.button("Generar Reporte Maestro", kind="primary"):
            di, dcap, drenta, dr, dc, ds, err = tabla_mensual_conceptos(anio_t)
            if err: st.error(err)
            else:
                for titulo,dft,fname in [
                    ("1. Capital Leasing (Amortización Principal)", dcap, f"capital_leasing_{anio_t}.xlsx"),
                    ("2. Intereses Leasing (Devengados)", di, f"intereses_leasing_{anio_t}.xlsx"),
                    ("3. Renta Neta (Capital + Interés)", drenta, f"renta_neta_{anio_t}.xlsx"),
                    ("4. Intereses Residual (Acumulación)", dr, f"int_residual_{anio_t}.xlsx"),
                    ("5. Amortización Comisión por Apertura", dc, f"amort_comision_{anio_t}.xlsx"),
                    ("6. Saldo del Valor Residual Activo", ds, f"saldo_residual_{anio_t}.xlsx"),
                ]:"""

content = ui_pattern.sub(new_ui, content)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)
print("Funciones de Reporte Maestro actualizadas!")
