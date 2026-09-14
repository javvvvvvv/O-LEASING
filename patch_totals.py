import re

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Update reporte_maestro_saldos
new_saldos_funcs = """def reporte_maestro_saldos(anio):
    import numpy as np
    df=obtener()
    if df.empty: return None,None,None,None,"Sin contratos registrados"
    MN=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
    ids=df['ID_Contrato'].tolist()
    d_cap = pd.DataFrame(index=ids,columns=range(1,13),dtype=float)
    d_int = d_cap.copy()
    d_tot = d_cap.copy()
    d_res = d_cap.copy()
    bar=st.progress(0,"Calculando Reporte de Saldos...")
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
        
        interes_array = dfa['Interes'].values
        rem_int = np.zeros(pl)
        for idx in range(pl):
            rem_int[idx] = np.sum(interes_array[idx+1:])
            
        for mes in range(1,13):
            fm=pd.Timestamp(anio,mes,1)
            if fm < pd.Timestamp(fa.year,fa.month,1): continue
                
            if fecha_baja is not None and fm > pd.Timestamp(fecha_baja.year,fecha_baja.month,1):
                d_cap.at[id_c,mes] = 0; d_int.at[id_c,mes] = 0; d_tot.at[id_c,mes] = 0; d_res.at[id_c,mes] = 0
                continue
                
            mc = (anio-fa.year)*12 + (mes-fa.month) + 1
            if mc > pl:
                d_cap.at[id_c,mes] = 0; d_int.at[id_c,mes] = 0; d_tot.at[id_c,mes] = 0; d_res.at[id_c,mes] = 0
                continue
                
            idx = mc - 1
            saldo_cap = dfa.iloc[idx]['Saldo']
            saldo_int = rem_int[idx]
            saldo_res = dfr.iloc[idx]['Saldo_Fin']
            
            d_cap.at[id_c,mes] = round(saldo_cap, 2)
            d_int.at[id_c,mes] = round(saldo_int, 2)
            d_tot.at[id_c,mes] = round(saldo_cap + saldo_int, 2)
            d_res.at[id_c,mes] = round(saldo_res, 2)
            
        bar.progress((i+1)/tot,text=f"Procesando {i+1}/{tot}...")
    bar.empty()
    
    df_meta = df.set_index('ID_Contrato')[['Cliente', 'Vehiculo', 'Estatus', 'Fecha_Alta', 'Plazo', 'Fecha_Baja', 'Valor_Sin_IVA', 'Tasa_Calculada']].copy()
    df_meta['Fecha_Alta'] = df_meta['Fecha_Alta'].dt.strftime('%Y-%m-%d')
    df_meta['Fecha_Baja'] = df_meta['Fecha_Baja'].dt.strftime('%Y-%m-%d').fillna('')
    df_meta['Tasa_Anual_%'] = (df_meta['Tasa_Calculada'] * 1200).round(2)
    df_meta.drop(columns=['Tasa_Calculada'], inplace=True)
    df_meta['Valor_Sin_IVA'] = df_meta['Valor_Sin_IVA'].round(2)
    
    out = []
    for d in [d_cap, d_int, d_tot, d_res]: 
        d.columns=MN
        d.dropna(how='all',inplace=True)
        d_merged = df_meta.join(d, how='right')
        
        # Agregar Columna TOTAL AÑO
        d_merged['Total Año'] = d_merged[MN].sum(axis=1)
        
        # Agregar Fila TOTAL CARTERA
        total_row = d_merged[MN + ['Valor_Sin_IVA', 'Total Año']].sum()
        total_row['Cliente'] = 'TOTAL CARTERA'
        d_merged.loc['TOTAL_000'] = total_row  # Use TOTAL_000 so it can sort or be distinct
        
        out.append(d_merged)
        
    return out[0], out[1], out[2], out[3], None


def exportar_saldos_completo(anio, dcap, dint, dtot, dres, avg_cap, avg_int, avg_tot):
    import io
    import pandas as pd
    from reports.excel import excel_con_formato
    import openpyxl
    from openpyxl.chart import AreaChart, Reference
    
    MN = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
    df_resumen = pd.DataFrame({
        'Mes': MN,
        'Saldo Capital': avg_cap.values,
        'Saldo Intereses': avg_int.values,
        'Saldo Total': avg_tot.values
    })
    
    skip_currency = ['ID_Contrato', 'Cliente', 'Vehiculo', 'Estatus', 'Fecha_Alta', 'Plazo', 'Fecha_Baja', 'Tasa_Anual_%']
    def pre(df): return df.reset_index().rename(columns={'index': 'ID_Contrato'})
        
    d1 = pre(dcap); d2 = pre(dint); d3 = pre(dtot); d4 = pre(dres)
    curr_cols_1 = [c for c in d1.columns if c not in skip_currency]
    
    hojas = {
        'Resumen Saldos': df_resumen,
        'Saldo Capital': d1,
        'Saldo Intereses': d2,
        'Saldo Total': d3,
        'Saldo Residual': d4
    }
    
    c_cols = {
        'Resumen Saldos': ['Saldo Capital', 'Saldo Intereses', 'Saldo Total'],
        'Saldo Capital': curr_cols_1,
        'Saldo Intereses': curr_cols_1,
        'Saldo Total': curr_cols_1,
        'Saldo Residual': curr_cols_1
    }
    
    p_cols = {
        'Saldo Capital': ['Tasa_Anual_%'],
        'Saldo Intereses': ['Tasa_Anual_%'],
        'Saldo Total': ['Tasa_Anual_%'],
        'Saldo Residual': ['Tasa_Anual_%']
    }
    
    buf = excel_con_formato(hojas, currency_cols=c_cols, pct_cols=p_cols)
    buf.seek(0)
    wb = openpyxl.load_workbook(buf)
    ws = wb['Resumen Saldos']
    
    chart = AreaChart()
    chart.title = "Evolución de Saldos de la Cartera"
    chart.style = 13
    chart.x_axis.title = "Mes"
    chart.y_axis.title = "Saldo Pendiente ($)"
    
    data = Reference(ws, min_col=2, min_row=3, max_col=3, max_row=15)
    cats = Reference(ws, min_col=1, min_row=4, max_row=15)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    
    ws.add_chart(chart, "F4")
    chart.width = 18
    chart.height = 10
    
    out_buf = io.BytesIO()
    wb.save(out_buf)
    out_buf.seek(0)
    return out_buf
"""

content = re.sub(r'def reporte_maestro_saldos\(.*return out_buf\n', new_saldos_funcs + '\n', content, flags=re.DOTALL)

# 2. Update reporte_maestro_mensual
new_maestro_out = """
    out = []
    for d in [di,dcap,drenta,dr,dc,ds]: 
        d.columns=MN
        d.dropna(how='all',inplace=True)
        d_merged = df_meta.join(d, how='right')
        
        # Columna Total
        d_merged['Total Año'] = d_merged[MN].sum(axis=1)
        # Fila Total
        total_row = d_merged[MN + ['Valor_Sin_IVA', 'Total Año']].sum()
        total_row['Cliente'] = 'TOTAL CARTERA'
        d_merged.loc['TOTAL_000'] = total_row
        
        out.append(d_merged)
        
    return out[0], out[1], out[2], out[3], out[4], out[5], None"""

content = re.sub(r'    out = \[\]\n    for d in \[di,dcap,drenta,dr,dc,ds\]:.*return out\[0\], out\[1\], out\[2\], out\[3\], out\[4\], out\[5\], None', new_maestro_out, content, flags=re.DOTALL)

# 3. Fix UI for Maestro Saldos
ui_saldos_patch = r"""elif menu=="Reporte Maestro Saldos":.*?elif menu=="Multiempresa":"""

new_ui_saldos = """elif menu=="Reporte Maestro Saldos":
        st.title("Reporte Maestro: Saldos Insolutos")
        st.markdown("Genera las tablas de **saldos pendientes** (lo que los clientes aún deben al final de cada mes) de toda la cartera activa.")
        anio_s=st.number_input("Año del reporte de saldos",min_value=2020,max_value=2050,value=datetime.now().year,step=1,key="ys")
        if st.button("Generar Reporte de Saldos", type="primary"):
            dcap, dint, dtot, dres, err = reporte_maestro_saldos(anio_s)
            if err: st.error(err)
            else:
                MN = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
                
                # Omitimos la fila TOTAL_000 para calcular totales de gráfica
                dcap_g = dcap[dcap.index != 'TOTAL_000']
                dint_g = dint[dint.index != 'TOTAL_000']
                dtot_g = dtot[dtot.index != 'TOTAL_000']
                
                total_cap = dcap_g[MN].sum().round(2)
                total_int = dint_g[MN].sum().round(2)
                total_tot = dtot_g[MN].sum().round(2)
                
                st.markdown("---")
                st.subheader(f"📊 Resumen de Cartera (Saldos a fin de mes) - {anio_s}")
                
                k1, k2, k3 = st.columns(3)
                k1.metric("Saldo Capital (Diciembre)", f"${total_cap.iloc[-1]:,.2f}")
                k2.metric("Saldo Intereses (Diciembre)", f"${total_int.iloc[-1]:,.2f}")
                k3.metric("Saldo Total (Diciembre)", f"${total_tot.iloc[-1]:,.2f}")
                
                fig_exec = go.Figure()
                fig_exec.add_trace(go.Scatter(x=MN, y=total_cap.values, name='Saldo Capital', fill='tonexty', mode='lines', line=dict(color=C['primary'], width=3)))
                fig_exec.add_trace(go.Scatter(x=MN, y=total_int.values, name='Saldo Intereses', fill='tonexty', mode='lines', line=dict(color=C['accent'], width=3)))
                fig_exec.add_trace(go.Scatter(x=MN, y=total_tot.values, name='Saldo Total', mode='lines+markers+text',
                                              text=[f"${v/1000:,.0f}k" if v > 0 else "" for v in total_tot.values],
                                              textposition="top center",
                                              line=dict(color=C['success'], width=3, dash='dot'),
                                              marker=dict(size=8, color=C['success'])))
                
                fig_exec.update_layout(
                    title=f"Curva de Abatimiento de Saldos de la Cartera en {anio_s}",
                    hovermode="x unified",
                    xaxis=dict(showgrid=False),
                    yaxis=dict(showgrid=True, gridcolor='rgba(200,200,200,0.2)'),
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    margin=dict(l=10, r=10, t=40, b=10),
                    legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1)
                )
                fig_exec = sfig(fig_exec, h=380)
                st.plotly_chart(fig_exec, width='stretch', key="pc_maestro_saldos")
                
                buf_completo = exportar_saldos_completo(anio_s, dcap, dint, dtot, dres, total_cap, total_int, total_tot)
                st.download_button("📥 Descargar Todo en un Solo Excel (con Gráficas)", buf_completo, f"reporte_saldos_completo_{anio_s}.xlsx", type="primary", use_container_width=True)
                
                st.markdown("---")
                
                for titulo,dft,fname in [
                    ("1. Saldo Capital (Insoluto al fin de mes)", dcap, f"saldo_capital_{anio_s}.xlsx"),
                    ("2. Saldo Intereses (Por devengar al fin de mes)", dint, f"saldo_intereses_{anio_s}.xlsx"),
                    ("3. Saldo Total (Lo que debe en total)", dtot, f"saldo_total_{anio_s}.xlsx"),
                    ("4. Saldo Residual", dres, f"saldo_residual_{anio_s}.xlsx"),
                ]:
                    st.subheader(titulo)
                    if dft is None or dft.empty: st.info("Sin datos.")
                    else:
                        fmt_dict = {m: "{:,.2f}" for m in MN}
                        fmt_dict['Valor_Sin_IVA'] = "{:,.2f}"
                        fmt_dict['Total Año'] = "{:,.2f}"
                        
                        dft_vis = dft.copy()
                        for col in MN + ['Valor_Sin_IVA', 'Tasa_Anual_%', 'Total Año']:
                            if col in dft_vis.columns:
                                dft_vis[col] = pd.to_numeric(dft_vis[col], errors='ignore').round(2)
                                
                        st.dataframe(dft_vis.style.format(fmt_dict).highlight_null("lightgray"), width='stretch', key=f"df_026_s_{fname}")
                        _dft_export = dft_vis.reset_index().rename(columns={'index': 'ID_Contrato'})
                        
                        skip_currency = ['ID_Contrato', 'Cliente', 'Vehiculo', 'Estatus', 'Fecha_Alta', 'Plazo', 'Fecha_Baja', 'Tasa_Anual_%']
                        from reports.excel import excel_con_formato
                        curr_cols = [c for c in _dft_export.columns if c not in skip_currency]
                        buf = excel_con_formato({titulo[:31]: _dft_export}, currency_cols=curr_cols, pct_cols=['Tasa_Anual_%'])
                        st.download_button(f"{titulo[:25]}", buf, fname, key=f"dl_s_{fname}")

    elif menu=="Multiempresa":"""

content = re.sub(ui_saldos_patch, new_ui_saldos, content, flags=re.DOTALL)

# 4. We also need to fix fmt_dict in `Reporte Maestro` to include 'Total Año'
ui_maestro = r"fmt_dict\['Valor_Sin_IVA'\] = \"\{:,\.2f\}\"\s*# Limitar a 2 decimales explícitamente"
ui_maestro_new = """fmt_dict['Valor_Sin_IVA'] = "{:,.2f}"
                        fmt_dict['Total Año'] = "{:,.2f}"
                        
                        # Limitar a 2 decimales explícitamente"""
content = re.sub(ui_maestro, ui_maestro_new, content)

ui_maestro_loop = r"for col in MN \+ \['Valor_Sin_IVA', 'Tasa_Anual_%'\]:"
ui_maestro_loop_new = "for col in MN + ['Valor_Sin_IVA', 'Tasa_Anual_%', 'Total Año']:"
content = re.sub(ui_maestro_loop, ui_maestro_loop_new, content)

# Remove TOTAL_000 for original maestro chart calculations
ui_maestro_chart = r"total_cap = dcap\[MN\]\.sum\(\)\.round\(2\)"
ui_maestro_chart_new = """dcap_g = dcap[dcap.index != 'TOTAL_000']
                di_g = di[di.index != 'TOTAL_000']
                drenta_g = drenta[drenta.index != 'TOTAL_000']
                
                total_cap = dcap_g[MN].sum().round(2)"""
content = re.sub(ui_maestro_chart, ui_maestro_chart_new, content)
content = re.sub(r"total_int = di\[MN\]\.sum\(\)\.round\(2\)", "total_int = di_g[MN].sum().round(2)", content)
content = re.sub(r"total_renta = drenta\[MN\]\.sum\(\)\.round\(2\)", "total_renta = drenta_g[MN].sum().round(2)", content)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)
print("Parche aplicado!")
