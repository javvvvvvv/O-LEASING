import re

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# I will replace the block inside "elif menu=='Reporte Maestro':" after the `if st.button("Generar Reporte Maestro"...`
# Actually, I will write a script to patch it safely.

ui_patch = """elif menu=="Reporte Maestro":
        st.title("Reporte Maestro: Capital e Intereses")
        st.markdown("Genera las tablas completas de amortización de toda la cartera activa, incluyendo metadatos de los contratos y resúmenes ejecutivos.")
        anio_m=st.number_input("Año del reporte maestro",min_value=2020,max_value=2050,value=datetime.now().year,step=1,key="ym")
        if st.button("Generar Reporte Maestro", type="primary"):
            di, dcap, drenta, dr, dc, ds, err = reporte_maestro_mensual(anio_m)
            if err: st.error(err)
            else:
                MN = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
                
                # Calcular totales globales para KPIs y gráficas (limitado a 2 decimales)
                total_cap = dcap[MN].sum().round(2)
                total_int = di[MN].sum().round(2)
                total_renta = drenta[MN].sum().round(2)
                
                st.markdown("---")
                st.subheader(f"📊 Resumen Ejecutivo {anio_m}")
                
                k1, k2, k3 = st.columns(3)
                k1.metric("Capital a Recuperar", f"${total_cap.sum():,.2f}")
                k2.metric("Intereses a Devengar", f"${total_int.sum():,.2f}")
                k3.metric("Flujo Total (Renta Neta)", f"${total_renta.sum():,.2f}")
                
                # Gráfica Coqueta (Composición mensual)
                fig_exec = go.Figure()
                fig_exec.add_trace(go.Bar(x=MN, y=total_cap.values, name='Amortización de Capital', marker_color=C['primary'], opacity=0.85))
                fig_exec.add_trace(go.Bar(x=MN, y=total_int.values, name='Intereses', marker_color=C['accent'], opacity=0.85))
                fig_exec.add_trace(go.Scatter(x=MN, y=total_renta.values, name='Flujo Total', mode='lines+markers+text',
                                              text=[f"${v/1000:,.0f}k" if v > 0 else "" for v in total_renta.values],
                                              textposition="top center",
                                              line=dict(color=C['success'], width=3),
                                              marker=dict(size=8, color=C['success'], line=dict(width=2, color='white'))))
                
                fig_exec.update_layout(
                    barmode='stack',
                    title=f"Evolución del Flujo de Efectivo en {anio_m}",
                    hovermode="x unified",
                    xaxis=dict(showgrid=False),
                    yaxis=dict(showgrid=True, gridcolor='rgba(200,200,200,0.2)'),
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    margin=dict(l=10, r=10, t=40, b=10),
                    legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1)
                )
                fig_exec = sfig(fig_exec, h=380)
                st.plotly_chart(fig_exec, width='stretch', key="pc_maestro_exec")
                
                st.markdown("---")
                
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
                        fmt_dict = {m: "{:,.2f}" for m in MN}
                        fmt_dict['Valor_Sin_IVA'] = "{:,.2f}"
                        
                        # Limitar a 2 decimales explícitamente en el dataframe visual
                        dft_vis = dft.copy()
                        for col in MN + ['Valor_Sin_IVA', 'Tasa_Anual_%']:
                            if col in dft_vis.columns:
                                dft_vis[col] = dft_vis[col].round(2)
                                
                        st.dataframe(dft_vis.style.format(fmt_dict).highlight_null("lightgray"), width='stretch', key=f"df_026_m_{fname}")
                        _dft_export = dft_vis.reset_index().rename(columns={'index': 'ID_Contrato'})
                        
                        skip_currency = ['ID_Contrato', 'Cliente', 'Vehiculo', 'Estatus', 'Fecha_Alta', 'Plazo', 'Fecha_Baja', 'Tasa_Anual_%']
                        curr_cols = [c for c in _dft_export.columns if c not in skip_currency]
                        
                        buf = excel_con_formato({titulo[:31]: _dft_export}, currency_cols=curr_cols, pct_cols=['Tasa_Anual_%'])
                        st.download_button(f"{titulo[:25]}", buf, fname, key=f"dl_m_{fname}")"""

start_str = 'elif menu=="Reporte Maestro":'
end_str = 'key=f"dl_m_{fname}")'

idx_start = content.find(start_str)
idx_end = content.find(end_str, idx_start)

if idx_start != -1 and idx_end != -1:
    content = content[:idx_start] + ui_patch + content[idx_end + len(end_str):]
    with open('app.py', 'w', encoding='utf-8') as f:
        f.write(content)
    print("Gráficas coquetas agregadas.")
else:
    print("No se encontró el bloque.")
