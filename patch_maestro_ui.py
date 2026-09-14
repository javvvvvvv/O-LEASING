import re

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Fix UI for Reporte Maestro to handle string/metadata columns properly
old_ui_block = r"""elif menu=="Reporte Maestro":
        st.title("Reporte Maestro: Capital e Intereses")
        st.markdown("Genera las tablas completas de amortización de toda la cartera activa: \*\*Capital\*\*, \*\*Intereses\*\*, \*\*Renta Neta\*\* y anexos \(Residual/Comisiones\).")
        anio_m=st.number_input("Año del reporte maestro",min_value=2020,max_value=2050,value=datetime.now().year,step=1,key="ym")
        if st.button("Generar Reporte Maestro", type="primary"):
            di, dcap, drenta, dr, dc, ds, err = reporte_maestro_mensual(anio_m)
            if err: st.error\(err\)
            else:
                for titulo,dft,fname in \[
                    \("1. Capital Leasing \(Amortización Principal\)", dcap, f"capital_leasing_\{anio_m\}\.xlsx"\),
                    \("2. Intereses Leasing \(Devengados\)", di, f"intereses_leasing_\{anio_m\}\.xlsx"\),
                    \("3. Renta Neta \(Capital \+ Interés\)", drenta, f"renta_neta_\{anio_m\}\.xlsx"\),
                    \("4. Intereses Residual \(Acumulación\)", dr, f"int_residual_\{anio_m\}\.xlsx"\),
                    \("5. Amortización Comisión por Apertura", dc, f"amort_comision_\{anio_m\}\.xlsx"\),
                    \("6. Saldo del Valor Residual Activo", ds, f"saldo_residual_\{anio_m\}\.xlsx"\),
                \]:
                    st.subheader\(titulo\)
                    if dft is None or dft.empty: st.info\("Sin datos."\)
                    else:
                        st.dataframe\(dft.style.format\("\{:,.2f\}"\).highlight_null\("lightgray"\),width='stretch', key=f"df_026_m_\{fname\}"\)
                        _dft_export = dft.reset_index\(\).rename\(columns=\{'index': 'ID_Contrato'\}\)
                        buf = excel_con_formato\(\{titulo\[:31\]: _dft_export\},
                            currency_cols=\[c for c in _dft_export.columns if c != 'ID_Contrato'\]\)
                        st.download_button\(f"\{titulo\[:25\]\}",buf,fname,key=f"dl_m_\{fname\}"\)"""

new_ui_block = """elif menu=="Reporte Maestro":
        st.title("Reporte Maestro: Capital e Intereses")
        st.markdown("Genera las tablas completas de amortización de toda la cartera activa, incluyendo metadatos de los contratos.")
        anio_m=st.number_input("Año del reporte maestro",min_value=2020,max_value=2050,value=datetime.now().year,step=1,key="ym")
        if st.button("Generar Reporte Maestro", type="primary"):
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
                        MN = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
                        fmt_dict = {m: "{:,.2f}" for m in MN}
                        fmt_dict['Valor_Sin_IVA'] = "{:,.2f}"
                        
                        st.dataframe(dft.style.format(fmt_dict).highlight_null("lightgray"), width='stretch', key=f"df_026_m_{fname}")
                        _dft_export = dft.reset_index().rename(columns={'index': 'ID_Contrato'})
                        
                        # No formatear como moneda los campos de texto
                        skip_currency = ['ID_Contrato', 'Cliente', 'Vehiculo', 'Estatus', 'Fecha_Alta', 'Plazo', 'Fecha_Baja', 'Tasa_Anual_%']
                        curr_cols = [c for c in _dft_export.columns if c not in skip_currency]
                        
                        buf = excel_con_formato({titulo[:31]: _dft_export}, currency_cols=curr_cols, pct_cols=['Tasa_Anual_%'])
                        st.download_button(f"{titulo[:25]}", buf, fname, key=f"dl_m_{fname}")"""

# We'll use index/replace to avoid regex syntax hell
start_str = 'elif menu=="Reporte Maestro":'
end_str = 'key=f"dl_m_{fname}")'

idx_start = content.find(start_str)
idx_end = content.find(end_str, idx_start)

if idx_start != -1 and idx_end != -1:
    content = content[:idx_start] + new_ui_block + content[idx_end + len(end_str):]
    with open('app.py', 'w', encoding='utf-8') as f:
        f.write(content)
    print("UI Reporte maestro actualizada.")
else:
    print("No se encontraron los bloques.")
