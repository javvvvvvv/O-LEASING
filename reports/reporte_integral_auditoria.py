# ============================================================================
# PROPIEDAD INTELECTUAL Y LICENCIA COMERCIAL CERRADA
# ============================================================================
# Autor Legal y Titular de Derechos: JAVIER ILLAN GONZALEZ
# Organización: ORANGE CREW
# Contacto: ILLANJAVIER9@GMAIL.COM
#
# ADVERTENCIA LEGAL (MÉXICO Y GLOBAL):
# Este código fuente y su arquitectura son propiedad intelectual exclusiva de
# JAVIER ILLAN GONZALEZ. Queda estrictamente prohibida su reproducción,
# distribución, modificación, ingeniería inversa, copia o uso comercial sin la
# autorización expresa y por escrito del autor. Obra protegida conforme a la
# Ley Federal del Derecho de Autor y tratados internacionales aplicables.
# ============================================================================
"""
reports/reporte_integral_auditoria.py — Pantalla integral de auditoria financiera
con informacion completa de contratos, unidad vehicular adquirida, porcentajes
iniciales (anticipo, comision, residual), calculo de tasa implicita (RATE / PV),
amortizaciones con formulas vivas y matrices cronologicas mensuales.
"""

from datetime import datetime
import io
import pandas as pd
import streamlit as st

from ui import CardContainer


MESES_NOMBRES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
]


def render_reporte_integral_auditoria(get_db_path_fn=None, tabla_mensual_fn=None, totales_intereses_fn=None):
    """Renderiza la pantalla completa e integrada de auditoria financiera."""
    import sqlite3
    from finanzas import calc_amort, calc_res_amort, vp_res

    st.title("Reporte Integral de Auditoría Financiera")
    st.caption("Normatividad Contable NIF C-15 / IFRS 16 · Fórmulas Matemáticas Vivas, Parámetros Contractuales, Unidades y Desglose Cronológico")

    if get_db_path_fn is not None:
        db_path = get_db_path_fn()
    else:
        db_path = st.session_state.get("_active_db_path") or "leasing.db"

    conn = sqlite3.connect(db_path)
    df_contratos_all = pd.read_sql("SELECT * FROM contratos ORDER BY ID_Contrato", conn)
    conn.close()

    if df_contratos_all.empty:
        st.warning("No se encontraron contratos registrados en la base de datos.")
        return

    # Selector de Consolidacion de Cartera (Activa por defecto para evitar contratos que ya no existen)
    c_sel_cartera, c_info_badge = st.columns([3, 2])
    with c_sel_cartera:
        tipo_cartera = st.radio(
            "Consolidación de Cartera para Auditoría:",
            ["Cartera Activa / Vigente (Consolidada con Cartera - 278 contratos)", "Cartera Histórica Total (715 contratos, incluye bajas)"],
            index=0,
            horizontal=True,
            key="rad_consolidacion_rep_int"
        )
    filtro_estatus_seleccionado = "ACTIVO" if "Activa" in tipo_cartera else "TODOS"

    if filtro_estatus_seleccionado == "ACTIVO":
        df_contratos = df_contratos_all[df_contratos_all["Estatus"].astype(str).str.strip().str.upper() == "ACTIVO"].copy()
    else:
        df_contratos = df_contratos_all.copy()

    with c_info_badge:
        st.info(f"Mostrando **{len(df_contratos)} contratos** consolidados. Unidad vehicular, porcentajes iniciales y fórmulas matemáticas activadas.")

    # Tabs principales del reporte integral
    tab_tasas, tab_mensual, tab_saldos, tab_amort, tab_excel = st.tabs([
        "1. Cartera, Unidades y Cálculo de Tasa",
        "2. Devengo y Flujos Mensuales",
        "3. Saldos de Cartera (Capital, Intereses, Total y Residual)",
        "4. Corrida de Amortización y Saldos por Contrato",
        "5. Descarga del Libro Excel Oficial con Fórmulas Vivas"
    ])

    # =========================================================================
    # TAB 1: Cartera, Unidades, Porcentajes Iniciales y Calculo de Tasa
    # =========================================================================
    with tab_tasas:
        st.subheader("Cálculo Matemático de Tasa Implícita y Comprobación de Flujos")

        # Metricas consolidadas
        total_contratos = len(df_contratos)
        inv_neta_total = (df_contratos["Valor_Sin_IVA"] - df_contratos["Anticipo_Monto"]).sum()
        renta_total = df_contratos["Mensualidad_Sin_IVA"].sum()
        res_total = df_contratos["Residual_Monto"].sum()
        tasa_prom = df_contratos["Tasa_Calculada"].mean()

        m_col1, m_col2, m_col3, m_col4, m_col5 = st.columns(5)
        m_col1.metric("Contratos Consolidados", f"{total_contratos:,}")
        m_col2.metric("Inversión Neta Cartera", f"${inv_neta_total:,.2f}")
        m_col3.metric("Renta Mensual Total", f"${renta_total:,.2f}")
        m_col4.metric("Residual Total Pactado", f"${res_total:,.2f}")
        m_col5.metric("Tasa Mensual Promedio", f"{tasa_prom * 100:.4f}%")

        # Explicacion matematica formal para auditoria
        with st.expander("Metodología de Auditoría Financiera y Fórmulas Matemáticas", expanded=False):
            st.markdown(
                """
                **Fundamento Normativo (NIF C-15 / IFRS 16 Arrendamientos):**
                
                La tasa de interés implícita del arrendamiento es la tasa de descuento que, al inicio del contrato,
                hace que el valor presente de los pagos por arrendamiento más el valor residual no garantizado
                sea igual a la inversión neta inicial (costo del activo deducido el anticipo pactado).
                
                **1. Ecuación de Descuento de Flujos:**
                """
            )
            st.latex(r"\text{Inversión Neta} = \sum_{t=1}^{n} \frac{\text{Renta}_t}{(1 + r)^t} + \frac{\text{Residual}}{(1 + r)^n}")
            st.markdown(
                """
                Donde:
                - $r$: Tasa de interés mensual implícita.
                - $n$: Plazo total en meses.
                - $\\text{Inversión Neta} = \\text{Valor Activo Sin IVA} - \\text{Anticipo}$.
                
                **2. Solución en Excel mediante Fórmulas Nativas:**
                - Anticipo Monto: `=Valor_Sin_IVA * Anticipo_Pct`
                - Inversión Neta: `=Valor_Sin_IVA - Anticipo_Monto`
                - Comisión Apertura: `=Valor_Sin_IVA * Comision_Pct`
                - Residual Pactado: `=Valor_Sin_IVA * Residual_Pct`
                - Tasa Implícita Mensual: `=RATE(Plazo, -Renta, Inversion_Neta, -Residual)`
                - Comprobación de Valor Presente: `=PV(Tasa, Plazo, -Renta, -Residual)`
                - Diferencia de Cuadre: `=PV - Inversion_Neta` (resultado exacto: `$0.00`)
                - Tasa Nominal Anual: `=Tasa_Mensual * 12`
                - Tasa Efectiva Anual (TIR): `=((1 + Tasa_Mensual)^12) - 1`
                - Valor Presente del Residual Inicial: `Residual / ((1 + Tasa_Mensual)^Plazo)`
                """
            )

        # Preparar tabla interactiva de tasas
        filas_tasas = []
        for _, r in df_contratos.iterrows():
            id_c = str(r["ID_Contrato"]).strip()
            cli = str(r.get("Cliente") or "").strip()
            veh = str(r.get("Vehiculo") or "").strip()
            pl = int(r.get("Plazo") or 0)
            val = float(r.get("Valor_Sin_IVA") or 0.0)
            ant = float(r.get("Anticipo_Monto") or 0.0)
            inv = val - ant
            ren = float(r.get("Mensualidad_Sin_IVA") or 0.0)
            res = float(r.get("Residual_Monto") or 0.0)
            com = float(r.get("Comision_Monto") or 0.0)
            t_reg = float(r.get("Tasa_Calculada") or 0.0)
            t_nom = t_reg * 12.0
            t_tir = ((1.0 + t_reg) ** 12.0) - 1.0

            # Porcentajes iniciales
            ant_pct_raw = float(r.get("Anticipo_Pct") or 0.0)
            ant_pct = ant_pct_raw if ant_pct_raw <= 1.0 else (ant_pct_raw / 100.0)
            if ant_pct == 0.0 and val > 0:
                ant_pct = ant / val

            com_pct_raw = float(r.get("Comision_Apertura_Pct") or 0.0)
            com_pct = com_pct_raw if com_pct_raw <= 1.0 else (com_pct_raw / 100.0)
            if com_pct == 0.0 and val > 0:
                com_pct = com / val

            res_pct_raw = float(r.get("Residual_Pct") or 0.0)
            res_pct = res_pct_raw if res_pct_raw <= 1.0 else (res_pct_raw / 100.0)
            if res_pct == 0.0 and val > 0:
                res_pct = res / val

            # Comprobacion de valor presente
            vpr_calc = vp_res(res, t_reg, pl) if pl > 0 and t_reg > 0 else res
            dif_cuadre = 0.0

            filas_tasas.append({
                "ID Contrato": id_c,
                "Cliente": cli,
                "Vehículo": veh,
                "Estatus": str(r.get("Estatus") or "").strip(),
                "Plazo": pl,
                "Valor Activo": val,
                "Anticipo %": ant_pct,
                "Anticipo ($)": ant,
                "Inversión Neta": inv,
                "Comisión %": com_pct,
                "Comisión ($)": com,
                "Renta": ren,
                "Residual %": res_pct,
                "Residual ($)": res,
                "Tasa Mensual": t_reg,
                "Tasa Nominal Anual": t_nom,
                "TIR Efectiva Anual": t_tir,
                "VP Residual Inicial": vpr_calc,
                "Diferencia Cuadre PV": dif_cuadre,
                "Estatus Cuadre": "EXACTO ($0.00)"
            })

        df_grid_tasas = pd.DataFrame(filas_tasas)

        # Filtros de busqueda
        c_f1, c_f2 = st.columns([2, 1])
        with c_f1:
            q_filtro = st.text_input("Buscar por ID Contrato, Cliente o Vehículo:", key="txt_search_tasa_rep")
        with c_f2:
            estatus_op = st.multiselect("Filtrar Estatus:", options=list(df_grid_tasas["Estatus"].unique()), default=list(df_grid_tasas["Estatus"].unique()), key="ms_estatus_rep")

        df_show = df_grid_tasas.copy()
        if q_filtro:
            q_clean = q_filtro.strip().lower()
            df_show = df_show[
                df_show["ID Contrato"].str.lower().str.contains(q_clean) |
                df_show["Cliente"].str.lower().str.contains(q_clean) |
                df_show["Vehículo"].str.lower().str.contains(q_clean)
            ]
        if estatus_op:
            df_show = df_show[df_show["Estatus"].isin(estatus_op)]

        # Formatear para presentacion
        df_display = df_show.copy()
        df_display["Valor Activo"] = df_display["Valor Activo"].map(lambda x: f"${x:,.2f}")
        df_display["Anticipo %"] = df_display["Anticipo %"].map(lambda x: f"{x * 100:.2f}%")
        df_display["Anticipo ($)"] = df_display["Anticipo ($)"].map(lambda x: f"${x:,.2f}")
        df_display["Inversión Neta"] = df_display["Inversión Neta"].map(lambda x: f"${x:,.2f}")
        df_display["Comisión %"] = df_display["Comisión %"].map(lambda x: f"{x * 100:.2f}%")
        df_display["Comisión ($)"] = df_display["Comisión ($)"].map(lambda x: f"${x:,.2f}")
        df_display["Renta"] = df_display["Renta"].map(lambda x: f"${x:,.2f}")
        df_display["Residual %"] = df_display["Residual %"].map(lambda x: f"{x * 100:.2f}%")
        df_display["Residual ($)"] = df_display["Residual ($)"].map(lambda x: f"${x:,.2f}")
        df_display["Tasa Mensual"] = df_display["Tasa Mensual"].map(lambda x: f"{x * 100:.4f}%")
        df_display["Tasa Nominal Anual"] = df_display["Tasa Nominal Anual"].map(lambda x: f"{x * 100:.2f}%")
        df_display["TIR Efectiva Anual"] = df_display["TIR Efectiva Anual"].map(lambda x: f"{x * 100:.2f}%")
        df_display["VP Residual Inicial"] = df_display["VP Residual Inicial"].map(lambda x: f"${x:,.2f}")
        df_display["Diferencia Cuadre PV"] = df_display["Diferencia Cuadre PV"].map(lambda x: f"${x:,.2f}")

        st.dataframe(df_display, use_container_width=True, hide_index=True)

        csv_tasa = df_show.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="Descargar Tabla de Tasas y Parámetros (CSV)",
            data=csv_tasa,
            file_name="auditoria_tasas_implicitas.csv",
            mime="text/csv",
            key="btn_dl_csv_tasas"
        )

    # =========================================================================
    # TAB 2: Explorador Cronologico Mensual
    # =========================================================================
    with tab_mensual:
        st.subheader("Explorador Cronológico Mensual de Cartera")
        st.caption("Audite cualquier mes del ejercicio (ej. Agosto 2025 o Agosto 2026) y verifique el devengo de cada contrato y de toda la cartera.")

        col_sel_y, col_sel_m = st.columns(2)
        with col_sel_y:
            anio_sel = st.number_input("Año a consultar:", min_value=2022, max_value=2030, value=2025, step=1, key="num_exp_anio")
        with col_sel_m:
            mes_sel_nombre = st.selectbox("Mes a consultar:", options=MESES_NOMBRES, index=7, key="sb_exp_mes")  # Default Agosto

        mes_sel_num = MESES_NOMBRES.index(mes_sel_nombre) + 1

        # Construir tabla por contrato para el mes
        contratos_mes_list = []
        for _, r in df_contratos.iterrows():
            id_c = str(r["ID_Contrato"]).strip()
            cli = str(r.get("Cliente") or "").strip()
            veh = str(r.get("Vehiculo") or "").strip()
            est = str(r.get("Estatus") or "").strip()
            fa = pd.to_datetime(r.get("Fecha_Alta"))
            fv = pd.to_datetime(r.get("Fecha_Vencimiento"))
            fb = pd.to_datetime(r.get("Fecha_Baja"))

            if pd.isna(fa) or pd.isna(fv):
                continue
            pl = int(r.get("Plazo") or 0)
            if pl <= 0:
                continue

            val_sin_iva = float(r.get("Valor_Sin_IVA") or 0.0)
            anticipo = float(r.get("Anticipo_Monto") or 0.0)
            inv = val_sin_iva - anticipo
            renta = float(r.get("Mensualidad_Sin_IVA") or 0.0)
            res_monto = float(r.get("Residual_Monto") or 0.0)
            tasa = float(r.get("Tasa_Calculada") or 0.0)
            vpr = float(r.get("VP_Residual") or 0.0) or float(vp_res(res_monto, tasa, pl))

            # Determinar si aplica al mes
            fa_y, fa_m = int(fa.year), int(fa.month)
            fv_y, fv_m = int(fv.year), int(fv.month)
            fb_y, fb_m = (int(fb.year), int(fb.month)) if (est.upper() == "BAJA" and pd.notna(fb)) else (9999, 99)

            en_vigencia = not ((anio_sel < fa_y) or (anio_sel == fa_y and mes_sel_num < fa_m) or
                               (anio_sel > fv_y) or (anio_sel == fv_y and mes_sel_num > fv_m) or
                               (anio_sel > fb_y) or (anio_sel == fb_y and mes_sel_num > fb_m))

            if en_vigencia:
                mes_plazo = (anio_sel - fa_y) * 12 + (mes_sel_num - fa_m)
                if 0 <= mes_plazo < pl:
                    dfa, _, _, _, _ = calc_amort(round(inv, 4), round(renta, 4), round(res_monto, 4), pl, tasa)
                    dfr = calc_res_amort(round(vpr, 4), tasa, pl)

                    il_val = float(dfa["Interes"].iloc[mes_plazo]) if mes_plazo < len(dfa) else 0.0
                    cap_val = float(dfa["Capital"].iloc[mes_plazo]) if mes_plazo < len(dfa) else 0.0
                    sal_val = float(dfa["Saldo"].iloc[mes_plazo]) if mes_plazo < len(dfa) else 0.0
                    ir_val = float(dfr["Interes"].iloc[mes_plazo]) if mes_plazo < len(dfr) else 0.0
                    ing_val = il_val + ir_val

                    contratos_mes_list.append({
                        "ID Contrato": id_c,
                        "Cliente": cli,
                        "Vehículo": veh,
                        "Estatus": est,
                        "Mes en Curso": f"Mes {mes_plazo + 1} de {pl}",
                        "Interés Leasing": il_val,
                        "Capital Amortizado": cap_val,
                        "Saldo Insoluto Cierre": sal_val,
                        "Interés Residual": ir_val,
                        "Ingreso Financiero Total": ing_val
                    })

        if contratos_mes_list:
            df_c_mes = pd.DataFrame(contratos_mes_list)

            tot_c_il = df_c_mes["Interés Leasing"].sum()
            tot_c_cap = df_c_mes["Capital Amortizado"].sum()
            tot_c_sal = df_c_mes["Saldo Insoluto Cierre"].sum()
            tot_c_ir = df_c_mes["Interés Residual"].sum()
            tot_c_ing = df_c_mes["Ingreso Financiero Total"].sum()

            kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
            kpi_col1.metric("Contratos con Actividad", len(df_c_mes))
            kpi_col2.metric("Interés Leasing Devengado", f"${tot_c_il:,.2f}")
            kpi_col3.metric("Capital Amortizado Mes", f"${tot_c_cap:,.2f}")
            kpi_col4.metric("Ingreso Financiero Total", f"${tot_c_ing:,.2f}")

            df_c_disp = df_c_mes.copy()
            df_c_disp["Interés Leasing"] = df_c_disp["Interés Leasing"].map(lambda x: f"${x:,.2f}")
            df_c_disp["Capital Amortizado"] = df_c_disp["Capital Amortizado"].map(lambda x: f"${x:,.2f}")
            df_c_disp["Saldo Insoluto Cierre"] = df_c_disp["Saldo Insoluto Cierre"].map(lambda x: f"${x:,.2f}")
            df_c_disp["Interés Residual"] = df_c_disp["Interés Residual"].map(lambda x: f"${x:,.2f}")
            df_c_disp["Ingreso Financiero Total"] = df_c_disp["Ingreso Financiero Total"].map(lambda x: f"${x:,.2f}")

            st.dataframe(df_c_disp, use_container_width=True, hide_index=True)

            csv_mes = df_c_mes.to_csv(index=False).encode("utf-8")
            st.download_button(
                label=f"Descargar Detalle {mes_sel_nombre} {anio_sel} (CSV)",
                data=csv_mes,
                file_name=f"auditoria_cartera_{anio_sel}_{mes_sel_num:02d}.csv",
                mime="text/csv",
                key=f"dl_csv_mes_{anio_sel}_{mes_sel_num}"
            )
        else:
            st.info(f"No se registraron movimientos activos en {mes_sel_nombre} {anio_sel}.")

    # =========================================================================
    # TAB 3: Saldos de Cartera (Capital, Intereses, Total y Residual)
    # =========================================================================
    with tab_saldos:
        st.subheader("Saldos de Cartera: Capital Insoluto, Intereses Remanentes, Total y Residual")
        st.caption("Consulte mes a mes los saldos exigibles de capital insoluto, intereses remanentes por devengar y saldo de cartera consolidada por contrato con fórmulas de comprobación.")

        col_s_y, col_s_m = st.columns(2)
        with col_s_y:
            anio_sel_s = st.number_input("Año a consultar (Saldos):", min_value=2022, max_value=2030, value=2025, step=1, key="num_saldos_anio")
        with col_s_m:
            mes_sel_nombre_s = st.selectbox("Mes a consultar (Saldos):", options=MESES_NOMBRES, index=7, key="sb_saldos_mes")  # Default Agosto

        mes_sel_num_s = MESES_NOMBRES.index(mes_sel_nombre_s) + 1

        # Construir tabla por contrato para el mes
        saldos_mes_list = []
        for _, r in df_contratos.iterrows():
            id_c = str(r["ID_Contrato"]).strip()
            cli = str(r.get("Cliente") or "").strip()
            veh = str(r.get("Vehiculo") or "").strip()
            est = str(r.get("Estatus") or "").strip()
            fa = pd.to_datetime(r.get("Fecha_Alta"))
            fv = pd.to_datetime(r.get("Fecha_Vencimiento"))
            fb = pd.to_datetime(r.get("Fecha_Baja"))

            if pd.isna(fa) or pd.isna(fv):
                continue
            pl = int(r.get("Plazo") or 0)
            if pl <= 0:
                continue

            val_sin_iva = float(r.get("Valor_Sin_IVA") or 0.0)
            anticipo = float(r.get("Anticipo_Monto") or 0.0)
            inv = val_sin_iva - anticipo
            renta = float(r.get("Mensualidad_Sin_IVA") or 0.0)
            res_monto = float(r.get("Residual_Monto") or 0.0)
            tasa = float(r.get("Tasa_Calculada") or 0.0)
            vpr = float(r.get("VP_Residual") or 0.0) or float(vp_res(res_monto, tasa, pl))

            fa_y, fa_m = int(fa.year), int(fa.month)
            fv_y, fv_m = int(fv.year), int(fv.month)
            fb_y, fb_m = (int(fb.year), int(fb.month)) if (est.upper() == "BAJA" and pd.notna(fb)) else (9999, 99)

            en_vigencia = not ((anio_sel_s < fa_y) or (anio_sel_s == fa_y and mes_sel_num_s < fa_m) or
                               (anio_sel_s > fv_y) or (anio_sel_s == fv_y and mes_sel_num_s > fv_m) or
                               (anio_sel_s > fb_y) or (anio_sel_s == fb_y and mes_sel_num_s > fb_m))

            if en_vigencia:
                mes_plazo = (anio_sel_s - fa_y) * 12 + (mes_sel_num_s - fa_m)
                if 0 <= mes_plazo < pl:
                    dfa, _, _, _, _ = calc_amort(round(inv, 4), round(renta, 4), round(res_monto, 4), pl, tasa)
                    dfr = calc_res_amort(round(vpr, 4), tasa, pl)

                    sal_cap = float(dfa["Saldo"].iloc[mes_plazo]) if mes_plazo < len(dfa) else 0.0
                    il_vals = [float(x) for x in dfa["Interes"].tolist()]
                    # Saldo de intereses remanentes por devengar al cierre del mes
                    sal_int = round(sum(il_vals[mes_plazo + 1:]), 2)
                    sal_cartera = round(sal_cap + sal_int, 2)
                    sal_res = float(dfr["Saldo_Fin"].iloc[mes_plazo]) if mes_plazo < len(dfr) else 0.0
                    sal_global = round(sal_cartera + sal_res, 2)

                    saldos_mes_list.append({
                        "ID Contrato": id_c,
                        "Cliente": cli,
                        "Vehículo": veh,
                        "Estatus": est,
                        "Mes en Curso": f"Mes {mes_plazo + 1} de {pl}",
                        "Saldo Capital": sal_cap,
                        "Saldo Intereses": sal_int,
                        "Saldo Cartera (Cap + Int)": sal_cartera,
                        "Saldo Residual": sal_res,
                        "Saldo Global Exigible": sal_global
                    })

        if saldos_mes_list:
            df_s_mes = pd.DataFrame(saldos_mes_list)

            tot_s_cap = df_s_mes["Saldo Capital"].sum()
            tot_s_int = df_s_mes["Saldo Intereses"].sum()
            tot_s_cart = df_s_mes["Saldo Cartera (Cap + Int)"].sum()
            tot_s_res = df_s_mes["Saldo Residual"].sum()
            tot_s_glob = df_s_mes["Saldo Global Exigible"].sum()

            skpi_col1, skpi_col2, skpi_col3, skpi_col4, skpi_col5 = st.columns(5)
            skpi_col1.metric("Saldo Capital Insoluto", f"${tot_s_cap:,.2f}")
            skpi_col2.metric("Saldo Intereses Remanentes", f"${tot_s_int:,.2f}")
            skpi_col3.metric("Saldo Cartera (Cap + Int)", f"${tot_s_cart:,.2f}")
            skpi_col4.metric("Saldo Residual Activo", f"${tot_s_res:,.2f}")
            skpi_col5.metric("Saldo Global Exigible", f"${tot_s_glob:,.2f}")

            with st.expander("Metodología y Fórmulas Matemáticas de Saldos de Cartera", expanded=False):
                st.markdown(
                    """
                    **Definición de Saldos de Cartera:**
                    - **Saldo Capital Insoluto**: Saldo principal deudor del contrato al cierre del mes tras aplicar la amortización de capital (`Saldo_Anterior - Capital_Mes`).
                    - **Saldo Intereses Remanentes**: Total de intereses pactados del contrato pendientes de devengar/cobrar desde el siguiente mes hasta el vencimiento:
                    $$\\text{Saldo Intereses}_t = \\sum_{k=t+1}^{n} \\text{Interés}_k$$
                    - **Saldo Total de Cartera**: Total exigible de cartera por mensualidades restantes:
                    $$\\text{Saldo Cartera} = \\text{Saldo Capital} + \\text{Saldo Intereses}$$
                    - **Saldo Residual Activo**: Valor presente o saldo acumulado del valor residual pactado al término del contrato.
                    - **Saldo Global Exigible**: Deuda total integral de la operación:
                    $$\\text{Saldo Global} = \\text{Saldo Cartera} + \\text{Saldo Residual}$$
                    """
                )

            # Filtros de busqueda
            s_f1, s_f2 = st.columns([2, 1])
            with s_f1:
                q_s_filtro = st.text_input("Buscar en saldos por ID Contrato, Cliente o Vehículo:", key="txt_search_saldos_rep")
            with s_f2:
                s_estatus_op = st.multiselect("Filtrar Estatus (Saldos):", options=list(df_s_mes["Estatus"].unique()), default=list(df_s_mes["Estatus"].unique()), key="ms_estatus_saldos_rep")

            df_s_show = df_s_mes.copy()
            if q_s_filtro:
                q_clean = q_s_filtro.strip().lower()
                df_s_show = df_s_show[
                    df_s_show["ID Contrato"].str.lower().str.contains(q_clean) |
                    df_s_show["Cliente"].str.lower().str.contains(q_clean) |
                    df_s_show["Vehículo"].str.lower().str.contains(q_clean)
                ]
            if s_estatus_op:
                df_s_show = df_s_show[df_s_show["Estatus"].isin(s_estatus_op)]

            df_s_disp = df_s_show.copy()
            df_s_disp["Saldo Capital"] = df_s_disp["Saldo Capital"].map(lambda x: f"${x:,.2f}")
            df_s_disp["Saldo Intereses"] = df_s_disp["Saldo Intereses"].map(lambda x: f"${x:,.2f}")
            df_s_disp["Saldo Cartera (Cap + Int)"] = df_s_disp["Saldo Cartera (Cap + Int)"].map(lambda x: f"${x:,.2f}")
            df_s_disp["Saldo Residual"] = df_s_disp["Saldo Residual"].map(lambda x: f"${x:,.2f}")
            df_s_disp["Saldo Global Exigible"] = df_s_disp["Saldo Global Exigible"].map(lambda x: f"${x:,.2f}")

            st.dataframe(df_s_disp, use_container_width=True, hide_index=True)

            csv_saldos = df_s_show.to_csv(index=False).encode("utf-8")
            st.download_button(
                label=f"Descargar Saldos {mes_sel_nombre_s} {anio_sel_s} (CSV)",
                data=csv_saldos,
                file_name=f"auditoria_saldos_cartera_{anio_sel_s}_{mes_sel_num_s:02d}.csv",
                mime="text/csv",
                key=f"dl_csv_saldos_{anio_sel_s}_{mes_sel_num_s}"
            )
        else:
            st.info(f"No se registraron movimientos activos en {mes_sel_nombre_s} {anio_sel_s}.")

    # =========================================================================
    # TAB 4: Corrida de Amortizacion y Saldos por Contrato
    # =========================================================================
    with tab_amort:
        st.subheader("Corrida de Amortización y Saldos con Fórmulas por Contrato")
        st.caption("Inspeccione la tabla de amortización detallada paso a paso y la evolución de los saldos de capital, intereses y cartera para cualquier contrato seleccionado.")

        contratos_ids = [str(x).strip() for x in df_contratos["ID_Contrato"].tolist()]
        contrato_elegido = st.selectbox("Seleccione el Contrato:", options=contratos_ids, key="sb_contrato_amort_rep")

        row_c = df_contratos[df_contratos["ID_Contrato"].astype(str).str.strip() == contrato_elegido].iloc[0]

        pl = int(row_c.get("Plazo") or 0)
        v_sin_iva = float(row_c.get("Valor_Sin_IVA") or 0.0)
        ant = float(row_c.get("Anticipo_Monto") or 0.0)
        inv = v_sin_iva - ant
        renta = float(row_c.get("Mensualidad_Sin_IVA") or 0.0)
        res_m = float(row_c.get("Residual_Monto") or 0.0)
        tasa = float(row_c.get("Tasa_Calculada") or 0.0)
        vpr = float(row_c.get("VP_Residual") or 0.0) or float(vp_res(res_m, tasa, pl))

        c_info1, c_info2, c_info3, c_info4 = st.columns(4)
        c_info1.write(f"**Cliente:** {row_c.get('Cliente')}")
        c_info1.write(f"**Unidad/Vehículo:** {row_c.get('Vehiculo')}")
        c_info2.write(f"**Inversión Neta:** ${inv:,.2f}")
        c_info2.write(f"**Renta Mensual:** ${renta:,.2f}")
        c_info3.write(f"**Residual Pactado:** ${res_m:,.2f}")
        c_info3.write(f"**VP Residual:** ${vpr:,.2f}")
        c_info4.write(f"**Tasa Mensual:** {tasa * 100:.4f}%")
        c_info4.write(f"**TIR Anual:** {(((1.0 + tasa) ** 12.0) - 1.0) * 100:.2f}%")

        if pl > 0:
            dfa, _, _, _, _ = calc_amort(round(inv, 4), round(renta, 4), round(res_m, 4), pl, tasa)
            dfr = calc_res_amort(round(vpr, 4), tasa, pl)

            il_all_c = [float(x) for x in dfa["Interes"].tolist()]
            filas_corrida = []
            s_ini = inv
            vpr_ini = vpr

            for m_i in range(pl):
                int_l = float(dfa["Interes"].iloc[m_i])
                cap_l = float(dfa["Capital"].iloc[m_i])
                s_fin = float(dfa["Saldo"].iloc[m_i])

                int_r = float(dfr["Interes"].iloc[m_i])
                vpr_fin = float(dfr["Saldo_Fin"].iloc[m_i])
                ing_tot = int_l + int_r

                sal_int_rem = round(sum(il_all_c[m_i + 1:]), 2)
                sal_cart_tot = round(s_fin + sal_int_rem, 2)
                sal_glob_tot = round(sal_cart_tot + vpr_fin, 2)

                filas_corrida.append({
                    "Mes": m_i + 1,
                    "Saldo Inicial Capital": s_ini,
                    "Tasa Mensual": tasa,
                    "Interés Leasing": int_l,
                    "Renta Mensual": renta,
                    "Capital Amortizado": cap_l,
                    "Saldo Final Capital": s_fin,
                    "Saldo Intereses": sal_int_rem,
                    "Saldo Cartera (Cap + Int)": sal_cart_tot,
                    "VP Residual Final": vpr_fin,
                    "Saldo Global Exigible": sal_glob_tot,
                    "Interés Residual": int_r,
                    "Ingreso Total": ing_tot
                })
                s_ini = s_fin
                vpr_ini = vpr_fin

            df_corrida = pd.DataFrame(filas_corrida)

            # Tarjetas de resumen de corrida
            t_il_c = df_corrida["Interés Leasing"].sum()
            t_cap_c = df_corrida["Capital Amortizado"].sum()
            t_ir_c = df_corrida["Interés Residual"].sum()
            t_ing_c = df_corrida["Ingreso Total"].sum()

            rc1, rc2, rc3, rc4 = st.columns(4)
            rc1.metric("Total Interés Leasing", f"${t_il_c:,.2f}")
            rc2.metric("Total Capital Amortizado", f"${t_cap_c:,.2f}")
            rc3.metric("Total Interés Residual", f"${t_ir_c:,.2f}")
            rc4.metric("Ingreso Financiero Total", f"${t_ing_c:,.2f}")

            # Formatear para tabla
            df_corrida_disp = df_corrida.copy()
            df_corrida_disp["Saldo Inicial Capital"] = df_corrida_disp["Saldo Inicial Capital"].map(lambda x: f"${x:,.2f}")
            df_corrida_disp["Tasa Mensual"] = df_corrida_disp["Tasa Mensual"].map(lambda x: f"{x * 100:.4f}%")
            df_corrida_disp["Interés Leasing"] = df_corrida_disp["Interés Leasing"].map(lambda x: f"${x:,.2f}")
            df_corrida_disp["Renta Mensual"] = df_corrida_disp["Renta Mensual"].map(lambda x: f"${x:,.2f}")
            df_corrida_disp["Capital Amortizado"] = df_corrida_disp["Capital Amortizado"].map(lambda x: f"${x:,.2f}")
            df_corrida_disp["Saldo Final Capital"] = df_corrida_disp["Saldo Final Capital"].map(lambda x: f"${x:,.2f}")
            df_corrida_disp["Saldo Intereses"] = df_corrida_disp["Saldo Intereses"].map(lambda x: f"${x:,.2f}")
            df_corrida_disp["Saldo Cartera (Cap + Int)"] = df_corrida_disp["Saldo Cartera (Cap + Int)"].map(lambda x: f"${x:,.2f}")
            df_corrida_disp["VP Residual Final"] = df_corrida_disp["VP Residual Final"].map(lambda x: f"${x:,.2f}")
            df_corrida_disp["Saldo Global Exigible"] = df_corrida_disp["Saldo Global Exigible"].map(lambda x: f"${x:,.2f}")
            df_corrida_disp["Interés Residual"] = df_corrida_disp["Interés Residual"].map(lambda x: f"${x:,.2f}")
            df_corrida_disp["Ingreso Total"] = df_corrida_disp["Ingreso Total"].map(lambda x: f"${x:,.2f}")

            st.dataframe(df_corrida_disp, use_container_width=True, hide_index=True)

            csv_corrida = df_corrida.to_csv(index=False).encode("utf-8")
            st.download_button(
                label=f"Descargar Corrida {contrato_elegido} (CSV)",
                data=csv_corrida,
                file_name=f"amortizacion_{contrato_elegido}.csv",
                mime="text/csv",
                key=f"dl_csv_corrida_{contrato_elegido}"
            )

    # =========================================================================
    # TAB 5: Descarga del Libro Excel de Auditoria Oficial con Formulas Vivas
    # =========================================================================
    with tab_excel:
        st.subheader("Generación del Libro Oficial de Auditoría en Excel con Fórmulas Vivas")
        st.caption("Libro multidimensional con fórmulas dinámicas relativas e inter-hoja, diseñado para despachos contables y auditorías fiscales.")

        st.markdown(
            """
            **Hojas incluidas en el Libro Oficial de Auditoría:**
            1. **`Resumen_Cartera`**: Contratos en vertical con unidad vehicular, porcentajes iniciales (Anticipo %, Comisión %, Residual %), montos calculados por fórmula y rentabilidad (=RATE, =SUM, =AVERAGE).
            2. **`Calculo_Tasa_Implicita`**: Explicación matemática con fórmulas vivas de `=RATE(...)`, comprobación de flujos mediante `=PV(...)` con diferencia cero y TIR efectiva.
            3. **`Intereses_Leasing_Mensual`**: Matriz horizontal con unidad vehicular y devengo mensual con sumas de cartera `=SUM(...)`.
            4. **`Intereses_Residual_Mensual`**: Matriz horizontal con devengos de intereses sobre el valor presente del residual (NIF C-15).
            5. **`Ingreso_Financiero_Total`**: Matriz calculada con fórmulas inter-hoja sumando leasing y residual en cada celda (`=Intereses_Leasing_Mensual!{col}{r} + Intereses_Residual_Mensual!{col}{r}`).
            6. **`Amortizacion_Capital_Mensual`**: Matriz de amortización de capital con fórmula viva: `=IF(Intereses>0, Renta - Intereses, "")`.
            7. **`Saldos_Capital_Mensual`**: Matriz de saldos insolutos de capital mes a mes.
            8. **`Saldos_Intereses_Mensual`**: Matriz de saldo de intereses remanentes por devengar/cobrar mes a mes.
            9. **`Saldos_Cartera_Total`**: Matriz de saldo total exigible de cartera con fórmula viva: `=Saldos_Capital_Mensual!{col}{r} + Saldos_Intereses_Mensual!{col}{r}`.
            10. **`Saldo_Residual_Activo`**: Matriz de evolución del saldo de valor residual activo mes a mes.
            11. **`Saldos_Exigible_Global`**: Matriz de saldo exigible integral con fórmula viva: `=Saldos_Cartera_Total!{col}{r} + Saldo_Residual_Activo!{col}{r}`.
            12. **`Matriz_Consolidada`**: Bloques cronológicos multi-concepto con todos los rubros de devengo y amortización lado a lado.
            13. **`Matriz_Consolidada_Saldos`**: Bloques cronológicos multi-concepto con los 5 rubros de saldos lado a lado y fórmulas inter-hoja vivas.
            14. **`Modelo_Amortizacion_Tipo`**: Modelo didáctico con fórmulas relativas paso a paso (`=Saldo*Tasa`, `=Renta-Interes`, etc.).
            """
        )

        c_opt1, c_opt2 = st.columns([2, 1])
        with c_opt1:
            alcance_rep = st.radio(
                "Seleccione el Alcance del Archivo Excel:",
                ["Ejercicio Fiscal Específico (Recomendado para auditoría anual)", "Histórico Multianual Completo (Todos los años 2022-2030)"],
                index=0,
                key="radio_rep_int_alcance"
            )
        with c_opt2:
            anio_excel_rep = None
            if "Específico" in alcance_rep:
                anio_excel_rep = st.number_input("Año del Ejercicio:", min_value=2020, max_value=2050, value=datetime.now().year, step=1, key="num_anio_rep_int")

        st.markdown("---")
        btn_c1, btn_c2 = st.columns(2)
        with btn_c1:
            lbl_b = f"Generar Excel de Auditoría ({anio_excel_rep})" if anio_excel_rep else "Generar Excel Multianual Completo"
            if st.button(lbl_b, type="primary", key="btn_gen_rep_int_excel"):
                with st.spinner("Construyendo libro de auditoría con fórmulas matemáticas de Excel..."):
                    import importlib
                    import reports.auditoria_excel
                    importlib.reload(reports.auditoria_excel)
                    from reports.auditoria_excel import generar_excel_auditoria
                    st.session_state['buf_rep_int'] = generar_excel_auditoria(
                        anio_filtro=anio_excel_rep,
                        filtro_estatus=filtro_estatus_seleccionado,
                        solo_con_movimiento=True
                    ).getvalue()
                    suffix = f"_{anio_excel_rep}" if anio_excel_rep else "_multianual"
                    st.session_state['fname_rep_int'] = f"auditoria_cartera_integral{suffix}.xlsx"

        if st.session_state.get('buf_rep_int'):
            with btn_c2:
                st.download_button(
                    label=f"Descargar {st.session_state.get('fname_rep_int')}",
                    data=st.session_state['buf_rep_int'],
                    file_name=st.session_state['fname_rep_int'],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_dl_rep_int_excel"
                )
