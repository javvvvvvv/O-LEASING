    elif menu=="Estado de Cuenta":
        st.title("Estado de Cuenta por Contrato")
        st.caption(f"Empresa activa: **{get_empresa_actual().get('nombre','—')}** — los contratos que ves abajo pertenecen solo a esta empresa.")
        df_todos = obtener()
        if df_todos.empty:
            estado_vacio("Todavía no hay contratos registrados",
                         "En cuanto des de alta el primer contrato desde 'Carga Masiva y Altas', aparecerá aquí su estado de cuenta completo.")
        else:
            # Los IDs se numeran por empresa (ej. "0635-0003"), así que el mismo
            # ID puede repetirse en otra empresa apuntando a OTRO cliente. Por
            # eso la selección de contrato se limpia cada vez que se cambia de
            # empresa (ver limpiar_seleccion_contrato) — nunca debe sobrevivir
            # un ID que no exista en la empresa activa.
            opts_ec = df_todos['ID_Contrato'].tolist()
            # Salto rápido desde otra pantalla ("Ver estado de cuenta — X"):
            # se fuerza la preselección ANTES de crear el widget, escribiendo
            # directamente en su estado — nunca junto con "index=", que es la
            # combinación que causaba que la casilla mostrara un contrato
            # mientras los datos de abajo correspondían a otro distinto.
            pre = st.session_state.pop('ec_contrato', None)
            if pre and pre in opts_ec:
                st.session_state['ec_sel'] = pre
            elif 'ec_sel' not in st.session_state or st.session_state['ec_sel'] not in opts_ec:
                st.session_state['ec_sel'] = opts_ec[0]
            sel_ec  = st.selectbox(
                "Contrato",
                opts_ec,
                format_func=lambda x: f"{x}  —  {df_todos.loc[df_todos['ID_Contrato']==x,'Cliente'].values[0]}",
                key="ec_sel"
            )
            _match = df_todos[df_todos['ID_Contrato']==sel_ec]
            if _match.empty:
                # Salvaguarda: si por cualquier motivo el contrato elegido ya no
                # existe en esta empresa, nunca mostramos datos de otro contrato
                # "por accidente" — se avisa y se fuerza a re-elegir.
                st.error("El contrato seleccionado no pertenece a la empresa activa. Vuelve a elegirlo.")
                st.session_state.pop('ec_sel', None)
                st.stop()
            row = _match.iloc[0]
            assert str(row['ID_Contrato']) == str(sel_ec), "Inconsistencia de selección de contrato detectada."

            with st.container(key=f"tabbox_header_{sel_ec}"):
                nm = int(row.get('Nivel_Morosidad',0) or 0)
                ML_ec = {0:"Al corriente",1:"Atraso",2:"Convenio",3:"Devuelve no paga",4:"Judicial"}
                MC_ec = {0:C['success'],1:C['gold'],2:C['warning'],3:C['accent'],4:"#7A1015"}
                excl_poliza = str(row.get('Fecha_Excl_Poliza','') or '').strip()
                excl_poliza = '' if excl_poliza in ('0','None','nan') else excl_poliza
                motivo_excl = str(row.get('Motivo_Excl_Poliza','') or '').strip()
                motivo_excl = '' if motivo_excl in ('0','None','nan') else motivo_excl
                _fb_hdr = str(row.get('Fecha_Baja','') or '').strip()
                _fb_hdr = '' if _fb_hdr in ('0','None','nan','NaT') else _fb_hdr
                es_baja_hdr = str(row.get('Estatus','')).upper() == 'BAJA'

                badge_mora  = f'<span style="background:{MC_ec[nm]};color:#fff;border-radius:6px;padding:2px 10px;font-size:.8rem;font-weight:700;">Nivel {nm} — {ML_ec[nm]}</span>'
                badge_baja  = (f'<span style="background:#B3261E;color:#fff;border-radius:4px;padding:2px 10px;font-size:.8rem;font-weight:700;">Dado de baja el {_fb_hdr[:10]}</span>'
                               if es_baja_hdr and _fb_hdr else
                               ('<span style="background:#96660C;color:#fff;border-radius:4px;padding:2px 10px;font-size:.8rem;font-weight:700;">Dado de baja (sin fecha capturada)</span>' if es_baja_hdr else ''))
                badge_excl  = (f'<span style="background:#B3261E;color:#fff;border-radius:4px;padding:2px 10px;font-size:.8rem;font-weight:700;">Excluido de pólizas desde {excl_poliza}</span>'
                               if excl_poliza else '')
                badge_avance = ''
                if not es_baja_hdr:
                    _avp = calcular_avance_pago(row)
                    if _avp['estado'] == 'ATRASADO':
                        _txt_falt = ', '.join(str(m) for m in _avp['meses_faltantes'][:8]) + \
                                    (f" y {len(_avp['meses_faltantes'])-8} más" if len(_avp['meses_faltantes']) > 8 else '')
                        badge_avance = (f'<span style="background:#B3261E;color:#fff;border-radius:4px;padding:2px 10px;font-size:.8rem;font-weight:700;" '
                                        f'title="Meses de plazo sin factura conciliada: {_txt_falt}">'
                                        f"Va atrasado {_avp['meses_atraso']} mes(es) — falta{'n' if _avp['meses_atraso']>1 else ''} el/los mes(es) {_txt_falt}</span>")
                    elif _avp['estado'] == 'ADELANTADO':
                        badge_avance = (f'<span style="background:#1E5C4F;color:#fff;border-radius:4px;padding:2px 10px;font-size:.8rem;font-weight:700;">'
                                        f"Va adelantado {_avp['meses_adelanto']} mes(es) — ya tiene facturado hasta el mes {_avp['mes_max_facturado']} de {int(row.get('Plazo',0))}</span>")
                _campo_vencimiento = (
                    f"<div><b>Fecha Vencimiento (original)</b><br>{str(row.get('Fecha_Vencimiento','—'))[:10]}</div>"
                    if es_baja_hdr else
                    f"<div><b>Fecha Vencimiento</b><br>{str(row.get('Fecha_Vencimiento','—'))[:10]}</div>"
                )
                st.markdown(f"""
                <div style="background:#FFFFFF;border:1px solid #DCE0E5;border-radius:6px;padding:18px 22px;
                            margin-bottom:16px;">
                  <div style="font-size:1.3rem;font-weight:700;color:#1E5C4F;margin-bottom:8px;">
                    {sel_ec} &nbsp;·&nbsp; {row['Cliente']}
                  </div>
                  <div style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:10px;">
                    {badge_mora} {badge_baja} {badge_excl} {badge_avance}
                  </div>
                  <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:8px;font-size:.85rem;color:#565E68;">
                    <div><b>Vehículo</b><br>{row.get('Vehiculo','—')}</div>
                    <div><b>Fecha Alta</b><br>{str(row.get('Fecha_Alta','—'))[:10]}</div>
                    {_campo_vencimiento}
                    <div><b>Estatus</b><br>{row.get('Estatus','—')}</div>
                    <div><b>Plazo</b><br>{int(row.get('Plazo',0))} meses</div>
                    <div><b>Valor (s/IVA)</b><br>${float(row.get('Valor_Sin_IVA',0)):,.2f}</div>
                    <div><b>Anticipo</b><br>${float(row.get('Anticipo_Monto',0)):,.2f} ({float(row.get('Anticipo_Pct',0)):.1f}%)</div>
                    <div><b>Comisión</b><br>${float(row.get('Comision_Monto',0)):,.2f}</div>
                    <div><b>Renta (s/IVA)</b><br>${float(row.get('Mensualidad_Sin_IVA',0)):,.2f}</div>
                    <div><b>Residual</b><br>${float(row.get('Residual_Monto',0)):,.2f}</div>
                    <div><b>Tasa anual impl.</b><br>{float(row.get('Tasa_Calculada',0))*1200:.2f}%</div>
                  </div>
                  {('<div style="margin-top:10px;font-size:.8rem;color:#B3261E;"><b>Motivo exclusi\u00f3n:</b> ' + motivo_excl + '</div>') if motivo_excl else ""}
                </div>
                """, unsafe_allow_html=True)

                try:
                    g, m, _pe_lineal = rentabilidad(row); t_ec = tir(row)
                    inv_neta = max(float(row['Valor_Sin_IVA']) - float(row['Anticipo_Monto']), 0.01)
                except: g=m=t_ec=inv_neta=0

                # Punto de equilibrio: mismo criterio que la página de
                # Reportes (renta acumulada vs. inversión neta). No se usa
                # solo la columna Capital de la amortización porque casi
                # nunca cuadra si el contrato tiene residual, ya que ese se
                # recupera aparte al final.
                try:
                    mes_pe = mes_pe_rentas(row)
                except Exception:
                    mes_pe = None

                km1,km2,km3,km4,km5 = st.columns(5)
                km1.metric("Inversión neta",   f"${inv_neta:,.2f}")
                km2.metric("Ganancia proy.",   f"${g:,.2f}")
                km3.metric("Margen",           f"{m:.2f}%")
                km4.metric("TIR Anual",        f"{t_ec:.2f}%" if t_ec else "N/A")
                km5.metric("Mes punto equil.", f"Mes {mes_pe}" if mes_pe else "N/A",
                           help="Mes en que el capital recuperado (vía la tabla de amortización) iguala la inversión neta.")

                # Segunda fila: lectura "de contador" — cobranza y saldo
                fa_hdr = pd.to_datetime(row.get('Fecha_Alta'))
                hoy_ts = pd.Timestamp(hoy_ref())
                estatus_hdr = str(row.get('Estatus','')).upper()
                plazo_hdr = int(row.get('Plazo',0) or 0)
                meses_transcurridos = 0
                saldo_insoluto = None
                utilidad_devengada = None
                if estatus_hdr == 'ACTIVO' and pd.notnull(fa_hdr):
                    meses_transcurridos = max((hoy_ts.year-fa_hdr.year)*12 + (hoy_ts.month-fa_hdr.month), 0)
                    dia_corte = fa_hdr.day
                    prox_pago = (fa_hdr + relativedelta(months=meses_transcurridos)).replace(day=min(dia_corte,28))
                    if prox_pago < hoy_ts:
                        prox_pago = prox_pago + relativedelta(months=1)
                    dias_prox = (prox_pago.normalize() - hoy_ts.normalize()).days
                    txt_prox = prox_pago.strftime('%Y-%m-%d')
                    txt_dias = f"En {dias_prox} días" if dias_prox >= 0 else f"{abs(dias_prox)} días de atraso"
                    try:
                        _inv_h = float(row['Valor_Sin_IVA']) - float(row['Anticipo_Monto'])
                        if _inv_h > 0 and plazo_hdr > 0:
                            _dfa_h, _, _, _, _ = calc_amort(round(_inv_h,4), float(row['Mensualidad_Sin_IVA']),
                                                              float(row['Residual_Monto']), plazo_hdr, float(row['Tasa_Calculada']))
                            _idx = min(meses_transcurridos, plazo_hdr) - 1
                            saldo_insoluto = float(_dfa_h.iloc[_idx]['Saldo']) if _idx >= 0 else _inv_h
                            utilidad_devengada = float(_dfa_h.iloc[:max(_idx+1,0)]['Interes'].sum()) if _idx >= 0 else 0.0
                    except Exception:
                        pass
                else:
                    txt_prox, txt_dias = "—", "Contrato no activo"
                pagos_restantes = max(plazo_hdr - meses_transcurridos, 0) if estatus_hdr=='ACTIVO' else 0

                kf1,kf2,kf3,kf4 = st.columns(4)
                kf1.metric("Próximo pago",          txt_prox)
                kf2.metric("Vence en / atraso",     txt_dias)
                kf3.metric("Pagos realizados / restantes", f"{meses_transcurridos} / {pagos_restantes}")
                kf4.metric("Saldo insoluto (capital)", f"${saldo_insoluto:,.2f}" if saldo_insoluto is not None else "—",
                           help="Capital que aún se debe hoy, según la tabla de amortización — no es lo mismo que la inversión neta inicial.")

                kg1,kg2 = st.columns(2)
                kg1.metric("Utilidad (interés) devengada a la fecha", f"${utilidad_devengada:,.2f}" if utilidad_devengada is not None else "—")
                _avance = (meses_transcurridos/plazo_hdr*100) if (estatus_hdr=='ACTIVO' and plazo_hdr>0) else 0
                kg2.metric("Avance del plazo", f"{_avance:.1f}%")
                if estatus_hdr == 'ACTIVO' and plazo_hdr > 0:
                    st.progress(min(max(_avance/100,0),1.0), text=f"Mes {meses_transcurridos} de {plazo_hdr}")

                # Insignia de Punto de Equilibrio: ¿ya se alcanzó o no?
                if estatus_hdr == 'ACTIVO' and mes_pe:
                    _avance_pe = min(meses_transcurridos/mes_pe, 1.0) if mes_pe else 0
                    if meses_transcurridos >= mes_pe:
                        st.markdown(f"""
                        <div class="milestone-banner">
                          <div class="mb-text">
                            <b>¡Punto de equilibrio alcanzado!</b> — se recuperó el capital invertido en el
                            <b>mes {mes_pe}</b> de {plazo_hdr} (llevas {meses_transcurridos} meses transcurridos).
                            Todo lo que se cobre de aquí en adelante es utilidad sobre el capital. 
                          </div>
                        </div>
                        """, unsafe_allow_html=True)
                    else:
                        _falta = mes_pe - meses_transcurridos
                        st.warning(
                            f"**Aún no llega a su punto de equilibrio** — se alcanza en el **mes {mes_pe}** de "
                            f"{plazo_hdr} (van {meses_transcurridos}, faltan {_falta} pago(s) más)."
                        )
                        st.progress(_avance_pe, text=f"Camino al punto de equilibrio: {_avance_pe*100:.0f}%")
                elif estatus_hdr == 'ACTIVO' and not mes_pe:
                    st.info("Con la renta y plazo pactados, este contrato no recupera el capital dentro del plazo por sí solo (depende del valor residual al final).")

                if estatus_hdr == 'ACTIVO' and pd.notnull(fa_hdr) and plazo_hdr > 0:
                    _dia_corte = fa_hdr.day
                    _prox_pagos = []
                    _m0 = meses_transcurridos
                    for _j in range(3):
                        _mm = min(_m0 + _j, plazo_hdr)
                        _f = (fa_hdr + relativedelta(months=_mm)).replace(day=min(_dia_corte,28))
                        if _f < hoy_ts and _j == 0:
                            _f = _f + relativedelta(months=1)
                        _prox_pagos.append({'Fecha': _f.strftime('%Y-%m-%d'), 'Monto programado': float(row.get('Mensualidad_Sin_IVA',0) or 0)})
                    with st.expander("Próximos 3 pagos programados (para planear cobranza)"):
                        st.dataframe(pd.DataFrame(_prox_pagos), width='stretch', height=140, key=f"df_prox3_{sel_ec}")

            st.markdown("---")

            tab_amort, tab_res, tab_evts, tab_anot, tab_cont = st.tabs([
                "Amortización Leasing",
                "Acumulación Residual",
                "Eventos Especiales",
                "Anotaciones",
                "Contabilidad"
            ])

            inv = float(row['Valor_Sin_IVA']) - float(row['Anticipo_Monto'])
            pl  = int(row['Plazo'])
            t   = round(float(row['Tasa_Calculada']), 8)
            r_m = round(float(row['Mensualidad_Sin_IVA']), 4)
            res = round(float(row['Residual_Monto']), 4)
            fa  = pd.to_datetime(row['Fecha_Alta'])

            # Si el contrato se dio de baja ANTES de terminar su plazo original,
            # las tablas de abajo deben cortarse en el mes real de la baja — de lo
            # contrario mostrarían meses de renta que nunca se llegaron a cobrar.
            fecha_baja_ec = None
            _fb_raw = row.get('Fecha_Baja')
            if str(row.get('Estatus','')).upper()=='BAJA' and pd.notnull(_fb_raw) and str(_fb_raw).strip() not in ('','None','NaT','0'):
                try: fecha_baja_ec = pd.to_datetime(_fb_raw)
                except Exception: fecha_baja_ec = None
            mes_corte_ec = None
            if fecha_baja_ec is not None:
                mes_corte_ec = max((fecha_baja_ec.year-fa.year)*12 + (fecha_baja_ec.month-fa.month), 0)
                if mes_corte_ec >= pl:
                    mes_corte_ec = None  # terminó justo en su plazo natural: no hay nada que truncar

            with tab_amort:
                with st.container(key=f"tabbox_amort_{sel_ec}"):
                    if inv > 0:
                        dfa, _, _, _, _ = calc_amort(round(inv,4), r_m, res, pl, t)
                        dfa = dfa.copy()
                        dfa['Fecha']     = dfa['Mes'].apply(lambda m: (fa + relativedelta(months=m)).strftime('%Y-%m'))
                        saldos_ini = [round(inv,4)] + list(dfa['Saldo'].iloc[:-1].round(4))
                        dfa.insert(dfa.columns.get_loc('Interes'), 'Saldo_Ini', saldos_ini)
                        dfa.rename(columns={'Saldo':'Saldo_Fin'}, inplace=True)

                        if mes_corte_ec is not None:
                            dfa = dfa[dfa['Mes'] <= mes_corte_ec].copy()
                            st.warning(
                                f"Este contrato se dio de **baja anticipada** el **{fecha_baja_ec.strftime('%Y-%m-%d')}** "
                                f"(mes {mes_corte_ec} de {pl} originalmente pactados). La tabla y la gráfica solo muestran "
                                "hasta ese punto — los meses posteriores nunca se cobraron."
                            )

                        mes_hoy = date.today().strftime('%Y-%m')
                        mes_act = dfa[dfa['Fecha'] <= mes_hoy]
                        mes_prox = dfa[dfa['Fecha'] > mes_hoy]
                        capital_amort = mes_act['Capital'].sum() if not mes_act.empty else 0
                        capital_pend  = mes_prox['Capital'].sum() if not mes_prox.empty else 0
                        saldo_actual  = mes_act['Saldo_Fin'].iloc[-1] if not mes_act.empty else inv

                        pa1,pa2,pa3 = st.columns(3)
                        pa1.metric("Capital amortizado", f"${capital_amort:,.2f}")
                        pa2.metric("Saldo vigente",       f"${saldo_actual:,.4f}")
                        pa3.metric("Capital pendiente",   f"${capital_pend:,.2f}")

                        fig_a = go.Figure()
                        fig_a.add_trace(go.Bar(x=dfa['Fecha'], y=dfa['Capital'],
                            name='Capital', marker_color=C_PASTEL['primary']))
                        fig_a.add_trace(go.Bar(x=dfa['Fecha'], y=dfa['Interes'],
                            name='Interés', marker_color=C_PASTEL['accent']))
                        fig_a.add_trace(go.Scatter(x=dfa['Fecha'], y=dfa['Saldo_Fin'],
                            name='Saldo', mode='lines+markers',
                            line=dict(color=C_PASTEL['success'], width=2.5), yaxis='y2'))
                        if mes_hoy in list(dfa['Fecha'].values):
                            idx_hoy = list(dfa['Fecha'].values).index(mes_hoy)
                            fig_a.add_shape(type='line', x0=idx_hoy-0.5, x1=idx_hoy-0.5,
                                y0=0, y1=1, xref='x', yref='paper',
                                line=dict(color=C['gold'], width=2, dash='dash'))
                            fig_a.add_annotation(x=idx_hoy, y=1.04, xref='x', yref='paper',
                                text='Hoy', showarrow=False, font=dict(color=C['gold'], size=11))
                        if mes_corte_ec is not None and not dfa.empty:
                            idx_baja = len(dfa) - 1
                            fig_a.add_shape(type='line', x0=idx_baja+0.5, x1=idx_baja+0.5,
                                y0=0, y1=1, xref='x', yref='paper',
                                line=dict(color=C['accent'], width=2, dash='dash'))
                            fig_a.add_annotation(x=idx_baja, y=1.04, xref='x', yref='paper',
                                text='Baja', showarrow=False, font=dict(color=C['accent'], size=11))
                        fig_a.update_layout(
                            barmode='stack', yaxis2=dict(overlaying='y',side='right',showgrid=False),
                            legend=dict(orientation='h',y=1.06)
                        )
                        fig_a = sfig(fig_a, h=300)
                        st.plotly_chart(fig_a, width='stretch', key=f"pc_007_{sel_ec}")

                        cols_a = ['Fecha','Mes','Saldo_Ini','Interes','Capital','Saldo_Fin']
                        fmt_a  = {c:'${:,.4f}' for c in ['Saldo_Ini','Interes','Capital','Saldo_Fin']}
                        _styler_a = dfa[cols_a].style.format(fmt_a)
                        if mes_corte_ec is not None:
                            _styler_a = _styler_a.set_properties(
                                subset=pd.IndexSlice[[dfa.index[-1]], :],
                                **{'background-color': '#FAE3E1', 'color': '#8A2019', 'font-weight': '700'}
                            )
                        st.dataframe(_styler_a, width='stretch', height=300, key=f"df_008_{sel_ec}")

                        buf_ea = excel_con_formato({'Amort_Leasing': dfa[cols_a]},
                            currency_cols=['Saldo_Ini','Interes','Capital','Saldo_Fin'])
                        st.download_button("Descargar Excel", buf_ea, f"amort_{sel_ec}.xlsx", key=f"dl_amort_{sel_ec}")
                    else:
                        st.info("Inversión neta cero — sin tabla de amortización.")

            with tab_res:
                with st.container(key=f"tabbox_res_{sel_ec}"):
                    if inv > 0 and res > 0:
                        vpr = vp_res(res, t, pl)
                        dfr = calc_res_amort(round(vpr,4), t, pl).copy()
                        dfr['Fecha'] = dfr['Mes'].apply(lambda m: (fa + relativedelta(months=m)).strftime('%Y-%m'))
                        dfr['Residual_Pactado'] = res
                        dfr['VP_Residual']      = round(vpr, 4)

                        if mes_corte_ec is not None:
                            dfr = dfr[dfr['Mes'] <= mes_corte_ec].copy()
                            st.warning(
                                f"Este contrato se dio de **baja anticipada** el **{fecha_baja_ec.strftime('%Y-%m-%d')}** "
                                f"(mes {mes_corte_ec} de {pl} originalmente pactados). La tabla y la gráfica solo muestran "
                                "hasta ese punto."
                            )

                        pr1,pr2,pr3 = st.columns(3)
                        pr1.metric("VP del Residual",    f"${vpr:,.4f}")
                        pr2.metric("Residual pactado",   f"${res:,.2f}")
                        pr3.metric("Intereses residual", f"${dfr['Interes'].sum():,.4f}")

                        fig_r = go.Figure()
                        fig_r.add_trace(go.Scatter(
                            x=dfr['Fecha'], y=dfr['Saldo_Fin'],
                            name='Saldo acumulado', fill='tozeroy', mode='lines',
                            line=dict(color=C_PASTEL['gold'], width=2.5)
                        ))
                        fig_r.add_hline(y=res, line_dash='dash', line_color=C['accent'],
                            annotation_text=f"Residual pactado ${res:,.2f}")
                        fig_r.add_hline(y=vpr, line_dash='dot', line_color=C['info'],
                            annotation_text=f"VP ${vpr:,.4f}")
                        if mes_corte_ec is not None and not dfr.empty:
                            idx_baja_r = len(dfr) - 1
                            fig_r.add_shape(type='line', x0=idx_baja_r, x1=idx_baja_r,
                                y0=0, y1=1, xref='x', yref='paper',
                                line=dict(color=C['accent'], width=2, dash='dash'))
                            fig_r.add_annotation(x=idx_baja_r, y=1.04, xref='x', yref='paper',
                                text='Baja', showarrow=False, font=dict(color=C['accent'], size=11))
                        fig_r = sfig(fig_r, h=280)
                        st.plotly_chart(fig_r, width='stretch', key=f"pc_008_{sel_ec}")

                        cols_r = ['Fecha','Mes','VP_Residual','Saldo_Ini','Interes','Saldo_Fin','Residual_Pactado']
                        fmt_r  = {c:'${:,.4f}' for c in ['VP_Residual','Saldo_Ini','Interes','Saldo_Fin','Residual_Pactado']}
                        _styler_r = dfr[cols_r].style.format(fmt_r)
                        if mes_corte_ec is not None:
                            _styler_r = _styler_r.set_properties(
                                subset=pd.IndexSlice[[dfr.index[-1]], :],
                                **{'background-color': '#FAE3E1', 'color': '#8A2019', 'font-weight': '700'}
                            )
                        st.dataframe(_styler_r, width='stretch', height=300, key=f"df_009_{sel_ec}")
                    else:
                        st.info("Sin residual o inversión neta cero.")

            with tab_evts:
                with st.container(key=f"tabbox_evts_{sel_ec}"):
                    df_ev_c = obtener_eventos_especiales(id_c=sel_ec)
                    if df_ev_c.empty:
                        st.info("Sin eventos especiales registrados para este contrato.")
                    else:
                        for _, ev in df_ev_c.iterrows():
                            est  = str(ev.get('Estatus_Seguro','') or '')
                            col_e = C['success'] if est=='CONFIRMADO' else (C['accent'] if est=='PENDIENTE' else C['warning'])
                            with st.container(key=f"evtcard_{sel_ec}_{int(ev['id'])}"):
                                st.markdown(f"""
                                <div style="border:1px solid #DCE0E5;border-left:4px solid {col_e};background:#FFFFFF;border-radius:0 4px 4px 0;
                                            padding:10px 14px;margin-bottom:8px;">
                                  <b>{ev['Tipo_Evento']}</b>
                                  <span style="float:right;background:{col_e};color:#fff;border-radius:4px;
                                               padding:1px 8px;font-size:.75rem;">{est}</span><br>
                                  <span style="font-size:.8rem;color:#565E68;">Fecha evento: {ev['Fecha_Evento']}
                                  &nbsp;·&nbsp; Registrado: {ev['Fecha_Registro']}</span><br>
                                  <span style="font-size:.82rem;color:#20242B;">
                                    Valor recuperable: <b>${float(ev.get('Valor_Recuperable',0)):,.2f}</b>
                                    &nbsp;·&nbsp; Reclamado: <b>${float(ev.get('Monto_Reclamado',0)):,.2f}</b>
                                    &nbsp;·&nbsp; Confirmado: <b>${float(ev.get('Monto_Confirmado',0)):,.2f}</b>
                                  </span>
                                  {"<br><span style='font-size:.78rem;color:#8A929C;'>"+str(ev.get('Observaciones',''))+"</span>" if ev.get('Observaciones') else ""}
                                </div>
                                """, unsafe_allow_html=True)

            with tab_anot:
                with st.container(key=f"tabbox_anot_{sel_ec}"):
                    df_an_c = obtener_anotaciones(id_c=sel_ec)
                    TIPO_COLORS_EC = {"General":"#3E6FA6","Ajuste contable":"#B3261E","Nota legal":"#6E5A9C",
                                      "Seguimiento":"#1E5C4F","Alerta":"#96660C","Acuerdo con cliente":"#1C7A4D","Otro":"#8A6D2F"}

                    with st.expander("Agregar anotación", expanded=df_an_c.empty):
                        TIPOS_AN = ["General","Ajuste contable","Nota legal","Seguimiento","Alerta","Acuerdo con cliente","Otro"]
                        ec_an1, ec_an2 = st.columns([3,1])
                        ec_txt  = ec_an1.text_area("Texto", height=70, key="ec_anot_txt", placeholder="Escribe aquí…")
                        ec_tipo = ec_an2.selectbox("Tipo", TIPOS_AN, key="ec_anot_tipo")
                        if st.button("Guardar", key="ec_anot_save"):
                            if ec_txt.strip():
                                agregar_anotacion(sel_ec, ec_txt.strip(), ec_tipo)
                                st.success("Anotación guardada.")
                                st.session_state['_refresh'] = True

                    if df_an_c.empty:
                        st.info("Sin anotaciones para este contrato.")
                    else:
                        df_an_c_s = df_an_c.sort_values('Fecha', ascending=False)
                        for _, an in df_an_c_s.iterrows():
                            tipo_an = str(an.get('Tipo','General') or 'General')
                            col_an  = TIPO_COLORS_EC.get(tipo_an,'#0369A1')
                            confirm_del_ec = (st.session_state.get('anot_confirm_del') == int(an['id']))
                            with st.container(key=f"anotcard_{sel_ec}_{int(an['id'])}"):
                                st.markdown(f"""
                                <div style="border:1px solid #DCE0E5;border-left:4px solid {col_an};background:#FFFFFF;border-radius:0 4px 4px 0;
                                            padding:10px 14px;margin-bottom:6px;">
                                  <span style="font-weight:700;color:{col_an};font-size:.82rem;">{tipo_an}</span>
                                  <span style="float:right;font-size:.75rem;color:#8A929C;">{an['Fecha']}</span><br>
                                  <span style="font-size:.88rem;color:#20242B;white-space:pre-wrap;">{an['Texto']}</span>
                                </div>
                                """, unsafe_allow_html=True)
                                bc1_ec, bc2_ec, _ = st.columns([1,1,5])
                                if confirm_del_ec:
                                    bc1_ec.warning("¿Borrar?")
                                    if bc2_ec.button("Sí", key=f"ec_del_ok_{an['id']}", width='stretch'):
                                        eliminar_anotacion(int(an['id']))
                                        st.session_state['anot_confirm_del'] = None
                                        st.session_state['_refresh'] = True
                                    if _.button("No", key=f"ec_del_no_{an['id']}", width='stretch'):
                                        st.session_state['anot_confirm_del'] = None
                                        st.session_state['_refresh'] = True
                                else:
                                    if bc2_ec.button("Eliminar", key=f"ec_del_{an['id']}", width='stretch'):
                                        st.session_state['anot_confirm_del'] = int(an['id'])
                                        st.session_state['_refresh'] = True

            with tab_cont:
                with st.container(key=f"tabbox_cont_{sel_ec}"):
                    st.markdown('<span class="section-label">Composición financiera del contrato</span>', unsafe_allow_html=True)
                    if inv > 0:
                        _dfa_c, _icp_c, _ilp_c, _r12_c, _rresto_c = calc_amort(round(inv,4), r_m, res, pl, t)
                        _cap_total = float(_dfa_c['Capital'].sum())
                        _int_total = _icp_c + _ilp_c
                        fc1,fc2,fc3,fc4 = st.columns(4)
                        fc1.metric("Capital a recuperar (plazo)", f"${_cap_total:,.2f}")
                        fc2.metric("Interés total del plazo",     f"${_int_total:,.2f}")
                        fc3.metric("Interés primeros 12 meses",   f"${_icp_c:,.2f}")
                        fc4.metric("Interés resto del plazo",     f"${_ilp_c:,.2f}")
                        st.caption("Capital + interés + residual pactado deben cuadrar contra la suma de rentas más el valor residual del contrato.")
                    else:
                        st.info("No se puede calcular la composición: inversión neta no positiva.")

                    st.markdown("---")
                    st.markdown('<span class="section-label">Cuentas contables asignadas a este contrato</span>', unsafe_allow_html=True)
                    st.caption(
                        "Estas son, cuenta por cuenta, las claves de tu catálogo maestro que aplican a este "
                        "contrato para Contpaqi. Las que dejaste en blanco en 'Cuentas (Macro)' no aparecen "
                        "aquí ni se exportan."
                    )
                    cat_ec = cargar_catalogo()
                    filas_cat = []
                    for k_ec, d_ec in cat_ec.items():
                        if k_ec == 'CANCELACION_ANTICIPO':
                            continue
                        cuenta_base = str(d_ec.get('cuenta') or '').strip()
                        if not cuenta_base:
                            continue
                        filas_cat.append({
                            'Clave': k_ec,
                            'Concepto': d_ec['nombre'],
                            'Cuenta Contpaqi (detalle)': fmt17(cuenta_base, sel_ec),
                        })
                    if filas_cat:
                        st.dataframe(pd.DataFrame(filas_cat), width='stretch', height=280, key=f"df_cont_{sel_ec}")
                    else:
                        st.warning("No hay cuentas activas en el catálogo maestro. Ve a 'Cuentas (Macro)' y captura al menos una.")

                    st.markdown("---")
                    st.markdown('<span class="section-label">Otros contratos del mismo cliente</span>', unsafe_allow_html=True)
                    _otros = df_todos[(df_todos['Cliente']==row['Cliente']) & (df_todos['ID_Contrato']!=sel_ec)]
                    if not _otros.empty:
                        st.dataframe(
                            _otros[['ID_Contrato','Vehiculo','Estatus','Plazo','Mensualidad_Sin_IVA','Nivel_Morosidad']]
                                .rename(columns={'Mensualidad_Sin_IVA':'Renta mensual'}),
                            width='stretch', height=180, key=f"df_otros_{sel_ec}"
                        )
                        st.caption("Útil para ver la exposición total de este cliente antes de aprobar un nuevo contrato o negociar una renovación.")
                    else:
                        st.caption("Este cliente no tiene otros contratos registrados en esta empresa.")

            st.markdown("---")
            if st.button("Exportar estado de cuenta completo (Excel)", width='stretch'):
                _hojas_ec = {}
                _hojas_ec['Datos Generales'] = pd.DataFrame([{
                    'Campo': k, 'Valor': str(row.get(k,''))
                } for k in ['ID_Contrato','Cliente','Vehiculo','Estatus','Fecha_Alta',
                             'Fecha_Vencimiento','Plazo','Valor_Sin_IVA','Anticipo_Monto',
                             'Mensualidad_Sin_IVA','Residual_Monto','Tasa_Calculada',
                             'Nivel_Morosidad','Fecha_Excl_Poliza','Motivo_Excl_Poliza']
                ])
                if inv > 0:
                    dfa_exp, _, _, _, _ = calc_amort(round(inv,4), r_m, res, pl, t)
                    dfa_exp['Fecha'] = dfa_exp['Mes'].apply(lambda m: (fa+relativedelta(months=m)).strftime('%Y-%m'))
                    _hojas_ec['Amort Leasing'] = dfa_exp
                if inv > 0 and res > 0:
                    vpr_e = vp_res(res, t, pl)
                    dfr_e = calc_res_amort(round(vpr_e,4), t, pl)
                    dfr_e['Fecha'] = dfr_e['Mes'].apply(lambda m: (fa+relativedelta(months=m)).strftime('%Y-%m'))
                    _hojas_ec['Acum Residual'] = dfr_e
                df_an_exp = obtener_anotaciones(id_c=sel_ec)
                if not df_an_exp.empty:
                    _hojas_ec['Anotaciones'] = df_an_exp
                df_ev_exp = obtener_eventos_especiales(id_c=sel_ec)
                if not df_ev_exp.empty:
                    _hojas_ec['Eventos Especiales'] = df_ev_exp
                _hojas_ec['Resumen Financiero'] = pd.DataFrame([
                    {'Concepto':'Próximo pago','Monto':txt_prox},
                    {'Concepto':'Pagos realizados','Monto':meses_transcurridos},
                    {'Concepto':'Pagos restantes','Monto':pagos_restantes},
                    {'Concepto':'Saldo insoluto (capital)','Monto':saldo_insoluto if saldo_insoluto is not None else ''},
                    {'Concepto':'Utilidad (interés) devengada a la fecha','Monto':utilidad_devengada if utilidad_devengada is not None else ''},
                ])
                if filas_cat:
                    _hojas_ec['Cuentas Contpaqi'] = pd.DataFrame(filas_cat)

                buf_ec = excel_con_formato(
                    _hojas_ec,
                    currency_cols={
                        'Amort Leasing': ['Saldo_Ini','Interes','Capital','Saldo'],
                        'Acum Residual': ['VP_Residual','Saldo_Ini','Interes','Saldo_Fin','Residual_Pactado'],
                        'Eventos Especiales': ['Valor_Recuperable','Monto_Reclamado','Monto_Confirmado'],
                    },
                )
                st.download_button(
                    "Descargar Excel completo",
                    buf_ec,
                    f"estado_cuenta_{sel_ec}.xlsx",
                    key="ec_download"
                )

            if inv > 0:
                with st.expander("Exportar a PDF", expanded=False):
                    pdf_key = f"pdf_buf_{sel_ec}"
                    if st.button("Generar Documento PDF", width='stretch', key="ec_pdf_gen_btn"):
                        with st.spinner("Creando PDF... (puede tomar un par de segundos)"):
                            dfa_pdf, _, _, _, _ = calc_amort(round(inv,4), r_m, res, pl, t)
                            dfa_pdf = dfa_pdf.copy()
                            dfa_pdf['Fecha'] = dfa_pdf['Mes'].apply(lambda m: (fa + relativedelta(months=m)).strftime('%Y-%m'))
                            saldos_ini_pdf = [round(inv,4)] + list(dfa_pdf['Saldo'].iloc[:-1].round(4))
                            dfa_pdf.insert(dfa_pdf.columns.get_loc('Interes'), 'Saldo_Ini', saldos_ini_pdf)
                            dfa_pdf.rename(columns={'Saldo':'Saldo_Fin'}, inplace=True)
                            # Se genera el PDF completo siempre a peticion del usuario
                            dfr_pdf = None
                            if res > 0:
                                vpr_pdf = vp_res(res, t, pl)
                                dfr_pdf = calc_res_amort(round(vpr_pdf,4), t, pl).copy()
                                dfr_pdf['Fecha'] = dfr_pdf['Mes'].apply(lambda m: (fa + relativedelta(months=m)).strftime('%Y-%m'))
                                dfr_pdf['Residual_Pactado'] = res
                                dfr_pdf['VP_Residual']      = round(vpr_pdf, 4)
                                # Se genera el PDF completo siempre a peticion del usuario
                            _avp_pdf = calcular_avance_pago(row) if str(row.get('Estatus','')).upper() != 'BAJA' else None
                            pdf_buf = pdf_estado_cuenta(
                                row, dfa_pdf, dfr=dfr_pdf, avance=_avp_pdf, mes_corte=mes_corte_ec,
                                fecha_hoy_txt=fecha_larga(hoy_ref())
                            )
                            st.session_state[pdf_key] = pdf_buf

                    if pdf_key in st.session_state:
                        st.download_button(
                            "Descargar PDF Ahora",
                            st.session_state[pdf_key],
                            f"estado_cuenta_{sel_ec}.pdf",
                            mime="application/pdf",
                            key=f"ec_pdf_download_{sel_ec}"
                        )


