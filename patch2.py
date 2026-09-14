import re

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Update the HTML grid in Estado de Cuenta
old_grid_pattern = re.compile(r'st\.markdown\(f"""\s*<div style="border:1px solid #DCE0E5;border-left:4px solid {col_estado};background:#FFFFFF;.*?</div>\s*""", unsafe_allow_html=True\)', re.DOTALL)

new_grid = """st.markdown(f'''
<div style="background:var(--bg-1); border:1px solid var(--border); border-left:4px solid {col_estado};
            padding:20px; border-radius:var(--radius-md); margin-bottom:20px;">
  <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px; border-bottom:1px solid var(--border); padding-bottom:12px;">
    <div>
        <span class="label-caps">Cliente</span><br>
        <strong style="font-size:1.1rem; color:var(--text-primary);">{row.get('Cliente','—')}</strong>
    </div>
    <div style="text-align:right;">
        <span class="label-caps">Contrato</span><br>
        <strong style="font-size:1.1rem; color:var(--brand);">{row.get('ID_Contrato','—')}</strong>
    </div>
  </div>
  
  <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(140px, 1fr)); gap:16px; font-size:13px; color:var(--text-secondary);">
    <div><b>Vehículo</b><br><span style="color:var(--text-primary);">{row.get('Vehiculo','—')}</span></div>
    <div><b>Fecha Alta</b><br><span style="color:var(--text-primary);">{str(row.get('Fecha_Alta','—'))[:10]}</span></div>
    {_campo_vencimiento}
    <div><b>Estatus</b><br><span style="color:{col_estado}; font-weight:600;">{row.get('Estatus','—')}</span></div>
    
    <div><b>Plazo</b><br><span style="color:var(--text-primary);">{int(row.get('Plazo',0))} meses</span></div>
    <div><b>Valor (s/IVA)</b><br><span style="color:var(--text-primary);">${float(row.get('Valor_Sin_IVA',0)):,.2f}</span></div>
    <div><b>Anticipo</b><br><span style="color:var(--text-primary);">${float(row.get('Anticipo_Monto',0)):,.2f} ({float(row.get('Anticipo_Pct',0)):.0f}%)</span></div>
    
    <div><b>Renta (s/IVA)</b><br><span style="color:var(--text-primary);">${float(row.get('Mensualidad_Sin_IVA',0)):,.2f}</span></div>
    <div><b>Residual</b><br><span style="color:var(--text-primary);">${float(row.get('Residual_Monto',0)):,.2f}</span></div>
    <div><b>Tasa Anual</b><br><span style="color:var(--text-primary);">{float(row.get('Tasa_Calculada',0))*1200:.2f}%</span></div>
  </div>
  {('<div style="margin-top:12px; padding:10px; background:var(--danger-soft); border-left:3px solid var(--danger); border-radius:4px; font-size:13px; color:var(--danger);"><b>Motivo exclusión:</b> ' + motivo_excl + '</div>') if motivo_excl else ""}
</div>
''', unsafe_allow_html=True)"""

content = old_grid_pattern.sub(new_grid, content)

# 2. Update the columns for metrics (from 5 to 4 to give space)
old_km = r'km1,km2,km3,km4,km5 = st\.columns\(5\)\s*km1\.metric\("Inversi.*?km5\.metric\("Mes punto equil.*?neta\."\)'
new_km = """km1,km2,km3,km4 = st.columns(4)
                km1.metric("Inversión neta", f"${inv_neta:,.2f}")
                km2.metric("Ganancia proy.", f"${g:,.2f}")
                km3.metric("Margen total",   f"{m:.1f}%")
                km4.metric("Punto Equil.",   f"Mes {mes_pe}" if mes_pe else "N/A", help="Mes en que se recupera la inversión")"""
content = re.sub(old_km, new_km, content, flags=re.DOTALL)

# 3. Update the leasing chart (fig_a) to adjust to plazo
old_fig_a = r"fig_a = go\.Figure\(\).*?fig_a = sfig\(fig_a, h=300\)\s*st\.plotly_chart\(fig_a, width='stretch', key=f\"pc_007_\{sel_ec\}\"\)"

new_fig_a = """fig_a = go.Figure()
                        fig_a.add_trace(go.Bar(x=dfa['Fecha'], y=dfa['Capital'], name='Capital', marker_color=C_PASTEL['primary']))
                        fig_a.add_trace(go.Bar(x=dfa['Fecha'], y=dfa['Interes'], name='Interés', marker_color=C_PASTEL['accent']))
                        fig_a.add_trace(go.Scatter(x=dfa['Fecha'], y=dfa['Saldo_Fin'], name='Saldo', mode='lines+markers',
                            line=dict(color=C_PASTEL['success'], width=2.5), yaxis='y2'))
                        
                        if mes_hoy in list(dfa['Fecha'].values):
                            idx_hoy = list(dfa['Fecha'].values).index(mes_hoy)
                            fig_a.add_shape(type='line', x0=idx_hoy-0.5, x1=idx_hoy-0.5, y0=0, y1=1, xref='x', yref='paper',
                                line=dict(color=C['gold'], width=2, dash='dash'))
                            fig_a.add_annotation(x=idx_hoy, y=1.04, xref='x', yref='paper', text='Hoy', showarrow=False, font=dict(color=C['gold'], size=11))
                        
                        if mes_corte_ec is not None and not dfa.empty:
                            idx_baja = len(dfa) - 1
                            fig_a.add_shape(type='line', x0=idx_baja+0.5, x1=idx_baja+0.5, y0=0, y1=1, xref='x', yref='paper',
                                line=dict(color=C['accent'], width=2, dash='dash'))
                            fig_a.add_annotation(x=idx_baja, y=1.04, xref='x', yref='paper', text='Baja', showarrow=False, font=dict(color=C['accent'], size=11))
                        
                        # Fix X-axis to display months cleanly without omitting ticks if possible, and adjust bar width
                        fig_a.update_layout(
                            barmode='stack', 
                            yaxis2=dict(overlaying='y', side='right', showgrid=False),
                            legend=dict(orientation='h', y=1.06, x=0),
                            xaxis=dict(
                                type='category', 
                                tickangle=-45,
                                dtick=1 if len(dfa) <= 36 else 2 # Adapt to plazo
                            ),
                            margin=dict(l=10, r=10, t=30, b=10)
                        )
                        fig_a = sfig(fig_a, h=350) # Make slightly taller to fit labels
                        st.plotly_chart(fig_a, width='stretch', key=f"pc_007_{sel_ec}", use_container_width=True)"""

content = re.sub(old_fig_a, new_fig_a, content, flags=re.DOTALL)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)

print("Estado de cuenta actualizado")
