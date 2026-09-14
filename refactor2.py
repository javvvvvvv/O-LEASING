import re

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

old_alerts_code = """        if alertas:
            st.markdown("#### Alertas del sistema")
            for _idx_alerta, (tipo, msg, detalle) in enumerate(alertas):
                if tipo == 'error':
                    st.error(msg)
                elif tipo == 'warning':
                    st.warning(msg)
                else:
                    st.info(msg)
                if detalle is not None:
                    with st.expander("Ver detalle", key=f"exp_alerta_{_idx_alerta}"):
                        st.dataframe(detalle, width='stretch', key=f"df_006_{_idx_alerta}")
            st.markdown("---")"""

new_alerts_code = """        if alertas:
            # Reorganizar visualmente como "Centro de Atención" (D.1)
            # Priorizamos 'error' (datos imposibles), luego 'warning', luego 'info'
            orden_severidad = {'error': 0, 'warning': 1, 'info': 2}
            alertas.sort(key=lambda x: orden_severidad.get(x[0], 99))
            
            st.markdown(f'''
                <div style="display:flex; justify-content:space-between; align-items:baseline; margin-bottom:10px;">
                    <h3 style="margin:0; color:var(--text-primary) !important;">Centro de Atención</h3>
                    <span style="color:var(--text-secondary); font-size:14px; font-weight:600;">{len(alertas)} pendientes</span>
                </div>
            ''', unsafe_allow_html=True)
            
            for _idx_alerta, (tipo, msg, detalle) in enumerate(alertas):
                if tipo == 'error':
                    icono = "🔴"
                    bg = "var(--danger-soft)"
                    color = "var(--danger)"
                elif tipo == 'warning':
                    icono = "🟡"
                    bg = "var(--warning-soft)"
                    color = "var(--warning)"
                else:
                    icono = "🔵"
                    bg = "var(--info-soft)"
                    color = "var(--info)"
                
                with st.container(key=f"ca_{_idx_alerta}"):
                    st.markdown(f'''
                        <div style="background:{bg}; border-left:4px solid {color}; border-radius:4px; padding:12px; margin-bottom:10px;">
                            <span style="font-size:16px; margin-right:8px;">{icono}</span>
                            <span style="font-weight:500; color:var(--text-primary);">{msg}</span>
                        </div>
                    ''', unsafe_allow_html=True)
                    if detalle is not None:
                        with st.expander("Ver detalle", key=f"exp_alerta_{_idx_alerta}"):
                            st.dataframe(detalle, width='stretch', key=f"df_006_{_idx_alerta}")
            st.markdown("---")"""

if old_alerts_code in content:
    content = content.replace(old_alerts_code, new_alerts_code)
    with open('app.py', 'w', encoding='utf-8') as f:
        f.write(content)
    print("Centro de Atención reemplazado.")
else:
    print("No se encontró el bloque de código de alertas viejo.")
