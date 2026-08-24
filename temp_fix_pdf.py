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
import sys

with open('app.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

start_idx = -1
end_idx = -1
for i, line in enumerate(lines):
    if 'if st.button("Generar PDF para imprimir' in line and 'ec_pdf_btn' in line:
        start_idx = i
    if 'key="ec_pdf_download"' in line and start_idx != -1:
        end_idx = i + 1
        break

if start_idx != -1 and end_idx != -1:
    replacement = """                with st.expander("🖨️ Exportar a PDF", expanded=False):
                    pdf_key = f"pdf_buf_{sel_ec}"
                    if st.button("Generar Documento PDF", width='stretch', key="ec_pdf_gen_btn"):
                        with st.spinner("Creando PDF... (puede tomar un par de segundos)"):
                            dfa_pdf, _, _, _, _ = calc_amort(round(inv,4), r_m, res, pl, t)
                            dfa_pdf = dfa_pdf.copy()
                            dfa_pdf['Fecha'] = dfa_pdf['Mes'].apply(lambda m: (fa + relativedelta(months=m)).strftime('%Y-%m'))
                            saldos_ini_pdf = [round(inv,4)] + list(dfa_pdf['Saldo'].iloc[:-1].round(4))
                            dfa_pdf.insert(dfa_pdf.columns.get_loc('Interes'), 'Saldo_Ini', saldos_ini_pdf)
                            dfa_pdf.rename(columns={'Saldo':'Saldo_Fin'}, inplace=True)
                            if mes_corte_ec is not None:
                                dfa_pdf = dfa_pdf[dfa_pdf['Mes'] <= mes_corte_ec].copy()
                            dfr_pdf = None
                            if res > 0:
                                vpr_pdf = vp_res(res, t, pl)
                                dfr_pdf = calc_res_amort(round(vpr_pdf,4), t, pl).copy()
                                dfr_pdf['Fecha'] = dfr_pdf['Mes'].apply(lambda m: (fa + relativedelta(months=m)).strftime('%Y-%m'))
                                dfr_pdf['Residual_Pactado'] = res
                                dfr_pdf['VP_Residual']      = round(vpr_pdf, 4)
                                if mes_corte_ec is not None:
                                    dfr_pdf = dfr_pdf[dfr_pdf['Mes'] <= mes_corte_ec].copy()
                            _avp_pdf = calcular_avance_pago(row) if str(row.get('Estatus','')).upper() != 'BAJA' else None
                            pdf_buf = pdf_estado_cuenta(
                                row, dfa_pdf, dfr=dfr_pdf, avance=_avp_pdf, mes_corte=mes_corte_ec,
                                fecha_hoy_txt=fecha_larga(hoy_ref())
                            )
                            st.session_state[pdf_key] = pdf_buf

                    if pdf_key in st.session_state:
                        st.download_button(
                            "📥 Descargar PDF Ahora",
                            st.session_state[pdf_key],
                            f"estado_cuenta_{sel_ec}.pdf",
                            mime="application/pdf",
                            key=f"ec_pdf_download_{sel_ec}"
                        )
"""
    lines = lines[:start_idx] + [replacement] + lines[end_idx+1:]
    with open('app.py', 'w', encoding='utf-8') as f:
        f.writelines(lines)
    print("PDF generation logic updated successfully")
else:
    print("Could not find PDF block")
