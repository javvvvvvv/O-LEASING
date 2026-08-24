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
import streamlit as st
import pandas as pd

def sfig(fig, title=None, h=300):
    fig.update_layout(
        height=h, margin=dict(l=10, r=10, t=30 if title else 10, b=10),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        title_text=title, title_font=dict(size=14, color="#888")
    )
    st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})

def explain(titulo: str, texto: str):
    st.markdown(f"""
        <div style='background-color: var(--secondary-background-color); padding: 12px; border-radius: 8px; margin-bottom: 15px;'>
            <strong>{titulo}</strong><br>
            <span style='font-size: 0.9em; color: var(--text-color); opacity: 0.8;'>{texto}</span>
        </div>
    """, unsafe_allow_html=True)

def estado_vacio(titulo: str, mensaje: str):
    st.markdown(f"""
        <div style='text-align: center; padding: 40px 20px; background-color: var(--secondary-background-color); border-radius: 10px; margin: 20px 0;'>
            <h4 style='color: var(--text-color); margin-bottom: 10px;'>{titulo}</h4>
            <p style='color: var(--text-color); opacity: 0.7; font-size: 1.1em;'>{mensaje}</p>
        </div>
    """, unsafe_allow_html=True)

def sz(series: pd.Series) -> pd.Series:
    return series.fillna(0)

def titled_chart(title: str, fig, use_container_width: bool = True):
    st.markdown(f"<div class='section-title'>{title}</div>", unsafe_allow_html=True)
    st.plotly_chart(fig, use_container_width=use_container_width, config={'displayModeBar': False})

def titled_table(title: str, df_show: pd.DataFrame, fmt_dict: dict = None, cmap_col: str = None, key: str = None):
    st.markdown(f"<div class='section-title'>{title}</div>", unsafe_allow_html=True)
    if df_show.empty:
        st.caption("No hay datos.")
    else:
        st.dataframe(df_show, use_container_width=True, hide_index=True)
