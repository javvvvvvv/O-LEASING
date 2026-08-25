from matplotlib.colors import LinearSegmentedColormap
CMAP_INDIGO = LinearSegmentedColormap.from_list("ledger_indigo", ["#161B2200", "#2F81F766", "#2f81f7"])
CMAP_CORAL  = LinearSegmentedColormap.from_list("ledger_coral",  ["#161B2200", "#DA363366", "#da3633"])
CMAP_TEAL   = LinearSegmentedColormap.from_list("ledger_teal",   ["#161B2200", "#23863666", "#238636"])

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
import numpy as np

PAL_PASTEL = ["#5E9587","#7C93AD","#D9A75C","#C4796C","#7FA8C9","#9C8AA5","#B49A5E","#C79AAA","#84AD79","#6E86AE"]

def sfig(fig, title=None, h=300):

    if fig is None:
        import plotly.graph_objects as go
        fig = go.Figure()
        fig.add_annotation(text="Sin datos para graficar", showarrow=False, font=dict(size=14, color="#8b949e"))
        fig.update_xaxes(visible=False)
        fig.update_yaxes(visible=False)
    if not hasattr(fig, 'update_layout'):
        return fig

    layout = dict(
        font=dict(family="Segoe UI, Helvetica Neue, Arial, sans-serif", color="#8b949e", size=12),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(t=56, r=18, b=42, l=18),
        hoverlabel=dict(bgcolor="#1E5C4F", font_color="#fff", bordercolor="#1E5C4F",
                        font_size=13, font_family="Segoe UI, Helvetica Neue, Arial, sans-serif",
                        align="left", namelength=-1),
        xaxis=dict(gridcolor="rgba(255,255,255,.05)", linecolor="rgba(255,255,255,.1)",
                   showspikes=True, spikethickness=1,
                   spikecolor="rgba(47,129,247,.35)", spikedash="dot"),
        yaxis=dict(gridcolor="rgba(255,255,255,.05)", linecolor="rgba(255,255,255,.1)"),
        height=h,
        bargap=0.22, bargroupgap=0.09,
        colorway=PAL_PASTEL,
        legend=dict(bgcolor="rgba(255,255,255,0)", bordercolor="rgba(255,255,255,.1)",
                    borderwidth=0, font=dict(size=11)),
    )
    _title = title
    if not _title:
        _title = fig.layout.title.text if fig.layout.title and fig.layout.title.text else None
    if _title:
        layout['title'] = dict(
            text=f"<span style='color:#2f81f7;'>●</span>&nbsp; <b>{_title}</b>",
            font=dict(size=14.5, color="#e6edf3"),
            x=0.012, xanchor='left', y=0.97, yanchor='top'
        )
    fig.update_layout(**layout)
    # Barras: esquinas rectas, sin relleno decorativo. Las líneas van rectas
    # de un dato al siguiente (nada de curvas spline) — en un reporte
    # financiero, suavizar la línea puede sugerir una tendencia que los
    # datos reales no tienen.
    fig.update_traces(
        selector=dict(type='bar'),
        marker_cornerradius=2,
    )
    fig.update_traces(
        selector=dict(type='scatter'),
        line_shape='linear',
        marker=dict(line=dict(width=1, color="#FFFFFF"), size=6),
    )
    fig.update_traces(
        selector=dict(type='pie'),
        marker=dict(line=dict(color="#F3F4F6", width=2)),
        textfont=dict(size=12, family="Segoe UI, Helvetica Neue, Arial, sans-serif"),
        rotation=30,
    )
    return fig


def explain(titulo, texto):
    st.markdown(f"""<div class="chart-explain">
      <span class="ce-title">{titulo}</span><br>
      <span class="ce-body">{texto}</span>
    </div>""", unsafe_allow_html=True)


def estado_vacio(titulo, mensaje):
    """Tarjeta centrada para pantallas sin datos — en vez de un st.info plano
    que se siente como un error a medias, deja claro que la pantalla cargó
    bien y solo no hay nada que mostrar todavía."""
    st.markdown(f"""<div class="empty-state">
      <span class="es-title">{titulo}</span>
      <span class="es-msg">{mensaje}</span>
    </div>""", unsafe_allow_html=True)


def sz(series):
    return pd.to_numeric(series, errors='coerce').fillna(1).replace([np.inf,-np.inf],1).clip(lower=0.1)


def titled_chart(title, fig, use_container_width=True):
    st.markdown(f'<span class="section-label">{title}</span>', unsafe_allow_html=True)
    st.plotly_chart(fig, width=('stretch' if use_container_width else 'content'), key="pc_001")


def titled_table(title, df_show, fmt_dict=None, cmap_col=None, key=None):
    st.markdown(f'<span class="section-label">{title}</span>', unsafe_allow_html=True)
    styled = df_show.style
    if fmt_dict:
        safe_fmt = {k:v for k,v in fmt_dict.items() if k in df_show.columns}
        if safe_fmt: styled = styled.format(safe_fmt, na_rep="-")
    if cmap_col and cmap_col in df_show.columns:
        styled = styled.background_gradient(subset=[cmap_col], cmap=CMAP_INDIGO)
    st.dataframe(styled, width='stretch', key="df_001")


