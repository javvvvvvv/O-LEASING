import streamlit as st

def inyectar_css():
    st.markdown("""
<style>
/* 
================================================================================
  SISTEMA DE DISEÑO O-LEASING v11
================================================================================
*/

@font-face {
  font-family: 'Inter';
  font-style: normal;
  font-weight: 100 900;
  font-display: swap;
  src: local('Inter'), url('https://fonts.gstatic.com/s/inter/v12/UcCO3FwrK3iLTeHuS_fvQtMwCp50KnMw2boKoduKmMEVuLyfMZhrib2Bg-4.ttf') format('truetype');
}

:root {
  /* Elevaciones y Fondos */
  --bg-0: #0D1117;
  --bg-1: #161B22;
  --bg-2: #1C222B;
  --bg-3: #232A34;
  
  /* Texto */
  --text-primary: #E6EDF3;
  --text-secondary: #8B949E;
  --text-disabled: #4B5563;
  --text-link: #FF8A4D;
  
  /* Marca e Interacción */
  --brand: #FF6619;
  --brand-hover: #E85A15;
  --brand-soft: rgba(255, 102, 25, 0.12);
  --brand-border: rgba(255, 102, 25, 0.4);
  
  /* Semánticos */
  --success: #238636;
  --success-soft: rgba(35, 134, 54, 0.12);
  --warning: #D29922;
  --warning-soft: rgba(210, 153, 34, 0.12);
  --danger: #DA3633;
  --danger-soft: rgba(218, 54, 51, 0.12);
  --info: #2F81F7;
  --info-soft: rgba(47, 129, 247, 0.12);
  
  /* Bordes */
  --border: rgba(255, 255, 255, 0.08);
  --border-hover: rgba(255, 255, 255, 0.16);
  --border-focus: var(--brand-border);
  
  /* Radios */
  --radius-sm: 6px;
  --radius-md: 10px;
  --radius-lg: 16px;
  
  /* Tipografía */
  --app-font: "Inter", "Segoe UI", -apple-system, sans-serif;
  
  /* Transiciones */
  --ease-micro: 120ms cubic-bezier(.2,.8,.2,1);
  --ease-std: 200ms cubic-bezier(.2,.8,.2,1);
  --ease-panel: 280ms cubic-bezier(.2,.8,.2,1);
}

/* =========================================================
   BASE & NUCLEO STREAMLIT
   ========================================================= */

.stApp {
  background-color: var(--bg-0) !important;
  font-family: var(--app-font);
  color: var(--text-primary);
}

.main .block-container { 
    padding: 2rem !important; 
    max-width: 1440px !important; /* Limitar estiramiento en monitores grandes */
    overflow-x: hidden !important; 
}

/* Typography Base */
h1 { color: var(--text-primary) !important; font-weight: 600 !important; font-size: 22px !important; line-height: 28px !important; margin-bottom: 0.5rem !important; }
h2 { color: var(--text-primary) !important; font-weight: 600 !important; font-size: 16px !important; line-height: 22px !important; }
h3 { color: var(--text-primary) !important; font-weight: 600 !important; font-size: 14px !important; line-height: 20px !important; }
.main p, .main .stMarkdown p { color: var(--text-secondary) !important; font-size: 14px; line-height: 20px; }
a { color: var(--text-link) !important; }

/* Tabular Nums para datos financieros y Ajuste de Tamaños */
[data-testid="stMetricValue"], [data-testid="stMetricDelta"], .stDataFrame, table, .tabular-nums { 
    font-variant-numeric: tabular-nums; 
}
[data-testid="stMetricValue"] {
    font-size: 1.4rem !important; /* Reducir el tamaño para que no se corten */
    word-wrap: break-word !important;
}
[data-testid="stMetricLabel"] {
    font-size: 0.85rem !important;
    white-space: normal !important; /* Permitir que los títulos largos usen 2 líneas */
    overflow: visible !important;
}

/* =========================================================
   SIDEBAR & TOP-BAR (Los únicos con Blur)
   ========================================================= */

[data-testid="stSidebar"] {
  background: rgba(22, 27, 34, 0.75) !important;
  backdrop-filter: blur(12px);
  -webkit-backdrop-filter: blur(12px);
  border-right: 1px solid var(--border) !important;
}

.top-header-bar {
  position: sticky; top: 0; z-index: 999;
  background: rgba(22, 27, 34, 0.75);
  backdrop-filter: blur(10px); -webkit-backdrop-filter: blur(10px);
  border: 1px solid var(--border); 
  border-radius: var(--radius-md);
  padding: .6rem 1rem; margin-bottom: 1rem;
  display: flex; align-items: center; justify-content: space-between; gap: 10px; flex-wrap: wrap;
  box-shadow: 0 4px 15px rgba(0,0,0,0.3);
}

.top-header-bar .thb-empresa { color: var(--text-primary) !important; font-weight: 600; }
.top-header-bar .thb-tag { display: block; font-size: 12px; color: var(--text-secondary) !important; text-transform: uppercase; letter-spacing: 0.04em; }

/* =========================================================
   CONTENEDORES & TARJETAS
   ========================================================= */

.section-header { 
    background: var(--bg-1); 
    border: 1px solid var(--border); 
    border-left: 3px solid var(--brand); 
    border-radius: var(--radius-md); 
    padding: 10px 16px; 
    margin: 12px 0 16px 0; 
}
.section-header h3 { color: var(--brand) !important; }

/* Cards (st.metric, form, expander, etc) */
div[data-testid="metric-container"], 
.stExpander,
div[data-testid="stForm"] {
    background: var(--bg-1) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius-md) !important;
    padding: 20px !important;
    box-shadow: none !important; /* Sin glow/sombras repetitivas */
    transition: border var(--ease-micro) !important;
}
div[data-testid="metric-container"]:hover { border-color: var(--border-hover) !important; }

/* =========================================================
   BOTONES E INPUTS
   ========================================================= */

/* Inputs genéricos */
.stTextInput input, .stDateInput input, .stNumberInput input, .stSelectbox > div > div {
  border-radius: var(--radius-sm) !important;
  border: 1px solid var(--border) !important;
  background-color: var(--bg-1) !important;
  transition: all var(--ease-micro) !important;
}
.stTextInput input:focus, .stDateInput input:focus, .stNumberInput input:focus, .stSelectbox > div > div:focus-within {
  border-color: var(--brand-border) !important;
  box-shadow: 0 0 0 2px var(--brand-soft) !important;
}

/* Botones genéricos (Secondary por default) */
.stButton button {
    background: var(--bg-1) !important;
    border: 1px solid var(--border) !important;
    color: var(--text-primary) !important;
    border-radius: var(--radius-md) !important;
    transition: all var(--ease-micro) !important;
    font-weight: 500 !important;
}
.stButton button:hover {
    border-color: var(--border-hover) !important;
    background: var(--bg-2) !important;
}
.stButton button:focus {
    border-color: var(--brand-border) !important;
    box-shadow: 0 0 0 2px var(--brand-soft) !important;
}

/* Botones primarios */
.stButton button[kind="primary"] {
    background: var(--brand) !important;
    border: 1px solid var(--brand) !important;
    color: #FFF !important; /* Texto claro siempre */
}
.stButton button[kind="primary"]:hover {
    background: var(--brand-hover) !important;
    border-color: var(--brand-hover) !important;
    transform: translateY(-1px);
}
.stButton button[kind="primary"]:active {
    transform: translateY(1px);
}

/* Botones disabled */
.stButton button:disabled {
    opacity: 0.4 !important;
    cursor: not-allowed !important;
    transform: none !important;
}

/* =========================================================
   BADGES (Estatus)
   ========================================================= */

.badge-success { background-color: var(--success-soft); color: var(--success); padding: 2px 8px; border-radius: var(--radius-sm); font-size: 12px; font-weight: 600; }
.badge-warning { background-color: var(--warning-soft); color: var(--warning); padding: 2px 8px; border-radius: var(--radius-sm); font-size: 12px; font-weight: 600; }
.badge-danger { background-color: var(--danger-soft); color: var(--danger); padding: 2px 8px; border-radius: var(--radius-sm); font-size: 12px; font-weight: 600; }
.badge-info { background-color: var(--info-soft); color: var(--info); padding: 2px 8px; border-radius: var(--radius-sm); font-size: 12px; font-weight: 600; }
.badge-neutral { background-color: rgba(255,255,255,0.05); color: var(--text-secondary); padding: 2px 8px; border-radius: var(--radius-sm); font-size: 12px; font-weight: 600; }

/* Punto pulsante para nivel 3-4 */
@keyframes pulseDanger { 0% { box-shadow: 0 0 0 0 rgba(218,54,51,0.4); } 70% { box-shadow: 0 0 0 6px rgba(218,54,51,0); } 100% { box-shadow: 0 0 0 0 rgba(218,54,51,0); } }
.pulse-dot { display: inline-block; width: 8px; height: 8px; background-color: var(--danger); border-radius: 50%; animation: pulseDanger 2s infinite; margin-right: 6px; }

/* =========================================================
   ANIMACIONES LIMITADAS
   ========================================================= */
/* Solo para splash screen */
.st-key-splash_stage [data-testid="stVideo"] {
    position: fixed !important;
    top: 50% !important;
    left: 50% !important;
    min-width: 100% !important;
    min-height: 100% !important;
    width: auto !important;
    height: auto !important;
    transform: translate(-50%, -50%) scale(1.1) !important;
}
.st-key-login_wrap {
    position: relative;
    z-index: 10;
    background: rgba(15, 15, 15, 0.4) !important;
    backdrop-filter: blur(15px) !important;
    -webkit-backdrop-filter: blur(15px) !important;
    padding: 40px;
    border-radius: var(--radius-lg);
    box-shadow: 0 10px 40px rgba(0, 0, 0, 0.6);
}

/* Ocultar elementos de UI cuando no hay login */
.hide-ui [data-testid="stSidebar"], .hide-ui [data-testid="stHeader"], .hide-ui footer { display: none !important; }

/* Utilitarias */
.label-caps { font-size: 11px; line-height: 14px; font-weight: 600; letter-spacing: 0.04em; text-transform: uppercase; color: var(--text-secondary); }

</style>
    """, unsafe_allow_html=True)
