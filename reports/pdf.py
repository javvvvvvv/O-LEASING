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
reports/pdf.py — Generación de PDFs (Estado de Cuenta y Pólizas).

Tercera pieza movida fuera de app.py en el refactor por capas (ver
docs/PLAN_REFACTOR_CAPAS.md, fase 1). Mismo patrón que reports/excel.py:
el nombre de la empresa activa se resuelve vía `set_resolver_titulo_empresa`,
nunca importando app.py ni models/ directamente.
"""
import io
import os
import sys

import pandas as pd
import matplotlib
matplotlib.use('Agg')  # backend sin ventana — el servidor no tiene pantalla
import matplotlib.pyplot as plt

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors as rl_colors

if getattr(sys, "frozen", False):
    _BASE_DIR = os.path.dirname(sys.executable)
else:
    _BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

LOGO_OLEASING_PNG_DEFAULT = os.path.join(_BASE_DIR, "assets", "o-leasing-logo.png")
LOGO_ORANGE_PNG_DEFAULT   = os.path.join(_BASE_DIR, "assets", "orange-crew-logo.png")

_resolver_titulo_empresa = None


def set_resolver_titulo_empresa(fn) -> None:
    """Igual que en reports/excel.py: app.py registra aquí, una sola vez,
    cómo obtener el nombre de la empresa activa."""
    global _resolver_titulo_empresa
    _resolver_titulo_empresa = fn


def _nombre_empresa_actual() -> str:
    if _resolver_titulo_empresa:
        try:
            return _resolver_titulo_empresa() or "Mi Empresa"
        except Exception:
            return "Mi Empresa"
    return "Mi Empresa"


def _grafica_amort_para_pdf(dfa, mes_corte=None):
    """Dibuja con matplotlib (no plotly, para no depender de kaleido) el
    mismo tipo de gráfica de barras apiladas (interés + capital) con la
    línea de saldo que ya se ve en pantalla, y regresa un PNG en memoria
    listo para meter al PDF."""
    fig, ax1 = plt.subplots(figsize=(7.2, 2.6), dpi=170)
    meses = dfa['Fecha'].tolist()
    x = range(len(meses))
    ax1.bar(x, dfa['Interes'], color="#1E5C4F", label="Interés")
    ax1.bar(x, dfa['Capital'], bottom=dfa['Interes'], color="#8FD9BE", label="Capital")
    ax1.set_ylabel("Interés + Capital (MXN)", fontsize=7)
    ax1.tick_params(axis='both', labelsize=6)
    paso = max(1, len(meses)//12)
    ax1.set_xticks(list(x)[::paso])
    ax1.set_xticklabels([meses[i] for i in x][::paso], rotation=45, ha='right', fontsize=6)

    ax2 = ax1.twinx()
    ax2.plot(x, dfa['Saldo_Fin'], color="#B3261E", linewidth=1.6, label="Saldo insoluto")
    ax2.set_ylabel("Saldo (MXN)", fontsize=7)
    ax2.tick_params(axis='y', labelsize=6)

    if mes_corte is not None and 0 <= mes_corte < len(meses):
        ax1.axvline(mes_corte, color="#96660C", linestyle='--', linewidth=1.2)
        ax1.text(mes_corte, ax1.get_ylim()[1], 'Baja', fontsize=6, color="#96660C", ha='center', va='bottom')

    h1, l1 = ax1.get_legend_handles_labels(); h2, l2 = ax2.get_legend_handles_labels()
    ax1.legend(h1+h2, l1+l2, fontsize=6.5, loc='lower center', bbox_to_anchor=(0.5, 1.02),
               ncol=3, frameon=False, columnspacing=1.4, handlelength=1.4)
    fig.tight_layout()

    buf_img = io.BytesIO()
    fig.savefig(buf_img, format='png', bbox_inches='tight')
    plt.close(fig)
    buf_img.seek(0)
    return buf_img


def _pie_pdf_marca(epeq, logo_path=None):
    """Pie de página estándar de los reportes PDF: logo de Orange (marca
    matriz) + crédito de O-Leasing, para que la marca quede visible en
    cualquier documento que salga del sistema."""
    logo_path = logo_path or LOGO_ORANGE_PNG_DEFAULT
    if os.path.exists(logo_path):
        try:
            _logo_pie = RLImage(logo_path, width=46, height=46 * 713 / 2437)
            t_pie = Table(
                [[_logo_pie, Paragraph("Generado con O-Leasing · una app de Orange", epeq)]],
                colWidths=[50, 300],
            )
            t_pie.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ]))
            return t_pie
        except Exception:
            pass
    return Paragraph("Generado con O-Leasing · una app de Orange", epeq)


def pdf_estado_cuenta(row, dfa, dfr=None, avance=None, mes_corte=None, fecha_hoy_txt="", logo_path=None):
    """Genera el Estado de Cuenta de un contrato en PDF, con la misma
    información que se ve en pantalla (datos generales, avisos de baja/
    avance de pago, y la gráfica de amortización) para imprimir o mandar
    por correo."""
    logo_path = logo_path or LOGO_OLEASING_PNG_DEFAULT
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, rightMargin=34, leftMargin=34, topMargin=34, bottomMargin=34)
    es = getSampleStyleSheet()
    et = es['Heading1']; et.textColor = rl_colors.HexColor("#1E5C4F"); et.alignment = 1
    esub = es['Heading2']; esub.textColor = rl_colors.HexColor("#20242B"); esub.fontSize = 13
    enota = es['Normal']; enota.fontSize = 8; enota.textColor = rl_colors.HexColor("#565E68")
    epeq = es['Normal']; epeq.fontSize = 7; epeq.textColor = rl_colors.HexColor("#8A929C"); epeq.alignment = 1

    _nombre_emp = _nombre_empresa_actual()
    elems = []
    if os.path.exists(logo_path):
        try:
            _logo_img = RLImage(logo_path, width=110, height=110 * 471 / 2119)
            _logo_img.hAlign = 'CENTER'
            elems.append(_logo_img)
            elems.append(Spacer(1, 4))
        except Exception:
            pass
    elems += [
        Paragraph(f"<b>{_nombre_emp}</b>", et),
        Paragraph("Estado de Cuenta", esub),
        Paragraph(f"<b>{row['ID_Contrato']}</b> — {row['Cliente']}", es['Normal']),
        Paragraph(f"Generado el {fecha_hoy_txt} · Cifras en Pesos Mexicanos (MXN)", enota),
        Spacer(1, 10),
    ]

    if str(row.get('Estatus','')).upper() == 'BAJA':
        _fb = str(row.get('Fecha_Baja','') or '')[:10]
        elems.append(Paragraph(f"<b><font color='#B3261E'>CONTRATO DADO DE BAJA</font></b> — {_fb or 'sin fecha capturada'}", es['Normal']))
        elems.append(Spacer(1, 6))
    if avance and avance.get('estado') == 'ATRASADO':
        elems.append(Paragraph(
            f"<b><font color='#B3261E'>Va atrasado {avance['meses_atraso']} mes(es)</font></b> — "
            f"faltan el/los mes(es) {', '.join(str(m) for m in avance['meses_faltantes'][:10])}", es['Normal']))
        elems.append(Spacer(1, 6))
    elif avance and avance.get('estado') == 'ADELANTADO':
        elems.append(Paragraph(
            f"<b><font color='#1E5C4F'>Va adelantado {avance['meses_adelanto']} mes(es)</font></b> — "
            f"ya tiene facturado hasta el mes {avance['mes_max_facturado']} de {int(row.get('Plazo',0))}", es['Normal']))
        elems.append(Spacer(1, 6))

    datos = [
        ["Vehículo", str(row.get('Vehiculo','—')), "Fecha Alta", str(row.get('Fecha_Alta',''))[:10]],
        ["Fecha Vencimiento", str(row.get('Fecha_Vencimiento',''))[:10], "Plazo", f"{int(row.get('Plazo',0))} meses"],
        ["Valor (s/IVA)", f"${float(row.get('Valor_Sin_IVA',0)):,.2f}", "Renta (s/IVA)", f"${float(row.get('Mensualidad_Sin_IVA',0)):,.2f}"],
        ["Anticipo", f"${float(row.get('Anticipo_Monto',0)):,.2f}", "Residual", f"${float(row.get('Residual_Monto',0)):,.2f}"],
        ["Tasa anual impl.", f"{float(row.get('Tasa_Calculada',0))*1200:.2f}%", "Estatus", str(row.get('Estatus',''))],
    ]
    t_datos = Table(datos, colWidths=[95, 155, 95, 155])
    t_datos.setStyle(TableStyle([
        ('FONTNAME',(0,0),(0,-1),'Helvetica-Bold'), ('FONTNAME',(2,0),(2,-1),'Helvetica-Bold'),
        ('FONTSIZE',(0,0),(-1,-1),8.3), ('TEXTCOLOR',(0,0),(0,-1),rl_colors.HexColor("#565E68")),
        ('TEXTCOLOR',(2,0),(2,-1),rl_colors.HexColor("#565E68")),
        ('ROWBACKGROUNDS',(0,0),(-1,-1),[rl_colors.white, rl_colors.HexColor("#F3F4F6")]),
        ('BOX',(0,0),(-1,-1),0.6,rl_colors.HexColor("#DCE0E5")),
        ('INNERGRID',(0,0),(-1,-1),0.4,rl_colors.HexColor("#DCE0E5")),
        ('TOPPADDING',(0,0),(-1,-1),4), ('BOTTOMPADDING',(0,0),(-1,-1),4),
    ]))
    elems += [t_datos, Spacer(1, 14)]

    if dfa is not None and not dfa.empty:
        elems.append(Paragraph("<b>Amortización</b>", esub))
        img_buf = _grafica_amort_para_pdf(dfa, mes_corte=mes_corte)
        elems.append(RLImage(img_buf, width=515, height=186))
        elems.append(Spacer(1, 8))

        filas_tabla = [["Fecha","Mes","Saldo Inicial","Interés","Capital","Saldo Final"]]
        for _, r in dfa.iterrows():
            filas_tabla.append([
                str(r['Fecha']), str(int(r['Mes'])),
                f"${r['Saldo_Ini']:,.2f}" if 'Saldo_Ini' in r else '',
                f"${r['Interes']:,.2f}", f"${r['Capital']:,.2f}", f"${r['Saldo_Fin']:,.2f}",
            ])
        t_amort = Table(filas_tabla, colWidths=[62,32,88,80,80,88], repeatRows=1)
        t_amort.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,0),rl_colors.HexColor("#1E5C4F")),('TEXTCOLOR',(0,0),(-1,0),rl_colors.whitesmoke),
            ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'), ('FONTSIZE',(0,0),(-1,-1),7),
            ('ALIGN',(2,0),(-1,-1),'RIGHT'), ('ALIGN',(0,0),(1,-1),'CENTER'),
            ('ROWBACKGROUNDS',(0,1),(-1,-1),[rl_colors.white, rl_colors.HexColor("#F3F4F6")]),
            ('TOPPADDING',(0,0),(-1,-1),2.5), ('BOTTOMPADDING',(0,0),(-1,-1),2.5),
        ]))
        elems.append(t_amort)

    if dfr is not None and not dfr.empty:
        elems.append(Spacer(1, 14))
        elems.append(Paragraph("<b>Acumulación del Residual</b>", esub))
        filas_res = [["Fecha","Mes","VP Residual","Saldo Inicial","Interés","Saldo Final","Residual Pactado"]]
        for _, r in dfr.iterrows():
            filas_res.append([
                str(r['Fecha']), str(int(r['Mes'])),
                f"${r['VP_Residual']:,.2f}" if 'VP_Residual' in r else '',
                f"${r['Saldo_Ini']:,.2f}" if 'Saldo_Ini' in r else '',
                f"${r['Interes']:,.2f}",
                f"${r['Saldo_Fin']:,.2f}" if 'Saldo_Fin' in r else '',
                f"${r['Residual_Pactado']:,.2f}" if 'Residual_Pactado' in r else '',
            ])
        t_res = Table(filas_res, colWidths=[52,26,72,72,64,72,80], repeatRows=1)
        t_res.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,0),rl_colors.HexColor("#1E5C4F")),('TEXTCOLOR',(0,0),(-1,0),rl_colors.whitesmoke),
            ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'), ('FONTSIZE',(0,0),(-1,-1),6.5),
            ('ALIGN',(2,0),(-1,-1),'RIGHT'), ('ALIGN',(0,0),(1,-1),'CENTER'),
            ('ROWBACKGROUNDS',(0,1),(-1,-1),[rl_colors.white, rl_colors.HexColor("#F3F4F6")]),
            ('TOPPADDING',(0,0),(-1,-1),2.5), ('BOTTOMPADDING',(0,0),(-1,-1),2.5),
        ]))
        elems.append(t_res)

    elems.append(Spacer(1, 16))
    elems.append(_pie_pdf_marca(epeq))

    doc.build(elems)
    buf.seek(0)
    return buf


def pdf_poliza(df, titulo, periodo, logo_path=None):
    logo_path = logo_path or LOGO_OLEASING_PNG_DEFAULT
    buf=io.BytesIO()
    doc=SimpleDocTemplate(buf,pagesize=letter,rightMargin=38,leftMargin=38,topMargin=38,bottomMargin=38)
    elems=[]; es=getSampleStyleSheet(); et=es['Heading1']
    et.textColor=rl_colors.HexColor("#1E5C4F"); et.alignment=1
    epeq = es['Normal']; epeq.fontSize = 7; epeq.textColor = rl_colors.HexColor("#8A929C")
    _nombre_emp_p = _nombre_empresa_actual()
    if os.path.exists(logo_path):
        try:
            _logo_img_p = RLImage(logo_path, width=100, height=100 * 471 / 2119)
            _logo_img_p.hAlign = 'CENTER'
            elems += [_logo_img_p, Spacer(1, 4)]
        except Exception:
            pass
    elems+=[Paragraph(f"<b>{_nombre_emp_p}</b>",et),
            Paragraph(f"<b>Póliza:</b> {titulo} | <b>Período:</b> {periodo}",es['Normal']),
            Paragraph("Cifras expresadas en Pesos Mexicanos (MXN)",es['Normal']),Spacer(1,12)]
    data=[["Cuenta","Descripción","Cargo","Abono","Concepto"]]; tc=ta=0.0
    for _,r in df.iterrows():
        data.append([r['Cuenta'],r['Descripcion'][:38],f"${r['Cargo']:,.2f}",f"${r['Abono']:,.2f}",r['Concepto'][:50]])
        tc+=r['Cargo']; ta+=r['Abono']
    data.append(["","SUMAS IGUALES",f"${tc:,.2f}",f"${ta:,.2f}",""])
    t=Table(data,colWidths=[88,168,78,78,118])
    t.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),rl_colors.HexColor("#1E5C4F")),('TEXTCOLOR',(0,0),(-1,0),rl_colors.whitesmoke),
        ('ALIGN',(0,0),(-1,-1),'CENTER'),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('ROWBACKGROUNDS',(0,1),(-1,-2),[rl_colors.white,rl_colors.HexColor("#F3F4F6")]),
        ('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),('LINEABOVE',(0,-1),(-1,-1),1.2,rl_colors.HexColor("#1E5C4F")),
        ('ALIGN',(2,1),(3,-1),'RIGHT')]))
    elems.append(t)
    elems.append(Spacer(1,14))
    elems.append(_pie_pdf_marca(epeq))
    doc.build(elems); buf.seek(0); return buf
