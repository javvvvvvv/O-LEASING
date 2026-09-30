# ============================================================================
# PROPIEDAD INTELECTUAL Y LICENCIA COMERCIAL CERRADA
# ============================================================================
# Autor Legal y Titular de Derechos: JAVIER ILLAN GONZALEZ
# Organización: ORANGE CREW
# Contacto: ILLANJAVIER9@GMAIL.COM
# ============================================================================
"""
reports/cierre_mensual_pdf.py — PDF de Cierre Mensual y Conciliación Contable.
Genera un informe gerencial imprimible en PDF con el resumen de cobranza, facturación,
devengamiento de intereses y movimientos del mes.
"""
import io
import os
import pandas as pd
from datetime import datetime

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors as rl_colors

def generar_pdf_cierre_mensual(
    mes: int,
    anio: int,
    kpis: dict,
    df_conciliados: pd.DataFrame,
    empresa_nombre: str = "O-Leasing"
) -> bytes:
    """
    Genera un informe en PDF de Cierre Mensual y Conciliación.
    """
    MN = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
    nombre_mes = MN[mes - 1]

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36
    )

    styles = getSampleStyleSheet()
    primary_color = rl_colors.HexColor('#1E5C4F')
    dark_gray = rl_colors.HexColor('#222222')
    light_bg = rl_colors.HexColor('#F4F6F6')

    title_style = ParagraphStyle(
        'CierreTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=16,
        leading=20,
        textColor=primary_color
    )
    subtitle_style = ParagraphStyle(
        'CierreSubtitle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9.5,
        leading=13,
        textColor=dark_gray
    )
    bold_cell = ParagraphStyle(
        'CierreBoldCell',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=10,
        textColor=dark_gray
    )
    norm_cell = ParagraphStyle(
        'CierreNormCell',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=10,
        textColor=dark_gray
    )

    story = []

    # Encabezado principal
    story.append(Paragraph(f"<b>{empresa_nombre.upper()}</b> — INFORME DE CIERRE MENSUAL Y CONCILIACIÓN", title_style))
    story.append(Spacer(1, 4))
    fecha_impresion = datetime.now().strftime("%d/%m/%Y %H:%M")
    story.append(Paragraph(f"<b>Período del Cierre:</b> {nombre_mes} {anio} &nbsp;|&nbsp; <b>Impresión:</b> {fecha_impresion}", subtitle_style))
    story.append(Spacer(1, 8))
    story.append(HRFlowable(width="100%", thickness=1.5, color=primary_color, spaceBefore=2, spaceAfter=8))

    # Resumen de KPIs del mes
    data_kpi = [
        [Paragraph("INDICADOR CONTABLE", bold_cell), Paragraph("MONTO EN PERÍODO", bold_cell)],
        [Paragraph("Intereses Leasing Devengados (Cta 208)", norm_cell), f"${kpis.get('interes_leasing',0):,.2f}"],
        [Paragraph("Intereses Residuales Acumulados", norm_cell), f"${kpis.get('interes_residual',0):,.2f}"],
        [Paragraph("Amortización de Comisión por Apertura", norm_cell), f"${kpis.get('comision',0):,.2f}"],
        [Paragraph("Facturación Total CFDI Conciliada", norm_cell), f"${kpis.get('facturado_total',0):,.2f}"],
        [Paragraph("Contratos Activos Procesados", norm_cell), str(kpis.get('contratos_activos', 0))],
        [Paragraph("Facturas Conciliadas / Discrepancias", norm_cell), f"{kpis.get('conciliadas',0)} OK / {kpis.get('discrepancias',0)} Pend."]
    ]
    t_kpi = Table(data_kpi, colWidths=[340, 200])
    t_kpi.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), primary_color),
        ('TEXTCOLOR', (0, 0), (-1, 0), rl_colors.white),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, rl_colors.HexColor('#DDDDDD')),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(t_kpi)
    story.append(Spacer(1, 14))

    # Detalle de Contratos y Conciliación
    story.append(Paragraph(f"<b>DETALLE DE CONTRATOS Y FACTURACIÓN — {nombre_mes.upper()} {anio}</b>", bold_cell))
    story.append(Spacer(1, 4))

    data_det = [[
        Paragraph("Contrato", bold_cell),
        Paragraph("Cliente", bold_cell),
        Paragraph("Int. Leasing", bold_cell),
        Paragraph("Int. Residual", bold_cell),
        Paragraph("Estatus", bold_cell)
    ]]

    if df_conciliados is not None and not df_conciliados.empty:
        for idx, row in df_conciliados.head(25).iterrows():
            data_det.append([
                str(row.get('ID_Contrato', idx)),
                Paragraph(str(row.get('Cliente', '—'))[:28], norm_cell),
                f"${float(row.get(nombre_mes, 0) or 0):,.2f}",
                f"${float(row.get('Int_Residual', 0) or 0):,.2f}",
                str(row.get('Estatus', 'OK'))
            ])

    t_det = Table(data_det, colWidths=[70, 210, 90, 90, 80])
    t_det.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), light_bg),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, rl_colors.HexColor('#EEEEEE')),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    story.append(t_det)
    story.append(Spacer(1, 20))

    # Pie de firmas
    f_data = [
        [Paragraph("__________________________________________<br><b>Elaboró / Contador General</b>", bold_cell),
         Paragraph("__________________________________________<br><b>Autorizó / Dirección General</b>", bold_cell)]
    ]
    t_firmas = Table(f_data, colWidths=[270, 270])
    t_firmas.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    story.append(t_firmas)

    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes
