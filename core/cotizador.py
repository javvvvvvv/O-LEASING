# ============================================================================
# PROPIEDAD INTELECTUAL Y LICENCIA COMERCIAL CERRADA
# ============================================================================
# Autor Legal y Titular de Derechos: JAVIER ILLAN GONZALEZ
# Organización: ORANGE CREW
# Contacto: ILLANJAVIER9@GMAIL.COM
# ============================================================================
"""
core/cotizador.py — Módulo de Cotizaciones y Simulación de Arrendamientos Puros.
Calcula amortizaciones proyectadas, pagos iniciales, TIR comercial y genera PDF de cotización.
"""
import io
import os
import math
import numpy as np
import pandas as pd
from datetime import datetime

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors as rl_colors

from finanzas import calc_amort, vp_res, calc_res_amort

def simular_cotizacion(
    valor_vehiculo_sin_iva: float,
    pct_anticipo: float,
    pct_residual: float,
    tasa_anual: float,
    plazo_meses: int,
    pct_comision: float = 2.0,
    incluir_iva: bool = True
) -> dict:
    """
    Simula una cotización comercial de Arrendamiento Puro.
    Retorna desglose de pago inicial, mensuaidad sin/con IVA, tabla amortización y métricas.
    """
    v_veh = float(valor_vehiculo_sin_iva)
    monto_anticipo = round(v_veh * (pct_anticipo / 100.0), 2)
    monto_residual = round(v_veh * (pct_residual / 100.0), 2)
    monto_comision = round(v_veh * (pct_comision / 100.0), 2)
    inv_neta = max(round(v_veh - monto_anticipo, 2), 0.01)

    tasa_dec = tasa_anual / 100.0
    tasa_mensual = tasa_dec / 12.0

    # Calcular renta mensual necesaria usando formula de amortización con residual
    vp_res_val = round(vp_res(monto_residual, tasa_dec, plazo_meses), 2)
    inv_a_amortizar = inv_neta - vp_res_val

    if tasa_mensual > 0:
        renta_mensual_sin_iva = round(
            (inv_a_amortizar * tasa_mensual) / (1.0 - math.pow(1.0 + tasa_mensual, -plazo_meses)), 2
        )
    else:
        renta_mensual_sin_iva = round(inv_a_amortizar / plazo_meses, 2)

    iva_factor = 0.16 if incluir_iva else 0.0
    renta_iva = round(renta_mensual_sin_iva * iva_factor, 2)
    renta_total_con_iva = round(renta_mensual_sin_iva + renta_iva, 2)

    anticipo_iva = round(monto_anticipo * iva_factor, 2)
    comision_iva = round(monto_comision * iva_factor, 2)
    
    pago_inicial_sin_iva = round(monto_anticipo + monto_comision + renta_mensual_sin_iva, 2)
    pago_inicial_con_iva = round(
        (monto_anticipo + anticipo_iva) + (monto_comision + comision_iva) + renta_total_con_iva, 2
    )

    # Tabla de amortización proyectada
    dfa, _, _, _, _ = calc_amort(inv_neta, renta_mensual_sin_iva, monto_residual, plazo_meses, tasa_dec)
    
    # TIR mensual y anualizada de la operación
    flujos = [-inv_neta] + [renta_mensual_sin_iva] * (plazo_meses - 1) + [renta_mensual_sin_iva + monto_residual]
    try:
        irr_m = np.irr(flujos) if hasattr(np, 'irr') else np.npv # fallback
    except Exception:
        irr_m = 0.0
    
    # Usar tir de numpy_financial si está disponible
    try:
        import numpy_financial as npf
        irr_m = npf.irr(flujos)
    except Exception:
        pass
        
    tir_anual = round(irr_m * 12.0 * 100.0, 2) if (irr_m and not np.isnan(irr_m)) else tasa_anual

    return {
        "valor_vehiculo_sin_iva": v_veh,
        "valor_vehiculo_con_iva": round(v_veh * (1 + iva_factor), 2),
        "pct_anticipo": pct_anticipo,
        "monto_anticipo": monto_anticipo,
        "monto_anticipo_con_iva": round(monto_anticipo * (1 + iva_factor), 2),
        "inversion_neta": inv_neta,
        "pct_residual": pct_residual,
        "monto_residual": monto_residual,
        "vp_residual": vp_res_val,
        "tasa_anual": tasa_anual,
        "plazo_meses": plazo_meses,
        "pct_comision": pct_comision,
        "monto_comision": monto_comision,
        "monto_comision_con_iva": round(monto_comision * (1 + iva_factor), 2),
        "renta_mensual_sin_iva": renta_mensual_sin_iva,
        "renta_iva": renta_iva,
        "renta_total_con_iva": renta_total_con_iva,
        "pago_inicial_sin_iva": pago_inicial_sin_iva,
        "pago_inicial_con_iva": pago_inicial_con_iva,
        "tir_anual": tir_anual,
        "tabla_amortizacion": dfa
    }


def generar_pdf_cotizacion(sim: dict, cliente_nombre: str, vehiculo_desc: str, empresa_nombre: str = "O-Leasing") -> bytes:
    """
    Genera un PDF membretado formal con la Cotización Comercial.
    """
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
        'CotTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=18,
        leading=22,
        textColor=primary_color
    )
    subtitle_style = ParagraphStyle(
        'CotSubtitle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        leading=14,
        textColor=dark_gray
    )
    bold_cell = ParagraphStyle(
        'CotBoldCell',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=11,
        textColor=dark_gray
    )
    norm_cell = ParagraphStyle(
        'CotNormCell',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=11,
        textColor=dark_gray
    )

    story = []

    # Encabezado
    story.append(Paragraph(f"<b>{empresa_nombre.upper()}</b> — COTIZACIÓN COMERCIAL DE ARRENDAMIENTO", title_style))
    story.append(Spacer(1, 4))
    fecha_hoy = datetime.now().strftime("%d/%m/%Y")
    story.append(Paragraph(f"<b>Fecha:</b> {fecha_hoy} &nbsp;|&nbsp; <b>Cliente / Prospecto:</b> {cliente_nombre}", subtitle_style))
    story.append(Paragraph(f"<b>Vehículo / Equipo:</b> {vehiculo_desc}", subtitle_style))
    story.append(Spacer(1, 10))
    story.append(HRFlowable(width="100%", thickness=1.5, color=primary_color, spaceBefore=2, spaceAfter=10))

    # Resumen Ejecutivo (Resumen Financiero)
    data_resumen = [
        [Paragraph("CONCEPTO COMERCIAL", bold_cell), Paragraph("SIN IVA", bold_cell), Paragraph("CON IVA (16%)", bold_cell)],
        [Paragraph("Valor de la Unidad", norm_cell), f"${sim['valor_vehiculo_sin_iva']:,.2f}", f"${sim['valor_vehiculo_con_iva']:,.2f}"],
        [Paragraph(f"Anticipo a Capital ({sim['pct_anticipo']}%)", norm_cell), f"${sim['monto_anticipo']:,.2f}", f"${sim['monto_anticipo_con_iva']:,.2f}"],
        [Paragraph("Inversión Neta a Financiar", norm_cell), f"${sim['inversion_neta']:,.2f}", "—"],
        [Paragraph(f"Comisión por Apertura ({sim['pct_comision']}%)", norm_cell), f"${sim['monto_comision']:,.2f}", f"${sim['monto_comision_con_iva']:,.2f}"],
        [Paragraph(f"<b>Renta Mensual ({sim['plazo_meses']} meses)</b>", bold_cell), f"<b>${sim['renta_mensual_sin_iva']:,.2f}</b>", f"<b>${sim['renta_total_con_iva']:,.2f}</b>"],
        [Paragraph(f"Valor Residual Pactado ({sim['pct_residual']}%)", norm_cell), f"${sim['monto_residual']:,.2f}", f"${sim['monto_residual']*1.16:,.2f}"],
        [Paragraph("<b>PAGO INICIAL REQUERIDO</b>", bold_cell), f"<b>${sim['pago_inicial_sin_iva']:,.2f}</b>", f"<b>${sim['pago_inicial_con_iva']:,.2f}</b>"]
    ]

    t_res = Table(data_resumen, colWidths=[240, 150, 150])
    t_res.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), primary_color),
        ('TEXTCOLOR', (0, 0), (-1, 0), rl_colors.white),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, rl_colors.HexColor('#DDDDDD')),
        ('BACKGROUND', (0, -1), (-1, -1), light_bg),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    story.append(t_res)
    story.append(Spacer(1, 12))

    # Parámetros Financieros
    story.append(Paragraph("<b>Condiciones Financieras:</b> Tasa Anual de Referencia: <b>{:.2f}%</b> | Plazo: <b>{} meses</b> | TIR Estimada: <b>{:.2f}%</b>".format(
        sim['tasa_anual'], sim['plazo_meses'], sim['tir_anual']
    ), subtitle_style))
    story.append(Spacer(1, 10))

    # Tabla de Amortización (Primeras 12 mensualidades + Resumen)
    story.append(Paragraph("<b>PROYECCIÓN DE MENSUALIDADES (PRIMEROS 12 MESES)</b>", bold_cell))
    story.append(Spacer(1, 4))

    dfa = sim['tabla_amortizacion']
    data_tab = [[
        Paragraph("Mes", bold_cell),
        Paragraph("Renta Sin IVA", bold_cell),
        Paragraph("Interés", bold_cell),
        Paragraph("Capital", bold_cell),
        Paragraph("Saldo Insoluto", bold_cell)
    ]]

    for i, row in dfa.head(12).iterrows():
        data_tab.append([
            str(int(row['Mes'])),
            f"${row['Renta']:,.2f}",
            f"${row['Interes']:,.2f}",
            f"${row['Capital']:,.2f}",
            f"${row['Saldo']:,.2f}"
        ])

    if len(dfa) > 12:
        data_tab.append(["...", "...", "...", "...", "..."])
        last_r = dfa.iloc[-1]
        data_tab.append([
            str(int(last_r['Mes'])),
            f"${last_r['Renta']:,.2f}",
            f"${last_r['Interes']:,.2f}",
            f"${last_r['Capital']:,.2f}",
            f"${last_r['Saldo']:,.2f}"
        ])

    t_amort = Table(data_tab, colWidths=[50, 115, 115, 115, 145])
    t_amort.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), light_bg),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, rl_colors.HexColor('#EEEEEE')),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    story.append(t_amort)
    story.append(Spacer(1, 20))

    # Notas legales / Firmas
    notas = ParagraphStyle('Notas', parent=styles['Normal'], fontName='Helvetica-Oblique', fontSize=7.5, leading=9, textColor=rl_colors.HexColor('#666666'))
    story.append(Paragraph("* Cotización informativa sujeta a aprobación de crédito y disponibilidad. Los valores presentados están calculados en Pesos Mexicanos (MXN). No constituye una oferta vinculante.", notas))
    story.append(Spacer(1, 30))

    # Firmas
    f_data = [
        [Paragraph("__________________________________________<br><b>Ejecutivo Comercial</b>", bold_cell),
         Paragraph("__________________________________________<br><b>Aceptación Prospecto / Cliente</b>", bold_cell)]
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
