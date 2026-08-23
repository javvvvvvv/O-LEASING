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
finanzas.py — Fórmulas financieras de O-Leasing (amortización, tasa
implícita, valor presente del residual, TIR).

Este archivo está separado de app.py a propósito: aquí NO se importa
Streamlit ni nada de la interfaz, solo las fórmulas puras. Así se pueden
probar con pytest (test_finanzas.py) sin levantar la app, y si un día se
corrige una fórmula, se corrige en un solo lugar y ya afecta a toda la app
(reportes, pólizas, dashboard, etc.). El cacheo (@st.cache_data) es cosa
de la interfaz, así que vive en app.py, no aquí.

Autor: Javier Illán
"""
import pandas as pd
import numpy as np
import numpy_financial as npf


def calc_amort(inv, renta, residual, plazo, tasa):
    """Tabla de amortización mensual de un contrato de arrendamiento.

    Cada mes: interés = saldo x tasa; capital = renta - interés;
    saldo nuevo = saldo - capital. El último mes se ajusta para que
    el saldo final cuadre exactamente contra el residual pactado
    (evita que quede $0.01-$0.05 de diferencia por redondeo acumulado).

    Devuelve (tabla, interés primeros 12 meses, interés resto del
    plazo, renta cobrada primeros 12 meses, renta cobrada el resto).
    """
    saldo = inv
    rows = []
    for m in range(1, plazo + 1):
        i = saldo * tasa
        k = renta - i
        saldo -= k
        rows.append({'Mes': m, 'Interes': i, 'Capital': k, 'Saldo': max(saldo, 0)})
    df = pd.DataFrame(rows)
    diff = df.iloc[-1]['Saldo'] - residual
    if abs(diff) > 0.01:
        df.at[df.index[-1], 'Capital'] += diff
        df.at[df.index[-1], 'Saldo'] = residual
    icp = df.head(12)['Interes'].sum()
    ilp = df.iloc[12:]['Interes'].sum()
    return df, icp, ilp, renta * min(12, plazo), renta * max(0, plazo - 12)


def calc_res_amort(vp, tasa, plazo):
    """Acumulación del valor presente del residual hacia su valor futuro
    pactado, mes a mes, a la misma tasa implícita del contrato (útil para
    reconocer el interés implícito del residual como ingreso financiero
    separado de la amortización de capital)."""
    saldo = vp
    rows = []
    for m in range(1, plazo + 1):
        i = saldo * tasa
        saldo += i
        rows.append({'Mes': m, 'Saldo_Ini': saldo - i, 'Interes': i, 'Saldo_Fin': saldo})
    return pd.DataFrame(rows)


def vp_res(res, tasa, plazo):
    """Valor presente del residual, descontado a la tasa implícita del
    contrato durante 'plazo' meses."""
    return res / ((1 + tasa) ** plazo) if tasa else 0


def calc_tasa(plazo, renta, inv, residual):
    """Tasa mensual implícita del contrato, o None si numpy_financial no
    logra calcularla (parámetros que no forman un flujo financiero válido:
    p. ej. renta insuficiente para nunca amortizar el residual pactado).

    Antes esta función devolvía 0.0 en cualquier fallo, lo que hacía que
    un contrato con datos mal capturados se guardara con 0% de tasa SIN
    AVISO — y a partir de ahí todos sus intereses, pólizas y reportes de
    rentabilidad salían en ceros sin que nadie lo notara. Ahora devuelve
    None explícitamente para que quien llama a esta función decida cómo
    avisar al usuario (ver guardar() en app.py)."""
    try:
        r = npf.rate(plazo, -renta, inv, -residual)
        if r is None or np.isnan(r):
            return None
        return float(r)
    except Exception:
        return None


def rentabilidad(row):
    """Ganancia nominal, margen (%) y payback simple (meses) de un
    contrato. 'row' es cualquier objeto tipo diccionario con las llaves
    Mensualidad_Sin_IVA, Plazo, Residual_Monto, Comision_Monto,
    Valor_Sin_IVA y Anticipo_Monto (una fila de pandas o un dict normal
    sirven igual — por eso esta función es fácil de probar)."""
    ing = row['Mensualidad_Sin_IVA'] * row['Plazo'] + row['Residual_Monto'] + row['Comision_Monto']
    eg = row['Valor_Sin_IVA'] - row['Anticipo_Monto']
    g = ing - eg
    m = (g / eg) * 100 if eg else 0
    pe = eg / row['Mensualidad_Sin_IVA'] if row['Mensualidad_Sin_IVA'] > 0 else 0
    return g, m, pe


def tir(row):
    """TIR ANUAL efectiva del contrato (flujo: -inversión neta, renta
    mensual x plazo, + residual en el último mes). Internamente
    numpy_financial calcula la tasa MENSUAL implícita; aquí se anualiza
    de forma compuesta ((1+mensual)^12 - 1) para que sea comparable con
    cualquier otra tasa anual del sistema (p. ej. 'Tasa_Anual'). Mostrar
    la tasa mensual sin convertir llevaría a leerla como si fuera 12
    veces menor de lo que realmente rinde el contrato."""
    fl = [-(row['Valor_Sin_IVA'] - row['Anticipo_Monto'])]
    for _ in range(int(row['Plazo'])):
        fl.append(row['Mensualidad_Sin_IVA'])
    fl[-1] += row['Residual_Monto']
    try:
        t_mensual = npf.irr(fl)
        if t_mensual is None or np.isnan(t_mensual):
            return None
        return (((1 + t_mensual) ** 12) - 1) * 100
    except Exception:
        return None
