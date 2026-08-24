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
test_finanzas.py — Pruebas de calc_amort, calc_tasa, tir y vp_res, las
fórmulas de las que depende cada peso que se muestra en intereses, pólizas
y reportes. Antes no había forma de comprobar que seguían bien después de
un cambio; un error de signo o redondeo se podía colar sin que nadie lo
notara hasta ver un contrato con números raros meses después.

Para correrlas: pytest test_finanzas.py -v
"""
import math
import pytest
from finanzas import calc_amort, calc_res_amort, vp_res, calc_tasa, rentabilidad, tir


def renta_teorica(inv, residual, tasa, plazo):
    """Fórmula clásica de anualidad con valor residual (pago fijo que
    amortiza 'inv' hasta dejar exactamente 'residual' al cabo de 'plazo'
    períodos a la tasa 'tasa'). Se usa como verdad independiente para
    comprobar calc_amort y calc_tasa — no reutiliza su algoritmo interno,
    así que si ambos coinciden es una señal fuerte de que están bien."""
    if tasa == 0:
        return (inv - residual) / plazo
    factor = (1 + tasa) ** plazo
    return (inv * factor - residual) * tasa / (factor - 1)


CASOS = [
    # (inversión neta, tasa mensual, plazo, residual)
    (100_000, 0.02, 24, 10_000),
    (500_000, 0.015, 36, 50_000),
    (1_000_000, 0.025, 48, 0),
    (250_000, 0.01, 12, 25_000),
    (80_000, 0.03, 6, 5_000),
]


@pytest.mark.parametrize("inv,tasa,plazo,residual", CASOS)
def test_calc_amort_termina_exacto_en_el_residual(inv, tasa, plazo, residual):
    """El saldo final de la tabla de amortización debe cuadrar exactamente
    contra el residual pactado (con la renta teórica correcta)."""
    renta = renta_teorica(inv, residual, tasa, plazo)
    df, *_ = calc_amort(inv, renta, residual, plazo, tasa)
    saldo_final = df.iloc[-1]['Saldo']
    assert saldo_final == pytest.approx(residual, abs=0.02)


@pytest.mark.parametrize("inv,tasa,plazo,residual", CASOS)
def test_calc_amort_capital_mas_interes_igual_a_renta(inv, tasa, plazo, residual):
    """En cada mes (salvo el ajuste final), Capital + Interés debe ser
    igual a la renta cobrada ese mes — si no, se estaría perdiendo o
    inventando dinero en algún lado."""
    renta = renta_teorica(inv, residual, tasa, plazo)
    df, *_ = calc_amort(inv, renta, residual, plazo, tasa)
    for _, fila in df.iloc[:-1].iterrows():
        assert fila['Capital'] + fila['Interes'] == pytest.approx(renta, abs=0.01)


@pytest.mark.parametrize("inv,tasa,plazo,residual", CASOS)
def test_calc_tasa_recupera_la_tasa_original(inv, tasa, plazo, residual):
    """Dada la renta teórica exacta para una tasa conocida, calc_tasa debe
    recuperar esa misma tasa (redondeando). Esta es la fórmula que se usa
    al dar de alta un contrato nuevo — si falla aquí, todos los contratos
    nuevos saldrían con la tasa mal calculada."""
    renta = renta_teorica(inv, residual, tasa, plazo)
    tasa_calculada = calc_tasa(plazo, renta, inv, residual)
    assert tasa_calculada is not None
    assert tasa_calculada == pytest.approx(tasa, abs=1e-4)


def test_calc_tasa_devuelve_none_si_es_imposible():
    """Con plazo=0 no hay ningún flujo que resolver (no hay meses sobre
    los que calcular una tasa) — numpy_financial no puede converger y
    calc_tasa debe devolver None en vez de tronar."""
    resultado = calc_tasa(plazo=0, renta=100, inv=1_000_000, residual=900_000)
    assert resultado is None


def test_calc_tasa_puede_devolver_tasas_negativas_o_absurdas():
    """OJO — hallazgo real, no un bug de esta prueba: numpy_financial.rate()
    SÍ converge (matemáticamente) incluso con datos de negocio absurdos,
    como una renta de $1 contra una inversión de $1,000,000 — solo que la
    tasa que encuentra es negativa. calc_tasa no tiene (todavía) ninguna
    validación de rango razonable de negocio; hoy mismo aceptaría y
    guardaría una tasa negativa sin avisar a nadie. Esta prueba documenta
    ese comportamiento a propósito, para que quede registrado como un
    pendiente conocido (ver recomendación de 'rango razonable de tasa'
    en la auditoría) y no se descubra por accidente en un contrato real."""
    resultado = calc_tasa(plazo=24, renta=1.0, inv=1_000_000, residual=900_000)
    assert resultado is not None
    assert resultado < 0  # confirma que hoy no hay tope de sanidad de negocio


def test_vp_res_y_calc_res_amort_son_inversas():
    """Si se descuenta un residual a valor presente con vp_res, y luego
    se hace crecer ese valor presente mes a mes con calc_res_amort a la
    misma tasa, se debe llegar EXACTAMENTE de vuelta al residual original
    — son operaciones inversas una de la otra."""
    residual, tasa, plazo = 45_000, 0.018, 36
    vp = vp_res(residual, tasa, plazo)
    df = calc_res_amort(vp, tasa, plazo)
    assert df.iloc[-1]['Saldo_Fin'] == pytest.approx(residual, abs=0.01)


def test_vp_res_con_tasa_cero_no_truena():
    """Con tasa 0%, vp_res no debe dividir entre cero ni tronar."""
    assert vp_res(10_000, 0, 24) == 0


def test_rentabilidad_margen_cero_cuando_no_hay_ganancia():
    fila = dict(Mensualidad_Sin_IVA=1000, Plazo=10, Residual_Monto=0,
                Comision_Monto=0, Valor_Sin_IVA=10_000, Anticipo_Monto=0)
    g, m, pe = rentabilidad(fila)
    assert g == pytest.approx(0, abs=0.01)
    assert m == pytest.approx(0, abs=0.01)
    assert pe == pytest.approx(10, abs=0.01)  # 10,000 / 1,000 = 10 meses


def test_tir_coincide_con_calc_tasa_sin_anticipo_ni_comision():
    """Cuando el contrato no tiene anticipo ni comisión, el flujo de la
    TIR es matemáticamente idéntico al que resuelve calc_tasa — deben
    coincidir (una vez anualizada la tasa mensual). Es una comprobación
    cruzada entre dos implementaciones distintas del mismo concepto
    actuarial: si una se desalinea de la otra, esta prueba lo detecta."""
    inv, tasa, plazo, residual = 300_000, 0.02, 24, 20_000
    renta = renta_teorica(inv, residual, tasa, plazo)
    fila = dict(Valor_Sin_IVA=inv, Anticipo_Monto=0, Mensualidad_Sin_IVA=renta,
                Plazo=plazo, Residual_Monto=residual, Comision_Monto=0)
    tir_anual = tir(fila)
    tasa_anual_equivalente = (((1 + tasa) ** 12) - 1) * 100
    assert tir_anual == pytest.approx(tasa_anual_equivalente, abs=0.05)


def test_tir_anualizada_es_mayor_que_la_mensual_simple():
    """Guarda contra el bug ya corregido de mostrar la TIR mensual como si
    fuera anual: para cualquier tasa positiva, la TIR anual compuesta
    debe ser considerablemente mayor que la tasa mensual x 100."""
    fila = dict(Valor_Sin_IVA=200_000, Anticipo_Monto=40_000, Mensualidad_Sin_IVA=6000,
                Plazo=36, Residual_Monto=15_000, Comision_Monto=0)
    tir_anual = tir(fila)
    assert tir_anual is not None
    assert tir_anual > 15  # muy por encima de una tasa mensual típica (~1-3%)
