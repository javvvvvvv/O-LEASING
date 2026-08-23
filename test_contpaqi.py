"""
test_contpaqi.py — Pruebas de core/contpaqi.py.

La prueba central de este archivo (test_reconstruye_archivo_real_completo)
no compara contra datos inventados: reconstruye las 2,904 líneas de
movimiento y las 3 líneas de encabezado de un TXT real, exportado
directamente del Contpaqi de la empresa, y verifica coincidencia carácter
por carácter (ignorando solo los UUID, que son aleatorios por diseño). Si
algo en el formato cambia sin querer, esta prueba truena.

Para correrlas: pytest test_contpaqi.py -v
"""
import re

import pandas as pd
import pytest

from core.contpaqi import poliza_txt_contpaqi, _pad, _num_txt

RUTA_MUESTRA_REAL = "tests_fixtures/poliza_contpaqi_muestra_real.txt"
_PAT_UUID = re.compile(r"[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}")


def _cargar_lineas_reales():
    with open(RUTA_MUESTRA_REAL, "rb") as f:
        contenido = f.read().decode("latin-1")
    return [l for l in contenido.split("\r\n") if l]


def test_pad_recorta_y_rellena():
    assert _pad("hola", 10) == "hola      "
    assert _pad("esto es muy largo", 5) == "esto "
    assert _pad(None, 3) == "   "


def test_num_txt_conserva_decimal_en_enteros():
    """Regresión directa de un bug real: usar formato %g quitaba el '.0'
    de los montos enteros, y la muestra real SIEMPRE lo trae."""
    assert _num_txt(422823.0, 21).strip() == "422823.0"
    assert _num_txt(109589.40, 21).strip() == "109589.4"
    assert _num_txt(3477.01, 21).strip() == "3477.01"


def test_encabezado_p_mide_185_caracteres():
    df = pd.DataFrame([{"Cuenta": "11800000546000900", "Concepto": "x", "Cargo": 1.0, "Abono": 0, "Descripcion": "d"}])
    out = poliza_txt_contpaqi(df, "Inicial", pd.Timestamp("2026-07-31"), 1)
    primera_linea = out.decode("latin-1").split("\r\n")[0]
    assert len(primera_linea) == 185


def test_linea_movimiento_mide_272_caracteres():
    df = pd.DataFrame([{"Cuenta": "11800000546000900", "Concepto": "x", "Cargo": 1.0, "Abono": 0, "Descripcion": "d"}])
    out = poliza_txt_contpaqi(df, "Inicial", pd.Timestamp("2026-07-31"), 1)
    segunda_linea = out.decode("latin-1").split("\r\n")[1]
    assert len(segunda_linea) == 272


def test_flag_cargo_es_0_y_abono_es_1():
    df = pd.DataFrame([
        {"Cuenta": "11800000546000900", "Concepto": "es cargo", "Cargo": 100.0, "Abono": 0, "Descripcion": "d"},
        {"Cuenta": "11700000546000900", "Concepto": "es abono", "Cargo": 0, "Abono": 200.0, "Descripcion": "d"},
    ])
    out = poliza_txt_contpaqi(df, "Inicial", pd.Timestamp("2026-07-31"), 1).decode("latin-1")
    lineas = out.split("\r\n")
    assert lineas[1][65] == "0"  # cargo
    assert lineas[2][65] == "1"  # abono


def test_usa_saltos_de_linea_estilo_windows():
    df = pd.DataFrame([{"Cuenta": "11800000546000900", "Concepto": "x", "Cargo": 1.0, "Abono": 0, "Descripcion": "d"}])
    out = poliza_txt_contpaqi(df, "Inicial", pd.Timestamp("2026-07-31"), 1)
    assert b"\r\n" in out
    assert b"\n\n" not in out.replace(b"\r\n", b"")


@pytest.mark.parametrize("tipo,esperado", [
    ("Inicial", "POLIZA DE APERTURA"),
    ("Parcialidad (mensual)", "POLIZA PARCIALIDADES"),
    ("Comisión por apertura", "POLIZA COMISIÓN POR APERTURA"),
])
def test_concepto_de_poliza_segun_tipo(tipo, esperado):
    df = pd.DataFrame([{"Cuenta": "11800000546000900", "Concepto": "x", "Cargo": 1.0, "Abono": 0, "Descripcion": "d"}])
    out = poliza_txt_contpaqi(df, tipo, pd.Timestamp("2026-07-31"), 1).decode("latin-1")
    assert esperado in out.split("\r\n")[0]


def test_reconstruye_archivo_real_completo():
    """La prueba de fuego: usando SOLO los datos que ya trae cada línea
    real (cuenta, concepto corto, cargo/abono, monto, descripción, fecha),
    se reconstruye esa misma línea con poliza_txt_contpaqi y debe salir
    IDÉNTICA, carácter por carácter, ignorando nada más el UUID (que es
    aleatorio por diseño, no parte del layout)."""
    lineas_reales = _cargar_lineas_reales()
    assert len(lineas_reales) == 2907  # 3 pólizas + 2904 movimientos, tal cual el archivo real

    p_lines = [l for l in lineas_reales if l.startswith("P")]
    m_lines = [l for l in lineas_reales if l.startswith("M")]
    assert len(p_lines) == 3
    assert len(m_lines) == 2904

    for real in p_lines:
        fecha = real[3:11]
        numero = int(real[19:26])
        concepto = real[40:141].rstrip()
        df_vacio = pd.DataFrame(columns=["Cuenta", "Concepto", "Cargo", "Abono", "Descripcion"])
        generado = poliza_txt_contpaqi(
            df_vacio, concepto, pd.Timestamp(fecha), numero,
            generador_uuid=lambda: "X" * 36,
        ).decode("latin-1").split("\r\n")[0]
        assert generado == _PAT_UUID.sub("X" * 36, real)

    for real in m_lines:
        cuenta = real[3:20]
        concepto = real[34:65].rstrip()
        flag = real[65]
        monto = float(real[67:88].rstrip())
        descripcion = real[120:226].rstrip()
        fecha = real[263:271]
        df = pd.DataFrame([{
            "Cuenta": cuenta, "Concepto": concepto,
            "Cargo": monto if flag == "0" else 0.0,
            "Abono": monto if flag == "1" else 0.0,
            "Descripcion": descripcion,
        }])
        generado = poliza_txt_contpaqi(
            df, "Inicial", pd.Timestamp(fecha), 1,
            generador_uuid=lambda: "X" * 36,
        ).decode("latin-1").split("\r\n")[1]
        assert generado == _PAT_UUID.sub("X" * 36, real), f"no coincide:\nreal: {real!r}\ngenerado: {generado!r}"
