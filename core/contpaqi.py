# ============================================================================
# PROPIEDAD INTELECTUAL Y LICENCIA COMERCIAL CERRADA
# ============================================================================
# Autor Legal y Titular de Derechos: JAVIER ILLAN GONZALEZ
# Organización: ORANGE CREW
# Contacto: ILLANJAVIER9@GMAIL.COM
# ============================================================================
"""
core/contpaqi.py — Generador del TXT de importación de pólizas a CONTPAQi
Contabilidad, ancho fijo por columna.

El layout de este archivo NO se adivinó: se reconstruyó carácter por
carácter comparando contra un TXT real, de 2,904 movimientos, exportado
directamente desde el CONTPAQi de la empresa (ver
tests_fixtures/poliza_contpaqi_muestra_real.txt y
test_contpaqi.py, que reconstruye esas 2,904 líneas exactas para
comprobarlo en cada corrida de pruebas).

Estructura de la línea de encabezado 'P' (185 caracteres):
    P + 2esp + fecha(8,AAAAMMDD) + 4esp + tipo('3'=Diario) + 3esp
    + numero_poliza(7) + esp + '1' + esp + '0' + 10esp
    + concepto(101) + '11 0 0 ' + UUID(36) + esp

Estructura de cada línea de movimiento 'M1' (272 caracteres):
    M1 + esp + cuenta(17, sin guiones) + 14esp + concepto_corto(31)
    + cargo_o_abono('0'=cargo,'1'=abono) + esp + monto(21)
    + '0' + 10esp + '0.0' + 18esp + descripcion(106) + UUID(36)
    + esp + fecha(8) + esp

Si algún día aparece un caso (otro tipo de póliza, otra empresa) con un
layout distinto, lo correcto es repetir el mismo proceso: conseguir un TXT
de muestra real y ajustar aquí — nunca adivinar un campo nuevo.
"""
import uuid as _uuid

import pandas as pd


def _pad(texto, ancho: int) -> str:
    return str(texto or "")[:ancho].ljust(ancho)


def _num_txt(valor, ancho: int) -> str:
    # str(float) en Python dejar "109589.4" (sin el cero de más) pero SÍ
    # conserva "422823.0" en los números enteros — es exactamente el
    # comportamiento que trae la muestra real, no hace falta forzar un
    # número fijo de decimales.
    return _pad(str(round(float(valor), 2)), ancho)


CONCEPTOS_POLIZA = {
    "Inicial": "POLIZA DE APERTURA",
    "Parcialidad (mensual)": "POLIZA PARCIALIDADES",
    "Comisión por apertura": "POLIZA COMISIÓN POR APERTURA",
    "Reglas Personalizadas": "POLIZA REGLAS PERSONALIZADAS",
}


def poliza_txt_contpaqi(
    df_fin: pd.DataFrame,
    tipo_poliza: str,
    fecha_poliza,
    numero_poliza: int = 1,
    generador_uuid=None,
) -> bytes:
    """Arma el TXT completo (encabezado + un renglón por movimiento).

    `generador_uuid` es inyectable solo para pruebas (para poder fijar
    los UUID y comparar el archivo completo contra la muestra real,
    carácter por carácter); en uso normal se deja en None y usa
    uuid.uuid4() de verdad."""
    gen = generador_uuid or (lambda: _uuid.uuid4())
    fecha_str = pd.Timestamp(fecha_poliza).strftime("%Y%m%d")
    concepto_poliza = CONCEPTOS_POLIZA.get(tipo_poliza, tipo_poliza.upper())

    linea_p = (
        "P  " + fecha_str + "    " + "3" + "   "
        + str(int(numero_poliza)).rjust(7, "0") + " 1 0" + " " * 10
        + _pad(concepto_poliza, 101) + "11 0 0 "
        + str(gen()).upper() + " "
    )

    lineas = [linea_p]
    for _, r in df_fin.iterrows():
        cargo = float(r.get("Cargo", 0) or 0)
        abono = float(r.get("Abono", 0) or 0)
        es_cargo = cargo > 0
        monto = cargo if es_cargo else abono
        linea_m = (
            "M1 " + _pad(r["Cuenta"], 17) + " " * 14
            + _pad(r.get("Concepto", ""), 31)
            + ("0" if es_cargo else "1") + " " + _num_txt(monto, 21)
            + "0" + " " * 10 + "0.0" + " " * 18
            + _pad(r.get("Descripcion", ""), 106)
            + str(gen()).upper() + " " + fecha_str + " "
        )
        lineas.append(linea_m)

    # El archivo real de referencia usa saltos de línea estilo Windows
    # (\r\n) y Latin-1 (Contpaqi es un programa de Windows; UTF-8 podría
    # descuadrar los acentos de la descripción).
    return ("\r\n".join(lineas) + "\r\n").encode("latin-1", errors="replace")
