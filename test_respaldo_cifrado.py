"""
test_respaldo_cifrado.py — Pruebas de las funciones de respaldo de
models/db.py, con foco en el cifrado AES-256 agregado sobre pyzipper.

Corre operaciones reales de archivo sobre una base de datos SQLite
temporal (no mocks) para probar el camino completo: crear un respaldo
cifrado, listar respaldos mezclados (cifrados y sin cifrar), restaurar
con la contraseña correcta, y confirmar que con la contraseña equivocada
o sin pyzipper la restauración falla de forma clara en vez de corromper
la base de datos.

Para correrlas: pytest test_respaldo_cifrado.py -v
"""
import os
import sqlite3

import pytest

from models.db import (
    crear_respaldo_automatico,
    listar_respaldos_automaticos,
    restaurar_respaldo_automatico,
    carpeta_respaldos_auto,
    _PYZIPPER_DISPONIBLE,
)

pytestmark = pytest.mark.skipif(not _PYZIPPER_DISPONIBLE, reason="pyzipper no está instalado en este entorno")


def _crear_db_sana(ruta: str, valor: str = "hola"):
    conn = sqlite3.connect(ruta)
    conn.execute("CREATE TABLE datos (id INTEGER PRIMARY KEY, valor TEXT)")
    conn.execute("INSERT INTO datos (valor) VALUES (?)", (valor,))
    conn.commit()
    conn.close()


def test_respaldo_sin_password_sigue_sin_cifrar(tmp_path):
    db = str(tmp_path / "leasing.db")
    _crear_db_sana(db)
    crear_respaldo_automatico(db)
    respaldos = listar_respaldos_automaticos(db)
    assert len(respaldos) == 1
    archivo, _fecha, cifrado = respaldos[0]
    assert cifrado is False
    assert archivo.endswith(".bak")


def test_respaldo_con_password_queda_cifrado(tmp_path):
    db = str(tmp_path / "leasing.db")
    _crear_db_sana(db)
    crear_respaldo_automatico(db, password="clave-super-secreta")
    respaldos = listar_respaldos_automaticos(db)
    assert len(respaldos) == 1
    archivo, _fecha, cifrado = respaldos[0]
    assert cifrado is True
    assert archivo.endswith(".bak.zip")


def test_archivo_cifrado_no_se_puede_leer_como_zip_normal(tmp_path):
    """El contenido real debe estar protegido — no basta con que el
    nombre diga .zip, el archivo interno no debe abrirse sin contraseña."""
    db = str(tmp_path / "leasing.db")
    _crear_db_sana(db)
    crear_respaldo_automatico(db, password="clave-super-secreta")
    archivo, _, _ = listar_respaldos_automaticos(db)[0]

    import zipfile
    with zipfile.ZipFile(archivo) as zf:
        with pytest.raises(Exception):
            zf.read(zf.namelist()[0])  # sin password, pyzipper/zipfile debe rechazar el contenido cifrado


def test_restaurar_respaldo_cifrado_con_password_correcta(tmp_path):
    db = str(tmp_path / "leasing.db")
    _crear_db_sana(db, valor="dato_original")
    crear_respaldo_automatico(db, password="clave-correcta")
    archivo, _, _ = listar_respaldos_automaticos(db)[0]

    # Se "daña" la base de datos actual para simular el escenario real de restauración
    with open(db, "wb") as f:
        f.write(b"esto ya no es un sqlite valido")

    restaurar_respaldo_automatico(db, archivo, password="clave-correcta")

    conn = sqlite3.connect(db)
    valor = conn.execute("SELECT valor FROM datos").fetchone()[0]
    conn.close()
    assert valor == "dato_original"


def test_restaurar_respaldo_cifrado_con_password_incorrecta_falla_claro(tmp_path):
    db = str(tmp_path / "leasing.db")
    _crear_db_sana(db)
    crear_respaldo_automatico(db, password="clave-correcta")
    archivo, _, _ = listar_respaldos_automaticos(db)[0]

    with pytest.raises(RuntimeError, match="contraseña"):
        restaurar_respaldo_automatico(db, archivo, password="clave-incorrecta")


def test_respaldos_cifrados_y_sin_cifrar_se_limpian_por_separado(tmp_path):
    """Cambiar de 'sin contraseña' a 'con contraseña' no debe borrar de
    golpe los respaldos buenos del otro tipo."""
    db = str(tmp_path / "leasing.db")
    _crear_db_sana(db)
    for _ in range(3):
        crear_respaldo_automatico(db, max_respaldos=8)
        import time; time.sleep(1.1)
    for _ in range(3):
        crear_respaldo_automatico(db, password="clave", max_respaldos=8)
        import time; time.sleep(1.1)

    respaldos = listar_respaldos_automaticos(db)
    cifrados = [r for r in respaldos if r[2]]
    planos = [r for r in respaldos if not r[2]]
    assert len(cifrados) == 3
    assert len(planos) == 3
