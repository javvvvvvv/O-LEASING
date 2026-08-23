"""
test_preparar_release.py — Pruebas de dev/preparar_release.py.

Lo único que de verdad importa aquí es una cosa: que la copia sin
comentarios se comporte EXACTAMENTE igual que el original. Cualquier otra
ganancia (menos bytes, menos pistas para alguien que quiera copiar el
código) no vale nada si por el camino se rompe o se altera el
comportamiento real del sistema — así que las pruebas comparan el árbol
sintáctico completo, no nada más "que compile".

Para correrlas: pytest test_preparar_release.py -v
"""
import ast
import subprocess
import sys
import textwrap

import pytest

from dev.preparar_release import preparar


def _ast_sin_docstrings(codigo: str):
    arbol = ast.parse(codigo)
    for nodo in ast.walk(arbol):
        cuerpo = getattr(nodo, "body", None)
        if isinstance(cuerpo, list):
            nodo.body = [
                s for s in cuerpo if not (
                    isinstance(s, ast.Expr)
                    and isinstance(s.value, ast.Constant)
                    and isinstance(s.value.value, str)
                )
            ]
    return ast.dump(arbol, annotate_fields=False)


def _preparar_en_memoria(tmp_path, codigo: str) -> str:
    origen = tmp_path / "origen.py"
    destino = tmp_path / "destino.py"
    origen.write_text(codigo, encoding="utf-8")
    preparar(str(origen), str(destino))
    return destino.read_text(encoding="utf-8")


def test_quita_docstring_de_modulo(tmp_path):
    codigo = '"""Esto es un docstring de módulo."""\nx = 1\n'
    resultado = _preparar_en_memoria(tmp_path, codigo)
    assert "docstring" not in resultado
    assert "x = 1" in resultado


def test_quita_comentarios_de_linea(tmp_path):
    codigo = "x = 1  # esto es un comentario\ny = 2\n# otro comentario\n"
    resultado = _preparar_en_memoria(tmp_path, codigo)
    assert "comentario" not in resultado
    assert "x = 1" in resultado
    assert "y = 2" in resultado


def test_no_toca_strings_que_son_argumentos(tmp_path):
    """El caso central: un string pasado como argumento (como el HTML de
    st.markdown) NO es un comentario — debe sobrevivir intacto."""
    codigo = 'contenido_html = "<div>Hola # no es un comentario</div>"\n'
    resultado = _preparar_en_memoria(tmp_path, codigo)
    assert "Hola # no es un comentario" in resultado


def test_no_toca_llaves_dobles_en_fstring(tmp_path):
    """Regresión directa del bug real que se encontró: una llave doble en
    un f-string (usada para producir una sola llave literal en CSS) se
    tokeniza distinto en Python 3.12 y una reconstrucción ingenua la
    corrompía."""
    codigo = 'valor = 5\nregla = f"<style>.clase{{color:red;}}</style> {valor}"\n'
    resultado = _preparar_en_memoria(tmp_path, codigo)
    ns = {}
    exec(compile(resultado, "test", "exec"), ns)
    assert ns["regla"] == "<style>.clase{color:red;}</style> 5"


def test_no_toca_espacios_dentro_de_string_multilinea(tmp_path):
    codigo = textwrap.dedent('''\
        texto = """primera línea   
        segunda línea"""
    ''')
    resultado = _preparar_en_memoria(tmp_path, codigo)
    ns = {}
    exec(compile(resultado, "test", "exec"), ns)
    assert ns["texto"] == "primera línea   \nsegunda línea"


def test_ast_identico_al_original_ignorando_docstrings(tmp_path):
    """La prueba más importante: el árbol sintáctico del resultado debe
    ser IDÉNTICO al original (una vez que a ambos se les quitan los
    docstrings) — eso garantiza el mismo comportamiento en tiempo de
    ejecución, no nada más "que no truene"."""
    codigo = textwrap.dedent('''\
        """Docstring de módulo."""
        import os

        def saluda(nombre):
            """Docstring de función."""
            # un comentario
            return f"Hola {nombre}"  # otro comentario

        class Cosa:
            """Docstring de clase."""
            valor = 1  # comentario en clase

        "string suelto usado como comentario"
        x = {"a": 1, "b": 2}
    ''')
    resultado = _preparar_en_memoria(tmp_path, codigo)
    assert _ast_sin_docstrings(codigo) == _ast_sin_docstrings(resultado)


def test_resultado_es_ejecutable_y_produce_el_mismo_resultado(tmp_path):
    codigo = textwrap.dedent('''\
        """Módulo de prueba."""
        def suma(a, b):
            """Suma dos números."""
            # lógica sencilla
            return a + b

        resultado = suma(2, 3)
    ''')
    resultado = _preparar_en_memoria(tmp_path, codigo)
    ns_original, ns_resultado = {}, {}
    exec(compile(codigo, "orig", "exec"), ns_original)
    exec(compile(resultado, "dest", "exec"), ns_resultado)
    assert ns_original["resultado"] == ns_resultado["resultado"] == 5


def test_falla_de_forma_ruidosa_si_el_origen_no_es_python_valido(tmp_path):
    origen = tmp_path / "origen.py"
    destino = tmp_path / "destino.py"
    origen.write_text("def esto_no_cierra(:\n", encoding="utf-8")
    with pytest.raises(SyntaxError):
        preparar(str(origen), str(destino))
    assert not destino.exists()


def test_app_py_real_produce_arbol_identico():
    """Corre el script contra el app.py real del proyecto (no un ejemplo
    de juguete) — es la prueba de fuego: si esto pasa, la copia que se
    entrega en el .exe se comporta exactamente igual que el código con el
    que de verdad se trabaja."""
    with open("app.py", encoding="utf-8") as f:
        original = f.read()

    import tempfile, os
    with tempfile.TemporaryDirectory() as tmp:
        destino = os.path.join(tmp, "app_release.py")
        preparar("app.py", destino)
        with open(destino, encoding="utf-8") as f:
            resultado = f.read()

    assert _ast_sin_docstrings(original) == _ast_sin_docstrings(resultado)
    # Debe quedar más chico (comentarios/docstrings reales sí se quitaron),
    # si no, algo no está funcionando aunque el AST haya salido igual.
    assert len(resultado) < len(original)
