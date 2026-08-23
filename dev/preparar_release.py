# ============================================================================
# PROPIEDAD INTELECTUAL Y LICENCIA COMERCIAL CERRADA
# ============================================================================
# Autor Legal y Titular de Derechos: JAVIER ILLAN GONZALEZ
# Organización: ORANGE CREW
# Contacto: ILLANJAVIER9@GMAIL.COM
# ============================================================================
"""
dev/preparar_release.py — Genera la copia de app.py que de verdad se
entrega dentro del .exe, sin comentarios ni docstrings.

POR QUÉ EXISTE ESTE ARCHIVO
----------------------------
El launcher de la app de escritorio (streamlit_desktop_app) necesita que
app.py exista como texto fuente real junto al .exe — así es como Streamlit
ejecuta el script. Eso significa que, tal como estaba antes, cualquiera que
reciba la carpeta "dist\\O-Leasing" puede abrir app.py con el Bloc de notas
y leer el código completo, comentarios incluidos. Los demás módulos
(models/, reports/, core/, finanzas.py) SÍ quedan protegidos porque
PyInstaller los compila dentro del propio .exe — el problema real estaba
concentrado nada más en app.py.

QUÉ HACE
--------
Lee app.py tal cual vive en el repositorio (con todos sus comentarios, útiles
para seguir trabajando en él) y escribe una copia — SOLO esa copia, nunca el
original — sin:
  - comentarios con #
  - docstrings y strings sueltos usados como comentario (no se toca ningún
    string que sea argumento de una función, como el HTML/CSS que se pasa
    a st.markdown() — eso es contenido real, no documentación, y se deja
    intacto)
  - líneas en blanco de sobra que quedan tras quitar lo anterior

Esto NO es una promesa de que el código se vuelve imposible de leer — quien
de verdad quiera invertir el tiempo puede seguir leyendo la lógica, nomás
que sin la ventaja de que alguien más ya se la explicó línea por línea. Para
una protección más fuerte (nombres de variables cambiados, lógica de
control disfrazada, bytecode cifrado) la herramienta hecha para esto es
PyArmor — ver la nota al final de construir_app_escritorio.bat.

USO
---
    python dev/preparar_release.py app.py dist_src/app.py

construir_app_escritorio.bat ya llama esto automáticamente antes de
compilar; normalmente no hace falta correrlo a mano.
"""
import ast
import io
import sys
import tokenize


def _quitar_docstrings_y_strings_sueltos(codigo_fuente: str) -> str:
    """Recorre el árbol sintáctico y vacía cualquier declaración que sea
    NADA MÁS un string suelto (docstrings de módulo/función/clase, o el
    típico "comentario largo" escrito como string en vez de con #). Un
    string suelto así no tiene ningún efecto en tiempo de ejecución —
    Python lo evalúa y lo descarta — así que quitarlo es seguro.

    A propósito NO toca strings que sean argumento de una llamada (como
    st.markdown("...") o las consultas SQL) — esos si hacen algo."""
    arbol = ast.parse(codigo_fuente)
    lineas = codigo_fuente.splitlines(keepends=True)

    objetivos = []
    for nodo in ast.walk(arbol):
        cuerpo = getattr(nodo, "body", None)
        if not isinstance(cuerpo, list):
            continue  # algunos nodos (como el "si...entonces" de una expresión
                       # ternaria) también tienen un atributo .body, pero no es
                       # una lista de sentencias — se ignoran esos casos.
        for stmt in cuerpo:
            if (
                isinstance(stmt, ast.Expr)
                and isinstance(stmt.value, ast.Constant)
                and isinstance(stmt.value.value, str)
            ):
                objetivos.append(stmt)

    # Se procesan de abajo hacia arriba para que borrar unas líneas no
    # corra los números de línea de las que faltan por procesar.
    for stmt in sorted(objetivos, key=lambda n: n.lineno, reverse=True):
        inicio, fin = stmt.lineno - 1, stmt.end_lineno
        for i in range(inicio, fin):
            lineas[i] = ""
    return "".join(lineas)


def _quitar_comentarios(codigo_fuente: str) -> str:
    """Quita los comentarios con # usando el tokenizador real de Python
    (no una expresión regular) para no arriesgarse a borrar un # que en
    realidad está dentro de un string.

    A propósito NO reconstruye el archivo token por token — reconstruir así
    resultó ser arriesgado: en Python 3.12 el tokenizador de f-strings
    puede reportar el texto de un token ya "desescapado" (una llave doble
    "{{" se reporta como una sola "{") mientras que su posición en el
    archivo sigue ocupando las dos columnas originales; reconstruir a mano
    con esa información deja un espacio de más justo ahí y convierte una
    llave literal en el inicio de una expresión de f-string — output roto.
    En vez de eso, aquí solo se recorta el rango exacto (fila, columna) de
    cada comentario directamente sobre el texto original; todo lo demás
    queda carácter por carácter idéntico al original, sin reconstrucción."""
    lineas = codigo_fuente.splitlines(keepends=True)
    tokens = tokenize.generate_tokens(io.StringIO(codigo_fuente).readline)
    for tok_tipo, tok_str, (fila_i, col_i), (fila_f, col_f), linea_completa in tokens:
        if tok_tipo != tokenize.COMMENT:
            continue
        # Un comentario siempre vive en una sola línea.
        idx = fila_i - 1
        original = lineas[idx]
        lineas[idx] = original[:col_i] + original[col_f:]
    return "".join(lineas)


def preparar(origen: str, destino: str) -> None:
    with open(origen, "r", encoding="utf-8") as f:
        codigo = f.read()

    sin_docstrings = _quitar_docstrings_y_strings_sueltos(codigo)
    resultado = _quitar_comentarios(sin_docstrings)

    # Verificación de seguridad: si por cualquier motivo el resultado ya no
    # es Python válido, NO se escribe nada — mejor que el build falle aquí
    # con un error claro, a que se entregue un .exe roto.
    compile(resultado, destino, "exec")

    with open(destino, "w", encoding="utf-8") as f:
        f.write(resultado)

    peso_antes = len(codigo.encode("utf-8"))
    peso_despues = len(resultado.encode("utf-8"))
    print(f"[preparar_release] {origen} -> {destino}")
    print(f"[preparar_release] {peso_antes:,} bytes -> {peso_despues:,} bytes "
          f"({100 - (peso_despues * 100 // peso_antes)}% menos)")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Uso: python dev/preparar_release.py <origen.py> <destino.py>")
        sys.exit(1)
    preparar(sys.argv[1], sys.argv[2])
