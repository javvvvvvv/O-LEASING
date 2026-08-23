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
Punto de entrada de la app de escritorio (O-Leasing.exe).

Sin este archivo, si algo falla al iniciar (falta un componente de Windows,
un módulo no se incluyó al compilar, etc.) la ventana se abre y se cierra
sola en una fracción de segundo y no queda rastro de qué pasó. Este
launcher atrapa el error, lo guarda en un .txt junto al programa y muestra
un mensaje en pantalla en vez de desaparecer sin explicación.

Autor: Javier Illán
"""
import os
import sys
import traceback


def _base_dir():
    if getattr(sys, "frozen", False):
        # PyInstaller: los datos (--contents-directory .) quedan junto al .exe
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR = _base_dir()
LOG_PATH = os.path.join(BASE_DIR, "error_al_iniciar.txt")


def _mostrar_error(mensaje: str):
    """Muestra un cuadro de diálogo nativo de Windows. Si por algún motivo
    ni siquiera esto funciona, al menos ya quedó el log escrito en disco."""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, mensaje, "O-Leasing — No se pudo iniciar", 0x10)
    except Exception:
        pass


def main():
    # Evita que matplotlib intente generar su caché de fuentes cada vez
    # (más rápido al abrir, y evita un posible fallo si la carpeta no es
    # escribible).
    os.environ.pop("MPLCONFIGDIR", None)

    from streamlit_desktop_app import start_desktop_app
    app_path = os.path.join(BASE_DIR, "app.py")
    start_desktop_app(app_path, title="O-Leasing")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        detalle = traceback.format_exc()
        try:
            with open(LOG_PATH, "w", encoding="utf-8") as f:
                f.write(detalle)
        except Exception:
            pass
        _mostrar_error(
            "O-Leasing no pudo iniciar.\n\n"
            f"Se guardó el detalle técnico en:\n{LOG_PATH}\n\n"
            "Comparte ese archivo para poder corregir el problema."
        )
        sys.exit(1)
