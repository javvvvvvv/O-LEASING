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
import sys
import os
import getpass

# Asegurarnos de que la ruta del proyecto esté en el path para importar los módulos
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from models.auth import crear_usuario, obtener_usuario_por_nombre, actualizar_rol_usuario
except ImportError as e:
    print(f"Error al importar módulos: {e}")
    print("Asegúrate de estar ejecutando este script desde la carpeta raíz del proyecto.")
    sys.exit(1)

def main():
    print("========================================")
    print("   CREACIÓN DE SUPER ADMINISTRADOR")
    print("========================================\n")

    username = input("Nombre de usuario para el Super Admin: ").strip()

    if not username:
        print("❌ El nombre de usuario no puede estar vacío.")
        return

    # Verificar si ya existe
    usuario_existente = obtener_usuario_por_nombre(username)

    if usuario_existente:
        print(f"\n⚠️  El usuario '{username}' ya existe.")
        confirmacion = input("¿Deseas convertirlo en Super Administrador? (s/n): ").strip().lower()
        if confirmacion == 's':
            exito = actualizar_rol_usuario(username, 'super_usuario')
            if exito:
                print(f"✅ ¡Éxito! El usuario '{username}' ahora es Super Administrador.")
            else:
                print("❌ Error al actualizar el rol.")
        else:
            print("Operación cancelada.")
        return

    password = getpass.getpass("Contraseña: ")
    confirm_password = getpass.getpass("Confirmar contraseña: ")

    if password != confirm_password:
        print("❌ Las contraseñas no coinciden.")
        return

    if len(password) < 4:
        print("❌ La contraseña debe tener al menos 4 caracteres.")
        return

    print("\nCreando usuario...")
    exito = crear_usuario(username, password, 'super_usuario')

    if exito:
        print(f"\n✅ ¡Éxito! Super Administrador '{username}' creado correctamente.")
        print("🔑 Ahora puedes iniciar sesión en la aplicación con estas credenciales.")
    else:
        print("\n❌ Error al crear el usuario. Verifica que la base de datos esté accesible.")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ Ocurrió un error inesperado: {e}")
        print("Asegúrate de estar en la carpeta correcta del proyecto.")
