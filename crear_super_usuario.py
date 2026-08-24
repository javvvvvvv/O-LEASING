import sqlite3
import hashlib
import os
import getpass
from pathlib import Path

def hash_password(password: str, salt: bytes = None) -> tuple[str, str]:
    """Genera el hash de la contraseña con sal."""
    if salt is None:
        salt = os.urandom(32)
    pwd_hash = hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt, 100000)
    return pwd_hash.hex(), salt.hex()

def crear_super_usuario():
    print("=== CREACIÓN DE SUPER USUARIO ===")
    print("Este script creará un usuario con rol 'super_usuario' en la base de datos.")
    print("Asegúrate de ejecutarlo desde la carpeta raíz del proyecto.\n")

    # Determinar ruta de la base de datos
    data_dir = Path("data")
    data_dir.mkdir(exist_ok=True)
    db_path = data_dir / "auth.db"

    username = input("Nombre de usuario para el super administrador: ").strip()
    if not username:
        print("❌ El nombre de usuario no puede estar vacío.")
        return

    nombre_completo = input("Nombre completo (ej. Administrador Principal): ").strip()
    if not nombre_completo:
        nombre_completo = username

    while True:
        password = getpass.getpass("Contraseña: ")
        if len(password) < 4:
            print("❌ La contraseña debe tener al menos 4 caracteres.")
            continue
        confirm = getpass.getpass("Confirmar contraseña: ")
        if password != confirm:
            print("❌ Las contraseñas no coinciden. Intenta de nuevo.")
            continue
        break

    # Conectar a la base de datos
    try:
        conn = sqlite3.connect(str(db_path))
        cursor = conn.cursor()

        # Verificar si la tabla existe, si no, crearla
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            nombre_completo TEXT NOT NULL,
            password_hash TEXT NOT NULL,
            salt TEXT NOT NULL,
            rol TEXT NOT NULL DEFAULT 'usuario',
            grupo_id INTEGER,
            activo BOOLEAN DEFAULT 1,
            creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """)

        # Verificar si el usuario ya existe
        cursor.execute("SELECT id, rol FROM usuarios WHERE username = ?", (username,))
        existente = cursor.fetchone()

        if existente:
            user_id, rol_actual = existente
            print(f"\n⚠️  El usuario '{username}' ya existe (ID: {user_id}, Rol actual: {rol_actual}).")
            opcion = input("¿Deseas actualizarlo a 'super_usuario'? (s/n): ").strip().lower()
            if opcion != 's':
                print("Operación cancelada.")
                conn.close()
                return
            
            # Actualizar contraseña y rol
            pwd_hash, salt = hash_password(password)
            cursor.execute("""
                UPDATE usuarios 
                SET password_hash = ?, salt = ?, rol = 'super_usuario', activo = 1 
                WHERE id = ?
            """, (pwd_hash, salt, user_id))
            print(f"✅ Usuario '{username}' actualizado exitosamente como SUPER USUARIO.")
        else:
            # Crear nuevo usuario
            pwd_hash, salt = hash_password(password)
            cursor.execute("""
                INSERT INTO usuarios (username, nombre_completo, password_hash, salt, rol, activo)
                VALUES (?, ?, ?, ?, 'super_usuario', 1)
            """, (username, nombre_completo, pwd_hash, salt))
            print(f"✅ Super usuario '{username}' creado exitosamente.")

        conn.commit()
        conn.close()
        print("\n🎉 ¡Listo! Ahora puedes iniciar sesión en la aplicación con este usuario.")
        print(f"   Usuario: {username}")
        print("   Rol: super_usuario (acceso total)")

    except sqlite3.IntegrityError as e:
        print(f"❌ Error de integridad: {e}")
    except Exception as e:
        print(f"❌ Error inesperado: {e}")

if __name__ == "__main__":
    crear_super_usuario()
