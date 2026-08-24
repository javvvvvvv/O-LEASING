import sqlite3
import hashlib
import os
import secrets
import getpass
from pathlib import Path
from datetime import datetime

def hash_password(password: str, salt: bytes = None) -> tuple[str, str]:
    """Genera el hash de la contraseña con sal usando PBKDF2-HMAC-SHA256."""
    if salt is None:
        salt = secrets.token_bytes(16)
    derivado = hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt, 200_000)
    return derivado.hex(), salt.hex()

def crear_super_usuario(username: str = None, password: str = None, nombre_completo: str = None):
    """
    Crea un super usuario en la base de datos.
    
    Si no se proporcionan argumentos, se ejecuta en modo interactivo.
    Si se proporcionan argumentos, se ejecuta en modo automático (para scripting).
    """
    print("=== CREACIÓN DE SUPER USUARIO ===")
    print("Este script creará un usuario con rol 'super_usuario' en la base de datos.")
    print("Asegúrate de ejecutarlo desde la carpeta raíz del proyecto.\n")

    # Determinar ruta de la base de datos (USUARIOS.DB, no AUTH.DB)
    data_dir = Path("data")
    data_dir.mkdir(exist_ok=True)
    db_path = data_dir / "usuarios.db"
    print(f"Base de datos: {db_path}\n")

    if username is None:
        username = input("Nombre de usuario para el super administrador: ").strip().lower()
        if not username:
            print("❌ El nombre de usuario no puede estar vacío.")
            return

        nombre_completo_input = input("Nombre completo (ej. Administrador Principal): ").strip()
        if not nombre_completo_input:
            nombre_completo_input = username
        
        while True:
            password = getpass.getpass("Contraseña (mínimo 6 caracteres): ")
            if len(password) < 6:
                print("❌ La contraseña debe tener al menos 6 caracteres.")
                continue
            confirm = getpass.getpass("Confirmar contraseña: ")
            if password != confirm:
                print("❌ Las contraseñas no coinciden. Intenta de nuevo.")
                continue
            break
    else:
        # Modo automático con argumentos
        username = username.strip().lower()
        if not username:
            print("❌ El nombre de usuario no puede estar vacío.")
            return
        if nombre_completo is None:
            nombre_completo = username
        else:
            nombre_completo = nombre_completo.strip()
        
        if password and len(password) < 6:
            print("❌ La contraseña debe tener al menos 6 caracteres.")
            return

    # Conectar a la base de datos
    try:
        conn = sqlite3.connect(str(db_path))
        cursor = conn.cursor()

        # Inicializar todas las tablas necesarias (igual que en models/auth.py)
        # Tabla de usuarios
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            nombre_completo TEXT NOT NULL,
            password_hash TEXT NOT NULL,
            password_salt TEXT NOT NULL,
            rol TEXT NOT NULL CHECK(rol IN ('super_usuario','admin','captura','lectura')),
            grupo_id INTEGER,
            activo INTEGER NOT NULL DEFAULT 1,
            creado_en TEXT NOT NULL,
            ultimo_login TEXT,
            FOREIGN KEY (grupo_id) REFERENCES grupos(id)
        )
        """)
        
        # Tabla de grupos
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS grupos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre_grupo TEXT UNIQUE NOT NULL,
            descripcion TEXT,
            creado_en TEXT NOT NULL
        )
        """)
        
        # Tabla intermedia grupo-empresas
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS grupo_empresas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            grupo_id INTEGER NOT NULL,
            empresa_id TEXT NOT NULL,
            UNIQUE(grupo_id, empresa_id),
            FOREIGN KEY (grupo_id) REFERENCES grupos(id) ON DELETE CASCADE
        )
        """)

        conn.commit()

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
            pwd_hash, salt_hex = hash_password(password)
            cursor.execute("""
                UPDATE usuarios 
                SET password_hash = ?, password_salt = ?, rol = 'super_usuario', activo = 1 
                WHERE id = ?
            """, (pwd_hash, salt_hex, user_id))
            print(f"✅ Usuario '{username}' actualizado exitosamente como SUPER USUARIO.")
        else:
            # Crear nuevo super usuario
            pwd_hash, salt_hex = hash_password(password)
            cursor.execute("""
                INSERT INTO usuarios (username, nombre_completo, password_hash, password_salt, rol, activo, creado_en)
                VALUES (?, ?, ?, ?, 'super_usuario', 1, ?)
            """, (username, nombre_completo, pwd_hash, salt_hex, datetime.now().isoformat()))
            print(f"✅ Super usuario '{username}' creado exitosamente.")

        conn.commit()
        conn.close()
        print("\n🎉 ¡Listo! Ahora puedes iniciar sesión en la aplicación con este usuario.")
        print(f"   Usuario: {username}")
        print("   Rol: super_usuario (acceso total a todos los usuarios, grupos y empresas)")

    except sqlite3.IntegrityError as e:
        print(f"❌ Error de integridad: {e}")
    except Exception as e:
        print(f"❌ Error inesperado: {e}")

if __name__ == "__main__":
    import sys
    
    # Modo de uso: python crear_super_usuario.py [username] [password] [nombre_completo]
    if len(sys.argv) >= 3:
        # Modo automático desde línea de comandos
        username_arg = sys.argv[1]
        password_arg = sys.argv[2]
        nombre_completo_arg = sys.argv[3] if len(sys.argv) > 3 else None
        crear_super_usuario(username=username_arg, password=password_arg, nombre_completo=nombre_completo_arg)
    else:
        # Modo interactivo
        crear_super_usuario()
