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
models/auth.py — Capa de acceso a datos de usuarios, grupos y sesiones.
"""
import os
import sqlite3
import hashlib
import secrets
from datetime import datetime

ROLES = ("admin", "captura", "lectura")
ROL_LABELS = {
    "admin": "Administrador (Superusuario - acceso total)",
    "captura": "Captura (puede registrar y editar, limitado a su grupo)",
    "lectura": "Solo lectura (limitado a su grupo)",
}

_PBKDF2_ITERACIONES = 200_000

def _hash_password(password: str, salt: bytes | None = None) -> tuple[str, str]:
    salt = salt or secrets.token_bytes(16)
    derivado = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, _PBKDF2_ITERACIONES)
    return derivado.hex(), salt.hex()

def _verificar_password(password: str, hash_guardado: str, salt_hex: str) -> bool:
    derivado, _ = _hash_password(password, bytes.fromhex(salt_hex))
    return secrets.compare_digest(derivado, hash_guardado)

def get_auth_db_path(data_dir: str) -> str:
    return os.path.join(data_dir, "usuarios.db")

def init_auth_db(db_path: str) -> None:
    conn = sqlite3.connect(db_path)
    try:
        # Grupos de acceso
        conn.execute("""
            CREATE TABLE IF NOT EXISTS grupos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre_grupo TEXT UNIQUE NOT NULL,
                descripcion TEXT,
                creado_en TEXT NOT NULL
            )
        """)
        # Relación Grupo - Empresa
        conn.execute("""
            CREATE TABLE IF NOT EXISTS grupo_empresas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                grupo_id INTEGER NOT NULL,
                empresa_id TEXT NOT NULL,
                UNIQUE(grupo_id, empresa_id),
                FOREIGN KEY (grupo_id) REFERENCES grupos(id) ON DELETE CASCADE
            )
        """)
        # Usuarios
        conn.execute("""
            CREATE TABLE IF NOT EXISTS usuarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                nombre_completo TEXT NOT NULL,
                password_hash TEXT NOT NULL,
                password_salt TEXT NOT NULL,
                rol TEXT NOT NULL,
                grupo_id INTEGER,
                activo INTEGER NOT NULL DEFAULT 1,
                creado_en TEXT NOT NULL,
                ultimo_login TEXT,
                FOREIGN KEY (grupo_id) REFERENCES grupos(id)
            )
        """)
        conn.commit()
    finally:
        conn.close()

def hay_usuarios(db_path: str) -> bool:
    conn = sqlite3.connect(db_path)
    try:
        n = conn.execute("SELECT COUNT(*) FROM usuarios").fetchone()[0]
        return n > 0
    finally:
        conn.close()

# --- GRUPOS ---

def crear_grupo(db_path: str, nombre_grupo: str, descripcion: str, empresa_ids: list[str]) -> tuple[bool, str]:
    if not nombre_grupo.strip(): return False, "El nombre del grupo es obligatorio."
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        cursor = conn.cursor()
        cursor.execute("INSERT INTO grupos (nombre_grupo, descripcion, creado_en) VALUES (?,?,?)",
                       (nombre_grupo.strip(), descripcion.strip(), datetime.now().isoformat()))
        grupo_id = cursor.lastrowid
        for emp_id in empresa_ids:
            cursor.execute("INSERT INTO grupo_empresas (grupo_id, empresa_id) VALUES (?,?)", (grupo_id, emp_id))
        conn.commit()
        return True, "Grupo creado."
    except sqlite3.IntegrityError:
        return False, "Ya existe un grupo con ese nombre."
    finally:
        conn.close()

def listar_grupos(db_path: str) -> list[dict]:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        filas = conn.execute("SELECT * FROM grupos ORDER BY nombre_grupo").fetchall()
        grupos = []
        for f in filas:
            g = dict(f)
            emps = conn.execute("SELECT empresa_id FROM grupo_empresas WHERE grupo_id=?", (g['id'],)).fetchall()
            g['empresas'] = [e['empresa_id'] for e in emps]
            grupos.append(g)
        return grupos
    finally:
        conn.close()

def eliminar_grupo(db_path: str, grupo_id: int) -> tuple[bool, str]:
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        n_usrs = conn.execute("SELECT COUNT(*) FROM usuarios WHERE grupo_id=?", (grupo_id,)).fetchone()[0]
        if n_usrs > 0:
            return False, f"No se puede eliminar: hay {n_usrs} usuarios en este grupo."
        
        conn.execute("DELETE FROM grupos WHERE id=?", (grupo_id,))
        conn.commit()
        return True, "Grupo eliminado."
    finally:
        conn.close()

def obtener_empresas_de_grupo(db_path: str, grupo_id: int) -> list[str]:
    conn = sqlite3.connect(db_path)
    try:
        filas = conn.execute("SELECT empresa_id FROM grupo_empresas WHERE grupo_id=?", (grupo_id,)).fetchall()
        return [f[0] for f in filas]
    finally:
        conn.close()

# --- USUARIOS ---

def crear_usuario(db_path: str, username: str, nombre_completo: str, password: str, rol: str, grupo_id: int = None) -> tuple[bool, str]:
    username = (username or "").strip().lower()
    nombre_completo = (nombre_completo or "").strip()
    if not username or not nombre_completo:
        return False, "El usuario y el nombre completo son obligatorios."
    if len(password) < 6:
        return False, "La contraseña debe tener al menos 6 caracteres."
    if rol not in ROLES:
        return False, "Rol inválido."
    if rol != "admin" and not grupo_id:
        return False, "Los usuarios que no son super administradores deben pertenecer a un grupo."
    
    hash_pw, salt_hex = _hash_password(password)
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        try:
            conn.execute(
                "INSERT INTO usuarios (username, nombre_completo, password_hash, password_salt, rol, grupo_id, activo, creado_en) "
                "VALUES (?,?,?,?,?,?,1,?)",
                (username, nombre_completo, hash_pw, salt_hex, rol, grupo_id if rol != 'admin' else None, datetime.now().isoformat()),
            )
            conn.commit()
            return True, "Usuario creado correctamente."
        except sqlite3.IntegrityError:
            return False, f"El usuario '{username}' ya existe."
    finally:
        conn.close()

def verificar_login(db_path: str, username: str, password: str) -> dict | None:
    username = (username or "").strip().lower()
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        fila = conn.execute(
            "SELECT * FROM usuarios WHERE username=? AND activo=1", (username,)
        ).fetchone()
        if not fila:
            return None
        if not _verificar_password(password, fila["password_hash"], fila["password_salt"]):
            return None
        conn.execute("UPDATE usuarios SET ultimo_login=? WHERE id=?", (datetime.now().isoformat(), fila["id"]))
        conn.commit()
        return {
            "id": fila["id"], 
            "username": fila["username"], 
            "nombre_completo": fila["nombre_completo"], 
            "rol": fila["rol"],
            "grupo_id": fila["grupo_id"]
        }
    finally:
        conn.close()

def listar_usuarios(db_path: str) -> list[dict]:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        filas = conn.execute(
            "SELECT u.id, u.username, u.nombre_completo, u.rol, u.activo, u.ultimo_login, u.grupo_id, g.nombre_grupo "
            "FROM usuarios u LEFT JOIN grupos g ON u.grupo_id = g.id ORDER BY u.nombre_completo"
        ).fetchall()
        return [dict(f) for f in filas]
    finally:
        conn.close()

def cambiar_estado_usuario(db_path: str, user_id: int, activo: bool) -> None:
    conn = sqlite3.connect(db_path)
    try:
        conn.execute("UPDATE usuarios SET activo=? WHERE id=?", (1 if activo else 0, user_id))
        conn.commit()
    finally:
        conn.close()

def cambiar_rol_usuario(db_path: str, user_id: int, rol: str, grupo_id: int = None) -> tuple[bool, str]:
    if rol not in ROLES:
        return False, "Rol inválido."
    if rol != "admin" and not grupo_id:
        return False, "Debe seleccionar un grupo para este rol."
    conn = sqlite3.connect(db_path)
    try:
        conn.execute("UPDATE usuarios SET rol=?, grupo_id=? WHERE id=?", (rol, grupo_id if rol != 'admin' else None, user_id))
        conn.commit()
        return True, "Rol actualizado."
    finally:
        conn.close()

def resetear_password(db_path: str, user_id: int, password_nueva: str, password_actual: str = None, require_actual: bool = False) -> tuple[bool, str]:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        if require_actual:
            if not password_actual: return False, "Contraseña actual requerida."
            fila = conn.execute("SELECT password_hash, password_salt FROM usuarios WHERE id=?", (user_id,)).fetchone()
            if not fila or not _verificar_password(password_actual, fila["password_hash"], fila["password_salt"]):
                return False, "La contraseña actual es incorrecta."
                
        if len(password_nueva) < 6:
            return False, "La nueva contraseña debe tener al menos 6 caracteres."
        hash_pw, salt_hex = _hash_password(password_nueva)
        conn.execute("UPDATE usuarios SET password_hash=?, password_salt=? WHERE id=?", (hash_pw, salt_hex, user_id))
        conn.commit()
        return True, "Contraseña actualizada."
    finally:
        conn.close()
