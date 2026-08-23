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
models/auth.py — Capa de acceso a datos de usuarios y sesiones.

Usuarios compartidos por el equipo/PC (no por empresa): un mismo usuario
puede operar cualquiera de las empresas dadas de alta en Multiempresa.
Base de datos propia (data/usuarios.db), separada de los .db de cada
empresa, para que borrar/restaurar la base de una empresa nunca afecte
las cuentas de acceso.

Reglas de esta capa (igual que models/db.py):
- Nada aquí lee `st.session_state`; todo recibe lo que necesita como
  parámetro. Las decisiones de "quién es el usuario en turno" viven en
  app.py (capa de UI).
- Contraseñas nunca se guardan en texto plano: PBKDF2-HMAC-SHA256 con
  salt aleatorio por usuario (200,000 iteraciones), usando únicamente
  la librería estándar de Python para no agregar dependencias nuevas.
"""
import os
import sqlite3
import hashlib
import secrets
from datetime import datetime

ROLES = ("admin", "captura", "lectura")
ROL_LABELS = {
    "admin": "Administrador (acceso total)",
    "captura": "Captura (puede registrar y editar, no puede cambiar configuración)",
    "lectura": "Solo lectura (puede consultar, no puede guardar cambios)",
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
        conn.execute("""
            CREATE TABLE IF NOT EXISTS usuarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                nombre_completo TEXT NOT NULL,
                password_hash TEXT NOT NULL,
                password_salt TEXT NOT NULL,
                rol TEXT NOT NULL CHECK(rol IN ('admin','captura','lectura')),
                activo INTEGER NOT NULL DEFAULT 1,
                creado_en TEXT NOT NULL,
                ultimo_login TEXT
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


def crear_usuario(db_path: str, username: str, nombre_completo: str, password: str, rol: str) -> tuple[bool, str]:
    username = (username or "").strip().lower()
    nombre_completo = (nombre_completo or "").strip()
    if not username or not nombre_completo:
        return False, "El usuario y el nombre completo son obligatorios."
    if len(password) < 6:
        return False, "La contraseña debe tener al menos 6 caracteres."
    if rol not in ROLES:
        return False, "Rol inválido."
    hash_pw, salt_hex = _hash_password(password)
    conn = sqlite3.connect(db_path)
    try:
        try:
            conn.execute(
                "INSERT INTO usuarios (username, nombre_completo, password_hash, password_salt, rol, activo, creado_en) "
                "VALUES (?,?,?,?,?,1,?)",
                (username, nombre_completo, hash_pw, salt_hex, rol, datetime.now().isoformat()),
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
        return {"id": fila["id"], "username": fila["username"], "nombre_completo": fila["nombre_completo"], "rol": fila["rol"]}
    finally:
        conn.close()


def listar_usuarios(db_path: str) -> list[dict]:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        filas = conn.execute(
            "SELECT id, username, nombre_completo, rol, activo, ultimo_login FROM usuarios ORDER BY nombre_completo"
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


def cambiar_rol_usuario(db_path: str, user_id: int, rol: str) -> tuple[bool, str]:
    if rol not in ROLES:
        return False, "Rol inválido."
    conn = sqlite3.connect(db_path)
    try:
        conn.execute("UPDATE usuarios SET rol=? WHERE id=?", (rol, user_id))
        conn.commit()
        return True, "Rol actualizado."
    finally:
        conn.close()


def resetear_password(db_path: str, user_id: int, password_nueva: str) -> tuple[bool, str]:
    if len(password_nueva) < 6:
        return False, "La contraseña debe tener al menos 6 caracteres."
    hash_pw, salt_hex = _hash_password(password_nueva)
    conn = sqlite3.connect(db_path)
    try:
        conn.execute("UPDATE usuarios SET password_hash=?, password_salt=? WHERE id=?", (hash_pw, salt_hex, user_id))
        conn.commit()
        return True, "Contraseña actualizada."
    finally:
        conn.close()
