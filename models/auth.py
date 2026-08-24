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

NUEVO: Sistema de grupos de usuarios y empresas asignadas
- Los usuarios pertenecen a grupos que definen qué empresas pueden ver
- El super_usuario puede administrar todos los usuarios y grupos
- Cada grupo tiene asignado un conjunto de empresas visibles

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

ROLES = ("super_usuario", "admin", "captura", "lectura")
ROL_LABELS = {
    "super_usuario": "Super Usuario (administración total de usuarios y grupos)",
    "admin": "Administrador (acceso total a empresas asignadas)",
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
        # Tabla de usuarios con grupo_id
        conn.execute("""
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
        # Tabla de grupos de usuarios
        conn.execute("""
            CREATE TABLE IF NOT EXISTS grupos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre_grupo TEXT UNIQUE NOT NULL,
                descripcion TEXT,
                creado_en TEXT NOT NULL
            )
        """)
        # Tabla intermedia grupo-empresas (qué empresas ve cada grupo)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS grupo_empresas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                grupo_id INTEGER NOT NULL,
                empresa_id TEXT NOT NULL,
                UNIQUE(grupo_id, empresa_id),
                FOREIGN KEY (grupo_id) REFERENCES grupos(id) ON DELETE CASCADE
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


def crear_usuario(db_path: str, username: str, nombre_completo: str, password: str, rol: str, grupo_id: int = None) -> tuple[bool, str]:
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
                "INSERT INTO usuarios (username, nombre_completo, password_hash, password_salt, rol, grupo_id, activo, creado_en) "
                "VALUES (?,?,?,?,?,?,1,?)",
                (username, nombre_completo, hash_pw, salt_hex, rol, grupo_id, datetime.now().isoformat()),
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
        filas = conn.execute("""
            SELECT u.id, u.username, u.nombre_completo, u.rol, u.activo, u.ultimo_login, u.grupo_id, g.nombre_grupo
            FROM usuarios u
            LEFT JOIN grupos g ON u.grupo_id = g.id
            ORDER BY u.nombre_completo
        """).fetchall()
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


def actualizar_grupo_usuario(db_path: str, user_id: int, grupo_id: int | None) -> tuple[bool, str]:
    conn = sqlite3.connect(db_path)
    try:
        conn.execute("UPDATE usuarios SET grupo_id=? WHERE id=?", (grupo_id, user_id))
        conn.commit()
        return True, "Grupo actualizado."
    finally:
        conn.close()


# ==================== GESTIÓN DE GRUPOS ====================

def crear_grupo(db_path: str, nombre_grupo: str, descripcion: str = "") -> tuple[bool, str]:
    nombre_grupo = (nombre_grupo or "").strip()
    if not nombre_grupo:
        return False, "El nombre del grupo es obligatorio."
    conn = sqlite3.connect(db_path)
    try:
        try:
            conn.execute(
                "INSERT INTO grupos (nombre_grupo, descripcion, creado_en) VALUES (?,?,?)",
                (nombre_grupo, descripcion, datetime.now().isoformat()),
            )
            conn.commit()
            return True, "Grupo creado correctamente."
        except sqlite3.IntegrityError:
            return False, f"El grupo '{nombre_grupo}' ya existe."
    finally:
        conn.close()


def listar_grupos(db_path: str) -> list[dict]:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        filas = conn.execute("SELECT * FROM grupos ORDER BY nombre_grupo").fetchall()
        return [dict(f) for f in filas]
    finally:
        conn.close()


def eliminar_grupo(db_path: str, grupo_id: int) -> tuple[bool, str]:
    conn = sqlite3.connect(db_path)
    try:
        # Verificar si hay usuarios en este grupo
        usuarios_en_grupo = conn.execute("SELECT COUNT(*) FROM usuarios WHERE grupo_id=?", (grupo_id,)).fetchone()[0]
        if usuarios_en_grupo > 0:
            return False, f"No se puede eliminar: hay {usuarios_en_grupo} usuario(s) en este grupo."
        conn.execute("DELETE FROM grupo_empresas WHERE grupo_id=?", (grupo_id,))
        conn.execute("DELETE FROM grupos WHERE id=?", (grupo_id,))
        conn.commit()
        return True, "Grupo eliminado."
    finally:
        conn.close()


# ==================== GESTIÓN DE EMPRESAS POR GRUPO ====================

def obtener_empresas_grupo(db_path: str, grupo_id: int) -> list[str]:
    """Retorna lista de empresa_ids que puede ver un grupo."""
    conn = sqlite3.connect(db_path)
    try:
        filas = conn.execute("SELECT empresa_id FROM grupo_empresas WHERE grupo_id=?", (grupo_id,)).fetchall()
        return [f[0] for f in filas]
    finally:
        conn.close()


def asignar_empresa_a_grupo(db_path: str, grupo_id: int, empresa_id: str) -> tuple[bool, str]:
    conn = sqlite3.connect(db_path)
    try:
        try:
            conn.execute("INSERT INTO grupo_empresas (grupo_id, empresa_id) VALUES (?,?)", (grupo_id, empresa_id))
            conn.commit()
            return True, "Empresa asignada al grupo."
        except sqlite3.IntegrityError:
            return False, "Esta empresa ya está asignada a este grupo."
    finally:
        conn.close()


def remover_empresa_de_grupo(db_path: str, grupo_id: int, empresa_id: str) -> tuple[bool, str]:
    conn = sqlite3.connect(db_path)
    try:
        conn.execute("DELETE FROM grupo_empresas WHERE grupo_id=? AND empresa_id=?", (grupo_id, empresa_id))
        conn.commit()
        return True, "Empresa removida del grupo."
    finally:
        conn.close()


def obtener_grupos_con_empresas(db_path: str) -> list[dict]:
    """Retorna todos los grupos con sus empresas asignadas."""
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        grupos = conn.execute("SELECT * FROM grupos ORDER BY nombre_grupo").fetchall()
        resultado = []
        for g in grupos:
            empresas = conn.execute("SELECT empresa_id FROM grupo_empresas WHERE grupo_id=?", (g["id"],)).fetchall()
            grupo_dict = dict(g)
            grupo_dict["empresas"] = [e["empresa_id"] for e in empresas]
            resultado.append(grupo_dict)
        return resultado
    finally:
        conn.close()
