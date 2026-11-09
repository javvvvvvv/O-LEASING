"""
O-LEASING · REST API BACKEND (FastAPI)
======================================
Servicio API REST para la aplicación móvil Android y clientes externos de O-Leasing.
Permite:
1. Consulta y búsqueda de contratos en tiempo real desde leasing.db.
2. Detalle financiero completo y amortización de cada contrato.
3. Cotizador / calculadora financiera instantánea con motor finanzas.py.
4. Alta / creación individual de contratos nuevos directamente en leasing.db.
5. Resumen ejecutivo de cartera y métricas para el dashboard móvil.
"""

import os
import sys
import sqlite3
import datetime
import hmac
import hashlib
import secrets
import threading
import time
from typing import List, Optional
from pydantic import BaseModel, Field
from fastapi import FastAPI, HTTPException, Query, status, Depends, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials

# Asegurar importación del motor financiero
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if _BASE_DIR not in sys.path:
    sys.path.insert(0, _BASE_DIR)

from finanzas import calc_tasa, vp_res, calc_amort, calc_rent_amort, calc_res_amort
from models.auth import (
    get_auth_db_path, init_auth_db, verificar_login,
    crear_sesion, validar_sesion, cerrar_sesion, resetear_password,
)

_DOCS_ON = os.environ.get("OLEASING_API_DOCS", "0") == "1"

app = FastAPI(
    title="O-Leasing Mobile API",
    description="API REST para consulta y creación individual de contratos de arrendamiento financiero",
    version="2.0.0",
    docs_url="/docs" if _DOCS_ON else None,
    redoc_url=None,
    openapi_url="/openapi.json" if _DOCS_ON else None,
)

# La app móvil nativa no necesita CORS. Solo se habilitan orígenes web
# si se declaran explícitamente en OLEASING_CORS_ORIGINS (separados por coma).
_cors = [o.strip() for o in os.environ.get("OLEASING_CORS_ORIGINS", "").split(",") if o.strip()]
if _cors:
    app.add_middleware(
        CORSMiddleware,
        allow_origins=_cors,
        allow_credentials=False,
        allow_methods=["GET", "POST", "PUT"],
        allow_headers=["Authorization", "Content-Type"],
    )


@app.middleware("http")
async def _cabeceras_seguridad(request: Request, call_next):
    resp = await call_next(request)
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["X-Frame-Options"] = "DENY"
    resp.headers["Referrer-Policy"] = "no-referrer"
    resp.headers["Cache-Control"] = "no-store"
    return resp


DB_PATH = os.path.join(_BASE_DIR, "leasing.db")
AUTH_DB = get_auth_db_path(os.path.join(_BASE_DIR, "data"))
init_auth_db(AUTH_DB)

# Duración de la sesión móvil (días). Al expirar, la app pide login otra vez.
SESION_MOVIL_DIAS = int(os.environ.get("OLEASING_SESION_DIAS", "7"))

# Llave para firmar enlaces temporales de PDF. Se genera una sola vez y se
# guarda en data/ (que nunca se sube al repositorio).
_SECRET_PATH = os.path.join(_BASE_DIR, "data", "api_secret.key")


def _cargar_secreto() -> bytes:
    try:
        if os.path.exists(_SECRET_PATH):
            with open(_SECRET_PATH, "rb") as f:
                data = f.read().strip()
                if len(data) >= 32:
                    return data
        os.makedirs(os.path.dirname(_SECRET_PATH), exist_ok=True)
        data = secrets.token_hex(32).encode()
        with open(_SECRET_PATH, "wb") as f:
            f.write(data)
        return data
    except Exception:
        return secrets.token_hex(32).encode()


_API_SECRET = _cargar_secreto()

# --- Protección contra fuerza bruta en el login -----------------------------
_INTENTOS_MAX = 5
_VENTANA_SEG = 15 * 60
_intentos: dict = {}
_intentos_lock = threading.Lock()


def _clave_intento(ip: str, usuario: str) -> str:
    return f"{ip}|{(usuario or '').strip().lower()}"


def _bloqueado(clave: str) -> int:
    ahora = time.time()
    with _intentos_lock:
        lst = [t for t in _intentos.get(clave, []) if ahora - t < _VENTANA_SEG]
        _intentos[clave] = lst
        if len(lst) >= _INTENTOS_MAX:
            return int(_VENTANA_SEG - (ahora - lst[0]))
    return 0


def _registrar_fallo(clave: str):
    with _intentos_lock:
        _intentos.setdefault(clave, []).append(time.time())


def _limpiar_intentos(clave: str):
    with _intentos_lock:
        _intentos.pop(clave, None)


# --- Dependencias de autenticación y permisos -------------------------------
_bearer = HTTPBearer(auto_error=False)


def usuario_actual(cred: Optional[HTTPAuthorizationCredentials] = Depends(_bearer)) -> dict:
    if not cred or not cred.credentials:
        raise HTTPException(status_code=401, detail="Sesión requerida. Inicia sesión en la app.")
    user = validar_sesion(AUTH_DB, cred.credentials)
    if not user:
        raise HTTPException(status_code=401, detail="Tu sesión expiró. Vuelve a iniciar sesión.")
    user["_token"] = cred.credentials
    return user


def usuario_escritura(user: dict = Depends(usuario_actual)) -> dict:
    if user.get("rol") not in ("admin", "captura"):
        raise HTTPException(status_code=403, detail="Tu usuario es de solo lectura.")
    return user


def _firmar_pdf(id_contrato: str, expira: int) -> str:
    msg = f"{id_contrato}|{expira}".encode()
    return hmac.new(_API_SECRET, msg, hashlib.sha256).hexdigest()


def get_db_connection():
    """Retorna una conexión a la base de datos leasing.db con row_factory como dict."""
    if not os.path.exists(DB_PATH):
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Base de datos no disponible."
        )
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


# =============================================================================
# MODELOS PYDANTIC
# =============================================================================

class LoginRequest(BaseModel):
    username: str = Field(..., min_length=1, max_length=60)
    password: str = Field(..., min_length=1, max_length=120)
    dias_sesion: int = Field(14, ge=1, le=30)


class UsuarioResponse(BaseModel):
    username: str
    nombre_completo: str
    rol: str
    puede_editar: bool


class LoginResponse(BaseModel):
    token: str
    expira_dias: int
    usuario: UsuarioResponse


class CambiarPasswordRequest(BaseModel):
    password_actual: str = Field(..., min_length=1, max_length=120)
    password_nueva: str = Field(..., min_length=6, max_length=120)


class ContratoResumen(BaseModel):
    id_contrato: str
    cliente: str
    vehiculo: Optional[str] = ""
    fecha_alta: Optional[str] = ""
    fecha_vencimiento: Optional[str] = ""
    estatus: str
    plazo: int
    valor_sin_iva: float
    anticipo_monto: float
    mensualidad_sin_iva: float
    residual_monto: float
    tasa_mensual_pct: float
    tasa_anual_pct: float
    nivel_morosidad: int = 0


class ContratoDetalle(BaseModel):
    id_contrato: str
    cliente: str
    vehiculo: Optional[str] = ""
    fecha_alta: Optional[str] = ""
    fecha_vencimiento: Optional[str] = ""
    fecha_baja: Optional[str] = None
    estatus: str
    plazo: int
    valor_sin_iva: float
    anticipo_pct: float
    anticipo_monto: float
    inversion_neta: float
    mensualidad_sin_iva: float
    residual_pct: float
    residual_monto: float
    vp_residual: float
    comision_apertura_pct: float
    comision_monto: float
    tasa_mensual_pct: float
    tasa_anual_pct: float
    anotaciones: Optional[str] = ""
    saldo_actual_estimado: float
    intereses_totales_estimados: float
    nivel_morosidad: int = 0


class PeriodoAmortizacion(BaseModel):
    periodo: int
    fecha_estimada: str
    renta_sin_iva: float
    capital: float
    interes: float
    iva: float
    total_con_iva: float
    saldo_insoluto: float
    estatus: str


class EstadoCuentaResponse(BaseModel):
    id_contrato: str
    cliente: str
    vehiculo: str
    estatus: str
    fecha_alta: str
    fecha_vencimiento: str
    plazo: int
    valor_sin_iva: float
    anticipo_monto: float
    inversion_neta: float
    mensualidad_sin_iva: float
    residual_monto: float
    tasa_mensual_pct: float
    tasa_anual_pct: float
    meses_transcurridos: int
    pagos_restantes: int
    saldo_actual_estimado: float
    capital_amortizado: float
    interes_devengado: float
    total_intereses_plazo: float
    porcentaje_avance: float
    tabla_amortizacion: List[PeriodoAmortizacion]


class CotizacionRequest(BaseModel):
    valor_sin_iva: float = Field(..., gt=0, description="Valor del activo sin IVA")
    plazo: int = Field(..., gt=0, le=120, description="Plazo en meses")
    mensualidad_sin_iva: Optional[float] = Field(None, description="Renta mensual sin IVA (si se calcula tasa)")
    tasa_anual_pct: Optional[float] = Field(None, description="Tasa anualizada en porcentaje (si se calcula renta)")
    anticipo_monto: Optional[float] = Field(0.0, ge=0)
    anticipo_pct: Optional[float] = Field(0.0, ge=0)
    residual_monto: Optional[float] = Field(0.0, ge=0)
    residual_pct: Optional[float] = Field(0.0, ge=0)
    comision_pct: Optional[float] = Field(0.0, ge=0)


class CotizacionResponse(BaseModel):
    valor_sin_iva: float
    anticipo_monto: float
    inversion_neta: float
    plazo: int
    mensualidad_sin_iva: float
    residual_monto: float
    vp_residual: float
    tasa_mensual_pct: float
    tasa_anual_pct: float
    comision_monto: float
    total_intereses: float
    total_rentas: float
    costo_total_leasing: float


class ContratoCreate(BaseModel):
    id_contrato: str = Field(..., min_length=3, max_length=20, description="Identificador único, ej. 0729-0001")
    cliente: str = Field(..., min_length=2, max_length=200, description="Nombre o razón social del cliente")
    vehiculo: Optional[str] = Field("", max_length=200, description="Descripción del activo o vehículo")
    fecha_alta: str = Field(..., pattern=r"^\d{4}-\d{2}-\d{2}$", description="Fecha de alta en formato YYYY-MM-DD")
    plazo: int = Field(..., gt=0, le=120, description="Plazo en meses")
    valor_sin_iva: float = Field(..., gt=0, description="Valor del activo sin IVA")
    anticipo_monto: Optional[float] = Field(0.0, ge=0, description="Monto del anticipo sin IVA")
    mensualidad_sin_iva: float = Field(..., gt=0, description="Renta mensual pactada sin IVA")
    residual_monto: Optional[float] = Field(0.0, ge=0, description="Monto residual pactado sin IVA")
    comision_pct: Optional[float] = Field(0.0, ge=0, description="Porcentaje de comisión por apertura")
    anotaciones: Optional[str] = Field("", max_length=500, description="Comentarios u observaciones operativas")


class ActualizarEstatusRequest(BaseModel):
    estatus: str = Field(..., description="ACTIVO, CONCLUIDO o BAJA")
    nivel_morosidad: Optional[int] = Field(0, ge=0, le=4)
    fecha_baja: Optional[str] = Field(None)


class AnotacionItem(BaseModel):
    id: int
    id_contrato: str
    fecha: str
    tipo: str
    texto: str


class AnotacionCreateRequest(BaseModel):
    tipo: str = Field("General", max_length=50)
    texto: str = Field(..., min_length=1, max_length=1000)


class CobranzaContratoItem(BaseModel):
    id_contrato: str
    cliente: str
    vehiculo: str
    mes_contrato: int
    plazo: int
    renta_sin_iva: float
    iva: float
    total_con_iva: float
    capital_mes: float
    interes_leasing_mes: float
    saldo_insoluto: float


class CobranzaMesResponse(BaseModel):
    anio: int
    mes: int
    periodo: str
    contratos_en_cobro: int
    renta_total_sin_iva: float
    iva_total: float
    total_cobranza_con_iva: float
    capital_total_mes: float
    interes_total_mes: float
    items: List[CobranzaContratoItem]


class ClienteResumen(BaseModel):
    cliente: str
    contratos_activos: int
    total_contratos: int
    inversion_activa: float
    renta_mensual_total: float
    residual_total: float


class ContratoUpdate(BaseModel):
    cliente: str = Field(..., min_length=2, max_length=200)
    vehiculo: Optional[str] = Field("", max_length=200)
    fecha_alta: str = Field(..., pattern=r"^\d{4}-\d{2}-\d{2}$")
    plazo: int = Field(..., gt=0, le=120)
    valor_sin_iva: float = Field(..., gt=0)
    anticipo_monto: Optional[float] = Field(0.0, ge=0)
    mensualidad_sin_iva: float = Field(..., gt=0)
    residual_monto: Optional[float] = Field(0.0, ge=0)
    comision_pct: Optional[float] = Field(0.0, ge=0)
    anotaciones: Optional[str] = Field("", max_length=500)



# =============================================================================
# ENDPOINTS DE AUTENTICACIÓN Y CONSULTA
# =============================================================================

@app.post("/api/auth/login", response_model=LoginResponse)
def login_movil(req: LoginRequest, request: Request):
    """Autentica al usuario contra usuarios.db con protección anti-fuerza bruta y emite token de sesión."""
    ip = request.client.host if request.client else "unknown"
    clave_ip = f"{ip}:{req.username.strip().lower()}"
    if _bloqueado(clave_ip):
        raise HTTPException(
            status_code=429,
            detail="Demasiados intentos fallidos. Espera 15 minutos antes de volver a intentar."
        )

    user = verificar_login(AUTH_DB, req.username.strip(), req.password)
    if not user:
        _registrar_fallo(clave_ip)
        raise HTTPException(status_code=401, detail="Usuario o contraseña incorrectos.")

    _limpiar_intentos(clave_ip)
    token = crear_sesion(AUTH_DB, user["id"], dias=req.dias_sesion)
    rol = user.get("rol", "lectura")
    return LoginResponse(
        token=token,
        expira_dias=req.dias_sesion,
        usuario=UsuarioResponse(
            username=user["username"],
            nombre_completo=user.get("nombre_completo") or user["username"],
            rol=rol,
            puede_editar=(rol in ("admin", "captura"))
        )
    )


@app.get("/api/auth/me", response_model=UsuarioResponse)
def sesion_actual(user: dict = Depends(usuario_actual)):
    """Retorna la información del usuario autenticado."""
    rol = user.get("rol", "lectura")
    return UsuarioResponse(
        username=user["username"],
        nombre_completo=user.get("nombre_completo") or user["username"],
        rol=rol,
        puede_editar=(rol in ("admin", "captura"))
    )


@app.post("/api/auth/logout")
def logout_movil(user: dict = Depends(usuario_actual)):
    """Revoca el token de sesión actual en la base de datos."""
    cerrar_sesion(AUTH_DB, user.get("_token", ""))
    return {"ok": True}


@app.post("/api/auth/cambiar_password")
def cambiar_password_movil(req: CambiarPasswordRequest, user: dict = Depends(usuario_actual)):
    """Permite al usuario autenticado actualizar su contraseña corporativa validando la actual."""
    ok, msg = resetear_password(
        AUTH_DB,
        user["id"],
        req.password_nueva,
        password_actual=req.password_actual,
        require_actual=True
    )
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {"ok": True, "message": msg}


@app.get("/api/health")
def health_check():
    """Verifica estado del servicio (sin exponer datos sensibles)."""
    return {
        "status": "online",
        "service": "O-Leasing Mobile API",
        "auth_required": True,
        "timestamp": datetime.datetime.now().isoformat()
    }


@app.get("/api/dashboard")
def get_dashboard_metrics(user: dict = Depends(usuario_actual)):
    """Retorna las métricas clave del portafolio y alertas de vencimiento para la pantalla de inicio móvil."""
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT COUNT(*),
               SUM(Valor_Sin_IVA - Anticipo_Monto),
               SUM(Mensualidad_Sin_IVA),
               SUM(Residual_Monto)
        FROM contratos
        WHERE Estatus = 'ACTIVO'
    """)
    row_act = cursor.fetchone()
    n_activos = row_act[0] or 0
    cartera_activa_monto = row_act[1] or 0.0
    renta_mensual_total = row_act[2] or 0.0
    residual_total_activo = row_act[3] or 0.0

    cursor.execute("SELECT COUNT(*) FROM contratos WHERE Estatus = 'CONCLUIDO'")
    n_concluidos = cursor.fetchone()[0] or 0

    cursor.execute("SELECT COUNT(*) FROM contratos WHERE Estatus = 'BAJA'")
    n_bajas = cursor.fetchone()[0] or 0

    cursor.execute("SELECT COUNT(*) FROM contratos")
    total_historico = cursor.fetchone()[0] or 0

    cursor.execute("SELECT COUNT(*) FROM contratos WHERE Estatus = 'ACTIVO' AND COALESCE(Nivel_Morosidad, 0) > 0")
    contratos_en_mora = cursor.fetchone()[0] or 0

    hoy = datetime.date.today()
    limite_90d = (hoy + datetime.timedelta(days=90)).isoformat()
    hoy_str = hoy.isoformat()

    cursor.execute("""
        SELECT ID_Contrato, Cliente, Vehiculo, Fecha_Vencimiento, Residual_Monto, Mensualidad_Sin_IVA
        FROM contratos
        WHERE Estatus = 'ACTIVO'
          AND Fecha_Vencimiento IS NOT NULL
          AND substr(Fecha_Vencimiento, 1, 10) >= ?
          AND substr(Fecha_Vencimiento, 1, 10) <= ?
        ORDER BY substr(Fecha_Vencimiento, 1, 10) ASC
    """, (hoy_str, limite_90d))
    rows_venc = cursor.fetchall()
    contratos_por_vencer_90d = len(rows_venc)

    vencimientos_proximos = []
    for rv in rows_venc[:6]:
        fv_s = str(rv["Fecha_Vencimiento"] or "")[:10]
        try:
            fv_d = datetime.date.fromisoformat(fv_s)
            dias_rest = (fv_d - hoy).days
        except Exception:
            dias_rest = 0
        vencimientos_proximos.append({
            "id_contrato": rv["ID_Contrato"],
            "cliente": rv["Cliente"],
            "vehiculo": rv["Vehiculo"] or "",
            "fecha_vencimiento": fv_s,
            "dias_restantes": max(dias_rest, 0),
            "residual_monto": round(float(rv["Residual_Monto"] or 0.0), 2),
            "mensualidad_sin_iva": round(float(rv["Mensualidad_Sin_IVA"] or 0.0), 2),
        })

    conn.close()

    return {
        "total_contratos": total_historico,
        "total_historico": total_historico,
        "contratos_activos": n_activos,
        "contratos_concluidos": n_concluidos,
        "contratos_baja": n_bajas,
        "contratos_bajas": n_bajas,
        "cartera_activa_valor": round(cartera_activa_monto, 2),
        "cartera_activa_monto": round(cartera_activa_monto, 2),
        "flujo_mensual_estimado": round(renta_mensual_total, 2),
        "flujo_mensual_rentas": round(renta_mensual_total, 2),
        "residual_total_activo": round(residual_total_activo, 2),
        "contratos_por_vencer_90d": contratos_por_vencer_90d,
        "contratos_en_mora": contratos_en_mora,
        "vencimientos_proximos": vencimientos_proximos,
        "fecha_corte": hoy_str
    }


@app.get("/api/contratos", response_model=List[ContratoResumen])
def listar_contratos(
    q: Optional[str] = Query(None, description="Texto de búsqueda por ID, Cliente o Vehículo"),
    estatus: Optional[str] = Query("ACTIVO", description="Filtro por estatus: ACTIVO, CONCLUIDO, BAJA, MORA o TODOS"),
    limit: int = Query(50, ge=1, le=200),
    offset: int = Query(0, ge=0),
    user: dict = Depends(usuario_actual)
):
    """Consulta paginada y filtrable de contratos para la lista principal de la app."""
    conn = get_db_connection()
    cursor = conn.cursor()

    sql = """
        SELECT ID_Contrato, Cliente, Vehiculo, Fecha_Alta, Fecha_Vencimiento,
               Estatus, Plazo, Valor_Sin_IVA, Anticipo_Monto, Mensualidad_Sin_IVA,
               Residual_Monto, Tasa_Calculada, COALESCE(Nivel_Morosidad, 0) AS Nivel_Morosidad
        FROM contratos
        WHERE 1=1
    """
    params = []

    if estatus:
        est_up = estatus.upper()
        if est_up == "MORA":
            sql += " AND COALESCE(Nivel_Morosidad, 0) > 0"
        elif est_up != "TODOS":
            sql += " AND Estatus = ?"
            params.append(est_up)

    if q:
        sql += " AND (ID_Contrato LIKE ? OR Cliente LIKE ? OR Vehiculo LIKE ?)"
        wildcard = f"%{q.strip()}%"
        params.extend([wildcard, wildcard, wildcard])

    sql += " ORDER BY Fecha_Alta DESC, ID_Contrato DESC LIMIT ? OFFSET ?"
    params.extend([limit, offset])

    cursor.execute(sql, params)
    rows = cursor.fetchall()
    conn.close()

    resultado = []
    for r in rows:
        tasa_m = float(r['Tasa_Calculada'] or 0.0)
        resultado.append(ContratoResumen(
            id_contrato=r['ID_Contrato'],
            cliente=r['Cliente'],
            vehiculo=r['Vehiculo'] or "",
            fecha_alta=str(r['Fecha_Alta'] or "")[:10],
            fecha_vencimiento=str(r['Fecha_Vencimiento'] or "")[:10],
            estatus=r['Estatus'] or "ACTIVO",
            plazo=int(r['Plazo'] or 0),
            valor_sin_iva=round(float(r['Valor_Sin_IVA'] or 0.0), 2),
            anticipo_monto=round(float(r['Anticipo_Monto'] or 0.0), 2),
            mensualidad_sin_iva=round(float(r['Mensualidad_Sin_IVA'] or 0.0), 2),
            residual_monto=round(float(r['Residual_Monto'] or 0.0), 2),
            tasa_mensual_pct=round(tasa_m * 100, 2),
            tasa_anual_pct=round(tasa_m * 12 * 100, 2),
            nivel_morosidad=int(r['Nivel_Morosidad'] or 0),
        ))

    return resultado


@app.get("/api/contratos/{id_contrato}", response_model=ContratoDetalle)
def obtener_detalle_contrato(id_contrato: str, user: dict = Depends(usuario_actual)):
    """Consulta a fondo de un contrato individual con desglose financiero de amortización."""
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT ID_Contrato, Cliente, Vehiculo, Fecha_Alta, Fecha_Vencimiento, Fecha_Baja,
               Estatus, Plazo, Valor_Sin_IVA, Anticipo_Pct, Anticipo_Monto,
               Mensualidad_Sin_IVA, Residual_Pct, Residual_Monto, VP_Residual,
               Comision_Apertura_Pct, Comision_Monto, Tasa_Calculada, Anotaciones,
               COALESCE(Nivel_Morosidad, 0) AS Nivel_Morosidad
        FROM contratos
        WHERE ID_Contrato = ?
    """, (id_contrato,))
    row = cursor.fetchone()
    conn.close()

    if not row:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f"Contrato con ID '{id_contrato}' no encontrado."
        )

    valor = float(row['Valor_Sin_IVA'] or 0.0)
    ant = float(row['Anticipo_Monto'] or 0.0)
    inv = valor - ant
    renta = float(row['Mensualidad_Sin_IVA'] or 0.0)
    plazo = int(row['Plazo'] or 1)
    res = float(row['Residual_Monto'] or 0.0)
    tasa = float(row['Tasa_Calculada'] or 0.0)
    vp = float(row['VP_Residual'] or 0.0) or (vp_res(res, tasa, plazo) if tasa else 0.0)

    # Calcular amortización estimada
    intereses_totales = 0.0
    saldo_actual = inv
    if tasa > 0 and inv > 0 and renta > 0:
        try:
            df_a, _, _, _, _ = calc_amort(round(inv, 4), round(renta, 4), round(res, 4), plazo, round(tasa, 8))
            intereses_totales = float(df_a['Interes'].sum())

            # Estimación del saldo al mes actual
            fa = datetime.datetime.strptime(str(row['Fecha_Alta'])[:10], "%Y-%m-%d")
            hoy = datetime.date.today()
            mc = (hoy.year - fa.year) * 12 + (hoy.month - fa.month) + 1
            if mc <= 0:
                saldo_actual = inv
            elif mc >= plazo:
                saldo_actual = res
            else:
                saldo_actual = float(df_a.iloc[mc - 1]['Saldo'])
        except Exception:
            saldo_actual = inv

    return ContratoDetalle(
        id_contrato=row['ID_Contrato'],
        cliente=row['Cliente'],
        vehiculo=row['Vehiculo'] or "",
        fecha_alta=str(row['Fecha_Alta'] or "")[:10],
        fecha_vencimiento=str(row['Fecha_Vencimiento'] or "")[:10],
        fecha_baja=str(row['Fecha_Baja'] or "")[:10] if row['Fecha_Baja'] else None,
        estatus=row['Estatus'] or "ACTIVO",
        plazo=plazo,
        valor_sin_iva=round(valor, 2),
        anticipo_pct=round(float(row['Anticipo_Pct'] or 0.0), 2),
        anticipo_monto=round(ant, 2),
        inversion_neta=round(inv, 2),
        mensualidad_sin_iva=round(renta, 2),
        residual_pct=round(float(row['Residual_Pct'] or 0.0), 2),
        residual_monto=round(res, 2),
        vp_residual=round(vp, 2),
        comision_apertura_pct=round(float(row['Comision_Apertura_Pct'] or 0.0), 2),
        comision_monto=round(float(row['Comision_Monto'] or 0.0), 2),
        tasa_mensual_pct=round(tasa * 100, 2),
        tasa_anual_pct=round(tasa * 12 * 100, 2),
        anotaciones=row['Anotaciones'] or "",
        saldo_actual_estimado=round(saldo_actual, 2),
        intereses_totales_estimados=round(intereses_totales, 2),
        nivel_morosidad=int(row['Nivel_Morosidad'] or 0)
    )


@app.get("/api/contratos/{id_contrato}/estado_cuenta", response_model=EstadoCuentaResponse)
def obtener_estado_cuenta(id_contrato: str, user: dict = Depends(usuario_actual)):
    """Retorna el estado de cuenta completo y la tabla de amortización periódica."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT ID_Contrato, Cliente, Vehiculo, Fecha_Alta, Fecha_Vencimiento, Fecha_Baja,
               Estatus, Plazo, Valor_Sin_IVA, Anticipo_Pct, Anticipo_Monto,
               Mensualidad_Sin_IVA, Residual_Pct, Residual_Monto, VP_Residual,
               Comision_Apertura_Pct, Comision_Monto, Tasa_Calculada, Anotaciones
        FROM contratos
        WHERE ID_Contrato = ?
    """, (id_contrato,))
    row = cursor.fetchone()
    conn.close()

    if not row:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f"Contrato con ID '{id_contrato}' no encontrado."
        )

    valor = float(row['Valor_Sin_IVA'] or 0.0)
    ant = float(row['Anticipo_Monto'] or 0.0)
    inv = valor - ant
    renta = float(row['Mensualidad_Sin_IVA'] or 0.0)
    plazo = int(row['Plazo'] or 1)
    res = float(row['Residual_Monto'] or 0.0)
    tasa = float(row['Tasa_Calculada'] or 0.0)

    from dateutil.relativedelta import relativedelta

    fa_str = str(row['Fecha_Alta'] or "")[:10]
    fa = datetime.datetime.strptime(fa_str, "%Y-%m-%d") if fa_str else datetime.datetime.now()
    hoy = datetime.date.today()
    meses_transcurridos = max((hoy.year - fa.year) * 12 + (hoy.month - fa.month), 0)
    pagos_restantes = max(plazo - meses_transcurridos, 0)
    porcentaje_avance = round(min(meses_transcurridos / max(plazo, 1) * 100, 100.0), 1)

    tabla = []
    capital_amortizado = 0.0
    interes_devengado = 0.0
    total_intereses = 0.0
    saldo_actual = inv

    if tasa > 0 and inv > 0 and renta > 0:
        try:
            df_a, _, _, _, _ = calc_amort(round(inv, 4), round(renta, 4), round(res, 4), plazo, round(tasa, 8))
            total_intereses = float(df_a['Interes'].sum())
            for idx, r_a in df_a.iterrows():
                p = int(r_a['Mes'])
                cap = float(r_a['Capital'])
                inte = float(r_a['Interes'])
                r_sin = cap + inte
                r_iva = round(r_sin * 0.16, 2)
                tot = round(r_sin + r_iva, 2)
                saldo = float(r_a['Saldo'])
                f_pago = (fa + relativedelta(months=p)).strftime("%Y-%m-%d")

                if p < meses_transcurridos:
                    st_pago = "PAGADO"
                    capital_amortizado += cap
                    interes_devengado += inte
                elif p == meses_transcurridos:
                    st_pago = "ACTUAL"
                    capital_amortizado += cap
                    interes_devengado += inte
                else:
                    st_pago = "PENDIENTE"

                tabla.append(PeriodoAmortizacion(
                    periodo=p,
                    fecha_estimada=f_pago,
                    renta_sin_iva=round(r_sin, 2),
                    capital=round(cap, 2),
                    interes=round(inte, 2),
                    iva=r_iva,
                    total_con_iva=tot,
                    saldo_insoluto=round(saldo, 2),
                    estatus=st_pago
                ))

            idx_s = min(meses_transcurridos, plazo) - 1
            if idx_s >= 0:
                saldo_actual = float(df_a.iloc[idx_s]['Saldo'])
        except Exception:
            pass

    return EstadoCuentaResponse(
        id_contrato=row['ID_Contrato'],
        cliente=row['Cliente'],
        vehiculo=row['Vehiculo'] or "",
        estatus=row['Estatus'] or "ACTIVO",
        fecha_alta=str(row['Fecha_Alta'] or "")[:10],
        fecha_vencimiento=str(row['Fecha_Vencimiento'] or "")[:10],
        plazo=plazo,
        valor_sin_iva=round(valor, 2),
        anticipo_monto=round(ant, 2),
        inversion_neta=round(inv, 2),
        mensualidad_sin_iva=round(renta, 2),
        residual_monto=round(res, 2),
        tasa_mensual_pct=round(tasa * 100, 2),
        tasa_anual_pct=round(tasa * 12 * 100, 2),
        meses_transcurridos=meses_transcurridos,
        pagos_restantes=pagos_restantes,
        saldo_actual_estimado=round(saldo_actual, 2),
        capital_amortizado=round(capital_amortizado, 2),
        interes_devengado=round(interes_devengado, 2),
        total_intereses_plazo=round(total_intereses, 2),
        porcentaje_avance=porcentaje_avance,
        tabla_amortizacion=tabla
    )


# =============================================================================
# ENDPOINTS DE COTIZACIÓN Y CREACIÓN DE CONTRATOS
# =============================================================================

@app.post("/api/cotizar", response_model=CotizacionResponse)
def cotizar_leasing(req: CotizacionRequest, user: dict = Depends(usuario_actual)):
    """
    Calculadora financiera instantánea con las fórmulas exactas de O-Leasing.
    Permite calcular la tasa implícita a partir de la renta, o la renta a partir de la tasa.
    """
    valor = req.valor_sin_iva
    plazo = req.plazo

    # Determinar anticipo
    anticipo = req.anticipo_monto or 0.0
    if req.anticipo_pct and req.anticipo_pct > 0 and anticipo == 0.0:
        anticipo = valor * (req.anticipo_pct / 100.0)

    inversion = valor - anticipo
    if inversion <= 0:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="La inversión neta (Valor - Anticipo) debe ser mayor a cero."
        )

    # Determinar residual
    residual = req.residual_monto or 0.0
    if req.residual_pct and req.residual_pct > 0 and residual == 0.0:
        residual = valor * (req.residual_pct / 100.0)

    renta = req.mensualidad_sin_iva
    tasa_anual = req.tasa_anual_pct

    if renta is not None and renta > 0:
        # Calcular tasa implícita con numpy_financial
        tasa_mensual = calc_tasa(plazo, renta, inversion, residual)
        if tasa_mensual is None or tasa_mensual < 0:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Los parámetros ingresados no forman un flujo financiero viable (la renta es insuficiente)."
            )
        tasa_anual = tasa_mensual * 12 * 100
    elif tasa_anual is not None and tasa_anual > 0:
        # Calcular renta a partir de la tasa
        tasa_mensual = (tasa_anual / 100.0) / 12.0
        import numpy_financial as npf
        renta = -npf.pmt(tasa_mensual, plazo, inversion, -residual)
    else:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Debe proporcionar la mensualidad_sin_iva o la tasa_anual_pct."
        )

    vp = vp_res(residual, tasa_mensual, plazo)
    comision_monto = valor * ((req.comision_pct or 0.0) / 100.0)

    # Calcular intereses totales
    df_a, _, _, _, _ = calc_amort(round(inversion, 4), round(renta, 4), round(residual, 4), plazo, round(tasa_mensual, 8))
    total_intereses = float(df_a['Interes'].sum())
    total_rentas = renta * plazo
    costo_total = total_rentas + residual + anticipo + comision_monto

    return CotizacionResponse(
        valor_sin_iva=round(valor, 2),
        anticipo_monto=round(anticipo, 2),
        inversion_neta=round(inversion, 2),
        plazo=plazo,
        mensualidad_sin_iva=round(renta, 2),
        residual_monto=round(residual, 2),
        vp_residual=round(vp, 2),
        tasa_mensual_pct=round(tasa_mensual * 100, 2),
        tasa_anual_pct=round(tasa_anual, 2),
        comision_monto=round(comision_monto, 2),
        total_intereses=round(total_intereses, 2),
        total_rentas=round(total_rentas, 2),
        costo_total_leasing=round(costo_total, 2)
    )


@app.post("/api/contratos", response_model=ContratoDetalle, status_code=status.HTTP_201_CREATED)
def crear_contrato_individual(contrato: ContratoCreate, user: dict = Depends(usuario_escritura)):
    """
    Crea un contrato nuevo de forma individual en leasing.db.
    Calcula tasa implícita, valor presente del residual y fecha de vencimiento automáticamente.
    """
    conn = get_db_connection()
    cursor = conn.cursor()

    # Validar que no exista el ID
    cursor.execute("SELECT ID_Contrato FROM contratos WHERE ID_Contrato = ?", (contrato.id_contrato.strip(),))
    if cursor.fetchone():
        conn.close()
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=f"El contrato '{contrato.id_contrato}' ya existe en la base de datos."
        )

    valor = contrato.valor_sin_iva
    ant = contrato.anticipo_monto or 0.0
    inv = valor - ant
    if inv <= 0:
        conn.close()
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="La inversión neta (Valor - Anticipo) debe ser positiva."
        )

    res = contrato.residual_monto or 0.0
    renta = contrato.mensualidad_sin_iva
    plazo = contrato.plazo

    # Calcular tasa implícita
    tasa = calc_tasa(plazo, renta, inv, res)
    if tasa is None or tasa <= 0:
        conn.close()
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Los valores ingresados no forman un flujo financiero válido. Verifique la renta y el plazo."
        )

    vp = vp_res(res, tasa, plazo)
    ant_pct = round((ant / valor) * 100, 2) if valor else 0.0
    res_pct = round((res / valor) * 100, 2) if valor else 0.0
    com_pct = contrato.comision_pct or 0.0
    com_monto = round(valor * (com_pct / 100.0), 2)

    # Calcular fecha de vencimiento
    fa_dt = datetime.datetime.strptime(contrato.fecha_alta, "%Y-%m-%d")
    # Sumar plazo meses aproximadamente
    m_venc = fa_dt.month + plazo
    y_venc = fa_dt.year + (m_venc - 1) // 12
    m_venc = (m_venc - 1) % 12 + 1
    d_venc = min(fa_dt.day, 28)
    fv_str = f"{y_venc:04d}-{m_venc:02d}-{d_venc:02d}"

    try:
        cursor.execute("""
            INSERT INTO contratos (
                ID_Contrato, Cliente, Vehiculo, Fecha_Alta, Fecha_Vencimiento,
                Valor_Sin_IVA, Mensualidad_Sin_IVA, Plazo,
                Comision_Apertura_Pct, Comision_Monto,
                Anticipo_Pct, Anticipo_Monto,
                Residual_Pct, Residual_Monto,
                Tasa_Calculada, Estatus, VP_Residual,
                Anotaciones, residual_transferred, Nivel_Morosidad
            ) VALUES (
                ?, ?, ?, ?, ?,
                ?, ?, ?,
                ?, ?,
                ?, ?,
                ?, ?,
                ?, 'ACTIVO', ?,
                ?, 0, 0
            )
        """, (
            contrato.id_contrato.strip(),
            contrato.cliente.strip(),
            contrato.vehiculo.strip() if contrato.vehiculo else "",
            contrato.fecha_alta,
            fv_str,
            valor,
            renta,
            plazo,
            com_pct,
            com_monto,
            ant_pct,
            ant,
            res_pct,
            res,
            round(tasa, 8),
            round(vp, 2),
            contrato.anotaciones or ""
        ))
        conn.commit()
    except Exception as e:
        conn.rollback()
        conn.close()
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error al guardar contrato en la base de datos: {str(e)}"
        )

    conn.close()

    # Retornar detalle del contrato recién creado
    return obtener_detalle_contrato(contrato.id_contrato.strip(), user=user)


@app.put("/api/contratos/{id_contrato}", response_model=ContratoDetalle)
def editar_contrato_individual(id_contrato: str, datos: ContratoUpdate, user: dict = Depends(usuario_escritura)):
    """Actualiza los parámetros financieros y generales de un contrato y recalcula tasa, VP residual y vencimiento."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT ID_Contrato FROM contratos WHERE ID_Contrato = ?", (id_contrato,))
    if not cursor.fetchone():
        conn.close()
        raise HTTPException(status_code=404, detail="Contrato no encontrado.")

    valor = datos.valor_sin_iva
    ant = datos.anticipo_monto or 0.0
    inv = valor - ant
    if inv <= 0:
        conn.close()
        raise HTTPException(status_code=400, detail="La inversión neta (Valor - Anticipo) debe ser positiva.")

    res = datos.residual_monto or 0.0
    renta = datos.mensualidad_sin_iva
    plazo = datos.plazo

    tasa = calc_tasa(plazo, renta, inv, res)
    if tasa is None or tasa <= 0:
        conn.close()
        raise HTTPException(status_code=400, detail="Los valores ingresados no forman un flujo financiero válido.")

    vp = vp_res(res, tasa, plazo)
    ant_pct = round((ant / valor) * 100, 2) if valor else 0.0
    res_pct = round((res / valor) * 100, 2) if valor else 0.0
    com_pct = datos.comision_pct or 0.0
    com_monto = round(valor * (com_pct / 100.0), 2)

    fa_dt = datetime.datetime.strptime(datos.fecha_alta, "%Y-%m-%d")
    m_venc = fa_dt.month + plazo
    y_venc = fa_dt.year + (m_venc - 1) // 12
    m_venc = (m_venc - 1) % 12 + 1
    d_venc = min(fa_dt.day, 28)
    fv_str = f"{y_venc:04d}-{m_venc:02d}-{d_venc:02d}"

    cursor.execute("""
        UPDATE contratos
        SET Cliente = ?,
            Vehiculo = ?,
            Fecha_Alta = ?,
            Fecha_Vencimiento = ?,
            Plazo = ?,
            Valor_Sin_IVA = ?,
            Anticipo_Pct = ?,
            Anticipo_Monto = ?,
            Mensualidad_Sin_IVA = ?,
            Residual_Pct = ?,
            Residual_Monto = ?,
            VP_Residual = ?,
            Comision_Apertura_Pct = ?,
            Comision_Monto = ?,
            Tasa_Calculada = ?,
            Anotaciones = ?
        WHERE ID_Contrato = ?
    """, (
        datos.cliente.strip(),
        (datos.vehiculo or "").strip(),
        datos.fecha_alta,
        fv_str,
        plazo,
        valor,
        ant_pct,
        ant,
        renta,
        res_pct,
        res,
        round(vp, 2),
        com_pct,
        com_monto,
        round(tasa, 8),
        datos.anotaciones or "",
        id_contrato
    ))

    # Registrar auditoría en bitácora
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS anotaciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ID_Contrato TEXT,
            Fecha TEXT,
            Tipo TEXT,
            Texto TEXT
        )
    """)
    autor = user.get("nombre_completo") or user.get("username") or "Usuario"
    ahora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    cursor.execute("""
        INSERT INTO anotaciones (ID_Contrato, Fecha, Tipo, Texto)
        VALUES (?, ?, 'Edición Móvil', ?)
    """, (id_contrato, ahora, f"[{autor}] Actualizó parámetros del contrato (Renta: ${renta:,.2f}, Plazo: {plazo}m)."))

    conn.commit()
    conn.close()
    return obtener_detalle_contrato(id_contrato, user=user)


@app.get("/api/clientes", response_model=List[ClienteResumen])
def listar_clientes_exposicion(
    q: Optional[str] = Query(None, description="Filtro por nombre de cliente"),
    user: dict = Depends(usuario_actual)
):
    """Agrupa la cartera por cliente mostrando contratos activos, inversión neta activa, renta mensual y residual."""
    conn = get_db_connection()
    cursor = conn.cursor()
    sql = """
        SELECT Cliente,
               SUM(CASE WHEN Estatus = 'ACTIVO' THEN 1 ELSE 0 END) AS activos,
               COUNT(*) AS total,
               SUM(CASE WHEN Estatus = 'ACTIVO' THEN (COALESCE(Valor_Sin_IVA,0) - COALESCE(Anticipo_Monto,0)) ELSE 0 END) AS inv_activa,
               SUM(CASE WHEN Estatus = 'ACTIVO' THEN COALESCE(Mensualidad_Sin_IVA,0) ELSE 0 END) AS renta_mensual,
               SUM(CASE WHEN Estatus = 'ACTIVO' THEN COALESCE(Residual_Monto,0) ELSE 0 END) AS residual_tot
        FROM contratos
        WHERE Cliente IS NOT NULL AND TRIM(Cliente) != ''
    """
    params = []
    if q and q.strip():
        sql += " AND Cliente LIKE ?"
        params.append(f"%{q.strip()}%")
    sql += " GROUP BY Cliente ORDER BY inv_activa DESC, activos DESC, total DESC LIMIT 100"
    cursor.execute(sql, params)
    rows = cursor.fetchall()
    conn.close()
    return [
        ClienteResumen(
            cliente=str(r["Cliente"]),
            contratos_activos=int(r["activos"] or 0),
            total_contratos=int(r["total"] or 0),
            inversion_activa=round(float(r["inv_activa"] or 0.0), 2),
            renta_mensual_total=round(float(r["renta_mensual"] or 0.0), 2),
            residual_total=round(float(r["residual_tot"] or 0.0), 2)
        )
        for r in rows
    ]


@app.put("/api/contratos/{id_contrato}/estatus", response_model=ContratoDetalle)
def actualizar_estatus_contrato(id_contrato: str, req: ActualizarEstatusRequest, user: dict = Depends(usuario_escritura)):
    """Actualiza el estatus (ACTIVO, CONCLUIDO, BAJA) y nivel de morosidad de un contrato."""
    est = req.estatus.strip().upper()
    if est not in ("ACTIVO", "CONCLUIDO", "BAJA"):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Estatus inválido. Use ACTIVO, CONCLUIDO o BAJA."
        )
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT ID_Contrato FROM contratos WHERE ID_Contrato = ?", (id_contrato,))
    if not cursor.fetchone():
        conn.close()
        raise HTTPException(status_code=404, detail="Contrato no encontrado.")

    fb = req.fecha_baja if est == "BAJA" else None
    if est == "BAJA" and not fb:
        fb = datetime.date.today().isoformat()

    cursor.execute("""
        UPDATE contratos
        SET Estatus = ?, Nivel_Morosidad = ?, Fecha_Baja = ?
        WHERE ID_Contrato = ?
    """, (est, req.nivel_morosidad or 0, fb, id_contrato))
    conn.commit()
    conn.close()
    return obtener_detalle_contrato(id_contrato, user=user)


@app.get("/api/contratos/{id_contrato}/anotaciones", response_model=List[AnotacionItem])
def listar_anotaciones_contrato(id_contrato: str, user: dict = Depends(usuario_actual)):
    """Lista la bitácora de anotaciones de un contrato."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS anotaciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ID_Contrato TEXT,
            Fecha TEXT,
            Tipo TEXT,
            Texto TEXT
        )
    """)
    cursor.execute("""
        SELECT id, ID_Contrato, Fecha, Tipo, Texto
        FROM anotaciones
        WHERE ID_Contrato = ?
        ORDER BY id DESC
    """, (id_contrato,))
    rows = cursor.fetchall()
    conn.close()
    return [
        AnotacionItem(
            id=int(r["id"]),
            id_contrato=str(r["ID_Contrato"] or id_contrato),
            fecha=str(r["Fecha"] or ""),
            tipo=str(r["Tipo"] or "General"),
            texto=str(r["Texto"] or "")
        )
        for r in rows
    ]


@app.post("/api/contratos/{id_contrato}/anotaciones", response_model=List[AnotacionItem])
def crear_anotacion_contrato(id_contrato: str, req: AnotacionCreateRequest, user: dict = Depends(usuario_escritura)):
    """Registra una nueva nota en la bitácora del contrato."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS anotaciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ID_Contrato TEXT,
            Fecha TEXT,
            Tipo TEXT,
            Texto TEXT
        )
    """)
    ahora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    autor = user.get("nombre_completo") or user.get("username") or "Usuario"
    texto_auditado = f"[{autor}] {req.texto.strip()}"
    cursor.execute("""
        INSERT INTO anotaciones (ID_Contrato, Fecha, Tipo, Texto)
        VALUES (?, ?, ?, ?)
    """, (id_contrato, ahora, req.tipo.strip() or "General", texto_auditado))
    conn.commit()
    conn.close()
    return listar_anotaciones_contrato(id_contrato, user=user)


@app.get("/api/contratos/{id_contrato}/pdf_link")
def generar_link_pdf_firmado(id_contrato: str, user: dict = Depends(usuario_actual)):
    """Genera un enlace temporal firmado con HMAC-SHA256 (válido por 5 minutos) para abrir el PDF."""
    expira = int(time.time()) + 300
    sig = _firmar_pdf(id_contrato, expira)
    return {
        "path": f"/api/contratos/{id_contrato}/pdf?e={expira}&s={sig}",
        "expira": expira
    }


@app.get("/api/contratos/{id_contrato}/pdf")
def descargar_pdf_estado_cuenta(
    id_contrato: str,
    e: int = Query(..., description="Timestamp de expiración"),
    s: str = Query(..., description="Firma HMAC-SHA256")
):
    """Genera y descarga el PDF oficial del Estado de Cuenta validando firma temporal."""
    if int(time.time()) > e:
        raise HTTPException(status_code=403, detail="El enlace del PDF expiró. Genéralo nuevamente desde la app.")
    esperado = _firmar_pdf(id_contrato, e)
    if not hmac.compare_digest(esperado, s):
        raise HTTPException(status_code=403, detail="Firma de descarga inválida.")

    from fastapi.responses import StreamingResponse
    from dateutil.relativedelta import relativedelta
    from reports.pdf import pdf_estado_cuenta

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM contratos WHERE ID_Contrato = ?", (id_contrato,))
    row = cursor.fetchone()
    conn.close()

    if not row:
        raise HTTPException(status_code=404, detail="Contrato no encontrado.")

    row_dict = dict(row)
    valor = float(row_dict.get("Valor_Sin_IVA") or 0.0)
    ant = float(row_dict.get("Anticipo_Monto") or 0.0)
    inv = max(valor - ant, 0.0)
    renta = float(row_dict.get("Mensualidad_Sin_IVA") or 0.0)
    plazo = int(row_dict.get("Plazo") or 1)
    res = float(row_dict.get("Residual_Monto") or 0.0)
    tasa = float(row_dict.get("Tasa_Calculada") or 0.0)
    vpr = float(row_dict.get("VP_Residual") or 0.0) or (vp_res(res, tasa, plazo) if tasa else 0.0)

    dfa, _, _, _, _ = calc_amort(round(inv, 4), round(renta, 4), round(res, 4), plazo, round(tasa, 8))
    fa_str = str(row_dict.get("Fecha_Alta") or "")[:10]
    fa = datetime.datetime.strptime(fa_str, "%Y-%m-%d") if fa_str else datetime.datetime.now()
    dfa["Fecha"] = [(fa + relativedelta(months=int(m))).strftime("%Y-%m") for m in dfa["Mes"]]
    dfa["Renta"] = dfa["Capital"] + dfa["Interes"]
    dfa["Saldo_Ini"] = dfa["Saldo"] + dfa["Capital"]
    dfa["Saldo_Fin"] = dfa["Saldo"]

    dfr = calc_res_amort(round(vpr, 4), round(tasa, 8), plazo)
    dfr["Fecha"] = dfa["Fecha"]
    dfr["VP_Residual"] = round(vpr, 2)
    dfr["Residual_Pactado"] = round(res, 2)

    buf = pdf_estado_cuenta(
        row_dict,
        dfa,
        dfr=dfr,
        fecha_hoy_txt=datetime.date.today().strftime("%d/%m/%Y"),
        nombre_empresa="O-LEASING"
    )
    headers = {
        "Content-Disposition": f'inline; filename="Estado_Cuenta_{id_contrato}.pdf"'
    }
    return StreamingResponse(buf, media_type="application/pdf", headers=headers)


@app.get("/api/cobranza_mes", response_model=CobranzaMesResponse)
def obtener_cobranza_mes(
    anio: Optional[int] = Query(None, description="Año de consulta"),
    mes: Optional[int] = Query(None, ge=1, le=12, description="Mes de consulta (1-12)"),
    user: dict = Depends(usuario_actual)
):
    """Calcula la cobranza y devengo mensual de todos los contratos activos en el mes indicado."""
    hoy = datetime.date.today()
    y = anio or hoy.year
    m = mes or hoy.month

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT ID_Contrato, Cliente, Vehiculo, Fecha_Alta, Fecha_Vencimiento, Estatus,
               Plazo, Valor_Sin_IVA, Anticipo_Monto, Mensualidad_Sin_IVA, Residual_Monto, Tasa_Calculada
        FROM contratos
        WHERE Estatus = 'ACTIVO'
        ORDER BY ID_Contrato ASC
    """)
    rows = cursor.fetchall()
    conn.close()

    items = []
    tot_renta = 0.0
    tot_iva = 0.0
    tot_con_iva = 0.0
    tot_cap = 0.0
    tot_int = 0.0

    for r in rows:
        fa_str = str(r["Fecha_Alta"] or "")[:10]
        if not fa_str:
            continue
        try:
            fa = datetime.datetime.strptime(fa_str, "%Y-%m-%d")
        except Exception:
            continue
        plazo = int(r["Plazo"] or 0)
        if plazo <= 0:
            continue
        mc = (y - fa.year) * 12 + (m - fa.month) + 1
        if mc < 1 or mc > plazo:
            continue

        valor = float(r["Valor_Sin_IVA"] or 0.0)
        ant = float(r["Anticipo_Monto"] or 0.0)
        inv = valor - ant
        renta = float(r["Mensualidad_Sin_IVA"] or 0.0)
        res = float(r["Residual_Monto"] or 0.0)
        tasa = float(r["Tasa_Calculada"] or 0.0)
        if inv <= 0 or renta <= 0 or tasa <= 0:
            continue

        try:
            dfa, _, _, _, _ = calc_amort(round(inv, 4), round(renta, 4), round(res, 4), plazo, round(tasa, 8))
            fila_m = dfa.iloc[mc - 1]
            cap_m = round(float(fila_m["Capital"]), 2)
            int_m = round(float(fila_m["Interes"]), 2)
            saldo_m = round(float(fila_m["Saldo"]), 2)
            r_sin = round(renta, 2)
            iva_m = round(r_sin * 0.16, 2)
            tot_m = round(r_sin + iva_m, 2)

            tot_renta += r_sin
            tot_iva += iva_m
            tot_con_iva += tot_m
            tot_cap += cap_m
            tot_int += int_m

            items.append(CobranzaContratoItem(
                id_contrato=r["ID_Contrato"],
                cliente=r["Cliente"],
                vehiculo=r["Vehiculo"] or "",
                mes_contrato=mc,
                plazo=plazo,
                renta_sin_iva=r_sin,
                iva=iva_m,
                total_con_iva=tot_m,
                capital_mes=cap_m,
                interes_leasing_mes=int_m,
                saldo_insoluto=saldo_m
            ))
        except Exception:
            continue

    return CobranzaMesResponse(
        anio=y,
        mes=m,
        periodo=f"{y:04d}-{m:02d}",
        contratos_en_cobro=len(items),
        renta_total_sin_iva=round(tot_renta, 2),
        iva_total=round(tot_iva, 2),
        total_cobranza_con_iva=round(tot_con_iva, 2),
        capital_total_mes=round(tot_cap, 2),
        interes_total_mes=round(tot_int, 2),
        items=items
    )


if __name__ == "__main__":
    import uvicorn
    print("Iniciando O-Leasing Mobile REST API en http://0.0.0.0:8008...")
    uvicorn.run("api:app", host="0.0.0.0", port=8008, reload=True)
