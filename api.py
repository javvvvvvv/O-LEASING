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
from typing import List, Optional
from pydantic import BaseModel, Field
from fastapi import FastAPI, HTTPException, Query, status
from fastapi.middleware.cors import CORSMiddleware

# Asegurar importación del motor financiero
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if _BASE_DIR not in sys.path:
    sys.path.insert(0, _BASE_DIR)

from finanzas import calc_tasa, vp_res, calc_amort, calc_rent_amort, calc_res_amort

app = FastAPI(
    title="O-Leasing Mobile API",
    description="API REST para consulta y creación individual de contratos de arrendamiento financiero",
    version="1.0.0"
)

# Habilitar CORS para peticiones desde app móvil, emulador y clientes web
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

DB_PATH = os.path.join(_BASE_DIR, "leasing.db")


def get_db_connection():
    """Retorna una conexión a la base de datos leasing.db con row_factory como dict."""
    if not os.path.exists(DB_PATH):
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Base de datos leasing.db no encontrada en {_BASE_DIR}"
        )
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


# =============================================================================
# MODELOS PYDANTIC
# =============================================================================

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


# =============================================================================
# ENDPOINTS DE CONSULTA Y DASHBOARD
# =============================================================================

@app.get("/api/health")
def health_check():
    """Verifica estado del servicio y conteo general de contratos en base de datos."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM contratos")
    total = cursor.fetchone()[0]
    conn.close()
    return {
        "status": "online",
        "service": "O-Leasing Mobile API",
        "database": os.path.basename(DB_PATH),
        "total_contratos_bd": total,
        "timestamp": datetime.datetime.now().isoformat()
    }


@app.get("/api/dashboard")
def get_dashboard_metrics():
    """Retorna las métricas clave del portafolio para la pantalla de inicio móvil."""
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*), SUM(Valor_Sin_IVA - Anticipo_Monto), SUM(Mensualidad_Sin_IVA) FROM contratos WHERE Estatus = 'ACTIVO'")
    row_act = cursor.fetchone()
    n_activos = row_act[0] or 0
    cartera_activa_monto = row_act[1] or 0.0
    renta_mensual_total = row_act[2] or 0.0

    cursor.execute("SELECT COUNT(*) FROM contratos WHERE Estatus = 'CONCLUIDO'")
    n_concluidos = cursor.fetchone()[0] or 0

    cursor.execute("SELECT COUNT(*) FROM contratos WHERE Estatus = 'BAJA'")
    n_bajas = cursor.fetchone()[0] or 0

    cursor.execute("SELECT COUNT(*) FROM contratos")
    total_historico = cursor.fetchone()[0] or 0

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
        "fecha_corte": datetime.date.today().isoformat()
    }


@app.get("/api/contratos", response_model=List[ContratoResumen])
def listar_contratos(
    q: Optional[str] = Query(None, description="Texto de búsqueda por ID, Cliente o Vehículo"),
    estatus: Optional[str] = Query("ACTIVO", description="Filtro por estatus: ACTIVO, CONCLUIDO, BAJA o TODOS"),
    limit: int = Query(50, ge=1, le=200),
    offset: int = Query(0, ge=0)
):
    """Consulta paginada y filtrable de contratos para la lista principal de la app."""
    conn = get_db_connection()
    cursor = conn.cursor()

    sql = """
        SELECT ID_Contrato, Cliente, Vehiculo, Fecha_Alta, Fecha_Vencimiento,
               Estatus, Plazo, Valor_Sin_IVA, Anticipo_Monto, Mensualidad_Sin_IVA,
               Residual_Monto, Tasa_Calculada
        FROM contratos
        WHERE 1=1
    """
    params = []

    if estatus and estatus.upper() != "TODOS":
        sql += " AND Estatus = ?"
        params.append(estatus.upper())

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
        ))

    return resultado


@app.get("/api/contratos/{id_contrato}", response_model=ContratoDetalle)
def obtener_detalle_contrato(id_contrato: str):
    """Consulta a fondo de un contrato individual con desglose financiero de amortización."""
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
        intereses_totales_estimados=round(intereses_totales, 2)
    )


@app.get("/api/contratos/{id_contrato}/estado_cuenta", response_model=EstadoCuentaResponse)
def obtener_estado_cuenta(id_contrato: str):
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
def cotizar_leasing(req: CotizacionRequest):
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
def crear_contrato_individual(contrato: ContratoCreate):
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
    return obtener_detalle_contrato(contrato.id_contrato.strip())


if __name__ == "__main__":
    import uvicorn
    print("Iniciando O-Leasing Mobile REST API en http://0.0.0.0:8000...")
    uvicorn.run("api:app", host="0.0.0.0", port=8000, reload=True)
