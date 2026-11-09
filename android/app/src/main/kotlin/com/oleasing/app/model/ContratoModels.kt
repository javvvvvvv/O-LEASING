package com.oleasing.app.model

data class UsuarioSesion(
    val username: String,
    val nombreCompleto: String,
    val rol: String,
    val puedeEditar: Boolean
)

data class LoginResponse(
    val token: String,
    val expiraDias: Int,
    val usuario: UsuarioSesion
)

data class CotizacionDraft(
    val valorSinIva: Double,
    val plazo: Int,
    val anticipoMonto: Double,
    val mensualidadSinIva: Double,
    val residualMonto: Double,
    val comisionPct: Double
)

data class VencimientoProximoItem(
    val idContrato: String,
    val cliente: String,
    val vehiculo: String,
    val fechaVencimiento: String,
    val diasRestantes: Int,
    val residualMonto: Double,
    val mensualidadSinIva: Double
)

data class DashboardMetrics(
    val totalContratos: Int,
    val contratosActivos: Int,
    val contratosConcluidos: Int,
    val contratosBaja: Int,
    val carteraActivaValor: Double,
    val flujoMensualEstimado: Double,
    val fechaCorte: String,
    val residualTotalActivo: Double = 0.0,
    val contratosPorVencer90d: Int = 0,
    val contratosEnMora: Int = 0,
    val vencimientosProximos: List<VencimientoProximoItem> = emptyList()
)

data class ContratoResumen(
    val idContrato: String,
    val cliente: String,
    val vehiculo: String,
    val fechaAlta: String,
    val fechaVencimiento: String,
    val estatus: String,
    val plazo: Int,
    val valorSinIva: Double,
    val anticipoMonto: Double,
    val mensualidadSinIva: Double,
    val residualMonto: Double,
    val tasaMensualPct: Double,
    val tasaAnualPct: Double,
    val nivelMorosidad: Int = 0
)

data class ContratoListResponse(
    val total: Int,
    val limit: Int,
    val offset: Int,
    val items: List<ContratoResumen>
)

data class ContratoDetalle(
    val idContrato: String,
    val cliente: String,
    val vehiculo: String,
    val fechaAlta: String,
    val fechaVencimiento: String,
    val fechaBaja: String?,
    val estatus: String,
    val plazo: Int,
    val valorSinIva: Double,
    val anticipoPct: Double,
    val anticipoMonto: Double,
    val inversionNeta: Double,
    val mensualidadSinIva: Double,
    val residualPct: Double,
    val residualMonto: Double,
    val vpResidual: Double,
    val comisionAperturaPct: Double,
    val comisionMonto: Double,
    val tasaMensualPct: Double,
    val tasaAnualPct: Double,
    val anotaciones: String,
    val saldoActualEstimado: Double,
    val interesesTotalesEstimados: Double,
    val nivelMorosidad: Int = 0
)

data class PeriodoAmortizacion(
    val periodo: Int,
    val fechaEstimada: String,
    val rentaSinIva: Double,
    val capital: Double,
    val interes: Double,
    val iva: Double,
    val totalConIva: Double,
    val saldoInsoluto: Double,
    val estatus: String
)

data class EstadoCuentaResponse(
    val idContrato: String,
    val cliente: String,
    val vehiculo: String,
    val estatus: String,
    val fechaAlta: String,
    val fechaVencimiento: String,
    val plazo: Int,
    val valorSinIva: Double,
    val anticipoMonto: Double,
    val inversionNeta: Double,
    val mensualidadSinIva: Double,
    val residualMonto: Double,
    val tasaMensualPct: Double,
    val tasaAnualPct: Double,
    val mesesTranscurridos: Int,
    val pagosRestantes: Int,
    val saldoActualEstimado: Double,
    val capitalAmortizado: Double,
    val interesDevengado: Double,
    val totalInteresesPlazo: Double,
    val porcentajeAvance: Double,
    val tablaAmortizacion: List<PeriodoAmortizacion>
)

data class CotizacionRequest(
    val valorSinIva: Double,
    val plazo: Int,
    val mensualidadSinIva: Double? = null,
    val tasaAnualPct: Double? = null,
    val anticipoMonto: Double = 0.0,
    val residualMonto: Double = 0.0,
    val comisionPct: Double = 0.0
)

data class CotizacionResponse(
    val valorSinIva: Double,
    val anticipoMonto: Double,
    val inversionNeta: Double,
    val plazo: Int,
    val mensualidadSinIva: Double,
    val residualMonto: Double,
    val vpResidual: Double,
    val tasaMensualPct: Double,
    val tasaAnualPct: Double,
    val comisionMonto: Double,
    val totalIntereses: Double,
    val totalRentas: Double,
    val costoTotalLeasing: Double
)

data class ContratoCreate(
    val idContrato: String,
    val cliente: String,
    val vehiculo: String,
    val fechaAlta: String,
    val plazo: Int,
    val valorSinIva: Double,
    val anticipoMonto: Double,
    val mensualidadSinIva: Double,
    val residualMonto: Double,
    val comisionPct: Double,
    val anotaciones: String
)

data class CreateContratoResponse(
    val message: String,
    val idContrato: String,
    val cliente: String,
    val tasaAnualPct: Double,
    val vpResidual: Double,
    val inversionNeta: Double,
    val fechaVencimiento: String
)

data class AnotacionItem(
    val id: Int,
    val idContrato: String,
    val fecha: String,
    val tipo: String,
    val texto: String
)

data class CobranzaContratoItem(
    val idContrato: String,
    val cliente: String,
    val vehiculo: String,
    val mesContrato: Int,
    val plazo: Int,
    val rentaSinIva: Double,
    val iva: Double,
    val totalConIva: Double,
    val capitalMes: Double,
    val interesLeasingMes: Double,
    val saldoInsoluto: Double
)

data class CobranzaMesResponse(
    val anio: Int,
    val mes: Int,
    val periodo: String,
    val contratosEnCobro: Int,
    val rentaTotalSinIva: Double,
    val ivaTotal: Double,
    val totalCobranzaConIva: Double,
    val capitalTotalMes: Double,
    val interesTotalMes: Double,
    val items: List<CobranzaContratoItem>
)
