package com.oleasing.app.model

data class DashboardMetrics(
    val totalContratos: Int,
    val contratosActivos: Int,
    val contratosConcluidos: Int,
    val contratosBaja: Int,
    val carteraActivaValor: Double,
    val flujoMensualEstimado: Double,
    val fechaCorte: String
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
    val tasaAnualPct: Double
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
    val interesesTotalesEstimados: Double
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
