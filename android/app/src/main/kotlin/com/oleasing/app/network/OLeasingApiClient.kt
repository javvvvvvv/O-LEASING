package com.oleasing.app.network

import com.oleasing.app.model.*
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import org.json.JSONArray
import org.json.JSONObject
import java.io.BufferedReader
import java.io.InputStreamReader
import java.io.OutputStreamWriter
import java.net.HttpURLConnection
import java.net.URL
import java.net.URLEncoder

object OLeasingApiClient {
    var baseUrl: String = "http://10.0.2.2:8008"

    private suspend fun <T> makeRequest(
        path: String,
        method: String = "GET",
        bodyJson: String? = null,
        parser: (String) -> T
    ): Result<T> = withContext(Dispatchers.IO) {
        var connection: HttpURLConnection? = null
        try {
            val url = URL(if (path.startsWith("http")) path else "$baseUrl$path")
            connection = (url.openConnection() as HttpURLConnection).apply {
                requestMethod = method
                connectTimeout = 8000
                readTimeout = 8000
                setRequestProperty("Accept", "application/json")
                setRequestProperty("Content-Type", "application/json")
                if (bodyJson != null) {
                    doOutput = true
                    val writer = OutputStreamWriter(outputStream, "UTF-8")
                    writer.write(bodyJson)
                    writer.flush()
                    writer.close()
                }
            }

            val statusCode = connection.responseCode
            val stream = if (statusCode in 200..299) connection.inputStream else connection.errorStream
            val reader = BufferedReader(InputStreamReader(stream ?: connection.inputStream, "UTF-8"))
            val responseText = reader.use { it.readText() }

            if (statusCode in 200..299) {
                Result.success(parser(responseText))
            } else {
                val detail = try {
                    val errorObj = JSONObject(responseText)
                    errorObj.optString("detail", responseText)
                } catch (e: Exception) {
                    responseText
                }
                Result.failure(Exception("Error $statusCode: $detail"))
            }
        } catch (e: Exception) {
            Result.failure(Exception("Fallo de red: ${e.message ?: "No se pudo conectar al servidor"}"))
        } finally {
            connection?.disconnect()
        }
    }

    suspend fun checkHealth(): Result<String> {
        return makeRequest("/api/health") { jsonStr ->
            val obj = JSONObject(jsonStr)
            obj.optString("status", "ok")
        }
    }

    suspend fun getDashboard(): Result<DashboardMetrics> {
        return makeRequest("/api/dashboard") { jsonStr ->
            val obj = JSONObject(jsonStr)
            DashboardMetrics(
                totalContratos = if (obj.has("total_contratos")) obj.optInt("total_contratos") else obj.optInt("total_historico"),
                contratosActivos = obj.optInt("contratos_activos"),
                contratosConcluidos = obj.optInt("contratos_concluidos"),
                contratosBaja = if (obj.has("contratos_baja")) obj.optInt("contratos_baja") else obj.optInt("contratos_bajas"),
                carteraActivaValor = if (obj.has("cartera_activa_valor")) obj.optDouble("cartera_activa_valor") else obj.optDouble("cartera_activa_monto"),
                flujoMensualEstimado = if (obj.has("flujo_mensual_estimado")) obj.optDouble("flujo_mensual_estimado") else obj.optDouble("flujo_mensual_rentas"),
                fechaCorte = obj.optString("fecha_corte")
            )
        }
    }

    suspend fun getContratos(
        query: String? = null,
        estatus: String? = null,
        limit: Int = 100,
        offset: Int = 0
    ): Result<ContratoListResponse> {
        val params = mutableListOf<String>()
        params.add("limit=$limit")
        params.add("offset=$offset")
        if (!query.isNullOrBlank()) {
            params.add("q=" + URLEncoder.encode(query.trim(), "UTF-8"))
        }
        if (!estatus.isNullOrBlank() && estatus != "Todos") {
            params.add("estatus=" + URLEncoder.encode(estatus.trim(), "UTF-8"))
        }
        val path = "/api/contratos?" + params.joinToString("&")

        return makeRequest(path) { jsonStr ->
            val trimmed = jsonStr.trim()
            val itemsArr = if (trimmed.startsWith("[")) {
                JSONArray(trimmed)
            } else {
                val obj = JSONObject(trimmed)
                obj.optJSONArray("items") ?: JSONArray()
            }
            val list = mutableListOf<ContratoResumen>()
            for (i in 0 until itemsArr.length()) {
                val item = itemsArr.getJSONObject(i)
                list.add(
                    ContratoResumen(
                        idContrato = item.optString("id_contrato"),
                        cliente = item.optString("cliente"),
                        vehiculo = item.optString("vehiculo", ""),
                        fechaAlta = item.optString("fecha_alta", ""),
                        fechaVencimiento = item.optString("fecha_vencimiento", ""),
                        estatus = item.optString("estatus", "Activo"),
                        plazo = item.optInt("plazo"),
                        valorSinIva = item.optDouble("valor_sin_iva"),
                        anticipoMonto = item.optDouble("anticipo_monto"),
                        mensualidadSinIva = item.optDouble("mensualidad_sin_iva"),
                        residualMonto = item.optDouble("residual_monto"),
                        tasaMensualPct = item.optDouble("tasa_mensual_pct"),
                        tasaAnualPct = item.optDouble("tasa_anual_pct")
                    )
                )
            }
            ContratoListResponse(
                total = list.size,
                limit = limit,
                offset = offset,
                items = list
            )
        }
    }

    suspend fun getContratoDetalle(idContrato: String): Result<ContratoDetalle> {
        val encodedId = URLEncoder.encode(idContrato, "UTF-8")
        return makeRequest("/api/contratos/$encodedId") { jsonStr ->
            val item = JSONObject(jsonStr)
            ContratoDetalle(
                idContrato = item.optString("id_contrato"),
                cliente = item.optString("cliente"),
                vehiculo = item.optString("vehiculo", ""),
                fechaAlta = item.optString("fecha_alta", ""),
                fechaVencimiento = item.optString("fecha_vencimiento", ""),
                fechaBaja = if (item.isNull("fecha_baja")) null else item.optString("fecha_baja"),
                estatus = item.optString("estatus", "Activo"),
                plazo = item.optInt("plazo"),
                valorSinIva = item.optDouble("valor_sin_iva"),
                anticipoPct = item.optDouble("anticipo_pct"),
                anticipoMonto = item.optDouble("anticipo_monto"),
                inversionNeta = item.optDouble("inversion_neta"),
                mensualidadSinIva = item.optDouble("mensualidad_sin_iva"),
                residualPct = item.optDouble("residual_pct"),
                residualMonto = item.optDouble("residual_monto"),
                vpResidual = item.optDouble("vp_residual"),
                comisionAperturaPct = item.optDouble("comision_apertura_pct"),
                comisionMonto = item.optDouble("comision_monto"),
                tasaMensualPct = item.optDouble("tasa_mensual_pct"),
                tasaAnualPct = item.optDouble("tasa_anual_pct"),
                anotaciones = item.optString("anotaciones", ""),
                saldoActualEstimado = item.optDouble("saldo_actual_estimado"),
                interesesTotalesEstimados = item.optDouble("intereses_totales_estimados")
            )
        }
    }

    suspend fun cotizar(req: CotizacionRequest): Result<CotizacionResponse> {
        val payload = JSONObject().apply {
            put("valor_sin_iva", req.valorSinIva)
            put("plazo", req.plazo)
            if (req.mensualidadSinIva != null) put("mensualidad_sin_iva", req.mensualidadSinIva)
            if (req.tasaAnualPct != null) put("tasa_anual_pct", req.tasaAnualPct)
            put("anticipo_monto", req.anticipoMonto)
            put("residual_monto", req.residualMonto)
            put("comision_pct", req.comisionPct)
        }

        return makeRequest("/api/cotizar", method = "POST", bodyJson = payload.toString()) { jsonStr ->
            val obj = JSONObject(jsonStr)
            CotizacionResponse(
                valorSinIva = obj.optDouble("valor_sin_iva"),
                anticipoMonto = obj.optDouble("anticipo_monto"),
                inversionNeta = obj.optDouble("inversion_neta"),
                plazo = obj.optInt("plazo"),
                mensualidadSinIva = obj.optDouble("mensualidad_sin_iva"),
                residualMonto = obj.optDouble("residual_monto"),
                vpResidual = obj.optDouble("vp_residual"),
                tasaMensualPct = obj.optDouble("tasa_mensual_pct"),
                tasaAnualPct = obj.optDouble("tasa_anual_pct"),
                comisionMonto = obj.optDouble("comision_monto"),
                totalIntereses = obj.optDouble("total_intereses"),
                totalRentas = obj.optDouble("total_rentas"),
                costoTotalLeasing = obj.optDouble("costo_total_leasing")
            )
        }
    }

    suspend fun crearContrato(req: ContratoCreate): Result<CreateContratoResponse> {
        val payload = JSONObject().apply {
            put("id_contrato", req.idContrato)
            put("cliente", req.cliente)
            put("vehiculo", req.vehiculo)
            put("fecha_alta", req.fechaAlta)
            put("plazo", req.plazo)
            put("valor_sin_iva", req.valorSinIva)
            put("anticipo_monto", req.anticipoMonto)
            put("mensualidad_sin_iva", req.mensualidadSinIva)
            put("residual_monto", req.residualMonto)
            put("comision_pct", req.comisionPct)
            put("anotaciones", req.anotaciones)
        }

        return makeRequest("/api/contratos", method = "POST", bodyJson = payload.toString()) { jsonStr ->
            val obj = JSONObject(jsonStr)
            CreateContratoResponse(
                message = obj.optString("message"),
                idContrato = obj.optString("id_contrato"),
                cliente = obj.optString("cliente"),
                tasaAnualPct = obj.optDouble("tasa_anual_pct"),
                vpResidual = obj.optDouble("vp_residual"),
                inversionNeta = obj.optDouble("inversion_neta"),
                fechaVencimiento = obj.optString("fecha_vencimiento")
            )
        }
    }
}
