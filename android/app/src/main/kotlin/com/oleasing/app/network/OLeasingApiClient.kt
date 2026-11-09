package com.oleasing.app.network

import android.content.Context
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
    private const val PREFS_NAME = "oleasing_prefs"
    private const val KEY_BASE_URL = "api_base_url"
    private const val KEY_AUTH_TOKEN = "auth_token"
    private const val KEY_USERNAME = "auth_username"
    private const val KEY_NOMBRE = "auth_nombre"
    private const val KEY_ROL = "auth_rol"
    private const val KEY_PUEDE_EDITAR = "auth_puede_editar"

    var baseUrl: String = "http://10.0.2.2:8008"
    var authToken: String? = null
    var currentUser: UsuarioSesion? = null
    var onUnauthorized: (() -> Unit)? = null

    fun init(context: Context) {
        val prefs = context.getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
        val saved = prefs.getString(KEY_BASE_URL, null)
        if (!saved.isNullOrBlank()) {
            baseUrl = saved.trim().removeSuffix("/")
        }
        val savedToken = prefs.getString(KEY_AUTH_TOKEN, null)
        val savedUser = prefs.getString(KEY_USERNAME, null)
        if (!savedToken.isNullOrBlank() && !savedUser.isNullOrBlank()) {
            authToken = savedToken
            currentUser = UsuarioSesion(
                username = savedUser,
                nombreCompleto = prefs.getString(KEY_NOMBRE, savedUser) ?: savedUser,
                rol = prefs.getString(KEY_ROL, "lectura") ?: "lectura",
                puedeEditar = prefs.getBoolean(KEY_PUEDE_EDITAR, false)
            )
        }
    }

    fun updateBaseUrl(context: Context, newUrl: String) {
        val cleaned = newUrl.trim().removeSuffix("/")
        baseUrl = cleaned
        val prefs = context.getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
        prefs.edit().putString(KEY_BASE_URL, cleaned).apply()
    }

    fun saveSession(context: Context, token: String, user: UsuarioSesion) {
        authToken = token
        currentUser = user
        val prefs = context.getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
        prefs.edit()
            .putString(KEY_AUTH_TOKEN, token)
            .putString(KEY_USERNAME, user.username)
            .putString(KEY_NOMBRE, user.nombreCompleto)
            .putString(KEY_ROL, user.rol)
            .putBoolean(KEY_PUEDE_EDITAR, user.puedeEditar)
            .apply()
    }

    fun clearSession(context: Context) {
        authToken = null
        currentUser = null
        val prefs = context.getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
        prefs.edit()
            .remove(KEY_AUTH_TOKEN)
            .remove(KEY_USERNAME)
            .remove(KEY_NOMBRE)
            .remove(KEY_ROL)
            .remove(KEY_PUEDE_EDITAR)
            .apply()
    }

    private val fallbackUrls = listOf(
        "http://10.0.2.2:8008",
        "http://192.168.1.97:8008",
        "http://127.0.0.1:8008"
    )

    private fun executeSingleHttp(
        candidateBaseUrl: String,
        path: String,
        method: String,
        bodyJson: String?
    ): Pair<Int, String> {
        var connection: HttpURLConnection? = null
        try {
            val fullUrl = if (path.startsWith("http")) path else "$candidateBaseUrl$path"
            val url = URL(fullUrl)
            connection = (url.openConnection() as HttpURLConnection).apply {
                requestMethod = method
                connectTimeout = 4500
                readTimeout = 6000
                setRequestProperty("Accept", "application/json")
                setRequestProperty("Content-Type", "application/json")
                val token = authToken
                if (!token.isNullOrBlank()) {
                    setRequestProperty("Authorization", "Bearer $token")
                }
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
            return Pair(statusCode, responseText)
        } finally {
            connection?.disconnect()
        }
    }

    private suspend fun <T> makeRequest(
        path: String,
        method: String = "GET",
        bodyJson: String? = null,
        parser: (String) -> T
    ): Result<T> = withContext(Dispatchers.IO) {
        val candidates = mutableListOf(baseUrl)
        for (fb in fallbackUrls) {
            if (!candidates.contains(fb)) {
                candidates.add(fb)
            }
        }

        var lastError: Exception? = null
        for (candidate in candidates) {
            try {
                val (statusCode, responseText) = executeSingleHttp(candidate, path, method, bodyJson)
                if (statusCode in 200..299) {
                    if (baseUrl != candidate) {
                        baseUrl = candidate
                    }
                    return@withContext Result.success(parser(responseText))
                } else {
                    val detail = try {
                        val errorObj = JSONObject(responseText)
                        errorObj.optString("detail", responseText)
                    } catch (e: Exception) {
                        responseText
                    }
                    if (statusCode == 401 && path != "/api/auth/login") {
                        authToken = null
                        currentUser = null
                        withContext(Dispatchers.Main) {
                            onUnauthorized?.invoke()
                        }
                    }
                    return@withContext Result.failure(Exception("Error $statusCode: $detail"))
                }
            } catch (e: Exception) {
                lastError = e
            }
        }
        Result.failure(Exception("Fallo de red: ${lastError?.message ?: "No se pudo conectar al servidor"}"))
    }

    suspend fun login(context: Context, username: String, password: String): Result<LoginResponse> {
        val payload = JSONObject().apply {
            put("username", username.trim())
            put("password", password)
            put("dias_sesion", 14)
        }
        val res = makeRequest("/api/auth/login", method = "POST", bodyJson = payload.toString()) { jsonStr ->
            val obj = JSONObject(jsonStr)
            val uObj = obj.getJSONObject("usuario")
            val user = UsuarioSesion(
                username = uObj.optString("username"),
                nombreCompleto = uObj.optString("nombre_completo"),
                rol = uObj.optString("rol", "lectura"),
                puedeEditar = uObj.optBoolean("puede_editar", false)
            )
            LoginResponse(
                token = obj.optString("token"),
                expiraDias = obj.optInt("expira_dias", 14),
                usuario = user
            )
        }
        res.onSuccess { loginResp ->
            saveSession(context, loginResp.token, loginResp.usuario)
        }
        return res
    }

    suspend fun logout(context: Context) {
        try {
            makeRequest("/api/auth/logout", method = "POST") { it }
        } catch (_: Exception) {
        }
        clearSession(context)
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

    suspend fun getEstadoCuenta(idContrato: String): Result<EstadoCuentaResponse> {
        val encodedId = URLEncoder.encode(idContrato, "UTF-8")
        return makeRequest("/api/contratos/$encodedId/estado_cuenta") { jsonStr ->
            val obj = JSONObject(jsonStr)
            val tablaArr = obj.optJSONArray("tabla_amortizacion") ?: JSONArray()
            val tabla = mutableListOf<PeriodoAmortizacion>()
            for (i in 0 until tablaArr.length()) {
                val p = tablaArr.getJSONObject(i)
                tabla.add(
                    PeriodoAmortizacion(
                        periodo = p.optInt("periodo"),
                        fechaEstimada = p.optString("fecha_estimada", ""),
                        rentaSinIva = p.optDouble("renta_sin_iva"),
                        capital = p.optDouble("capital"),
                        interes = p.optDouble("interes"),
                        iva = p.optDouble("iva"),
                        totalConIva = p.optDouble("total_con_iva"),
                        saldoInsoluto = p.optDouble("saldo_insoluto"),
                        estatus = p.optString("estatus", "PENDIENTE")
                    )
                )
            }
            EstadoCuentaResponse(
                idContrato = obj.optString("id_contrato"),
                cliente = obj.optString("cliente"),
                vehiculo = obj.optString("vehiculo", ""),
                estatus = obj.optString("estatus", "ACTIVO"),
                fechaAlta = obj.optString("fecha_alta", ""),
                fechaVencimiento = obj.optString("fecha_vencimiento", ""),
                plazo = obj.optInt("plazo"),
                valorSinIva = obj.optDouble("valor_sin_iva"),
                anticipoMonto = obj.optDouble("anticipo_monto"),
                inversionNeta = obj.optDouble("inversion_neta"),
                mensualidadSinIva = obj.optDouble("mensualidad_sin_iva"),
                residualMonto = obj.optDouble("residual_monto"),
                tasaMensualPct = obj.optDouble("tasa_mensual_pct"),
                tasaAnualPct = obj.optDouble("tasa_anual_pct"),
                mesesTranscurridos = obj.optInt("meses_transcurridos"),
                pagosRestantes = obj.optInt("pagos_restantes"),
                saldoActualEstimado = obj.optDouble("saldo_actual_estimado"),
                capitalAmortizado = obj.optDouble("capital_amortizado"),
                interesDevengado = obj.optDouble("interes_devengado"),
                totalInteresesPlazo = obj.optDouble("total_intereses_plazo"),
                porcentajeAvance = obj.optDouble("porcentaje_avance"),
                tablaAmortizacion = tabla
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

    suspend fun requestPdfSignedUrl(idContrato: String): Result<String> {
        val encodedId = URLEncoder.encode(idContrato, "UTF-8")
        return makeRequest("/api/contratos/$encodedId/pdf_link") { jsonStr ->
            val obj = JSONObject(jsonStr)
            val relPath = obj.optString("path", "/api/contratos/$encodedId/pdf")
            "$baseUrl$relPath"
        }
    }

    suspend fun actualizarEstatusContrato(
        idContrato: String,
        nuevoEstatus: String,
        nivelMorosidad: Int = 0
    ): Result<ContratoDetalle> {
        val encodedId = URLEncoder.encode(idContrato, "UTF-8")
        val payload = JSONObject().apply {
            put("estatus", nuevoEstatus)
            put("nivel_morosidad", nivelMorosidad)
        }
        return makeRequest("/api/contratos/$encodedId/estatus", method = "PUT", bodyJson = payload.toString()) { jsonStr ->
            val item = JSONObject(jsonStr)
            ContratoDetalle(
                idContrato = item.optString("id_contrato"),
                cliente = item.optString("cliente"),
                vehiculo = item.optString("vehiculo", ""),
                fechaAlta = item.optString("fecha_alta", ""),
                fechaVencimiento = item.optString("fecha_vencimiento", ""),
                fechaBaja = if (item.isNull("fecha_baja")) null else item.optString("fecha_baja"),
                estatus = item.optString("estatus", "ACTIVO"),
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

    suspend fun getAnotaciones(idContrato: String): Result<List<AnotacionItem>> {
        val encodedId = URLEncoder.encode(idContrato, "UTF-8")
        return makeRequest("/api/contratos/$encodedId/anotaciones") { jsonStr ->
            val arr = JSONArray(jsonStr.trim())
            val list = mutableListOf<AnotacionItem>()
            for (i in 0 until arr.length()) {
                val o = arr.getJSONObject(i)
                list.add(
                    AnotacionItem(
                        id = o.optInt("id"),
                        idContrato = o.optString("id_contrato"),
                        fecha = o.optString("fecha"),
                        tipo = o.optString("tipo", "General"),
                        texto = o.optString("texto", "")
                    )
                )
            }
            list
        }
    }

    suspend fun crearAnotacion(idContrato: String, tipo: String, texto: String): Result<List<AnotacionItem>> {
        val encodedId = URLEncoder.encode(idContrato, "UTF-8")
        val payload = JSONObject().apply {
            put("tipo", tipo)
            put("texto", texto)
        }
        return makeRequest("/api/contratos/$encodedId/anotaciones", method = "POST", bodyJson = payload.toString()) { jsonStr ->
            val arr = JSONArray(jsonStr.trim())
            val list = mutableListOf<AnotacionItem>()
            for (i in 0 until arr.length()) {
                val o = arr.getJSONObject(i)
                list.add(
                    AnotacionItem(
                        id = o.optInt("id"),
                        idContrato = o.optString("id_contrato"),
                        fecha = o.optString("fecha"),
                        tipo = o.optString("tipo", "General"),
                        texto = o.optString("texto", "")
                    )
                )
            }
            list
        }
    }

    suspend fun getCobranzaMes(anio: Int, mes: Int): Result<CobranzaMesResponse> {
        return makeRequest("/api/cobranza_mes?anio=$anio&mes=$mes") { jsonStr ->
            val obj = JSONObject(jsonStr)
            val arr = obj.optJSONArray("items") ?: JSONArray()
            val list = mutableListOf<CobranzaContratoItem>()
            for (i in 0 until arr.length()) {
                val o = arr.getJSONObject(i)
                list.add(
                    CobranzaContratoItem(
                        idContrato = o.optString("id_contrato"),
                        cliente = o.optString("cliente"),
                        vehiculo = o.optString("vehiculo", ""),
                        mesContrato = o.optInt("mes_contrato"),
                        plazo = o.optInt("plazo"),
                        rentaSinIva = o.optDouble("renta_sin_iva"),
                        iva = o.optDouble("iva"),
                        totalConIva = o.optDouble("total_con_iva"),
                        capitalMes = o.optDouble("capital_mes"),
                        interesLeasingMes = o.optDouble("interes_leasing_mes"),
                        saldoInsoluto = o.optDouble("saldo_insoluto")
                    )
                )
            }
            CobranzaMesResponse(
                anio = obj.optInt("anio", anio),
                mes = obj.optInt("mes", mes),
                periodo = obj.optString("periodo", ""),
                contratosEnCobro = obj.optInt("contratos_en_cobro"),
                rentaTotalSinIva = obj.optDouble("renta_total_sin_iva"),
                ivaTotal = obj.optDouble("iva_total"),
                totalCobranzaConIva = obj.optDouble("total_cobranza_con_iva"),
                capitalTotalMes = obj.optDouble("capital_total_mes"),
                interesTotalMes = obj.optDouble("interes_total_mes"),
                items = list
            )
        }
    }
}
