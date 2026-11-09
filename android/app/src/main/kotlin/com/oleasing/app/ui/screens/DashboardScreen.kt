package com.oleasing.app.ui.screens

import androidx.compose.foundation.BorderStroke
import androidx.compose.foundation.Image
import androidx.compose.foundation.background
import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.filled.ExitToApp
import androidx.compose.material.icons.filled.Add
import androidx.compose.material.icons.filled.Lock
import androidx.compose.material.icons.filled.Refresh
import androidx.compose.material.icons.filled.Search
import androidx.compose.material.icons.filled.Settings
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.draw.clip
import androidx.compose.ui.graphics.Brush
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.layout.ContentScale
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.res.painterResource
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.input.PasswordVisualTransformation
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import com.oleasing.app.R
import com.oleasing.app.model.ContratoResumen
import com.oleasing.app.model.DashboardMetrics
import com.oleasing.app.network.OLeasingApiClient
import com.oleasing.app.ui.components.MorosidadBadge
import com.oleasing.app.ui.components.StatCard
import com.oleasing.app.ui.components.StatusBadge
import com.oleasing.app.ui.components.formatCurrency
import com.oleasing.app.ui.theme.AppCardBorder
import com.oleasing.app.ui.theme.OrangePrimary
import com.oleasing.app.ui.theme.SlateDark
import com.oleasing.app.ui.theme.SlateSecondary
import kotlinx.coroutines.launch

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun DashboardScreen(
    onNavigateToList: () -> Unit,
    onNavigateToNew: () -> Unit,
    onSelectContrato: (String) -> Unit,
    onNavigateToEstadoCuenta: (String) -> Unit,
    onNavigateToCobranza: () -> Unit,
    onNavigateToCotizador: () -> Unit,
    onLogout: () -> Unit = {}
) {
    val context = LocalContext.current
    val coroutineScope = rememberCoroutineScope()
    val currentUser = OLeasingApiClient.currentUser
    val canEdit = currentUser?.puedeEditar != false

    var metrics by remember { mutableStateOf<DashboardMetrics?>(null) }
    var recentContratos by remember { mutableStateOf<List<ContratoResumen>>(emptyList()) }
    var isLoading by remember { mutableStateOf(false) }
    var errorMessage by remember { mutableStateOf<String?>(null) }

    var showServerDialog by remember { mutableStateOf(false) }
    var serverHostInput by remember { mutableStateOf(OLeasingApiClient.baseUrl) }
    var pingStatusMessage by remember { mutableStateOf<String?>(null) }

    var showQuickJumpDialog by remember { mutableStateOf(false) }
    var quickJumpId by remember { mutableStateOf("") }

    var showPasswordDialog by remember { mutableStateOf(false) }
    var currentPwInput by remember { mutableStateOf("") }
    var newPwInput by remember { mutableStateOf("") }
    var pwFeedback by remember { mutableStateOf<String?>(null) }
    var pwSuccess by remember { mutableStateOf(false) }

    fun loadData() {
        coroutineScope.launch {
            isLoading = true
            errorMessage = null
            val dashResult = OLeasingApiClient.getDashboard()
            if (dashResult.isSuccess) {
                metrics = dashResult.getOrNull()
            } else {
                errorMessage = dashResult.exceptionOrNull()?.message
            }

            val listResult = OLeasingApiClient.getContratos(limit = 5, estatus = "Activo")
            if (listResult.isSuccess) {
                recentContratos = listResult.getOrNull()?.items ?: emptyList()
            }
            isLoading = false
        }
    }

    LaunchedEffect(Unit) {
        loadData()
    }

    Scaffold(
        topBar = {
            TopAppBar(
                title = {
                    Row(verticalAlignment = Alignment.CenterVertically) {
                        Image(
                            painter = painterResource(id = R.drawable.oleasing_logo),
                            contentDescription = "Logotipo O-Leasing",
                            modifier = Modifier.height(30.dp),
                            contentScale = ContentScale.FillHeight
                        )
                    }
                },
                actions = {
                    IconButton(onClick = {
                        currentPwInput = ""
                        newPwInput = ""
                        pwFeedback = null
                        pwSuccess = false
                        showPasswordDialog = true
                    }) {
                        Icon(
                            imageVector = Icons.Default.Lock,
                            contentDescription = "Seguridad y Contraseña",
                            tint = SlateSecondary
                        )
                    }
                    IconButton(onClick = {
                        serverHostInput = OLeasingApiClient.baseUrl
                        pingStatusMessage = null
                        showServerDialog = true
                    }) {
                        Icon(
                            imageVector = Icons.Default.Settings,
                            contentDescription = "Configurar Servidor",
                            tint = SlateSecondary
                        )
                    }
                    IconButton(onClick = { loadData() }) {
                        Icon(
                            imageVector = Icons.Default.Refresh,
                            contentDescription = "Actualizar",
                            tint = OrangePrimary
                        )
                    }
                    IconButton(onClick = {
                        coroutineScope.launch {
                            OLeasingApiClient.logout(context)
                            onLogout()
                        }
                    }) {
                        Icon(
                            imageVector = Icons.AutoMirrored.Filled.ExitToApp,
                            contentDescription = "Cerrar Sesión",
                            tint = Color(0xFFDC2626)
                        )
                    }
                },
                colors = TopAppBarDefaults.topAppBarColors(
                    containerColor = MaterialTheme.colorScheme.surface
                )
            )
        }
    ) { padding ->
        LazyColumn(
            modifier = Modifier
                .fillMaxSize()
                .padding(padding)
                .background(MaterialTheme.colorScheme.background)
                .padding(horizontal = 16.dp),
            verticalArrangement = Arrangement.spacedBy(14.dp)
        ) {
            item {
                Spacer(modifier = Modifier.height(4.dp))
                // Executive Hero Card
                Card(
                    modifier = Modifier.fillMaxWidth(),
                    shape = RoundedCornerShape(18.dp),
                    elevation = CardDefaults.cardElevation(defaultElevation = 4.dp)
                ) {
                    Box(
                        modifier = Modifier
                            .fillMaxWidth()
                            .background(
                                Brush.horizontalGradient(
                                    colors = listOf(SlateDark, SlateSecondary)
                                )
                            )
                            .padding(18.dp)
                    ) {
                        Column {
                            Row(
                                modifier = Modifier.fillMaxWidth(),
                                horizontalArrangement = Arrangement.SpaceBetween,
                                verticalAlignment = Alignment.CenterVertically
                            ) {
                                Column {
                                    Text(
                                        text = "PORTAFOLIO EJECUTIVO O-LEASING",
                                        color = Color(0xFF94A3B8),
                                        fontSize = 10.sp,
                                        fontWeight = FontWeight.Bold,
                                        letterSpacing = 0.8.sp
                                    )
                                    Text(
                                        text = currentUser?.nombreCompleto ?: "Sesión Activa",
                                        color = Color.White,
                                        fontSize = 15.sp,
                                        fontWeight = FontWeight.Bold
                                    )
                                }
                                Box(
                                    modifier = Modifier
                                        .background(OrangePrimary, RoundedCornerShape(8.dp))
                                        .padding(horizontal = 10.dp, vertical = 4.dp)
                                ) {
                                    Text(
                                        text = (currentUser?.rol ?: "ADMIN").uppercase(),
                                        color = Color.White,
                                        fontSize = 10.sp,
                                        fontWeight = FontWeight.ExtraBold
                                    )
                                }
                            }

                            Spacer(modifier = Modifier.height(14.dp))

                            Row(
                                modifier = Modifier.fillMaxWidth(),
                                horizontalArrangement = Arrangement.SpaceBetween,
                                verticalAlignment = Alignment.Bottom
                            ) {
                                Column {
                                    Text(
                                        text = "Valor Cartera Activa (s/IVA)",
                                        color = Color(0xFFCBD5E1),
                                        fontSize = 11.sp
                                    )
                                    Text(
                                        text = metrics?.let { formatCurrency(it.carteraActivaValor) } ?: "$0.00",
                                        color = Color.White,
                                        fontSize = 23.sp,
                                        fontWeight = FontWeight.ExtraBold
                                    )
                                }
                                Column(horizontalAlignment = Alignment.End) {
                                    Text(
                                        text = "Flujo Mensual",
                                        color = Color(0xFFCBD5E1),
                                        fontSize = 11.sp
                                    )
                                    Text(
                                        text = metrics?.let { formatCurrency(it.flujoMensualEstimado) } ?: "$0.00",
                                        color = Color(0xFFFFB380),
                                        fontSize = 16.sp,
                                        fontWeight = FontWeight.Bold
                                    )
                                }
                            }

                            if ((metrics?.residualTotalActivo ?: 0.0) > 0.0) {
                                HorizontalDivider(
                                    modifier = Modifier.padding(vertical = 10.dp),
                                    color = Color(0xFF334155)
                                )
                                Row(
                                    modifier = Modifier.fillMaxWidth(),
                                    horizontalArrangement = Arrangement.SpaceBetween
                                ) {
                                    Text(
                                        text = "Valor Residual Pactado en Activos:",
                                        color = Color(0xFF94A3B8),
                                        fontSize = 11.sp
                                    )
                                    Text(
                                        text = formatCurrency(metrics?.residualTotalActivo ?: 0.0),
                                        color = Color.White,
                                        fontSize = 12.sp,
                                        fontWeight = FontWeight.Bold
                                    )
                                }
                            }
                        }
                    }
                }
            }

            if (errorMessage != null) {
                item {
                    Card(
                        colors = CardDefaults.cardColors(containerColor = Color(0xFFFEE2E2)),
                        shape = RoundedCornerShape(12.dp),
                        modifier = Modifier.fillMaxWidth()
                    ) {
                        Column(modifier = Modifier.padding(14.dp)) {
                            Text(
                                text = "Aviso de Conexión",
                                fontWeight = FontWeight.Bold,
                                color = Color(0xFFB91C1C),
                                fontSize = 14.sp
                            )
                            Spacer(modifier = Modifier.height(4.dp))
                            Text(
                                text = errorMessage ?: "",
                                color = Color(0xFF7F1D1D),
                                fontSize = 12.sp
                            )
                            Spacer(modifier = Modifier.height(8.dp))
                            Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                                Button(
                                    onClick = { loadData() },
                                    colors = ButtonDefaults.buttonColors(containerColor = Color(0xFFDC2626)),
                                    contentPadding = PaddingValues(horizontal = 12.dp, vertical = 6.dp)
                                ) {
                                    Text("Reintentar", fontSize = 12.sp)
                                }
                                OutlinedButton(
                                    onClick = {
                                        serverHostInput = OLeasingApiClient.baseUrl
                                        showServerDialog = true
                                    },
                                    contentPadding = PaddingValues(horizontal = 12.dp, vertical = 6.dp)
                                ) {
                                    Text("Configurar URL", fontSize = 12.sp, color = Color(0xFFB91C1C))
                                }
                            }
                        }
                    }
                }
            }

            // KPIs y barra de composición de cartera
            item {
                Row(
                    modifier = Modifier.fillMaxWidth(),
                    horizontalArrangement = Arrangement.spacedBy(12.dp)
                ) {
                    StatCard(
                        title = "Contratos Activos",
                        value = metrics?.contratosActivos?.toString() ?: "0",
                        subtitle = "Por vencer (90d): ${metrics?.contratosPorVencer90d ?: 0}",
                        colorAccent = Color(0xFF16A34A),
                        modifier = Modifier.weight(1f)
                    )
                    StatCard(
                        title = "Total Histórico",
                        value = metrics?.totalContratos?.toString() ?: "0",
                        subtitle = "Bajas: ${metrics?.contratosBaja ?: 0} · Mora: ${metrics?.contratosEnMora ?: 0}",
                        colorAccent = SlateSecondary,
                        modifier = Modifier.weight(1f)
                    )
                }
            }

            // Barra visual de composición de cartera
            metrics?.let { m ->
                val tot = (m.totalContratos).coerceAtLeast(1).toFloat()
                val wAct = (m.contratosActivos / tot).coerceIn(0.05f, 1f)
                val wBaja = (m.contratosBaja / tot).coerceIn(0.05f, 1f)
                item {
                    Card(
                        modifier = Modifier.fillMaxWidth(),
                        shape = RoundedCornerShape(14.dp),
                        colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
                        border = BorderStroke(1.dp, AppCardBorder)
                    ) {
                        Column(modifier = Modifier.padding(14.dp)) {
                            Row(
                                modifier = Modifier.fillMaxWidth(),
                                horizontalArrangement = Arrangement.SpaceBetween
                            ) {
                                Text(
                                    text = "Distribución del Portafolio",
                                    fontSize = 12.sp,
                                    fontWeight = FontWeight.Bold,
                                    color = SlateSecondary
                                )
                                Text(
                                    text = "${((m.contratosActivos * 100) / m.totalContratos.coerceAtLeast(1))}% Vigente",
                                    fontSize = 12.sp,
                                    fontWeight = FontWeight.ExtraBold,
                                    color = Color(0xFF16A34A)
                                )
                            }
                            Spacer(modifier = Modifier.height(8.dp))
                            Row(
                                modifier = Modifier
                                    .fillMaxWidth()
                                    .height(10.dp)
                                    .clip(RoundedCornerShape(5.dp))
                            ) {
                                Box(
                                    modifier = Modifier
                                        .weight(wAct)
                                        .fillMaxHeight()
                                        .background(Color(0xFF16A34A))
                                )
                                if (m.contratosConcluidos > 0) {
                                    Box(
                                        modifier = Modifier
                                            .weight((m.contratosConcluidos / tot).coerceAtLeast(0.03f))
                                            .fillMaxHeight()
                                            .background(Color(0xFF2563EB))
                                    )
                                }
                                Box(
                                    modifier = Modifier
                                        .weight(wBaja)
                                        .fillMaxHeight()
                                        .background(Color(0xFF94A3B8))
                                )
                            }
                        }
                    }
                }
            }

            item {
                Text(
                    text = "Operaciones y Accesos Directos",
                    fontSize = 15.sp,
                    fontWeight = FontWeight.Bold,
                    color = MaterialTheme.colorScheme.onBackground
                )
                Spacer(modifier = Modifier.height(8.dp))
                Row(
                    modifier = Modifier.fillMaxWidth(),
                    horizontalArrangement = Arrangement.spacedBy(10.dp)
                ) {
                    Button(
                        onClick = onNavigateToList,
                        modifier = Modifier
                            .weight(1f)
                            .height(50.dp),
                        shape = RoundedCornerShape(12.dp),
                        colors = ButtonDefaults.buttonColors(containerColor = SlateSecondary)
                    ) {
                        Icon(imageVector = Icons.Default.Search, contentDescription = null)
                        Spacer(modifier = Modifier.width(6.dp))
                        Text("Cartera", fontWeight = FontWeight.Bold)
                    }

                    if (canEdit) {
                        Button(
                            onClick = onNavigateToNew,
                            modifier = Modifier
                                .weight(1f)
                                .height(50.dp),
                            shape = RoundedCornerShape(12.dp),
                            colors = ButtonDefaults.buttonColors(containerColor = OrangePrimary)
                        ) {
                            Icon(imageVector = Icons.Default.Add, contentDescription = null)
                            Spacer(modifier = Modifier.width(6.dp))
                            Text("Alta Contrato", fontWeight = FontWeight.Bold)
                        }
                    }
                }
                Spacer(modifier = Modifier.height(8.dp))
                OutlinedButton(
                    onClick = { showQuickJumpDialog = true },
                    modifier = Modifier
                        .fillMaxWidth()
                        .height(46.dp),
                    shape = RoundedCornerShape(12.dp),
                    border = BorderStroke(1.dp, OrangePrimary)
                ) {
                    Text(
                        text = "Abrir Estado de Cuenta por ID o Escaneo QR",
                        fontWeight = FontWeight.Bold,
                        color = OrangePrimary,
                        fontSize = 13.sp
                    )
                }
            }

            // Alertas de Vencimientos Próximos (90 días)
            val vencimientos = metrics?.vencimientosProximos ?: emptyList()
            if (vencimientos.isNotEmpty()) {
                item {
                    Text(
                        text = "Alertas de Vencimiento Próximo (90 días)",
                        fontSize = 15.sp,
                        fontWeight = FontWeight.Bold,
                        color = MaterialTheme.colorScheme.onBackground
                    )
                }
                items(vencimientos.take(3)) { v ->
                    Card(
                        modifier = Modifier
                            .fillMaxWidth()
                            .clickable { onNavigateToEstadoCuenta(v.idContrato) },
                        shape = RoundedCornerShape(12.dp),
                        colors = CardDefaults.cardColors(containerColor = Color(0xFFFFFBEB)),
                        border = BorderStroke(1.dp, Color(0xFFFDE68A))
                    ) {
                        Row(
                            modifier = Modifier
                                .fillMaxWidth()
                                .padding(12.dp),
                            horizontalArrangement = Arrangement.SpaceBetween,
                            verticalAlignment = Alignment.CenterVertically
                        ) {
                            Column(modifier = Modifier.weight(1f)) {
                                Row(verticalAlignment = Alignment.CenterVertically) {
                                    Text(
                                        text = v.idContrato,
                                        fontWeight = FontWeight.ExtraBold,
                                        fontSize = 13.sp,
                                        color = Color(0xFFB45309)
                                    )
                                    Spacer(modifier = Modifier.width(8.dp))
                                    Text(
                                        text = "Vence: ${v.fechaVencimiento}",
                                        fontSize = 11.sp,
                                        color = Color(0xFF92400E)
                                    )
                                }
                                Text(
                                    text = v.cliente,
                                    fontWeight = FontWeight.Bold,
                                    fontSize = 13.sp,
                                    color = SlateDark
                                )
                                Text(
                                    text = "Residual: ${formatCurrency(v.residualMonto)}",
                                    fontSize = 11.sp,
                                    fontWeight = FontWeight.SemiBold,
                                    color = SlateSecondary
                                )
                            }
                            Box(
                                modifier = Modifier
                                    .background(Color(0xFFFEF3C7), RoundedCornerShape(8.dp))
                                    .padding(horizontal = 10.dp, vertical = 6.dp)
                            ) {
                                Text(
                                    text = "${v.diasRestantes} días",
                                    fontWeight = FontWeight.ExtraBold,
                                    fontSize = 11.sp,
                                    color = Color(0xFFB45309)
                                )
                            }
                        }
                    }
                }
            }

            item {
                Row(
                    modifier = Modifier.fillMaxWidth(),
                    horizontalArrangement = Arrangement.SpaceBetween,
                    verticalAlignment = Alignment.CenterVertically
                ) {
                    Text(
                        text = "Contratos Activos Recientes",
                        fontSize = 15.sp,
                        fontWeight = FontWeight.Bold,
                        color = MaterialTheme.colorScheme.onBackground
                    )
                    TextButton(onClick = onNavigateToList) {
                        Text("Ver Cartera Completa", color = OrangePrimary, fontSize = 12.sp, fontWeight = FontWeight.Bold)
                    }
                }
            }

            if (isLoading && recentContratos.isEmpty()) {
                item {
                    Box(
                        modifier = Modifier
                            .fillMaxWidth()
                            .padding(24.dp),
                        contentAlignment = Alignment.Center
                    ) {
                        CircularProgressIndicator(color = OrangePrimary)
                    }
                }
            } else if (recentContratos.isEmpty() && !isLoading) {
                item {
                    Card(
                        modifier = Modifier.fillMaxWidth(),
                        shape = RoundedCornerShape(12.dp),
                        colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
                        border = BorderStroke(1.dp, AppCardBorder)
                    ) {
                        Text(
                            text = "No hay contratos cargados o el servidor no responde.",
                            modifier = Modifier.padding(16.dp),
                            color = MaterialTheme.colorScheme.outline,
                            fontSize = 13.sp
                        )
                    }
                }
            } else {
                items(recentContratos) { c ->
                    Card(
                        modifier = Modifier
                            .fillMaxWidth()
                            .clickable { onNavigateToEstadoCuenta(c.idContrato) },
                        shape = RoundedCornerShape(14.dp),
                        colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
                        elevation = CardDefaults.cardElevation(defaultElevation = 2.dp),
                        border = BorderStroke(1.dp, AppCardBorder)
                    ) {
                        Column(modifier = Modifier.padding(14.dp)) {
                            Row(
                                modifier = Modifier.fillMaxWidth(),
                                horizontalArrangement = Arrangement.SpaceBetween,
                                verticalAlignment = Alignment.CenterVertically
                            ) {
                                Text(
                                    text = c.idContrato,
                                    fontWeight = FontWeight.ExtraBold,
                                    fontSize = 14.sp,
                                    color = OrangePrimary
                                )
                                Row(horizontalArrangement = Arrangement.spacedBy(6.dp)) {
                                    MorosidadBadge(nivel = c.nivelMorosidad)
                                    StatusBadge(estatus = c.estatus)
                                }
                            }
                            Spacer(modifier = Modifier.height(4.dp))
                            Text(
                                text = c.cliente,
                                fontWeight = FontWeight.Bold,
                                fontSize = 14.sp,
                                color = MaterialTheme.colorScheme.onSurface
                            )
                            if (c.vehiculo.isNotBlank()) {
                                Text(
                                    text = c.vehiculo,
                                    fontSize = 12.sp,
                                    color = MaterialTheme.colorScheme.outline
                                )
                            }
                            HorizontalDivider(
                                modifier = Modifier.padding(vertical = 8.dp),
                                color = MaterialTheme.colorScheme.surfaceVariant
                            )
                            Row(
                                modifier = Modifier.fillMaxWidth(),
                                horizontalArrangement = Arrangement.SpaceBetween,
                                verticalAlignment = Alignment.CenterVertically
                            ) {
                                Text(
                                    text = "Renta: ${formatCurrency(c.mensualidadSinIva)}/mes",
                                    fontSize = 12.sp,
                                    fontWeight = FontWeight.SemiBold,
                                    color = SlateSecondary
                                )
                                Text(
                                    text = "Plazo ${c.plazo} m · Estado Cta.",
                                    fontSize = 12.sp,
                                    color = OrangePrimary,
                                    fontWeight = FontWeight.Bold
                                )
                            }
                        }
                    }
                }
            }

            item {
                Spacer(modifier = Modifier.height(16.dp))
            }
        }
    }

    if (showServerDialog) {
        AlertDialog(
            onDismissRequest = { showServerDialog = false },
            title = { Text("Conexión del Servidor O-Leasing", fontWeight = FontWeight.Bold) },
            text = {
                Column {
                    Text(
                        "Ingresa tu URL de Túnel Web (ej. https://tu-tunel.trycloudflare.com) para conectarte desde cualquier lugar sin usar IP local, o http://10.0.2.2:8008 en emulador:",
                        fontSize = 12.sp,
                        color = MaterialTheme.colorScheme.outline
                    )
                    Spacer(modifier = Modifier.height(10.dp))
                    OutlinedTextField(
                        value = serverHostInput,
                        onValueChange = { serverHostInput = it },
                        label = { Text("URL Base del Servidor") },
                        singleLine = true,
                        modifier = Modifier.fillMaxWidth()
                    )
                    Spacer(modifier = Modifier.height(8.dp))
                    OutlinedButton(
                        onClick = {
                            coroutineScope.launch {
                                OLeasingApiClient.updateBaseUrl(context, serverHostInput)
                                val res = OLeasingApiClient.checkHealth()
                                pingStatusMessage = if (res.isSuccess) {
                                    "Conexión exitosa con el servidor (${OLeasingApiClient.baseUrl})"
                                } else {
                                    "Sin respuesta: ${res.exceptionOrNull()?.message}"
                                }
                            }
                        },
                        modifier = Modifier.fillMaxWidth()
                    ) {
                        Text("Probar Conexión Ahora", fontSize = 12.sp, fontWeight = FontWeight.Bold)
                    }
                    if (pingStatusMessage != null) {
                        Spacer(modifier = Modifier.height(6.dp))
                        Text(
                            text = pingStatusMessage ?: "",
                            fontSize = 11.sp,
                            fontWeight = FontWeight.SemiBold,
                            color = if (pingStatusMessage?.startsWith("Conexión exitosa") == true) Color(0xFF16A34A) else Color(0xFFDC2626)
                        )
                    }
                }
            },
            confirmButton = {
                Button(
                    onClick = {
                        OLeasingApiClient.updateBaseUrl(context, serverHostInput)
                        showServerDialog = false
                        loadData()
                    },
                    colors = ButtonDefaults.buttonColors(containerColor = OrangePrimary)
                ) {
                    Text("Guardar Permanente")
                }
            },
            dismissButton = {
                TextButton(onClick = { showServerDialog = false }) {
                    Text("Cancelar")
                }
            }
        )
    }

    if (showPasswordDialog) {
        AlertDialog(
            onDismissRequest = { showPasswordDialog = false },
            title = { Text("Seguridad de la Cuenta", fontWeight = FontWeight.Bold) },
            text = {
                Column(verticalArrangement = Arrangement.spacedBy(8.dp)) {
                    Text(
                        text = "Usuario: ${currentUser?.username ?: ""} (${currentUser?.rol?.uppercase() ?: ""})",
                        fontSize = 12.sp,
                        fontWeight = FontWeight.SemiBold,
                        color = SlateSecondary
                    )
                    OutlinedTextField(
                        value = currentPwInput,
                        onValueChange = { currentPwInput = it },
                        label = { Text("Contraseña Actual") },
                        visualTransformation = PasswordVisualTransformation(),
                        singleLine = true,
                        modifier = Modifier.fillMaxWidth()
                    )
                    OutlinedTextField(
                        value = newPwInput,
                        onValueChange = { newPwInput = it },
                        label = { Text("Nueva Contraseña (mín. 6 caracteres)") },
                        visualTransformation = PasswordVisualTransformation(),
                        singleLine = true,
                        modifier = Modifier.fillMaxWidth()
                    )
                    if (pwFeedback != null) {
                        Text(
                            text = pwFeedback ?: "",
                            fontSize = 12.sp,
                            fontWeight = FontWeight.SemiBold,
                            color = if (pwSuccess) Color(0xFF16A34A) else Color(0xFFDC2626)
                        )
                    }
                }
            },
            confirmButton = {
                Button(
                    onClick = {
                        coroutineScope.launch {
                            val res = OLeasingApiClient.cambiarPassword(currentPwInput, newPwInput)
                            if (res.isSuccess) {
                                pwSuccess = true
                                pwFeedback = res.getOrNull() ?: "Contraseña actualizada."
                                currentPwInput = ""
                                newPwInput = ""
                            } else {
                                pwSuccess = false
                                pwFeedback = res.exceptionOrNull()?.message ?: "Error al cambiar contraseña"
                            }
                        }
                    },
                    colors = ButtonDefaults.buttonColors(containerColor = OrangePrimary)
                ) {
                    Text("Actualizar Contraseña")
                }
            },
            dismissButton = {
                TextButton(onClick = { showPasswordDialog = false }) {
                    Text("Cerrar")
                }
            }
        )
    }

    if (showQuickJumpDialog) {
        AlertDialog(
            onDismissRequest = { showQuickJumpDialog = false },
            title = { Text("Abrir Estado de Cuenta", fontWeight = FontWeight.Bold) },
            text = {
                Column {
                    Text(
                        "También puedes escanear el código QR del Estado de Cuenta directamente con la cámara de tu teléfono y se abrirá automáticamente aquí.",
                        fontSize = 12.sp,
                        color = MaterialTheme.colorScheme.outline
                    )
                    Spacer(modifier = Modifier.height(10.dp))
                    OutlinedTextField(
                        value = quickJumpId,
                        onValueChange = { quickJumpId = it.uppercase() },
                        label = { Text("ID Contrato (ej. 0786-0001)") },
                        singleLine = true,
                        modifier = Modifier.fillMaxWidth()
                    )
                }
            },
            confirmButton = {
                Button(
                    onClick = {
                        val targetId = quickJumpId.trim()
                        if (targetId.isNotEmpty()) {
                            showQuickJumpDialog = false
                            quickJumpId = ""
                            onNavigateToEstadoCuenta(targetId)
                        }
                    },
                    colors = ButtonDefaults.buttonColors(containerColor = OrangePrimary)
                ) {
                    Text("Abrir Estado de Cuenta")
                }
            },
            dismissButton = {
                TextButton(onClick = { showQuickJumpDialog = false }) {
                    Text("Cancelar")
                }
            }
        )
    }
}
