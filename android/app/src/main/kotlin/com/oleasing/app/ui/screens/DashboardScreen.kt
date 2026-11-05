package com.oleasing.app.ui.screens

import androidx.compose.foundation.background
import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.Add
import androidx.compose.material.icons.filled.Refresh
import androidx.compose.material.icons.filled.Search
import androidx.compose.material.icons.filled.Settings
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import com.oleasing.app.model.ContratoResumen
import com.oleasing.app.model.DashboardMetrics
import com.oleasing.app.network.OLeasingApiClient
import com.oleasing.app.ui.components.StatCard
import com.oleasing.app.ui.components.StatusBadge
import com.oleasing.app.ui.components.formatCurrency
import com.oleasing.app.ui.theme.OrangePrimary
import com.oleasing.app.ui.theme.SlateSecondary
import kotlinx.coroutines.launch

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun DashboardScreen(
    onNavigateToList: () -> Unit,
    onNavigateToNew: () -> Unit,
    onSelectContrato: (String) -> Unit
) {
    val coroutineScope = rememberCoroutineScope()
    var metrics by remember { mutableStateOf<DashboardMetrics?>(null) }
    var recentContratos by remember { mutableStateOf<List<ContratoResumen>>(emptyList()) }
    var isLoading by remember { mutableStateOf(false) }
    var errorMessage by remember { mutableStateOf<String?>(null) }
    var showServerDialog by remember { mutableStateOf(false) }
    var serverHostInput by remember { mutableStateOf(OLeasingApiClient.baseUrl) }

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
                    Column {
                        Text(
                            text = "O-LEASING",
                            fontWeight = FontWeight.Black,
                            fontSize = 20.sp,
                            color = OrangePrimary
                        )
                        Text(
                            text = "Sistema de Cartera & Contratos",
                            fontSize = 11.sp,
                            color = MaterialTheme.colorScheme.outline
                        )
                    }
                },
                actions = {
                    IconButton(onClick = { showServerDialog = true }) {
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
            verticalArrangement = Arrangement.spacedBy(16.dp)
        ) {
            item {
                Spacer(modifier = Modifier.height(4.dp))
                if (errorMessage != null) {
                    Card(
                        colors = CardDefaults.cardColors(containerColor = Color(0xFFFEE2E2)),
                        shape = RoundedCornerShape(10.dp),
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
                            Button(
                                onClick = { loadData() },
                                colors = ButtonDefaults.buttonColors(containerColor = Color(0xFFDC2626)),
                                contentPadding = PaddingValues(horizontal = 12.dp, vertical = 6.dp)
                            ) {
                                Text("Reintentar Conexión", fontSize = 12.sp)
                            }
                        }
                    }
                }
            }

            item {
                Text(
                    text = "Resumen de Cartera",
                    fontSize = 16.sp,
                    fontWeight = FontWeight.Bold,
                    color = MaterialTheme.colorScheme.onBackground
                )
            }

            item {
                Row(
                    modifier = Modifier.fillMaxWidth(),
                    horizontalArrangement = Arrangement.spacedBy(12.dp)
                ) {
                    StatCard(
                        title = "Cartera Activa",
                        value = metrics?.let { formatCurrency(it.carteraActivaValor) } ?: "$0.00",
                        subtitle = "Valor en colocación",
                        colorAccent = OrangePrimary,
                        modifier = Modifier.weight(1f)
                    )
                    StatCard(
                        title = "Flujo Mensual",
                        value = metrics?.let { formatCurrency(it.flujoMensualEstimado) } ?: "$0.00",
                        subtitle = "Renta mensual bruta",
                        colorAccent = SlateSecondary,
                        modifier = Modifier.weight(1f)
                    )
                }
            }

            item {
                Row(
                    modifier = Modifier.fillMaxWidth(),
                    horizontalArrangement = Arrangement.spacedBy(12.dp)
                ) {
                    StatCard(
                        title = "Contratos Activos",
                        value = metrics?.contratosActivos?.toString() ?: "0",
                        subtitle = "Vigentes en sistema",
                        colorAccent = Color(0xFF16A34A),
                        modifier = Modifier.weight(1f)
                    )
                    StatCard(
                        title = "Total Histórico",
                        value = metrics?.totalContratos?.toString() ?: "0",
                        subtitle = "Contratos registrados",
                        colorAccent = SlateSecondary,
                        modifier = Modifier.weight(1f)
                    )
                }
            }

            item {
                Text(
                    text = "Operaciones Rápidas",
                    fontSize = 16.sp,
                    fontWeight = FontWeight.Bold,
                    color = MaterialTheme.colorScheme.onBackground
                )
                Spacer(modifier = Modifier.height(8.dp))
                Row(
                    modifier = Modifier.fillMaxWidth(),
                    horizontalArrangement = Arrangement.spacedBy(12.dp)
                ) {
                    Button(
                        onClick = onNavigateToList,
                        modifier = Modifier
                            .weight(1f)
                            .height(52.dp),
                        shape = RoundedCornerShape(12.dp),
                        colors = ButtonDefaults.buttonColors(containerColor = SlateSecondary)
                    ) {
                        Icon(imageVector = Icons.Default.Search, contentDescription = null)
                        Spacer(modifier = Modifier.width(8.dp))
                        Text("Consultar", fontWeight = FontWeight.SemiBold)
                    }

                    Button(
                        onClick = onNavigateToNew,
                        modifier = Modifier
                            .weight(1f)
                            .height(52.dp),
                        shape = RoundedCornerShape(12.dp),
                        colors = ButtonDefaults.buttonColors(containerColor = OrangePrimary)
                    ) {
                        Icon(imageVector = Icons.Default.Add, contentDescription = null)
                        Spacer(modifier = Modifier.width(8.dp))
                        Text("Nuevo", fontWeight = FontWeight.SemiBold)
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
                        fontSize = 16.sp,
                        fontWeight = FontWeight.Bold,
                        color = MaterialTheme.colorScheme.onBackground
                    )
                    TextButton(onClick = onNavigateToList) {
                        Text("Ver Todos", color = OrangePrimary, fontSize = 13.sp)
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
                        shape = RoundedCornerShape(10.dp),
                        colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface)
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
                            .clickable { onSelectContrato(c.idContrato) },
                        shape = RoundedCornerShape(10.dp),
                        colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
                        elevation = CardDefaults.cardElevation(defaultElevation = 1.dp)
                    ) {
                        Column(modifier = Modifier.padding(14.dp)) {
                            Row(
                                modifier = Modifier.fillMaxWidth(),
                                horizontalArrangement = Arrangement.SpaceBetween,
                                verticalAlignment = Alignment.CenterVertically
                            ) {
                                Text(
                                    text = c.idContrato,
                                    fontWeight = FontWeight.Bold,
                                    fontSize = 14.sp,
                                    color = OrangePrimary
                                )
                                StatusBadge(estatus = c.estatus)
                            }
                            Spacer(modifier = Modifier.height(4.dp))
                            Text(
                                text = c.cliente,
                                fontWeight = FontWeight.SemiBold,
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
                            Spacer(modifier = Modifier.height(6.dp))
                            Row(
                                modifier = Modifier.fillMaxWidth(),
                                horizontalArrangement = Arrangement.SpaceBetween
                            ) {
                                Text(
                                    text = "Renta: ${formatCurrency(c.mensualidadSinIva)}/mes",
                                    fontSize = 12.sp,
                                    fontWeight = FontWeight.Medium
                                )
                                Text(
                                    text = "Plazo: ${c.plazo} m",
                                    fontSize = 12.sp,
                                    color = MaterialTheme.colorScheme.outline
                                )
                            }
                        }
                    }
                }
            }

            item {
                Spacer(modifier = Modifier.height(24.dp))
            }
        }
    }

    if (showServerDialog) {
        AlertDialog(
            onDismissRequest = { showServerDialog = false },
            title = { Text("Configuración de Servidor API", fontWeight = FontWeight.Bold) },
            text = {
                Column {
                    Text(
                        "Ingresa la dirección base del backend FastAPI (ej. http://10.0.2.2:8000 para el emulador o la IP local de tu PC para dispositivo físico):",
                        fontSize = 13.sp,
                        color = MaterialTheme.colorScheme.outline
                    )
                    Spacer(modifier = Modifier.height(10.dp))
                    OutlinedTextField(
                        value = serverHostInput,
                        onValueChange = { serverHostInput = it },
                        label = { Text("URL Base") },
                        singleLine = true,
                        modifier = Modifier.fillMaxWidth()
                    )
                }
            },
            confirmButton = {
                Button(
                    onClick = {
                        OLeasingApiClient.baseUrl = serverHostInput.trim()
                        showServerDialog = false
                        loadData()
                    },
                    colors = ButtonDefaults.buttonColors(containerColor = OrangePrimary)
                ) {
                    Text("Guardar y Probar")
                }
            },
            dismissButton = {
                TextButton(onClick = { showServerDialog = false }) {
                    Text("Cancelar")
                }
            }
        )
    }
}
