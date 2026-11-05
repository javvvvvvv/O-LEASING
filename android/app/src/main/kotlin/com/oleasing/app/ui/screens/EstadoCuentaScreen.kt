package com.oleasing.app.ui.screens

import android.content.Intent
import android.net.Uri
import androidx.compose.foundation.background
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.LazyRow
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.filled.ArrowBack
import androidx.compose.material.icons.filled.Refresh
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import com.oleasing.app.model.EstadoCuentaResponse
import com.oleasing.app.network.OLeasingApiClient
import com.oleasing.app.ui.components.InfoRow
import com.oleasing.app.ui.components.StatusBadge
import com.oleasing.app.ui.components.formatCurrency
import com.oleasing.app.ui.components.formatPercent
import com.oleasing.app.ui.theme.OrangePrimary
import com.oleasing.app.ui.theme.SlateSecondary
import com.oleasing.app.ui.theme.StatusActivo
import com.oleasing.app.ui.theme.StatusActivoBg
import com.oleasing.app.ui.theme.StatusConcluido
import com.oleasing.app.ui.theme.StatusConcluidoBg
import kotlinx.coroutines.launch

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun EstadoCuentaScreen(
    idContrato: String,
    onBack: () -> Unit,
    onNavigateToDetail: (String) -> Unit
) {
    val context = LocalContext.current
    val coroutineScope = rememberCoroutineScope()
    var estadoCuenta by remember { mutableStateOf<EstadoCuentaResponse?>(null) }
    var isLoading by remember { mutableStateOf(false) }
    var errorMessage by remember { mutableStateOf<String?>(null) }
    var filtroPeriodo by remember { mutableStateOf("Todos") }
    val filtros = listOf("Todos", "Pagados", "Pendientes")

    fun loadEstadoCuenta() {
        coroutineScope.launch {
            isLoading = true
            errorMessage = null
            val res = OLeasingApiClient.getEstadoCuenta(idContrato)
            if (res.isSuccess) {
                estadoCuenta = res.getOrNull()
            } else {
                errorMessage = res.exceptionOrNull()?.message
            }
            isLoading = false
        }
    }

    LaunchedEffect(idContrato) {
        loadEstadoCuenta()
    }

    Scaffold(
        topBar = {
            TopAppBar(
                title = {
                    Column {
                        Text(
                            text = "Estado de Cuenta · $idContrato",
                            fontWeight = FontWeight.Bold,
                            fontSize = 17.sp,
                            color = OrangePrimary
                        )
                        Text(
                            text = estadoCuenta?.cliente ?: "Cargando expediente...",
                            fontSize = 11.sp,
                            color = MaterialTheme.colorScheme.outline
                        )
                    }
                },
                navigationIcon = {
                    IconButton(onClick = onBack) {
                        Icon(imageVector = Icons.AutoMirrored.Filled.ArrowBack, contentDescription = "Volver")
                    }
                },
                actions = {
                    IconButton(onClick = { loadEstadoCuenta() }) {
                        Icon(imageVector = Icons.Default.Refresh, contentDescription = "Actualizar", tint = OrangePrimary)
                    }
                },
                colors = TopAppBarDefaults.topAppBarColors(containerColor = MaterialTheme.colorScheme.surface)
            )
        }
    ) { padding ->
        if (isLoading && estadoCuenta == null) {
            Box(
                modifier = Modifier
                    .fillMaxSize()
                    .padding(padding),
                contentAlignment = Alignment.Center
            ) {
                CircularProgressIndicator(color = OrangePrimary)
            }
        } else if (errorMessage != null && estadoCuenta == null) {
            Box(
                modifier = Modifier
                    .fillMaxSize()
                    .padding(padding)
                    .padding(16.dp),
                contentAlignment = Alignment.Center
            ) {
                Column(horizontalAlignment = Alignment.CenterHorizontally) {
                    Text("No se pudo cargar el Estado de Cuenta:", fontWeight = FontWeight.Bold, color = Color(0xFFDC2626))
                    Spacer(modifier = Modifier.height(6.dp))
                    Text(errorMessage ?: "", fontSize = 13.sp, color = MaterialTheme.colorScheme.outline)
                    Spacer(modifier = Modifier.height(12.dp))
                    Button(onClick = { loadEstadoCuenta() }, colors = ButtonDefaults.buttonColors(containerColor = OrangePrimary)) {
                        Text("Reintentar")
                    }
                }
            }
        } else if (estadoCuenta != null) {
            val ec = estadoCuenta!!
            val periodosFiltrados = ec.tablaAmortizacion.filter { p ->
                when (filtroPeriodo) {
                    "Pagados" -> p.estatus == "PAGADO" || p.estatus == "ACTUAL"
                    "Pendientes" -> p.estatus == "PENDIENTE"
                    else -> true
                }
            }

            LazyColumn(
                modifier = Modifier
                    .fillMaxSize()
                    .padding(padding)
                    .background(MaterialTheme.colorScheme.background)
                    .padding(horizontal = 16.dp),
                verticalArrangement = Arrangement.spacedBy(12.dp)
            ) {
                item { Spacer(modifier = Modifier.height(4.dp)) }

                // Cabecera y barra de progreso
                item {
                    Card(
                        modifier = Modifier.fillMaxWidth(),
                        shape = RoundedCornerShape(14.dp),
                        colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
                        elevation = CardDefaults.cardElevation(defaultElevation = 2.dp)
                    ) {
                        Column(modifier = Modifier.padding(16.dp)) {
                            Row(
                                modifier = Modifier.fillMaxWidth(),
                                horizontalArrangement = Arrangement.SpaceBetween,
                                verticalAlignment = Alignment.CenterVertically
                            ) {
                                Column(modifier = Modifier.weight(1f)) {
                                    Text(
                                        text = ec.cliente,
                                        fontWeight = FontWeight.Bold,
                                        fontSize = 16.sp,
                                        color = MaterialTheme.colorScheme.onSurface
                                    )
                                    if (ec.vehiculo.isNotBlank()) {
                                        Text(
                                            text = ec.vehiculo,
                                            fontSize = 12.sp,
                                            color = MaterialTheme.colorScheme.outline
                                        )
                                    }
                                }
                                StatusBadge(estatus = ec.estatus)
                            }

                            Spacer(modifier = Modifier.height(12.dp))

                            Row(
                                modifier = Modifier.fillMaxWidth(),
                                horizontalArrangement = Arrangement.SpaceBetween
                            ) {
                                Text(
                                    text = "Avance del contrato: Mes ${ec.mesesTranscurridos} de ${ec.plazo}",
                                    fontSize = 12.sp,
                                    fontWeight = FontWeight.SemiBold,
                                    color = SlateSecondary
                                )
                                Text(
                                    text = "${ec.porcentajeAvance}%",
                                    fontSize = 12.sp,
                                    fontWeight = FontWeight.Bold,
                                    color = OrangePrimary
                                )
                            }
                            Spacer(modifier = Modifier.height(6.dp))
                            LinearProgressIndicator(
                                progress = { (ec.porcentajeAvance.toFloat() / 100f).coerceIn(0f, 1f) },
                                modifier = Modifier
                                    .fillMaxWidth()
                                    .height(8.dp),
                                color = OrangePrimary,
                                trackColor = MaterialTheme.colorScheme.surfaceVariant,
                            )
                            Spacer(modifier = Modifier.height(8.dp))
                            Row(
                                modifier = Modifier.fillMaxWidth(),
                                horizontalArrangement = Arrangement.SpaceBetween
                            ) {
                                Text("Alta: ${ec.fechaAlta}", fontSize = 11.sp, color = MaterialTheme.colorScheme.outline)
                                Text("Restan: ${ec.pagosRestantes} pagos", fontSize = 11.sp, fontWeight = FontWeight.Medium, color = SlateSecondary)
                                Text("Vence: ${ec.fechaVencimiento}", fontSize = 11.sp, color = MaterialTheme.colorScheme.outline)
                            }
                        }
                    }
                }

                // Resumen Ejecutivo del Estado de Cuenta
                item {
                    Card(
                        modifier = Modifier.fillMaxWidth(),
                        shape = RoundedCornerShape(14.dp),
                        colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
                        elevation = CardDefaults.cardElevation(defaultElevation = 2.dp)
                    ) {
                        Column(modifier = Modifier.padding(16.dp)) {
                            Text(
                                text = "Saldos y Conciliación del Contrato",
                                fontWeight = FontWeight.Bold,
                                fontSize = 14.sp,
                                color = SlateSecondary
                            )
                            HorizontalDivider(modifier = Modifier.padding(vertical = 8.dp), color = MaterialTheme.colorScheme.surfaceVariant)
                            InfoRow(label = "Inversión Neta Inicial", value = formatCurrency(ec.inversionNeta))
                            InfoRow(label = "Capital Amortizado a la Fecha", value = formatCurrency(ec.capitalAmortizado))
                            InfoRow(label = "Saldo Insoluto Actual (Capital)", value = formatCurrency(ec.saldoActualEstimado), isBold = true)
                            InfoRow(label = "Intereses Devengados a la Fecha", value = formatCurrency(ec.interesDevengado))
                            InfoRow(label = "Renta Mensual sin IVA", value = formatCurrency(ec.mensualidadSinIva), isBold = true)
                            InfoRow(label = "Valor Residual al Vencimiento", value = formatCurrency(ec.residualMonto))
                            InfoRow(label = "Tasa Anual Implícita", value = formatPercent(ec.tasaAnualPct))

                            Spacer(modifier = Modifier.height(8.dp))
                            Row(
                                modifier = Modifier.fillMaxWidth(),
                                horizontalArrangement = Arrangement.spacedBy(8.dp)
                            ) {
                                OutlinedButton(
                                    onClick = { onNavigateToDetail(ec.idContrato) },
                                    modifier = Modifier.weight(1f),
                                    shape = RoundedCornerShape(10.dp)
                                ) {
                                    Text("Ficha Técnica", color = SlateSecondary, fontSize = 12.sp, fontWeight = FontWeight.SemiBold)
                                }
                                Button(
                                    onClick = {
                                        val pdfUrl = OLeasingApiClient.getPdfEstadoCuentaUrl(ec.idContrato)
                                        val intent = Intent(Intent.ACTION_VIEW, Uri.parse(pdfUrl))
                                        context.startActivity(intent)
                                    },
                                    modifier = Modifier.weight(1f),
                                    shape = RoundedCornerShape(10.dp),
                                    colors = ButtonDefaults.buttonColors(containerColor = OrangePrimary)
                                ) {
                                    Text("Descargar PDF", color = Color.White, fontSize = 12.sp, fontWeight = FontWeight.Bold)
                                }
                            }
                        }
                    }
                }

                // Selector de filtro de tabla de amortización
                item {
                    Row(
                        modifier = Modifier.fillMaxWidth(),
                        horizontalArrangement = Arrangement.SpaceBetween,
                        verticalAlignment = Alignment.CenterVertically
                    ) {
                        Text(
                            text = "Tabla de Amortización (${periodosFiltrados.size})",
                            fontWeight = FontWeight.Bold,
                            fontSize = 15.sp,
                            color = MaterialTheme.colorScheme.onBackground
                        )
                    }
                    Spacer(modifier = Modifier.height(6.dp))
                    LazyRow(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                        items(filtros) { f ->
                            val sel = f == filtroPeriodo
                            FilterChip(
                                selected = sel,
                                onClick = { filtroPeriodo = f },
                                label = { Text(f, fontSize = 12.sp, fontWeight = if (sel) FontWeight.Bold else FontWeight.Normal) },
                                colors = FilterChipDefaults.filterChipColors(
                                    selectedContainerColor = OrangePrimary,
                                    selectedLabelColor = Color.White
                                )
                            )
                        }
                    }
                }

                // Filas de la tabla de amortización
                items(periodosFiltrados) { p ->
                    val (badgeBg, badgeColor) = when (p.estatus) {
                        "PAGADO" -> Pair(StatusActivoBg, StatusActivo)
                        "ACTUAL" -> Pair(Color(0xFFFFEDD5), OrangePrimary)
                        else -> Pair(StatusConcluidoBg, StatusConcluido)
                    }

                    Card(
                        modifier = Modifier.fillMaxWidth(),
                        shape = RoundedCornerShape(10.dp),
                        colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
                        elevation = CardDefaults.cardElevation(defaultElevation = 1.dp)
                    ) {
                        Column(modifier = Modifier.padding(12.dp)) {
                            Row(
                                modifier = Modifier.fillMaxWidth(),
                                horizontalArrangement = Arrangement.SpaceBetween,
                                verticalAlignment = Alignment.CenterVertically
                            ) {
                                Row(verticalAlignment = Alignment.CenterVertically) {
                                    Text(
                                        text = "Mes ${p.periodo}",
                                        fontWeight = FontWeight.Bold,
                                        fontSize = 14.sp,
                                        color = SlateSecondary
                                    )
                                    Spacer(modifier = Modifier.width(10.dp))
                                    Text(
                                        text = p.fechaEstimada,
                                        fontSize = 12.sp,
                                        color = MaterialTheme.colorScheme.outline
                                    )
                                }
                                Box(
                                    modifier = Modifier
                                        .background(badgeBg, shape = RoundedCornerShape(6.dp))
                                        .padding(horizontal = 8.dp, vertical = 2.dp)
                                ) {
                                    Text(
                                        text = p.estatus,
                                        color = badgeColor,
                                        fontSize = 10.sp,
                                        fontWeight = FontWeight.Bold
                                    )
                                }
                            }

                            HorizontalDivider(modifier = Modifier.padding(vertical = 6.dp), color = MaterialTheme.colorScheme.surfaceVariant)

                            Row(
                                modifier = Modifier.fillMaxWidth(),
                                horizontalArrangement = Arrangement.SpaceBetween
                            ) {
                                Column {
                                    Text("Capital", fontSize = 10.sp, color = MaterialTheme.colorScheme.outline)
                                    Text(formatCurrency(p.capital), fontSize = 12.sp, fontWeight = FontWeight.Medium)
                                }
                                Column {
                                    Text("Interés", fontSize = 10.sp, color = MaterialTheme.colorScheme.outline)
                                    Text(formatCurrency(p.interes), fontSize = 12.sp, fontWeight = FontWeight.Medium)
                                }
                                Column {
                                    Text("IVA (16%)", fontSize = 10.sp, color = MaterialTheme.colorScheme.outline)
                                    Text(formatCurrency(p.iva), fontSize = 12.sp, fontWeight = FontWeight.Medium)
                                }
                                Column(horizontalAlignment = Alignment.End) {
                                    Text("Total c/IVA", fontSize = 10.sp, color = MaterialTheme.colorScheme.outline)
                                    Text(formatCurrency(p.totalConIva), fontSize = 12.sp, fontWeight = FontWeight.Bold, color = OrangePrimary)
                                }
                            }

                            Spacer(modifier = Modifier.height(6.dp))
                            Row(
                                modifier = Modifier.fillMaxWidth(),
                                horizontalArrangement = Arrangement.SpaceBetween
                            ) {
                                Text(
                                    text = "Renta s/IVA: ${formatCurrency(p.rentaSinIva)}",
                                    fontSize = 11.sp,
                                    color = MaterialTheme.colorScheme.outline
                                )
                                Text(
                                    text = "Saldo Insoluto: ${formatCurrency(p.saldoInsoluto)}",
                                    fontSize = 11.sp,
                                    fontWeight = FontWeight.SemiBold,
                                    color = SlateSecondary
                                )
                            }
                        }
                    }
                }

                item { Spacer(modifier = Modifier.height(28.dp)) }
            }
        }
    }
}
