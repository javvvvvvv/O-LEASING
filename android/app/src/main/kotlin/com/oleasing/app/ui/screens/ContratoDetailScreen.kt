package com.oleasing.app.ui.screens

import androidx.compose.foundation.background
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.filled.ArrowBack
import androidx.compose.material.icons.filled.Refresh
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import com.oleasing.app.model.ContratoDetalle
import com.oleasing.app.network.OLeasingApiClient
import com.oleasing.app.ui.components.InfoRow
import com.oleasing.app.ui.components.StatusBadge
import com.oleasing.app.ui.components.formatCurrency
import com.oleasing.app.ui.components.formatPercent
import com.oleasing.app.ui.theme.OrangePrimary
import com.oleasing.app.ui.theme.SlateSecondary
import kotlinx.coroutines.launch

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun ContratoDetailScreen(
    idContrato: String,
    onBack: () -> Unit
) {
    val coroutineScope = rememberCoroutineScope()
    var contrato by remember { mutableStateOf<ContratoDetalle?>(null) }
    var isLoading by remember { mutableStateOf(false) }
    var errorMessage by remember { mutableStateOf<String?>(null) }

    fun loadDetalle() {
        coroutineScope.launch {
            isLoading = true
            errorMessage = null
            val result = OLeasingApiClient.getContratoDetalle(idContrato)
            if (result.isSuccess) {
                contrato = result.getOrNull()
            } else {
                errorMessage = result.exceptionOrNull()?.message
            }
            isLoading = false
        }
    }

    LaunchedEffect(idContrato) {
        loadDetalle()
    }

    Scaffold(
        topBar = {
            TopAppBar(
                title = {
                    Column {
                        Text(idContrato, fontWeight = FontWeight.Bold, fontSize = 18.sp, color = OrangePrimary)
                        Text(contrato?.cliente ?: "Detalle del Contrato", fontSize = 11.sp, color = MaterialTheme.colorScheme.outline)
                    }
                },
                navigationIcon = {
                    IconButton(onClick = onBack) {
                        Icon(imageVector = Icons.AutoMirrored.Filled.ArrowBack, contentDescription = "Volver")
                    }
                },
                actions = {
                    IconButton(onClick = { loadDetalle() }) {
                        Icon(imageVector = Icons.Default.Refresh, contentDescription = "Actualizar", tint = OrangePrimary)
                    }
                },
                colors = TopAppBarDefaults.topAppBarColors(containerColor = MaterialTheme.colorScheme.surface)
            )
        }
    ) { padding ->
        if (isLoading && contrato == null) {
            Box(
                modifier = Modifier
                    .fillMaxSize()
                    .padding(padding),
                contentAlignment = Alignment.Center
            ) {
                CircularProgressIndicator(color = OrangePrimary)
            }
        } else if (errorMessage != null) {
            Box(
                modifier = Modifier
                    .fillMaxSize()
                    .padding(padding)
                    .padding(16.dp),
                contentAlignment = Alignment.Center
            ) {
                Column(horizontalAlignment = Alignment.CenterHorizontally) {
                    Text("Error al cargar contrato:", fontWeight = FontWeight.Bold, color = Color(0xFFDC2626))
                    Spacer(modifier = Modifier.height(4.dp))
                    Text(errorMessage ?: "", color = MaterialTheme.colorScheme.outline, fontSize = 13.sp)
                    Spacer(modifier = Modifier.height(12.dp))
                    Button(onClick = { loadDetalle() }, colors = ButtonDefaults.buttonColors(containerColor = OrangePrimary)) {
                        Text("Reintentar")
                    }
                }
            }
        } else if (contrato != null) {
            val c = contrato!!
            LazyColumn(
                modifier = Modifier
                    .fillMaxSize()
                    .padding(padding)
                    .background(MaterialTheme.colorScheme.background)
                    .padding(16.dp),
                verticalArrangement = Arrangement.spacedBy(14.dp)
            ) {
                // Cabecera Principal
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
                                Text(
                                    text = c.idContrato,
                                    fontSize = 20.sp,
                                    fontWeight = FontWeight.Black,
                                    color = OrangePrimary
                                )
                                StatusBadge(estatus = c.estatus)
                            }
                            Spacer(modifier = Modifier.height(6.dp))
                            Text(
                                text = c.cliente,
                                fontSize = 16.sp,
                                fontWeight = FontWeight.Bold,
                                color = MaterialTheme.colorScheme.onSurface
                            )
                            if (c.vehiculo.isNotBlank()) {
                                Spacer(modifier = Modifier.height(2.dp))
                                Text(
                                    text = c.vehiculo,
                                    fontSize = 13.sp,
                                    color = MaterialTheme.colorScheme.outline
                                )
                            }
                        }
                    }
                }

                // Estructura Financiera
                item {
                    Card(
                        modifier = Modifier.fillMaxWidth(),
                        shape = RoundedCornerShape(14.dp),
                        colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
                        elevation = CardDefaults.cardElevation(defaultElevation = 2.dp)
                    ) {
                        Column(modifier = Modifier.padding(16.dp)) {
                            Text(
                                text = "Estructura Financiera (sin IVA)",
                                fontWeight = FontWeight.Bold,
                                fontSize = 15.sp,
                                color = SlateSecondary
                            )
                            HorizontalDivider(modifier = Modifier.padding(vertical = 8.dp), color = MaterialTheme.colorScheme.surfaceVariant)
                            InfoRow(label = "Valor del Activo", value = formatCurrency(c.valorSinIva), isBold = true)
                            InfoRow(label = "Anticipo Pagado", value = "${formatCurrency(c.anticipoMonto)} (${formatPercent(c.anticipoPct)})")
                            InfoRow(label = "Inversión Neta Financiada", value = formatCurrency(c.inversionNeta), isBold = true)
                            InfoRow(label = "Valor Residual Pactado", value = "${formatCurrency(c.residualMonto)} (${formatPercent(c.residualPct)})")
                            InfoRow(label = "Valor Presente Residual (VP)", value = formatCurrency(c.vpResidual))
                            if (c.comisionMonto > 0) {
                                InfoRow(label = "Comisión por Apertura", value = "${formatCurrency(c.comisionMonto)} (${formatPercent(c.comisionAperturaPct)})")
                            }
                        }
                    }
                }

                // Condiciones y Tasas
                item {
                    Card(
                        modifier = Modifier.fillMaxWidth(),
                        shape = RoundedCornerShape(14.dp),
                        colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
                        elevation = CardDefaults.cardElevation(defaultElevation = 2.dp)
                    ) {
                        Column(modifier = Modifier.padding(16.dp)) {
                            Text(
                                text = "Condiciones de Renta y Tasa",
                                fontWeight = FontWeight.Bold,
                                fontSize = 15.sp,
                                color = SlateSecondary
                            )
                            HorizontalDivider(modifier = Modifier.padding(vertical = 8.dp), color = MaterialTheme.colorScheme.surfaceVariant)
                            InfoRow(label = "Renta Mensual sin IVA", value = formatCurrency(c.mensualidadSinIva), isBold = true)
                            InfoRow(label = "Plazo del Arrendamiento", value = "${c.plazo} meses")
                            InfoRow(label = "Tasa Mensual Implícita", value = formatPercent(c.tasaMensualPct))
                            InfoRow(label = "Tasa Anual Implícita", value = formatPercent(c.tasaAnualPct), isBold = true)
                        }
                    }
                }

                // Balances y Fechas
                item {
                    Card(
                        modifier = Modifier.fillMaxWidth(),
                        shape = RoundedCornerShape(14.dp),
                        colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
                        elevation = CardDefaults.cardElevation(defaultElevation = 2.dp)
                    ) {
                        Column(modifier = Modifier.padding(16.dp)) {
                            Text(
                                text = "Cronología y Saldos Proyectados",
                                fontWeight = FontWeight.Bold,
                                fontSize = 15.sp,
                                color = SlateSecondary
                            )
                            HorizontalDivider(modifier = Modifier.padding(vertical = 8.dp), color = MaterialTheme.colorScheme.surfaceVariant)
                            InfoRow(label = "Fecha de Alta", value = c.fechaAlta.ifEmpty { "N/D" })
                            InfoRow(label = "Fecha de Vencimiento", value = c.fechaVencimiento.ifEmpty { "N/D" })
                            if (!c.fechaBaja.isNullOrBlank()) {
                                InfoRow(label = "Fecha de Baja", value = c.fechaBaja)
                            }
                            InfoRow(label = "Saldo Actual Insoluto Est.", value = formatCurrency(c.saldoActualEstimado), isBold = true)
                            InfoRow(label = "Intereses Totales Proyectados", value = formatCurrency(c.interesesTotalesEstimados))
                        }
                    }
                }

                // Anotaciones
                if (c.anotaciones.isNotBlank()) {
                    item {
                        Card(
                            modifier = Modifier.fillMaxWidth(),
                            shape = RoundedCornerShape(14.dp),
                            colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
                            elevation = CardDefaults.cardElevation(defaultElevation = 2.dp)
                        ) {
                            Column(modifier = Modifier.padding(16.dp)) {
                                Text(
                                    text = "Observaciones",
                                    fontWeight = FontWeight.Bold,
                                    fontSize = 15.sp,
                                    color = SlateSecondary
                                )
                                Spacer(modifier = Modifier.height(6.dp))
                                Text(
                                    text = c.anotaciones,
                                    fontSize = 13.sp,
                                    color = MaterialTheme.colorScheme.onSurface
                                )
                            }
                        }
                    }
                }

                item { Spacer(modifier = Modifier.height(24.dp)) }
            }
        }
    }
}
