package com.oleasing.app.ui.screens

import androidx.compose.foundation.background
import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.LazyRow
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.filled.ArrowBack
import androidx.compose.material.icons.filled.Refresh
import androidx.compose.material.icons.filled.Search
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import com.oleasing.app.model.CobranzaMesResponse
import com.oleasing.app.network.OLeasingApiClient
import com.oleasing.app.ui.components.formatCurrency
import com.oleasing.app.ui.theme.OrangePrimary
import com.oleasing.app.ui.theme.SlateSecondary
import kotlinx.coroutines.launch
import java.util.Calendar

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun CobranzaMesScreen(
    onBack: () -> Unit,
    onNavigateToEstadoCuenta: (String) -> Unit
) {
    val coroutineScope = rememberCoroutineScope()
    val cal = remember { Calendar.getInstance() }
    var anioSel by remember { mutableStateOf(cal.get(Calendar.YEAR)) }
    var mesSel by remember { mutableStateOf(cal.get(Calendar.MONTH) + 1) }
    var filtroTexto by remember { mutableStateOf("") }

    var cobranza by remember { mutableStateOf<CobranzaMesResponse?>(null) }
    var isLoading by remember { mutableStateOf(false) }
    var errorMessage by remember { mutableStateOf<String?>(null) }

    val nombresMeses = listOf(
        "Ene", "Feb", "Mar", "Abr", "May", "Jun",
        "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"
    )
    val aniosDisponibles = listOf(2025, 2026, 2027, 2028)

    fun cargarCobranza() {
        coroutineScope.launch {
            isLoading = true
            errorMessage = null
            val res = OLeasingApiClient.getCobranzaMes(anioSel, mesSel)
            if (res.isSuccess) {
                cobranza = res.getOrNull()
            } else {
                errorMessage = res.exceptionOrNull()?.message
            }
            isLoading = false
        }
    }

    LaunchedEffect(anioSel, mesSel) {
        cargarCobranza()
    }

    Scaffold(
        topBar = {
            TopAppBar(
                title = {
                    Column {
                        Text(
                            text = "Cobranza y Devengo del Mes",
                            fontWeight = FontWeight.Bold,
                            fontSize = 17.sp,
                            color = OrangePrimary
                        )
                        Text(
                            text = "Periodo ${nombresMeses[mesSel - 1]} $anioSel",
                            fontSize = 12.sp,
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
                    IconButton(onClick = { cargarCobranza() }) {
                        Icon(imageVector = Icons.Default.Refresh, contentDescription = "Actualizar", tint = OrangePrimary)
                    }
                },
                colors = TopAppBarDefaults.topAppBarColors(containerColor = MaterialTheme.colorScheme.surface)
            )
        }
    ) { padding ->
        Column(
            modifier = Modifier
                .fillMaxSize()
                .padding(padding)
                .background(MaterialTheme.colorScheme.background)
        ) {
            // Selector de Año y Mes
            LazyRow(
                modifier = Modifier
                    .fillMaxWidth()
                    .padding(horizontal = 16.dp, vertical = 4.dp),
                horizontalArrangement = Arrangement.spacedBy(6.dp)
            ) {
                items(aniosDisponibles) { y ->
                    FilterChip(
                        selected = anioSel == y,
                        onClick = { anioSel = y },
                        label = { Text("$y", fontSize = 12.sp, fontWeight = FontWeight.Bold) },
                        colors = FilterChipDefaults.filterChipColors(
                            selectedContainerColor = SlateSecondary,
                            selectedLabelColor = Color.White
                        )
                    )
                }
            }

            LazyRow(
                modifier = Modifier
                    .fillMaxWidth()
                    .padding(horizontal = 16.dp, vertical = 2.dp),
                horizontalArrangement = Arrangement.spacedBy(6.dp)
            ) {
                items(12) { idx ->
                    val m = idx + 1
                    FilterChip(
                        selected = mesSel == m,
                        onClick = { mesSel = m },
                        label = { Text(nombresMeses[idx], fontSize = 12.sp) },
                        colors = FilterChipDefaults.filterChipColors(
                            selectedContainerColor = OrangePrimary,
                            selectedLabelColor = Color.White
                        )
                    )
                }
            }

            if (isLoading && cobranza == null) {
                Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                    CircularProgressIndicator(color = OrangePrimary)
                }
            } else if (errorMessage != null && cobranza == null) {
                Box(modifier = Modifier.fillMaxSize().padding(20.dp), contentAlignment = Alignment.Center) {
                    Column(horizontalAlignment = Alignment.CenterHorizontally) {
                        Text("Error al consultar cobranza", fontWeight = FontWeight.Bold, color = Color(0xFFDC2626))
                        Spacer(modifier = Modifier.height(6.dp))
                        Text(errorMessage ?: "", fontSize = 12.sp, color = MaterialTheme.colorScheme.outline)
                        Spacer(modifier = Modifier.height(12.dp))
                        Button(onClick = { cargarCobranza() }, colors = ButtonDefaults.buttonColors(containerColor = OrangePrimary)) {
                            Text("Reintentar")
                        }
                    }
                }
            } else if (cobranza != null) {
                val data = cobranza!!
                val filtrados = data.items.filter {
                    filtroTexto.isBlank() ||
                        it.idContrato.contains(filtroTexto, ignoreCase = true) ||
                        it.cliente.contains(filtroTexto, ignoreCase = true) ||
                        it.vehiculo.contains(filtroTexto, ignoreCase = true)
                }

                LazyColumn(
                    modifier = Modifier
                        .fillMaxSize()
                        .padding(horizontal = 16.dp),
                    verticalArrangement = Arrangement.spacedBy(10.dp)
                ) {
                    item {
                        Card(
                            modifier = Modifier.fillMaxWidth(),
                            shape = RoundedCornerShape(14.dp),
                            colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
                            elevation = CardDefaults.cardElevation(defaultElevation = 2.dp)
                        ) {
                            Column(modifier = Modifier.padding(14.dp)) {
                                Row(
                                    modifier = Modifier.fillMaxWidth(),
                                    horizontalArrangement = Arrangement.SpaceBetween
                                ) {
                                    Column {
                                        Text("Cobranza Total c/IVA", fontSize = 11.sp, color = MaterialTheme.colorScheme.outline)
                                        Text(
                                            text = formatCurrency(data.totalCobranzaConIva),
                                            fontWeight = FontWeight.ExtraBold,
                                            fontSize = 19.sp,
                                            color = OrangePrimary
                                        )
                                        Text(
                                            text = "Sin IVA: ${formatCurrency(data.rentaTotalSinIva)}",
                                            fontSize = 11.sp,
                                            color = SlateSecondary
                                        )
                                    }
                                    Column(horizontalAlignment = Alignment.End) {
                                        Text("Contratos Activos", fontSize = 11.sp, color = MaterialTheme.colorScheme.outline)
                                        Text(
                                            text = "${data.contratosEnCobro}",
                                            fontWeight = FontWeight.ExtraBold,
                                            fontSize = 19.sp,
                                            color = SlateSecondary
                                        )
                                        Text(
                                            text = "IVA: ${formatCurrency(data.ivaTotal)}",
                                            fontSize = 11.sp,
                                            color = MaterialTheme.colorScheme.outline
                                        )
                                    }
                                }
                                HorizontalDivider(modifier = Modifier.padding(vertical = 8.dp))
                                Row(
                                    modifier = Modifier.fillMaxWidth(),
                                    horizontalArrangement = Arrangement.SpaceBetween
                                ) {
                                    Column {
                                        Text("Capital del Mes", fontSize = 11.sp, color = MaterialTheme.colorScheme.outline)
                                        Text(formatCurrency(data.capitalTotalMes), fontWeight = FontWeight.Bold, fontSize = 13.sp)
                                    }
                                    Column(horizontalAlignment = Alignment.End) {
                                        Text("Interés Leasing del Mes", fontSize = 11.sp, color = MaterialTheme.colorScheme.outline)
                                        Text(formatCurrency(data.interesTotalMes), fontWeight = FontWeight.Bold, fontSize = 13.sp, color = SlateSecondary)
                                    }
                                }
                            }
                        }
                    }

                    item {
                        OutlinedTextField(
                            value = filtroTexto,
                            onValueChange = { filtroTexto = it },
                            placeholder = { Text("Filtrar contrato o cliente en este mes...", fontSize = 13.sp) },
                            leadingIcon = { Icon(Icons.Default.Search, contentDescription = null, tint = OrangePrimary) },
                            singleLine = true,
                            shape = RoundedCornerShape(10.dp),
                            modifier = Modifier.fillMaxWidth()
                        )
                    }

                    items(filtrados) { item ->
                        Card(
                            modifier = Modifier
                                .fillMaxWidth()
                                .clickable { onNavigateToEstadoCuenta(item.idContrato) },
                            shape = RoundedCornerShape(12.dp),
                            colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
                            elevation = CardDefaults.cardElevation(defaultElevation = 1.dp)
                        ) {
                            Column(modifier = Modifier.padding(12.dp)) {
                                Row(
                                    modifier = Modifier.fillMaxWidth(),
                                    horizontalArrangement = Arrangement.SpaceBetween,
                                    verticalAlignment = Alignment.CenterVertically
                                ) {
                                    Text(
                                        text = item.idContrato,
                                        fontWeight = FontWeight.Bold,
                                        fontSize = 14.sp,
                                        color = OrangePrimary
                                    )
                                    Surface(
                                        color = Color(0xFFFFF7ED),
                                        shape = RoundedCornerShape(6.dp)
                                    ) {
                                        Text(
                                            text = "Pago ${item.mesContrato} de ${item.plazo}",
                                            color = OrangePrimary,
                                            fontWeight = FontWeight.Bold,
                                            fontSize = 11.sp,
                                            modifier = Modifier.padding(horizontal = 8.dp, vertical = 3.dp)
                                        )
                                    }
                                }
                                Text(item.cliente, fontWeight = FontWeight.SemiBold, fontSize = 13.sp)
                                if (item.vehiculo.isNotBlank()) {
                                    Text(item.vehiculo, fontSize = 11.sp, color = MaterialTheme.colorScheme.outline)
                                }
                                Spacer(modifier = Modifier.height(6.dp))
                                Row(
                                    modifier = Modifier.fillMaxWidth(),
                                    horizontalArrangement = Arrangement.SpaceBetween
                                ) {
                                    Column {
                                        Text("Renta s/IVA", fontSize = 10.sp, color = MaterialTheme.colorScheme.outline)
                                        Text(formatCurrency(item.rentaSinIva), fontWeight = FontWeight.SemiBold, fontSize = 12.sp)
                                    }
                                    Column {
                                        Text("Capital", fontSize = 10.sp, color = MaterialTheme.colorScheme.outline)
                                        Text(formatCurrency(item.capitalMes), fontWeight = FontWeight.SemiBold, fontSize = 12.sp)
                                    }
                                    Column {
                                        Text("Interés", fontSize = 10.sp, color = MaterialTheme.colorScheme.outline)
                                        Text(formatCurrency(item.interesLeasingMes), fontWeight = FontWeight.SemiBold, fontSize = 12.sp)
                                    }
                                    Column(horizontalAlignment = Alignment.End) {
                                        Text("Total c/IVA", fontSize = 10.sp, color = MaterialTheme.colorScheme.outline)
                                        Text(formatCurrency(item.totalConIva), fontWeight = FontWeight.Bold, fontSize = 12.sp, color = OrangePrimary)
                                    }
                                }
                            }
                        }
                    }

                    item { Spacer(modifier = Modifier.height(24.dp)) }
                }
            }
        }
    }
}
