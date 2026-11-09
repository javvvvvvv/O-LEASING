package com.oleasing.app.ui.screens

import androidx.compose.foundation.background
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.LazyRow
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.foundation.text.KeyboardOptions
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.filled.ArrowBack
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.input.KeyboardType
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import com.oleasing.app.model.CotizacionDraft
import com.oleasing.app.model.CotizacionRequest
import com.oleasing.app.model.CotizacionResponse
import com.oleasing.app.network.OLeasingApiClient
import com.oleasing.app.ui.components.InfoRow
import com.oleasing.app.ui.components.formatCurrency
import com.oleasing.app.ui.components.formatPercent
import com.oleasing.app.ui.theme.OrangePrimary
import com.oleasing.app.ui.theme.SlateSecondary
import kotlinx.coroutines.launch

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun CotizadorScreen(
    onBack: () -> Unit,
    onNavigateToNewContrato: (CotizacionDraft) -> Unit
) {
    val coroutineScope = rememberCoroutineScope()
    var valorSinIva by remember { mutableStateOf("450000") }
    var plazo by remember { mutableStateOf("36") }
    var anticipoMonto by remember { mutableStateOf("45000") }
    var residualMonto by remember { mutableStateOf("90000") }
    var comisionPct by remember { mutableStateOf("2.0") }
    var modoCalculo by remember { mutableStateOf("TASA_A_RENTA") } // TASA_A_RENTA o RENTA_A_TASA
    var tasaAnualInput by remember { mutableStateOf("28.0") }
    var rentaSinIvaInput by remember { mutableStateOf("16500") }

    var resultado by remember { mutableStateOf<CotizacionResponse?>(null) }
    var isLoading by remember { mutableStateOf(false) }
    var errorMessage by remember { mutableStateOf<String?>(null) }

    val plazosRapidos = listOf(12, 24, 36, 48, 60)

    fun calcularCotizacion() {
        val v = valorSinIva.toDoubleOrNull() ?: 0.0
        val p = plazo.toIntOrNull() ?: 0
        val ant = anticipoMonto.toDoubleOrNull() ?: 0.0
        val res = residualMonto.toDoubleOrNull() ?: 0.0
        val com = comisionPct.toDoubleOrNull() ?: 0.0

        if (v <= 0 || p <= 0) {
            errorMessage = "Ingresa un valor del activo y un plazo mayores a cero."
            return
        }

        val req = if (modoCalculo == "TASA_A_RENTA") {
            val t = tasaAnualInput.toDoubleOrNull() ?: 0.0
            CotizacionRequest(
                valorSinIva = v,
                plazo = p,
                tasaAnualPct = t,
                anticipoMonto = ant,
                residualMonto = res,
                comisionPct = com
            )
        } else {
            val r = rentaSinIvaInput.toDoubleOrNull() ?: 0.0
            CotizacionRequest(
                valorSinIva = v,
                plazo = p,
                mensualidadSinIva = r,
                anticipoMonto = ant,
                residualMonto = res,
                comisionPct = com
            )
        }

        coroutineScope.launch {
            isLoading = true
            errorMessage = null
            val resp = OLeasingApiClient.cotizar(req)
            if (resp.isSuccess) {
                resultado = resp.getOrNull()
            } else {
                errorMessage = resp.exceptionOrNull()?.message
            }
            isLoading = false
        }
    }

    LaunchedEffect(Unit) {
        calcularCotizacion()
    }

    Scaffold(
        topBar = {
            TopAppBar(
                title = {
                    Column {
                        Text(
                            text = "Cotizador Financiero",
                            fontWeight = FontWeight.Bold,
                            fontSize = 18.sp,
                            color = OrangePrimary
                        )
                        Text(
                            text = "Simulador de Arrendamiento O-Leasing",
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
                colors = TopAppBarDefaults.topAppBarColors(containerColor = MaterialTheme.colorScheme.surface)
            )
        }
    ) { padding ->
        LazyColumn(
            modifier = Modifier
                .fillMaxSize()
                .padding(padding)
                .background(MaterialTheme.colorScheme.background)
                .padding(horizontal = 16.dp),
            verticalArrangement = Arrangement.spacedBy(12.dp)
        ) {
            item { Spacer(modifier = Modifier.height(4.dp)) }

            item {
                Card(
                    modifier = Modifier.fillMaxWidth(),
                    shape = RoundedCornerShape(14.dp),
                    colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
                    elevation = CardDefaults.cardElevation(defaultElevation = 2.dp)
                ) {
                    Column(modifier = Modifier.padding(16.dp), verticalArrangement = Arrangement.spacedBy(10.dp)) {
                        Text(
                            text = "Modo de Simulación",
                            fontWeight = FontWeight.Bold,
                            fontSize = 14.sp,
                            color = SlateSecondary
                        )
                        Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                            FilterChip(
                                selected = modoCalculo == "TASA_A_RENTA",
                                onClick = { modoCalculo = "TASA_A_RENTA" },
                                label = { Text("Calcular Renta (desde Tasa)", fontSize = 12.sp) },
                                colors = FilterChipDefaults.filterChipColors(
                                    selectedContainerColor = OrangePrimary,
                                    selectedLabelColor = Color.White
                                )
                            )
                            FilterChip(
                                selected = modoCalculo == "RENTA_A_TASA",
                                onClick = { modoCalculo = "RENTA_A_TASA" },
                                label = { Text("Calcular Tasa (desde Renta)", fontSize = 12.sp) },
                                colors = FilterChipDefaults.filterChipColors(
                                    selectedContainerColor = OrangePrimary,
                                    selectedLabelColor = Color.White
                                )
                            )
                        }

                        OutlinedTextField(
                            value = valorSinIva,
                            onValueChange = { valorSinIva = it },
                            label = { Text("Valor del Activo sin IVA (MXN)") },
                            keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                            singleLine = true,
                            modifier = Modifier.fillMaxWidth()
                        )

                        Text("Plazo (meses)", fontSize = 12.sp, color = MaterialTheme.colorScheme.outline)
                        LazyRow(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                            items(plazosRapidos) { p ->
                                val sel = plazo == p.toString()
                                FilterChip(
                                    selected = sel,
                                    onClick = { plazo = p.toString() },
                                    label = { Text("$p m", fontWeight = if (sel) FontWeight.Bold else FontWeight.Normal) },
                                    colors = FilterChipDefaults.filterChipColors(
                                        selectedContainerColor = SlateSecondary,
                                        selectedLabelColor = Color.White
                                    )
                                )
                            }
                        }

                        Row(horizontalArrangement = Arrangement.spacedBy(10.dp)) {
                            OutlinedTextField(
                                value = anticipoMonto,
                                onValueChange = { anticipoMonto = it },
                                label = { Text("Anticipo s/IVA") },
                                keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                                singleLine = true,
                                modifier = Modifier.weight(1f)
                            )
                            OutlinedTextField(
                                value = residualMonto,
                                onValueChange = { residualMonto = it },
                                label = { Text("Residual s/IVA") },
                                keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                                singleLine = true,
                                modifier = Modifier.weight(1f)
                            )
                        }

                        Row(horizontalArrangement = Arrangement.spacedBy(10.dp)) {
                            if (modoCalculo == "TASA_A_RENTA") {
                                OutlinedTextField(
                                    value = tasaAnualInput,
                                    onValueChange = { tasaAnualInput = it },
                                    label = { Text("Tasa Anual Objetivo (%)") },
                                    keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                                    singleLine = true,
                                    modifier = Modifier.weight(1f)
                                )
                            } else {
                                OutlinedTextField(
                                    value = rentaSinIvaInput,
                                    onValueChange = { rentaSinIvaInput = it },
                                    label = { Text("Renta Mensual s/IVA") },
                                    keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                                    singleLine = true,
                                    modifier = Modifier.weight(1f)
                                )
                            }
                            OutlinedTextField(
                                value = comisionPct,
                                onValueChange = { comisionPct = it },
                                label = { Text("Comisión (%)") },
                                keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                                singleLine = true,
                                modifier = Modifier.weight(0.8f)
                            )
                        }

                        Button(
                            onClick = { calcularCotizacion() },
                            modifier = Modifier.fillMaxWidth().height(48.dp),
                            shape = RoundedCornerShape(10.dp),
                            colors = ButtonDefaults.buttonColors(containerColor = OrangePrimary)
                        ) {
                            if (isLoading) {
                                CircularProgressIndicator(color = Color.White, modifier = Modifier.size(20.dp), strokeWidth = 2.dp)
                            } else {
                                Text("Calcular Cotización", fontWeight = FontWeight.Bold, color = Color.White)
                            }
                        }

                        if (errorMessage != null) {
                            Text(errorMessage ?: "", color = Color(0xFFDC2626), fontSize = 12.sp)
                        }
                    }
                }
            }

            resultado?.let { res ->
                val ivaRenta = res.mensualidadSinIva * 0.16
                val rentaConIva = res.mensualidadSinIva + ivaRenta

                item {
                    Card(
                        modifier = Modifier.fillMaxWidth(),
                        shape = RoundedCornerShape(14.dp),
                        colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
                        elevation = CardDefaults.cardElevation(defaultElevation = 3.dp)
                    ) {
                        Column(modifier = Modifier.padding(16.dp)) {
                            Text(
                                text = "Resultado de la Cotización",
                                fontWeight = FontWeight.Bold,
                                fontSize = 15.sp,
                                color = SlateSecondary
                            )
                            Spacer(modifier = Modifier.height(10.dp))
                            Row(
                                modifier = Modifier
                                    .fillMaxWidth()
                                    .background(Color(0xFFFFF7ED), RoundedCornerShape(10.dp))
                                    .padding(14.dp),
                                horizontalArrangement = Arrangement.SpaceBetween,
                                verticalAlignment = Alignment.CenterVertically
                            ) {
                                Column {
                                    Text("Renta Mensual s/IVA", fontSize = 12.sp, color = MaterialTheme.colorScheme.outline)
                                    Text(
                                        text = formatCurrency(res.mensualidadSinIva),
                                        fontWeight = FontWeight.ExtraBold,
                                        fontSize = 20.sp,
                                        color = OrangePrimary
                                    )
                                    Text(
                                        text = "Con IVA (16%): ${formatCurrency(rentaConIva)}",
                                        fontSize = 12.sp,
                                        fontWeight = FontWeight.SemiBold,
                                        color = SlateSecondary
                                    )
                                }
                                Column(horizontalAlignment = Alignment.End) {
                                    Text("Tasa Anual Impl.", fontSize = 12.sp, color = MaterialTheme.colorScheme.outline)
                                    Text(
                                        text = formatPercent(res.tasaAnualPct),
                                        fontWeight = FontWeight.ExtraBold,
                                        fontSize = 18.sp,
                                        color = SlateSecondary
                                    )
                                    Text(
                                        text = "Mensual: ${formatPercent(res.tasaMensualPct)}",
                                        fontSize = 11.sp,
                                        color = MaterialTheme.colorScheme.outline
                                    )
                                }
                            }

                            Spacer(modifier = Modifier.height(10.dp))
                            InfoRow(label = "Valor del Activo s/IVA", value = formatCurrency(res.valorSinIva))
                            InfoRow(label = "Anticipo s/IVA", value = formatCurrency(res.anticipoMonto))
                            InfoRow(label = "Inversión Neta Financiada", value = formatCurrency(res.inversionNeta), isBold = true)
                            InfoRow(label = "Plazo", value = "${res.plazo} meses")
                            InfoRow(label = "Valor Residual Pactado", value = formatCurrency(res.residualMonto))
                            InfoRow(label = "Valor Presente del Residual", value = formatCurrency(res.vpResidual))
                            InfoRow(label = "Comisión por Apertura", value = formatCurrency(res.comisionMonto))
                            InfoRow(label = "Total Intereses en el Plazo", value = formatCurrency(res.totalIntereses), isBold = true)
                            InfoRow(label = "Total Rentas Acumuladas", value = formatCurrency(res.totalRentas))
                            InfoRow(label = "Costo Total del Arrendamiento", value = formatCurrency(res.costoTotalLeasing), isBold = true)

                            Spacer(modifier = Modifier.height(12.dp))
                            if (OLeasingApiClient.currentUser?.puedeEditar != false) {
                                Button(
                                    onClick = {
                                        onNavigateToNewContrato(
                                            CotizacionDraft(
                                                valorSinIva = res.valorSinIva,
                                                plazo = res.plazo,
                                                anticipoMonto = res.anticipoMonto,
                                                mensualidadSinIva = res.mensualidadSinIva,
                                                residualMonto = res.residualMonto,
                                                comisionPct = comisionPct.toDoubleOrNull() ?: 0.0
                                            )
                                        )
                                    },
                                    modifier = Modifier.fillMaxWidth().height(48.dp),
                                    shape = RoundedCornerShape(10.dp),
                                    colors = ButtonDefaults.buttonColors(containerColor = SlateSecondary)
                                ) {
                                    Text("Dar de Alta Contrato con esta Cotización", fontWeight = FontWeight.Bold, color = Color.White)
                                }
                            }
                        }
                    }
                }
            }

            item { Spacer(modifier = Modifier.height(32.dp)) }
        }
    }
}
