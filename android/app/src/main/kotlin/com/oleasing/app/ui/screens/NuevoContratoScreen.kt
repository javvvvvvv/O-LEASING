package com.oleasing.app.ui.screens

import androidx.compose.foundation.background
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.foundation.text.KeyboardOptions
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.filled.ArrowBack
import androidx.compose.material.icons.filled.Check
import androidx.compose.material.icons.filled.Info
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.input.KeyboardType
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import com.oleasing.app.model.ContratoCreate
import com.oleasing.app.model.CotizacionRequest
import com.oleasing.app.model.CotizacionResponse
import com.oleasing.app.network.OLeasingApiClient
import com.oleasing.app.ui.components.InfoRow
import com.oleasing.app.ui.components.formatCurrency
import com.oleasing.app.ui.components.formatPercent
import com.oleasing.app.ui.theme.OrangePrimary
import com.oleasing.app.ui.theme.SlateSecondary
import kotlinx.coroutines.launch
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun NuevoContratoScreen(
    onBack: () -> Unit,
    onContratoCreated: (String) -> Unit
) {
    val coroutineScope = rememberCoroutineScope()
    val todayStr = remember { SimpleDateFormat("yyyy-MM-dd", Locale.US).format(Date()) }

    var idContrato by remember { mutableStateOf("") }
    var cliente by remember { mutableStateOf("") }
    var vehiculo by remember { mutableStateOf("") }
    var fechaAlta by remember { mutableStateOf(todayStr) }
    var plazoStr by remember { mutableStateOf("36") }
    var valorActivoStr by remember { mutableStateOf("") }
    var anticipoStr by remember { mutableStateOf("0") }
    var mensualidadStr by remember { mutableStateOf("") }
    var residualStr by remember { mutableStateOf("0") }
    var comisionStr by remember { mutableStateOf("0") }
    var anotaciones by remember { mutableStateOf("") }

    var isSimulating by remember { mutableStateOf(false) }
    var isSaving by remember { mutableStateOf(false) }
    var cotizacionResult by remember { mutableStateOf<CotizacionResponse?>(null) }
    var errorMessage by remember { mutableStateOf<String?>(null) }
    var successDialogMessage by remember { mutableStateOf<String?>(null) }
    var createdId by remember { mutableStateOf("") }

    fun simularCalculo() {
        val valor = valorActivoStr.toDoubleOrNull() ?: 0.0
        val plazo = plazoStr.toIntOrNull() ?: 0
        val renta = mensualidadStr.toDoubleOrNull() ?: 0.0
        val anticipo = anticipoStr.toDoubleOrNull() ?: 0.0
        val residual = residualStr.toDoubleOrNull() ?: 0.0
        val comision = comisionStr.toDoubleOrNull() ?: 0.0

        if (valor <= 0) {
            errorMessage = "Ingresa un valor del activo mayor a 0 para simular."
            return
        }
        if (plazo <= 0) {
            errorMessage = "Ingresa un plazo en meses válido (ej. 12, 24, 36, 48)."
            return
        }
        if (renta <= 0) {
            errorMessage = "Ingresa una renta mensual mayor a 0 para calcular la tasa implícita."
            return
        }

        coroutineScope.launch {
            isSimulating = true
            errorMessage = null
            val req = CotizacionRequest(
                valorSinIva = valor,
                plazo = plazo,
                mensualidadSinIva = renta,
                anticipoMonto = anticipo,
                residualMonto = residual,
                comisionPct = comision
            )
            val result = OLeasingApiClient.cotizar(req)
            if (result.isSuccess) {
                cotizacionResult = result.getOrNull()
            } else {
                errorMessage = result.exceptionOrNull()?.message
            }
            isSimulating = false
        }
    }

    fun guardarContrato() {
        val id = idContrato.trim()
        val cli = cliente.trim()
        val valor = valorActivoStr.toDoubleOrNull() ?: 0.0
        val plazo = plazoStr.toIntOrNull() ?: 0
        val renta = mensualidadStr.toDoubleOrNull() ?: 0.0
        val anticipo = anticipoStr.toDoubleOrNull() ?: 0.0
        val residual = residualStr.toDoubleOrNull() ?: 0.0
        val comision = comisionStr.toDoubleOrNull() ?: 0.0

        if (id.isEmpty()) {
            errorMessage = "Ingresa el ID del contrato (ej. 0729-0001)."
            return
        }
        if (cli.isEmpty()) {
            errorMessage = "Ingresa el nombre o razón social del cliente."
            return
        }
        if (valor <= 0 || renta <= 0 || plazo <= 0) {
            errorMessage = "Verifica que el Valor del Activo, Plazo y Renta Mensual sean mayores a cero."
            return
        }

        coroutineScope.launch {
            isSaving = true
            errorMessage = null
            val req = ContratoCreate(
                idContrato = id,
                cliente = cli,
                vehiculo = vehiculo.trim(),
                fechaAlta = fechaAlta.trim(),
                plazo = plazo,
                valorSinIva = valor,
                anticipoMonto = anticipo,
                mensualidadSinIva = renta,
                residualMonto = residual,
                comisionPct = comision,
                anotaciones = anotaciones.trim()
            )
            val result = OLeasingApiClient.crearContrato(req)
            if (result.isSuccess) {
                val resp = result.getOrNull()!!
                createdId = resp.idContrato
                successDialogMessage = "Contrato ${resp.idContrato} registrado exitosamente.\n\n" +
                        "Cliente: ${resp.cliente}\n" +
                        "Tasa Anual: ${formatPercent(resp.tasaAnualPct)}\n" +
                        "Inversión Neta: ${formatCurrency(resp.inversionNeta)}\n" +
                        "Vencimiento: ${resp.fechaVencimiento}"
            } else {
                errorMessage = result.exceptionOrNull()?.message
            }
            isSaving = false
        }
    }

    Scaffold(
        topBar = {
            TopAppBar(
                title = { Text("Nuevo Contrato Individual", fontWeight = FontWeight.Bold, fontSize = 18.sp) },
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
                .padding(16.dp),
            verticalArrangement = Arrangement.spacedBy(14.dp)
        ) {
            if (errorMessage != null) {
                item {
                    Card(
                        colors = CardDefaults.cardColors(containerColor = Color(0xFFFEE2E2)),
                        shape = RoundedCornerShape(10.dp),
                        modifier = Modifier.fillMaxWidth()
                    ) {
                        Column(modifier = Modifier.padding(12.dp)) {
                            Text("Atención:", fontWeight = FontWeight.Bold, color = Color(0xFFB91C1C), fontSize = 13.sp)
                            Spacer(modifier = Modifier.height(2.dp))
                            Text(errorMessage ?: "", color = Color(0xFF7F1D1D), fontSize = 12.sp)
                        }
                    }
                }
            }

            // Datos Generales
            item {
                Card(
                    modifier = Modifier.fillMaxWidth(),
                    shape = RoundedCornerShape(12.dp),
                    colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
                    elevation = CardDefaults.cardElevation(defaultElevation = 2.dp)
                ) {
                    Column(modifier = Modifier.padding(16.dp), verticalArrangement = Arrangement.spacedBy(10.dp)) {
                        Text("1. Identificación del Contrato", fontWeight = FontWeight.Bold, fontSize = 15.sp, color = SlateSecondary)

                        OutlinedTextField(
                            value = idContrato,
                            onValueChange = { idContrato = it.uppercase() },
                            label = { Text("ID Contrato (ej. 0729-0001)") },
                            singleLine = true,
                            modifier = Modifier.fillMaxWidth()
                        )

                        OutlinedTextField(
                            value = cliente,
                            onValueChange = { cliente = it },
                            label = { Text("Cliente / Razón Social") },
                            singleLine = true,
                            modifier = Modifier.fillMaxWidth()
                        )

                        OutlinedTextField(
                            value = vehiculo,
                            onValueChange = { vehiculo = it },
                            label = { Text("Vehículo / Equipo / Activo") },
                            singleLine = true,
                            modifier = Modifier.fillMaxWidth()
                        )

                        OutlinedTextField(
                            value = fechaAlta,
                            onValueChange = { fechaAlta = it },
                            label = { Text("Fecha de Alta (YYYY-MM-DD)") },
                            singleLine = true,
                            modifier = Modifier.fillMaxWidth()
                        )
                    }
                }
            }

            // Parámetros Financieros
            item {
                Card(
                    modifier = Modifier.fillMaxWidth(),
                    shape = RoundedCornerShape(12.dp),
                    colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
                    elevation = CardDefaults.cardElevation(defaultElevation = 2.dp)
                ) {
                    Column(modifier = Modifier.padding(16.dp), verticalArrangement = Arrangement.spacedBy(10.dp)) {
                        Text("2. Parámetros Financieros (sin IVA)", fontWeight = FontWeight.Bold, fontSize = 15.sp, color = SlateSecondary)

                        Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(10.dp)) {
                            OutlinedTextField(
                                value = valorActivoStr,
                                onValueChange = { valorActivoStr = it },
                                label = { Text("Valor Activo ($)") },
                                keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Decimal),
                                singleLine = true,
                                modifier = Modifier.weight(1f)
                            )
                            OutlinedTextField(
                                value = plazoStr,
                                onValueChange = { plazoStr = it },
                                label = { Text("Plazo (meses)") },
                                keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                                singleLine = true,
                                modifier = Modifier.weight(1f)
                            )
                        }

                        Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(10.dp)) {
                            OutlinedTextField(
                                value = anticipoStr,
                                onValueChange = { anticipoStr = it },
                                label = { Text("Anticipo ($)") },
                                keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Decimal),
                                singleLine = true,
                                modifier = Modifier.weight(1f)
                            )
                            OutlinedTextField(
                                value = residualStr,
                                onValueChange = { residualStr = it },
                                label = { Text("Residual ($)") },
                                keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Decimal),
                                singleLine = true,
                                modifier = Modifier.weight(1f)
                            )
                        }

                        Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(10.dp)) {
                            OutlinedTextField(
                                value = mensualidadStr,
                                onValueChange = { mensualidadStr = it },
                                label = { Text("Renta Mensual ($)") },
                                keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Decimal),
                                singleLine = true,
                                modifier = Modifier.weight(1f)
                            )
                            OutlinedTextField(
                                value = comisionStr,
                                onValueChange = { comisionStr = it },
                                label = { Text("Comisión (%)") },
                                keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Decimal),
                                singleLine = true,
                                modifier = Modifier.weight(1f)
                            )
                        }

                        OutlinedTextField(
                            value = anotaciones,
                            onValueChange = { anotaciones = it },
                            label = { Text("Observaciones / Notas") },
                            maxLines = 3,
                            modifier = Modifier.fillMaxWidth()
                        )

                        Spacer(modifier = Modifier.height(4.dp))
                        OutlinedButton(
                            onClick = { simularCalculo() },
                            modifier = Modifier.fillMaxWidth().height(48.dp),
                            shape = RoundedCornerShape(10.dp),
                            colors = ButtonDefaults.outlinedButtonColors(contentColor = OrangePrimary)
                        ) {
                            if (isSimulating) {
                                CircularProgressIndicator(modifier = Modifier.size(20.dp), strokeWidth = 2.dp, color = OrangePrimary)
                                Spacer(modifier = Modifier.width(8.dp))
                                Text("Calculando...")
                            } else {
                                Icon(Icons.Default.Info, contentDescription = null)
                                Spacer(modifier = Modifier.width(8.dp))
                                Text("Simular / Calcular Tasa Implícita", fontWeight = FontWeight.SemiBold)
                            }
                        }
                    }
                }
            }

            // Vista previa de simulación
            if (cotizacionResult != null) {
                val cot = cotizacionResult!!
                item {
                    Card(
                        modifier = Modifier.fillMaxWidth(),
                        shape = RoundedCornerShape(12.dp),
                        colors = CardDefaults.cardColors(containerColor = Color(0xFFF0FDF4)),
                        elevation = CardDefaults.cardElevation(defaultElevation = 2.dp)
                    ) {
                        Column(modifier = Modifier.padding(16.dp)) {
                            Row(
                                modifier = Modifier.fillMaxWidth(),
                                horizontalArrangement = Arrangement.SpaceBetween,
                                verticalAlignment = Alignment.CenterVertically
                            ) {
                                Text("Resultados de la Simulación", fontWeight = FontWeight.Bold, fontSize = 14.sp, color = Color(0xFF166534))
                                Text(formatPercent(cot.tasaAnualPct), fontWeight = FontWeight.Black, fontSize = 16.sp, color = OrangePrimary)
                            }
                            HorizontalDivider(modifier = Modifier.padding(vertical = 8.dp), color = Color(0xFFBBF7D0))
                            InfoRow(label = "Inversión Neta", value = formatCurrency(cot.inversionNeta), isBold = true)
                            InfoRow(label = "Tasa Mensual", value = formatPercent(cot.tasaMensualPct))
                            InfoRow(label = "Tasa Anualizada", value = formatPercent(cot.tasaAnualPct), isBold = true)
                            InfoRow(label = "VP Depósito Residual", value = formatCurrency(cot.vpResidual))
                            InfoRow(label = "Total Rentas en el Plazo", value = formatCurrency(cot.totalRentas))
                            InfoRow(label = "Intereses Totales Proyectados", value = formatCurrency(cot.totalIntereses), isBold = true)
                        }
                    }
                }
            }

            // Botón Guardar
            item {
                Button(
                    onClick = { guardarContrato() },
                    modifier = Modifier.fillMaxWidth().height(54.dp),
                    shape = RoundedCornerShape(12.dp),
                    colors = ButtonDefaults.buttonColors(containerColor = OrangePrimary)
                ) {
                    if (isSaving) {
                        CircularProgressIndicator(modifier = Modifier.size(24.dp), strokeWidth = 2.dp, color = Color.White)
                        Spacer(modifier = Modifier.width(8.dp))
                        Text("Registrando Contrato en leasing.db...", fontWeight = FontWeight.Bold)
                    } else {
                        Icon(Icons.Default.Check, contentDescription = null)
                        Spacer(modifier = Modifier.width(8.dp))
                        Text("Guardar Contrato", fontWeight = FontWeight.Bold, fontSize = 16.sp)
                    }
                }
                Spacer(modifier = Modifier.height(30.dp))
            }
        }
    }

    if (successDialogMessage != null) {
        AlertDialog(
            onDismissRequest = { /* Modal */ },
            title = { Text("Contrato Creado", fontWeight = FontWeight.Bold, color = Color(0xFF166534)) },
            text = { Text(successDialogMessage ?: "") },
            confirmButton = {
                Button(
                    onClick = {
                        val id = createdId
                        successDialogMessage = null
                        onContratoCreated(id)
                    },
                    colors = ButtonDefaults.buttonColors(containerColor = OrangePrimary)
                ) {
                    Text("Ver Detalle")
                }
            },
            dismissButton = {
                TextButton(
                    onClick = {
                        successDialogMessage = null
                        onBack()
                    }
                ) {
                    Text("Volver a Cartera")
                }
            }
        )
    }
}
