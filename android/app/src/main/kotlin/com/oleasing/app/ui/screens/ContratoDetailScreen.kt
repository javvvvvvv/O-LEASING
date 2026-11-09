package com.oleasing.app.ui.screens

import android.content.Intent
import android.net.Uri
import androidx.compose.foundation.background
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.foundation.text.KeyboardOptions
import androidx.compose.foundation.verticalScroll
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.filled.ArrowBack
import androidx.compose.material.icons.filled.Edit
import androidx.compose.material.icons.filled.Refresh
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.input.KeyboardType
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import com.oleasing.app.model.AnotacionItem
import com.oleasing.app.model.ContratoCreate
import com.oleasing.app.model.ContratoDetalle
import com.oleasing.app.network.OLeasingApiClient
import com.oleasing.app.ui.components.InfoRow
import com.oleasing.app.ui.components.MorosidadBadge
import com.oleasing.app.ui.components.StatusBadge
import com.oleasing.app.ui.components.formatCurrency
import com.oleasing.app.ui.components.formatPercent
import com.oleasing.app.ui.components.shareExecutiveText
import com.oleasing.app.ui.theme.OrangePrimary
import com.oleasing.app.ui.theme.SlateSecondary
import kotlinx.coroutines.launch

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun ContratoDetailScreen(
    idContrato: String,
    onBack: () -> Unit,
    onNavigateToEstadoCuenta: (String) -> Unit
) {
    val context = LocalContext.current
    val coroutineScope = rememberCoroutineScope()
    var contrato by remember { mutableStateOf<ContratoDetalle?>(null) }
    var anotacionesList by remember { mutableStateOf<List<AnotacionItem>>(emptyList()) }
    var nuevaNotaTexto by remember { mutableStateOf("") }
    var isLoading by remember { mutableStateOf(false) }
    var errorMessage by remember { mutableStateOf<String?>(null) }

    var showEditDialog by remember { mutableStateOf(false) }
    var editCliente by remember { mutableStateOf("") }
    var editVehiculo by remember { mutableStateOf("") }
    var editFechaAlta by remember { mutableStateOf("") }
    var editPlazo by remember { mutableStateOf("") }
    var editValor by remember { mutableStateOf("") }
    var editAnticipo by remember { mutableStateOf("") }
    var editRenta by remember { mutableStateOf("") }
    var editResidual by remember { mutableStateOf("") }
    var editComisionPct by remember { mutableStateOf("") }
    var editAnotaciones by remember { mutableStateOf("") }
    var editError by remember { mutableStateOf<String?>(null) }
    var isSavingEdit by remember { mutableStateOf(false) }

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
            val notesRes = OLeasingApiClient.getAnotaciones(idContrato)
            if (notesRes.isSuccess) {
                anotacionesList = notesRes.getOrNull() ?: emptyList()
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
                                Row(horizontalArrangement = Arrangement.spacedBy(6.dp)) {
                                    MorosidadBadge(nivel = c.nivelMorosidad)
                                    StatusBadge(estatus = c.estatus)
                                }
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
                            Spacer(modifier = Modifier.height(12.dp))
                            Button(
                                onClick = { onNavigateToEstadoCuenta(c.idContrato) },
                                modifier = Modifier
                                    .fillMaxWidth()
                                    .height(48.dp),
                                shape = RoundedCornerShape(10.dp),
                                colors = ButtonDefaults.buttonColors(containerColor = OrangePrimary)
                            ) {
                                Text(
                                    text = "Ver Estado de Cuenta y Amortización",
                                    fontWeight = FontWeight.Bold,
                                    fontSize = 14.sp
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

                // Anotaciones y Gestión de Estatus
                item {
                    Card(
                        modifier = Modifier.fillMaxWidth(),
                        shape = RoundedCornerShape(14.dp),
                        colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
                        elevation = CardDefaults.cardElevation(defaultElevation = 2.dp)
                    ) {
                        Column(modifier = Modifier.padding(16.dp)) {
                            Text(
                                text = "Gestión Rápida y Documentos",
                                fontWeight = FontWeight.Bold,
                                fontSize = 15.sp,
                                color = SlateSecondary
                            )
                            Spacer(modifier = Modifier.height(8.dp))
                            val canEdit = OLeasingApiClient.currentUser?.puedeEditar != false
                            if (canEdit) {
                                Text("Cambiar Estatus del Contrato:", fontSize = 12.sp, color = MaterialTheme.colorScheme.outline)
                                Row(
                                    modifier = Modifier.fillMaxWidth().padding(vertical = 6.dp),
                                    horizontalArrangement = Arrangement.spacedBy(8.dp)
                                ) {
                                    listOf("ACTIVO", "CONCLUIDO", "BAJA").forEach { est ->
                                        val isCurrent = c.estatus.equals(est, ignoreCase = true)
                                        FilterChip(
                                            selected = isCurrent,
                                            onClick = {
                                                if (!isCurrent) {
                                                    coroutineScope.launch {
                                                        val upd = OLeasingApiClient.actualizarEstatusContrato(c.idContrato, est, c.nivelMorosidad)
                                                        if (upd.isSuccess) {
                                                            contrato = upd.getOrNull()
                                                        }
                                                    }
                                                }
                                            },
                                            label = { Text(est, fontSize = 12.sp, fontWeight = if (isCurrent) FontWeight.Bold else FontWeight.Normal) },
                                            colors = FilterChipDefaults.filterChipColors(
                                                selectedContainerColor = OrangePrimary,
                                                selectedLabelColor = Color.White
                                            )
                                        )
                                    }
                                }
                                Spacer(modifier = Modifier.height(4.dp))
                                Text("Nivel de Seguimiento / Morosidad:", fontSize = 12.sp, color = MaterialTheme.colorScheme.outline)
                                Row(
                                    modifier = Modifier.fillMaxWidth().padding(vertical = 4.dp),
                                    horizontalArrangement = Arrangement.spacedBy(6.dp)
                                ) {
                                    listOf(0 to "Normal", 1 to "N1", 2 to "N2", 3 to "N3", 4 to "N4").forEach { (niv, lbl) ->
                                        val sel = c.nivelMorosidad == niv
                                        FilterChip(
                                            selected = sel,
                                            onClick = {
                                                if (!sel) {
                                                    coroutineScope.launch {
                                                        val upd = OLeasingApiClient.actualizarEstatusContrato(c.idContrato, c.estatus, niv)
                                                        if (upd.isSuccess) {
                                                            contrato = upd.getOrNull()
                                                        }
                                                    }
                                                }
                                            },
                                            label = { Text(lbl, fontSize = 11.sp, fontWeight = if (sel) FontWeight.Bold else FontWeight.Normal) },
                                            colors = FilterChipDefaults.filterChipColors(
                                                selectedContainerColor = if (niv == 0) SlateSecondary else Color(0xFFDC2626),
                                                selectedLabelColor = Color.White
                                            )
                                        )
                                    }
                                }
                                Spacer(modifier = Modifier.height(6.dp))
                                OutlinedButton(
                                    onClick = {
                                        editCliente = c.cliente
                                        editVehiculo = c.vehiculo
                                        editFechaAlta = c.fechaAlta
                                        editPlazo = c.plazo.toString()
                                        editValor = c.valorSinIva.toString()
                                        editAnticipo = c.anticipoMonto.toString()
                                        editRenta = c.mensualidadSinIva.toString()
                                        editResidual = c.residualMonto.toString()
                                        editComisionPct = c.comisionAperturaPct.toString()
                                        editAnotaciones = c.anotaciones
                                        editError = null
                                        showEditDialog = true
                                    },
                                    modifier = Modifier.fillMaxWidth(),
                                    shape = RoundedCornerShape(10.dp)
                                ) {
                                    Icon(Icons.Default.Edit, contentDescription = null, tint = OrangePrimary, modifier = Modifier.size(16.dp))
                                    Spacer(modifier = Modifier.width(6.dp))
                                    Text("Editar Parámetros del Contrato", fontWeight = FontWeight.Bold, fontSize = 12.sp, color = OrangePrimary)
                                }
                                Spacer(modifier = Modifier.height(6.dp))
                            }
                            Row(
                                modifier = Modifier.fillMaxWidth(),
                                horizontalArrangement = Arrangement.spacedBy(8.dp)
                            ) {
                                OutlinedButton(
                                    onClick = {
                                        val resumen = buildString {
                                            appendLine("O-LEASING · FICHA DE CONTRATO ${c.idContrato}")
                                            appendLine("Cliente: ${c.cliente}")
                                            if (c.vehiculo.isNotBlank()) appendLine("Activo: ${c.vehiculo}")
                                            appendLine("Estatus: ${c.estatus}")
                                            appendLine("Valor s/IVA: ${formatCurrency(c.valorSinIva)}")
                                            appendLine("Inversión Neta: ${formatCurrency(c.inversionNeta)}")
                                            appendLine("Renta Mensual s/IVA: ${formatCurrency(c.mensualidadSinIva)}")
                                            appendLine("Plazo: ${c.plazo} meses (${c.fechaAlta} a ${c.fechaVencimiento})")
                                            appendLine("Residual Pactado: ${formatCurrency(c.residualMonto)}")
                                            appendLine("Saldo Insoluto Est.: ${formatCurrency(c.saldoActualEstimado)}")
                                        }
                                        shareExecutiveText(context, "Ficha Contrato ${c.idContrato}", resumen)
                                    },
                                    modifier = Modifier.weight(1f),
                                    shape = RoundedCornerShape(10.dp)
                                ) {
                                    Text("Compartir Resumen", fontWeight = FontWeight.Bold, fontSize = 12.sp, color = SlateSecondary)
                                }
                                Button(
                                    onClick = {
                                        coroutineScope.launch {
                                            val linkRes = OLeasingApiClient.requestPdfSignedUrl(c.idContrato)
                                            if (linkRes.isSuccess) {
                                                val intent = Intent(Intent.ACTION_VIEW, Uri.parse(linkRes.getOrNull()!!))
                                                context.startActivity(intent)
                                            }
                                        }
                                    },
                                    modifier = Modifier.weight(1f),
                                    shape = RoundedCornerShape(10.dp),
                                    colors = ButtonDefaults.buttonColors(containerColor = OrangePrimary)
                                ) {
                                    Text("Abrir PDF Oficial", fontWeight = FontWeight.Bold, fontSize = 12.sp, color = Color.White)
                                }
                            }
                        }
                    }
                }

                // Bitácora de Notas y Seguimiento
                item {
                    Card(
                        modifier = Modifier.fillMaxWidth(),
                        shape = RoundedCornerShape(14.dp),
                        colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
                        elevation = CardDefaults.cardElevation(defaultElevation = 2.dp)
                    ) {
                        Column(modifier = Modifier.padding(16.dp)) {
                            Text(
                                text = "Bitácora y Notas del Contrato (${anotacionesList.size})",
                                fontWeight = FontWeight.Bold,
                                fontSize = 15.sp,
                                color = SlateSecondary
                            )
                            if (c.anotaciones.isNotBlank()) {
                                Spacer(modifier = Modifier.height(4.dp))
                                Text(
                                    text = "Nota inicial: ${c.anotaciones}",
                                    fontSize = 12.sp,
                                    color = MaterialTheme.colorScheme.outline
                                )
                            }
                            if (OLeasingApiClient.currentUser?.puedeEditar != false) {
                                Spacer(modifier = Modifier.height(8.dp))
                                Row(
                                    modifier = Modifier.fillMaxWidth(),
                                    horizontalArrangement = Arrangement.spacedBy(8.dp),
                                    verticalAlignment = Alignment.CenterVertically
                                ) {
                                    OutlinedTextField(
                                        value = nuevaNotaTexto,
                                        onValueChange = { nuevaNotaTexto = it },
                                        placeholder = { Text("Agregar nota de seguimiento...", fontSize = 12.sp) },
                                        singleLine = true,
                                        modifier = Modifier.weight(1f)
                                    )
                                    Button(
                                        onClick = {
                                            if (nuevaNotaTexto.isNotBlank()) {
                                                val txt = nuevaNotaTexto.trim()
                                                nuevaNotaTexto = ""
                                                coroutineScope.launch {
                                                    val res = OLeasingApiClient.crearAnotacion(c.idContrato, "Móvil", txt)
                                                    if (res.isSuccess) {
                                                        anotacionesList = res.getOrNull() ?: emptyList()
                                                    }
                                                }
                                            }
                                        },
                                        colors = ButtonDefaults.buttonColors(containerColor = OrangePrimary),
                                        shape = RoundedCornerShape(10.dp)
                                    ) {
                                        Text("Guardar", fontSize = 12.sp, color = Color.White, fontWeight = FontWeight.Bold)
                                    }
                                }
                            }
                            Spacer(modifier = Modifier.height(8.dp))
                            anotacionesList.forEach { n ->
                                HorizontalDivider(modifier = Modifier.padding(vertical = 6.dp), color = MaterialTheme.colorScheme.surfaceVariant)
                                Row(
                                    modifier = Modifier.fillMaxWidth(),
                                    horizontalArrangement = Arrangement.SpaceBetween
                                ) {
                                    Text(n.tipo, fontWeight = FontWeight.Bold, fontSize = 11.sp, color = OrangePrimary)
                                    Text(n.fecha, fontSize = 11.sp, color = MaterialTheme.colorScheme.outline)
                                }
                                Text(n.texto, fontSize = 13.sp, color = MaterialTheme.colorScheme.onSurface)
                            }
                        }
                    }
                }

                item { Spacer(modifier = Modifier.height(24.dp)) }
            }
        }
    }

    if (showEditDialog && contrato != null) {
        val c = contrato!!
        AlertDialog(
            onDismissRequest = { if (!isSavingEdit) showEditDialog = false },
            title = { Text("Editar Contrato ${c.idContrato}", fontWeight = FontWeight.Bold) },
            text = {
                Column(
                    modifier = Modifier.verticalScroll(rememberScrollState()),
                    verticalArrangement = Arrangement.spacedBy(8.dp)
                ) {
                    if (editError != null) {
                        Text(editError ?: "", color = Color(0xFFDC2626), fontSize = 12.sp, fontWeight = FontWeight.SemiBold)
                    }
                    OutlinedTextField(
                        value = editCliente,
                        onValueChange = { editCliente = it },
                        label = { Text("Cliente / Razón Social") },
                        singleLine = true,
                        modifier = Modifier.fillMaxWidth()
                    )
                    OutlinedTextField(
                        value = editVehiculo,
                        onValueChange = { editVehiculo = it },
                        label = { Text("Vehículo / Activo") },
                        singleLine = true,
                        modifier = Modifier.fillMaxWidth()
                    )
                    Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                        OutlinedTextField(
                            value = editFechaAlta,
                            onValueChange = { editFechaAlta = it },
                            label = { Text("Fecha Alta (YYYY-MM-DD)") },
                            singleLine = true,
                            modifier = Modifier.weight(1.2f)
                        )
                        OutlinedTextField(
                            value = editPlazo,
                            onValueChange = { editPlazo = it },
                            label = { Text("Plazo (m)") },
                            keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                            singleLine = true,
                            modifier = Modifier.weight(0.8f)
                        )
                    }
                    Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                        OutlinedTextField(
                            value = editValor,
                            onValueChange = { editValor = it },
                            label = { Text("Valor Activo s/IVA") },
                            keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                            singleLine = true,
                            modifier = Modifier.weight(1f)
                        )
                        OutlinedTextField(
                            value = editAnticipo,
                            onValueChange = { editAnticipo = it },
                            label = { Text("Anticipo s/IVA") },
                            keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                            singleLine = true,
                            modifier = Modifier.weight(1f)
                        )
                    }
                    Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                        OutlinedTextField(
                            value = editRenta,
                            onValueChange = { editRenta = it },
                            label = { Text("Renta Mensual s/IVA") },
                            keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                            singleLine = true,
                            modifier = Modifier.weight(1f)
                        )
                        OutlinedTextField(
                            value = editResidual,
                            onValueChange = { editResidual = it },
                            label = { Text("Residual s/IVA") },
                            keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                            singleLine = true,
                            modifier = Modifier.weight(1f)
                        )
                    }
                    OutlinedTextField(
                        value = editAnotaciones,
                        onValueChange = { editAnotaciones = it },
                        label = { Text("Nota Inicial / Observaciones") },
                        singleLine = true,
                        modifier = Modifier.fillMaxWidth()
                    )
                }
            },
            confirmButton = {
                Button(
                    enabled = !isSavingEdit,
                    onClick = {
                        val v = editValor.toDoubleOrNull() ?: 0.0
                        val p = editPlazo.toIntOrNull() ?: 0
                        val r = editRenta.toDoubleOrNull() ?: 0.0
                        val ant = editAnticipo.toDoubleOrNull() ?: 0.0
                        val res = editResidual.toDoubleOrNull() ?: 0.0
                        val com = editComisionPct.toDoubleOrNull() ?: 0.0
                        if (editCliente.isBlank() || v <= 0 || p <= 0 || r <= 0) {
                            editError = "Verifica cliente, valor, plazo y renta mensual."
                            return@Button
                        }
                        coroutineScope.launch {
                            isSavingEdit = true
                            editError = null
                            val req = ContratoCreate(
                                idContrato = c.idContrato,
                                cliente = editCliente.trim(),
                                vehiculo = editVehiculo.trim(),
                                fechaAlta = editFechaAlta.trim(),
                                plazo = p,
                                valorSinIva = v,
                                anticipoMonto = ant,
                                mensualidadSinIva = r,
                                residualMonto = res,
                                comisionPct = com,
                                anotaciones = editAnotaciones.trim()
                            )
                            val upd = OLeasingApiClient.editarContrato(c.idContrato, req)
                            if (upd.isSuccess) {
                                contrato = upd.getOrNull()
                                showEditDialog = false
                                val notesRes = OLeasingApiClient.getAnotaciones(c.idContrato)
                                if (notesRes.isSuccess) {
                                    anotacionesList = notesRes.getOrNull() ?: emptyList()
                                }
                            } else {
                                editError = upd.exceptionOrNull()?.message ?: "Error al actualizar contrato."
                            }
                            isSavingEdit = false
                        }
                    },
                    colors = ButtonDefaults.buttonColors(containerColor = OrangePrimary)
                ) {
                    Text(if (isSavingEdit) "Guardando..." else "Guardar Cambios")
                }
            },
            dismissButton = {
                TextButton(onClick = { showEditDialog = false }) {
                    Text("Cancelar")
                }
            }
        )
    }
}
