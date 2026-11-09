package com.oleasing.app.ui.screens

import androidx.compose.foundation.BorderStroke
import androidx.compose.foundation.background
import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.LazyRow
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.filled.ArrowBack
import androidx.compose.material.icons.filled.Add
import androidx.compose.material.icons.filled.Close
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
import com.oleasing.app.model.ClienteResumen
import com.oleasing.app.model.ContratoResumen
import com.oleasing.app.network.OLeasingApiClient
import com.oleasing.app.ui.components.MorosidadBadge
import com.oleasing.app.ui.components.StatusBadge
import com.oleasing.app.ui.components.formatCurrency
import com.oleasing.app.ui.components.formatPercent
import com.oleasing.app.ui.theme.AppCardBorder
import com.oleasing.app.ui.theme.OrangePrimary
import com.oleasing.app.ui.theme.SlateSecondary
import kotlinx.coroutines.Job
import kotlinx.coroutines.delay
import kotlinx.coroutines.launch

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun ContratosListScreen(
    onBack: () -> Unit,
    onSelectContrato: (String) -> Unit,
    onNavigateToEstadoCuenta: (String) -> Unit,
    onNavigateToNew: () -> Unit
) {
    val coroutineScope = rememberCoroutineScope()
    var vistaModo by remember { mutableStateOf("CONTRATOS") } // CONTRATOS o CLIENTES
    var searchQuery by remember { mutableStateOf("") }
    var selectedEstatus by remember { mutableStateOf("Todos") }
    var contratosList by remember { mutableStateOf<List<ContratoResumen>>(emptyList()) }
    var clientesList by remember { mutableStateOf<List<ClienteResumen>>(emptyList()) }
    var totalCount by remember { mutableStateOf(0) }
    var isLoading by remember { mutableStateOf(false) }
    var errorMessage by remember { mutableStateOf<String?>(null) }
    var searchJob by remember { mutableStateOf<Job?>(null) }

    val statusOptions = listOf("Todos", "Activo", "Concluido", "Baja", "Mora")

    fun fetchData(query: String = searchQuery, estatus: String = selectedEstatus) {
        searchJob?.cancel()
        searchJob = coroutineScope.launch {
            isLoading = true
            errorMessage = null
            if (vistaModo == "CONTRATOS") {
                val result = OLeasingApiClient.getContratos(query = query, estatus = estatus, limit = 100)
                if (result.isSuccess) {
                    val data = result.getOrNull()
                    contratosList = data?.items ?: emptyList()
                    totalCount = data?.total ?: 0
                } else {
                    errorMessage = result.exceptionOrNull()?.message
                }
            } else {
                val resCli = OLeasingApiClient.getClientes(query = query)
                if (resCli.isSuccess) {
                    clientesList = resCli.getOrNull() ?: emptyList()
                    totalCount = clientesList.size
                } else {
                    errorMessage = resCli.exceptionOrNull()?.message
                }
            }
            isLoading = false
        }
    }

    LaunchedEffect(selectedEstatus, vistaModo) {
        fetchData(query = searchQuery, estatus = selectedEstatus)
    }

    Scaffold(
        topBar = {
            TopAppBar(
                title = {
                    Column {
                        Text(
                            text = if (vistaModo == "CONTRATOS") "Cartera de Contratos" else "Directorio de Clientes",
                            fontWeight = FontWeight.Bold,
                            fontSize = 18.sp
                        )
                        Text(
                            text = if (vistaModo == "CONTRATOS") "$totalCount contratos encontrados" else "$totalCount clientes registrados",
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
                    IconButton(onClick = { fetchData() }) {
                        Icon(imageVector = Icons.Default.Refresh, contentDescription = "Actualizar", tint = OrangePrimary)
                    }
                },
                colors = TopAppBarDefaults.topAppBarColors(containerColor = MaterialTheme.colorScheme.surface)
            )
        },
        floatingActionButton = {
            if (OLeasingApiClient.currentUser?.puedeEditar != false) {
                FloatingActionButton(
                    onClick = onNavigateToNew,
                    containerColor = OrangePrimary,
                    contentColor = Color.White,
                    shape = RoundedCornerShape(16.dp)
                ) {
                    Icon(imageVector = Icons.Default.Add, contentDescription = "Nuevo Contrato")
                }
            }
        }
    ) { padding ->
        Column(
            modifier = Modifier
                .fillMaxSize()
                .padding(padding)
                .background(MaterialTheme.colorScheme.background)
        ) {
            // Selector de modo: Por Contratos vs Por Clientes
            Row(
                modifier = Modifier
                    .fillMaxWidth()
                    .padding(horizontal = 16.dp, vertical = 4.dp),
                horizontalArrangement = Arrangement.spacedBy(8.dp)
            ) {
                FilterChip(
                    selected = vistaModo == "CONTRATOS",
                    onClick = { vistaModo = "CONTRATOS" },
                    label = { Text("Vista por Contratos", fontSize = 12.sp, fontWeight = FontWeight.Bold) },
                    modifier = Modifier.weight(1f),
                    colors = FilterChipDefaults.filterChipColors(
                        selectedContainerColor = SlateSecondary,
                        selectedLabelColor = Color.White
                    )
                )
                FilterChip(
                    selected = vistaModo == "CLIENTES",
                    onClick = { vistaModo = "CLIENTES" },
                    label = { Text("Por Clientes (Exposición)", fontSize = 12.sp, fontWeight = FontWeight.Bold) },
                    modifier = Modifier.weight(1f),
                    colors = FilterChipDefaults.filterChipColors(
                        selectedContainerColor = OrangePrimary,
                        selectedLabelColor = Color.White
                    )
                )
            }

            // Buscador
            OutlinedTextField(
                value = searchQuery,
                onValueChange = {
                    searchQuery = it
                    searchJob?.cancel()
                    searchJob = coroutineScope.launch {
                        delay(350)
                        fetchData(query = it, estatus = selectedEstatus)
                    }
                },
                placeholder = {
                    Text(
                        text = if (vistaModo == "CONTRATOS") "Buscar contrato, cliente, vehículo..." else "Buscar cliente por nombre o razón social...",
                        fontSize = 14.sp
                    )
                },
                leadingIcon = { Icon(Icons.Default.Search, contentDescription = null, tint = OrangePrimary) },
                trailingIcon = {
                    if (searchQuery.isNotEmpty()) {
                        IconButton(onClick = {
                            searchQuery = ""
                            fetchData(query = "", estatus = selectedEstatus)
                        }) {
                            Icon(Icons.Default.Close, contentDescription = "Limpiar")
                        }
                    }
                },
                singleLine = true,
                shape = RoundedCornerShape(12.dp),
                colors = OutlinedTextFieldDefaults.colors(
                    focusedBorderColor = OrangePrimary,
                    unfocusedContainerColor = MaterialTheme.colorScheme.surface,
                    focusedContainerColor = MaterialTheme.colorScheme.surface
                ),
                modifier = Modifier
                    .fillMaxWidth()
                    .padding(horizontal = 16.dp, vertical = 6.dp)
            )

            // Filtros de estatus (solo en vista de Contratos)
            if (vistaModo == "CONTRATOS") {
                LazyRow(
                    modifier = Modifier
                        .fillMaxWidth()
                        .padding(horizontal = 16.dp, vertical = 2.dp),
                    horizontalArrangement = Arrangement.spacedBy(8.dp)
                ) {
                    items(statusOptions) { opt ->
                        val isSelected = opt == selectedEstatus
                        FilterChip(
                            selected = isSelected,
                            onClick = { selectedEstatus = opt },
                            label = { Text(opt, fontSize = 12.sp, fontWeight = if (isSelected) FontWeight.Bold else FontWeight.Normal) },
                            colors = FilterChipDefaults.filterChipColors(
                                selectedContainerColor = OrangePrimary,
                                selectedLabelColor = Color.White
                            )
                        )
                    }
                }
            }

            Spacer(modifier = Modifier.height(4.dp))

            if (isLoading && ((vistaModo == "CONTRATOS" && contratosList.isEmpty()) || (vistaModo == "CLIENTES" && clientesList.isEmpty()))) {
                Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                    CircularProgressIndicator(color = OrangePrimary)
                }
            } else if (errorMessage != null && contratosList.isEmpty() && clientesList.isEmpty()) {
                Box(modifier = Modifier.fillMaxSize().padding(16.dp), contentAlignment = Alignment.Center) {
                    Column(horizontalAlignment = Alignment.CenterHorizontally) {
                        Text(text = "Error al cargar información:", fontWeight = FontWeight.Bold, color = Color(0xFFDC2626))
                        Spacer(modifier = Modifier.height(6.dp))
                        Text(text = errorMessage ?: "", fontSize = 13.sp, color = MaterialTheme.colorScheme.outline)
                        Spacer(modifier = Modifier.height(12.dp))
                        Button(onClick = { fetchData() }, colors = ButtonDefaults.buttonColors(containerColor = OrangePrimary)) {
                            Text("Reintentar")
                        }
                    }
                }
            } else if (vistaModo == "CLIENTES") {
                LazyColumn(
                    modifier = Modifier
                        .fillMaxSize()
                        .padding(horizontal = 16.dp),
                    verticalArrangement = Arrangement.spacedBy(10.dp)
                ) {
                    item { Spacer(modifier = Modifier.height(4.dp)) }
                    items(clientesList) { cli ->
                        Card(
                            modifier = Modifier
                                .fillMaxWidth()
                                .clickable {
                                    searchQuery = cli.cliente
                                    selectedEstatus = "Todos"
                                    vistaModo = "CONTRATOS"
                                },
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
                                        text = cli.cliente,
                                        fontWeight = FontWeight.ExtraBold,
                                        fontSize = 14.sp,
                                        color = SlateSecondary,
                                        modifier = Modifier.weight(1f)
                                    )
                                    Box(
                                        modifier = Modifier
                                            .background(Color(0xFFFFEDD5), RoundedCornerShape(8.dp))
                                            .padding(horizontal = 8.dp, vertical = 4.dp)
                                    ) {
                                        Text(
                                            text = "${cli.contratosActivos} activos / ${cli.totalContratos} tot",
                                            color = OrangePrimary,
                                            fontWeight = FontWeight.ExtraBold,
                                            fontSize = 11.sp
                                        )
                                    }
                                }
                                HorizontalDivider(
                                    modifier = Modifier.padding(vertical = 8.dp),
                                    color = MaterialTheme.colorScheme.surfaceVariant
                                )
                                Row(
                                    modifier = Modifier.fillMaxWidth(),
                                    horizontalArrangement = Arrangement.SpaceBetween
                                ) {
                                    Column {
                                        Text("Exposición Activa", fontSize = 10.sp, color = MaterialTheme.colorScheme.outline)
                                        Text(
                                            text = formatCurrency(cli.inversionActiva),
                                            fontWeight = FontWeight.ExtraBold,
                                            fontSize = 13.sp,
                                            color = OrangePrimary
                                        )
                                    }
                                    Column(horizontalAlignment = Alignment.CenterHorizontally) {
                                        Text("Renta Mensual", fontSize = 10.sp, color = MaterialTheme.colorScheme.outline)
                                        Text(
                                            text = formatCurrency(cli.rentaMensualTotal),
                                            fontWeight = FontWeight.Bold,
                                            fontSize = 13.sp,
                                            color = SlateSecondary
                                        )
                                    }
                                    Column(horizontalAlignment = Alignment.End) {
                                        Text("Residual Total", fontSize = 10.sp, color = MaterialTheme.colorScheme.outline)
                                        Text(
                                            text = formatCurrency(cli.residualTotal),
                                            fontWeight = FontWeight.SemiBold,
                                            fontSize = 12.sp
                                        )
                                    }
                                }
                            }
                        }
                    }
                    item { Spacer(modifier = Modifier.height(80.dp)) }
                }
            } else if (contratosList.isEmpty()) {
                Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                    Text("No se encontraron contratos con estos filtros.", color = MaterialTheme.colorScheme.outline, fontSize = 14.sp)
                }
            } else {
                LazyColumn(
                    modifier = Modifier
                        .fillMaxSize()
                        .padding(horizontal = 16.dp),
                    verticalArrangement = Arrangement.spacedBy(10.dp)
                ) {
                    item { Spacer(modifier = Modifier.height(4.dp)) }
                    items(contratosList) { c ->
                        Card(
                            modifier = Modifier
                                .fillMaxWidth()
                                .clickable { onSelectContrato(c.idContrato) },
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
                                        fontSize = 15.sp,
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
                                    horizontalArrangement = Arrangement.SpaceBetween
                                ) {
                                    Column {
                                        Text("Renta mensual", fontSize = 11.sp, color = MaterialTheme.colorScheme.outline)
                                        Text(formatCurrency(c.mensualidadSinIva), fontWeight = FontWeight.Bold, fontSize = 13.sp)
                                    }
                                    Column(horizontalAlignment = Alignment.CenterHorizontally) {
                                        Text("Plazo", fontSize = 11.sp, color = MaterialTheme.colorScheme.outline)
                                        Text("${c.plazo} meses", fontWeight = FontWeight.Medium, fontSize = 13.sp)
                                    }
                                    Column(horizontalAlignment = Alignment.End) {
                                        Text("Tasa Anual", fontSize = 11.sp, color = MaterialTheme.colorScheme.outline)
                                        Text(formatPercent(c.tasaAnualPct), fontWeight = FontWeight.Bold, fontSize = 13.sp, color = SlateSecondary)
                                    }
                                }
                                Spacer(modifier = Modifier.height(10.dp))
                                Row(
                                    modifier = Modifier.fillMaxWidth(),
                                    horizontalArrangement = Arrangement.spacedBy(8.dp)
                                ) {
                                    OutlinedButton(
                                        onClick = { onSelectContrato(c.idContrato) },
                                        modifier = Modifier.weight(1f).height(34.dp),
                                        shape = RoundedCornerShape(8.dp),
                                        contentPadding = PaddingValues(horizontal = 8.dp, vertical = 0.dp)
                                    ) {
                                        Text("Ficha Técnica", fontSize = 11.sp, fontWeight = FontWeight.SemiBold, color = SlateSecondary)
                                    }
                                    Button(
                                        onClick = { onNavigateToEstadoCuenta(c.idContrato) },
                                        modifier = Modifier.weight(1f).height(34.dp),
                                        shape = RoundedCornerShape(8.dp),
                                        colors = ButtonDefaults.buttonColors(containerColor = OrangePrimary),
                                        contentPadding = PaddingValues(horizontal = 8.dp, vertical = 0.dp)
                                    ) {
                                        Text("Estado de Cuenta", fontSize = 11.sp, fontWeight = FontWeight.Bold, color = Color.White)
                                    }
                                }
                            }
                        }
                    }
                    item { Spacer(modifier = Modifier.height(80.dp)) }
                }
            }
        }
    }
}
