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
import com.oleasing.app.model.ContratoResumen
import com.oleasing.app.network.OLeasingApiClient
import com.oleasing.app.ui.components.MorosidadBadge
import com.oleasing.app.ui.components.StatusBadge
import com.oleasing.app.ui.components.formatCurrency
import com.oleasing.app.ui.components.formatPercent
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
    var searchQuery by remember { mutableStateOf("") }
    var selectedEstatus by remember { mutableStateOf("Todos") }
    var contratosList by remember { mutableStateOf<List<ContratoResumen>>(emptyList()) }
    var totalCount by remember { mutableStateOf(0) }
    var isLoading by remember { mutableStateOf(false) }
    var errorMessage by remember { mutableStateOf<String?>(null) }
    var searchJob by remember { mutableStateOf<Job?>(null) }

    val statusOptions = listOf("Todos", "Activo", "Concluido", "Baja", "Mora")

    fun fetchContratos(query: String = searchQuery, estatus: String = selectedEstatus) {
        searchJob?.cancel()
        searchJob = coroutineScope.launch {
            isLoading = true
            errorMessage = null
            val result = OLeasingApiClient.getContratos(query = query, estatus = estatus, limit = 100)
            if (result.isSuccess) {
                val data = result.getOrNull()
                contratosList = data?.items ?: emptyList()
                totalCount = data?.total ?: 0
            } else {
                errorMessage = result.exceptionOrNull()?.message
            }
            isLoading = false
        }
    }

    LaunchedEffect(selectedEstatus) {
        fetchContratos(query = searchQuery, estatus = selectedEstatus)
    }

    Scaffold(
        topBar = {
            TopAppBar(
                title = {
                    Column {
                        Text("Cartera de Contratos", fontWeight = FontWeight.Bold, fontSize = 18.sp)
                        Text("$totalCount contratos encontrados", fontSize = 11.sp, color = MaterialTheme.colorScheme.outline)
                    }
                },
                navigationIcon = {
                    IconButton(onClick = onBack) {
                        Icon(imageVector = Icons.AutoMirrored.Filled.ArrowBack, contentDescription = "Volver")
                    }
                },
                actions = {
                    IconButton(onClick = { fetchContratos() }) {
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
            // Buscador
            OutlinedTextField(
                value = searchQuery,
                onValueChange = {
                    searchQuery = it
                    searchJob?.cancel()
                    searchJob = coroutineScope.launch {
                        delay(350)
                        fetchContratos(query = it, estatus = selectedEstatus)
                    }
                },
                placeholder = { Text("Buscar contrato, cliente, vehículo...", fontSize = 14.sp) },
                leadingIcon = { Icon(Icons.Default.Search, contentDescription = null, tint = OrangePrimary) },
                trailingIcon = {
                    if (searchQuery.isNotEmpty()) {
                        IconButton(onClick = {
                            searchQuery = ""
                            fetchContratos(query = "", estatus = selectedEstatus)
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
                    .padding(horizontal = 16.dp, vertical = 8.dp)
            )

            // Filtros de estatus (Chips)
            LazyRow(
                modifier = Modifier
                    .fillMaxWidth()
                    .padding(horizontal = 16.dp, vertical = 4.dp),
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

            Spacer(modifier = Modifier.height(4.dp))

            if (isLoading && contratosList.isEmpty()) {
                Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                    CircularProgressIndicator(color = OrangePrimary)
                }
            } else if (errorMessage != null && contratosList.isEmpty()) {
                Box(modifier = Modifier.fillMaxSize().padding(16.dp), contentAlignment = Alignment.Center) {
                    Column(horizontalAlignment = Alignment.CenterHorizontally) {
                        Text(text = "Error al cargar contratos:", fontWeight = FontWeight.Bold, color = Color(0xFFDC2626))
                        Spacer(modifier = Modifier.height(6.dp))
                        Text(text = errorMessage ?: "", fontSize = 13.sp, color = MaterialTheme.colorScheme.outline)
                        Spacer(modifier = Modifier.height(12.dp))
                        Button(onClick = { fetchContratos() }, colors = ButtonDefaults.buttonColors(containerColor = OrangePrimary)) {
                            Text("Reintentar")
                        }
                    }
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
                            shape = RoundedCornerShape(12.dp),
                            colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
                            elevation = CardDefaults.cardElevation(defaultElevation = 2.dp)
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
