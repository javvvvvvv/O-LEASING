package com.oleasing.app

import android.content.Intent
import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.BackHandler
import androidx.activity.compose.setContent
import androidx.activity.enableEdgeToEdge
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.padding
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.DateRange
import androidx.compose.material.icons.filled.Home
import androidx.compose.material.icons.filled.Info
import androidx.compose.material.icons.filled.Search
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import com.oleasing.app.model.CotizacionDraft
import com.oleasing.app.network.OLeasingApiClient
import com.oleasing.app.ui.screens.CobranzaMesScreen
import com.oleasing.app.ui.screens.ContratoDetailScreen
import com.oleasing.app.ui.screens.ContratosListScreen
import com.oleasing.app.ui.screens.CotizadorScreen
import com.oleasing.app.ui.screens.DashboardScreen
import com.oleasing.app.ui.screens.EstadoCuentaScreen
import com.oleasing.app.ui.screens.LoginScreen
import com.oleasing.app.ui.screens.NuevoContratoScreen
import com.oleasing.app.ui.theme.OLeasingMobileTheme
import com.oleasing.app.ui.theme.OrangePrimary
import com.oleasing.app.ui.theme.SlateSecondary

sealed class Screen {
    data object Dashboard : Screen()
    data object ContratosList : Screen()
    data class ContratoDetail(val idContrato: String) : Screen()
    data class EstadoCuenta(val idContrato: String) : Screen()
    data class NuevoContrato(val draft: CotizacionDraft? = null) : Screen()
    data object CobranzaMes : Screen()
    data object Cotizador : Screen()
}

class MainActivity : ComponentActivity() {

    private var deepLinkContratoId by mutableStateOf<String?>(null)

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        OLeasingApiClient.init(applicationContext)
        enableEdgeToEdge()

        deepLinkContratoId = extractContratoIdFromIntent(intent)

        setContent {
            OLeasingMobileTheme {
                var isLoggedIn by remember { mutableStateOf(!OLeasingApiClient.authToken.isNullOrBlank()) }

                DisposableEffect(Unit) {
                    OLeasingApiClient.onUnauthorized = {
                        OLeasingApiClient.clearSession(applicationContext)
                        isLoggedIn = false
                    }
                    onDispose {
                        OLeasingApiClient.onUnauthorized = null
                    }
                }

                Surface(
                    modifier = Modifier.fillMaxSize(),
                    color = MaterialTheme.colorScheme.background
                ) {
                    if (!isLoggedIn) {
                        LoginScreen(
                            onLoginSuccess = {
                                isLoggedIn = true
                            }
                        )
                    } else {
                        val navigationStack = remember {
                            val initial = mutableStateListOf<Screen>(Screen.Dashboard)
                            deepLinkContratoId?.let { id ->
                                initial.add(Screen.EstadoCuenta(id))
                                deepLinkContratoId = null
                            }
                            initial
                        }

                        LaunchedEffect(deepLinkContratoId) {
                            val targetId = deepLinkContratoId
                            if (!targetId.isNullOrBlank()) {
                                val current = navigationStack.lastOrNull()
                                if (current !is Screen.EstadoCuenta || current.idContrato != targetId) {
                                    navigationStack.add(Screen.EstadoCuenta(targetId))
                                }
                                deepLinkContratoId = null
                            }
                        }

                        val currentScreen = navigationStack.lastOrNull() ?: Screen.Dashboard

                        fun navigateTo(screen: Screen) {
                            navigationStack.add(screen)
                        }

                        fun switchRootTab(tabScreen: Screen) {
                            navigationStack.clear()
                            navigationStack.add(Screen.Dashboard)
                            if (tabScreen !is Screen.Dashboard) {
                                navigationStack.add(tabScreen)
                            }
                        }

                        fun navigateBack() {
                            if (navigationStack.size > 1) {
                                navigationStack.removeAt(navigationStack.lastIndex)
                            }
                        }

                        BackHandler(enabled = navigationStack.size > 1) {
                            navigateBack()
                        }

                        val showBottomBar = currentScreen is Screen.Dashboard ||
                                currentScreen is Screen.ContratosList ||
                                currentScreen is Screen.CobranzaMes ||
                                currentScreen is Screen.Cotizador

                        Scaffold(
                            bottomBar = {
                                if (showBottomBar) {
                                    NavigationBar(
                                        containerColor = MaterialTheme.colorScheme.surface,
                                        tonalElevation = 6.dp
                                    ) {
                                        NavigationBarItem(
                                            selected = currentScreen is Screen.Dashboard,
                                            onClick = { switchRootTab(Screen.Dashboard) },
                                            icon = { Icon(Icons.Default.Home, contentDescription = "Inicio") },
                                            label = { Text("Inicio", fontSize = 11.sp, fontWeight = FontWeight.SemiBold) },
                                            colors = NavigationBarItemDefaults.colors(
                                                selectedIconColor = OrangePrimary,
                                                selectedTextColor = OrangePrimary,
                                                indicatorColor = Color(0xFFFFEDD5),
                                                unselectedIconColor = SlateSecondary,
                                                unselectedTextColor = SlateSecondary
                                            )
                                        )
                                        NavigationBarItem(
                                            selected = currentScreen is Screen.ContratosList,
                                            onClick = { switchRootTab(Screen.ContratosList) },
                                            icon = { Icon(Icons.Default.Search, contentDescription = "Cartera") },
                                            label = { Text("Cartera", fontSize = 11.sp, fontWeight = FontWeight.SemiBold) },
                                            colors = NavigationBarItemDefaults.colors(
                                                selectedIconColor = OrangePrimary,
                                                selectedTextColor = OrangePrimary,
                                                indicatorColor = Color(0xFFFFEDD5),
                                                unselectedIconColor = SlateSecondary,
                                                unselectedTextColor = SlateSecondary
                                            )
                                        )
                                        NavigationBarItem(
                                            selected = currentScreen is Screen.CobranzaMes,
                                            onClick = { switchRootTab(Screen.CobranzaMes) },
                                            icon = { Icon(Icons.Default.DateRange, contentDescription = "Cobranza") },
                                            label = { Text("Cobranza", fontSize = 11.sp, fontWeight = FontWeight.SemiBold) },
                                            colors = NavigationBarItemDefaults.colors(
                                                selectedIconColor = OrangePrimary,
                                                selectedTextColor = OrangePrimary,
                                                indicatorColor = Color(0xFFFFEDD5),
                                                unselectedIconColor = SlateSecondary,
                                                unselectedTextColor = SlateSecondary
                                            )
                                        )
                                        NavigationBarItem(
                                            selected = currentScreen is Screen.Cotizador,
                                            onClick = { switchRootTab(Screen.Cotizador) },
                                            icon = { Icon(Icons.Default.Info, contentDescription = "Cotizador") },
                                            label = { Text("Cotizador", fontSize = 11.sp, fontWeight = FontWeight.SemiBold) },
                                            colors = NavigationBarItemDefaults.colors(
                                                selectedIconColor = OrangePrimary,
                                                selectedTextColor = OrangePrimary,
                                                indicatorColor = Color(0xFFFFEDD5),
                                                unselectedIconColor = SlateSecondary,
                                                unselectedTextColor = SlateSecondary
                                            )
                                        )
                                    }
                                }
                            }
                        ) { innerPadding ->
                            Box(
                                modifier = Modifier
                                    .fillMaxSize()
                                    .padding(bottom = innerPadding.calculateBottomPadding())
                            ) {
                                when (val screen = currentScreen) {
                                    is Screen.Dashboard -> {
                                        DashboardScreen(
                                            onNavigateToList = { switchRootTab(Screen.ContratosList) },
                                            onNavigateToNew = { navigateTo(Screen.NuevoContrato()) },
                                            onSelectContrato = { id -> navigateTo(Screen.ContratoDetail(id)) },
                                            onNavigateToEstadoCuenta = { id -> navigateTo(Screen.EstadoCuenta(id)) },
                                            onNavigateToCobranza = { switchRootTab(Screen.CobranzaMes) },
                                            onNavigateToCotizador = { switchRootTab(Screen.Cotizador) },
                                            onLogout = { isLoggedIn = false }
                                        )
                                    }
                                    is Screen.CobranzaMes -> {
                                        CobranzaMesScreen(
                                            onBack = { navigateBack() },
                                            onNavigateToEstadoCuenta = { id -> navigateTo(Screen.EstadoCuenta(id)) }
                                        )
                                    }
                                    is Screen.Cotizador -> {
                                        CotizadorScreen(
                                            onBack = { navigateBack() },
                                            onNavigateToNewContrato = { draft ->
                                                navigateTo(Screen.NuevoContrato(draft))
                                            }
                                        )
                                    }
                                    is Screen.ContratosList -> {
                                        ContratosListScreen(
                                            onBack = { navigateBack() },
                                            onSelectContrato = { id -> navigateTo(Screen.ContratoDetail(id)) },
                                            onNavigateToEstadoCuenta = { id -> navigateTo(Screen.EstadoCuenta(id)) },
                                            onNavigateToNew = { navigateTo(Screen.NuevoContrato()) }
                                        )
                                    }
                                    is Screen.ContratoDetail -> {
                                        ContratoDetailScreen(
                                            idContrato = screen.idContrato,
                                            onBack = { navigateBack() },
                                            onNavigateToEstadoCuenta = { id -> navigateTo(Screen.EstadoCuenta(id)) }
                                        )
                                    }
                                    is Screen.EstadoCuenta -> {
                                        EstadoCuentaScreen(
                                            idContrato = screen.idContrato,
                                            onBack = { navigateBack() },
                                            onNavigateToDetail = { id -> navigateTo(Screen.ContratoDetail(id)) }
                                        )
                                    }
                                    is Screen.NuevoContrato -> {
                                        NuevoContratoScreen(
                                            draft = screen.draft,
                                            onBack = { navigateBack() },
                                            onContratoCreated = { id ->
                                                navigationStack.removeAt(navigationStack.lastIndex)
                                                navigateTo(Screen.ContratoDetail(id))
                                            }
                                        )
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    override fun onNewIntent(intent: Intent) {
        super.onNewIntent(intent)
        setIntent(intent)
        extractContratoIdFromIntent(intent)?.let { id ->
            deepLinkContratoId = id
        }
    }

    private fun extractContratoIdFromIntent(intent: Intent?): String? {
        val uri = intent?.data ?: return null
        val apiParam = uri.getQueryParameter("api")?.trim()
        if (!apiParam.isNullOrBlank() && (apiParam.startsWith("http://") || apiParam.startsWith("https://"))) {
            OLeasingApiClient.updateBaseUrl(applicationContext, apiParam)
        }
        val lastSeg = uri.lastPathSegment?.trim()
        if (!lastSeg.isNullOrBlank() && !lastSeg.equals("contrato", ignoreCase = true) && !lastSeg.equals("conectar", ignoreCase = true)) {
            return lastSeg
        }
        val host = uri.host?.trim()
        if (!host.isNullOrBlank() && !host.equals("contrato", ignoreCase = true) && !host.equals("conectar", ignoreCase = true) && !host.equals("oleasing.app", ignoreCase = true)) {
            return host
        }
        return null
    }
}
