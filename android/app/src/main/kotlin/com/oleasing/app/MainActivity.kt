package com.oleasing.app

import android.content.Intent
import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.BackHandler
import androidx.activity.compose.setContent
import androidx.activity.enableEdgeToEdge
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.Surface
import androidx.compose.runtime.*
import androidx.compose.ui.Modifier
import com.oleasing.app.network.OLeasingApiClient
import com.oleasing.app.ui.screens.ContratoDetailScreen
import com.oleasing.app.ui.screens.ContratosListScreen
import com.oleasing.app.ui.screens.DashboardScreen
import com.oleasing.app.ui.screens.EstadoCuentaScreen
import com.oleasing.app.ui.screens.NuevoContratoScreen
import com.oleasing.app.ui.theme.OLeasingMobileTheme

sealed class Screen {
    data object Dashboard : Screen()
    data object ContratosList : Screen()
    data class ContratoDetail(val idContrato: String) : Screen()
    data class EstadoCuenta(val idContrato: String) : Screen()
    data object NuevoContrato : Screen()
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
                Surface(
                    modifier = Modifier.fillMaxSize(),
                    color = MaterialTheme.colorScheme.background
                ) {
                    val navigationStack = remember {
                        val initial = mutableStateListOf<Screen>(Screen.Dashboard)
                        deepLinkContratoId?.let { id ->
                            initial.add(Screen.EstadoCuenta(id))
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

                    fun navigateBack() {
                        if (navigationStack.size > 1) {
                            navigationStack.removeAt(navigationStack.lastIndex)
                        }
                    }

                    BackHandler(enabled = navigationStack.size > 1) {
                        navigateBack()
                    }

                    when (val screen = currentScreen) {
                        is Screen.Dashboard -> {
                            DashboardScreen(
                                onNavigateToList = { navigateTo(Screen.ContratosList) },
                                onNavigateToNew = { navigateTo(Screen.NuevoContrato) },
                                onSelectContrato = { id -> navigateTo(Screen.ContratoDetail(id)) },
                                onNavigateToEstadoCuenta = { id -> navigateTo(Screen.EstadoCuenta(id)) }
                            )
                        }
                        is Screen.ContratosList -> {
                            ContratosListScreen(
                                onBack = { navigateBack() },
                                onSelectContrato = { id -> navigateTo(Screen.ContratoDetail(id)) },
                                onNavigateToEstadoCuenta = { id -> navigateTo(Screen.EstadoCuenta(id)) },
                                onNavigateToNew = { navigateTo(Screen.NuevoContrato) }
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
