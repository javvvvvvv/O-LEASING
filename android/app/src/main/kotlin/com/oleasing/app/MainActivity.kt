package com.oleasing.app

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
import com.oleasing.app.ui.screens.ContratoDetailScreen
import com.oleasing.app.ui.screens.ContratosListScreen
import com.oleasing.app.ui.screens.DashboardScreen
import com.oleasing.app.ui.screens.NuevoContratoScreen
import com.oleasing.app.ui.theme.OLeasingMobileTheme

sealed class Screen {
    data object Dashboard : Screen()
    data object ContratosList : Screen()
    data class ContratoDetail(val idContrato: String) : Screen()
    data object NuevoContrato : Screen()
}

class MainActivity : ComponentActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()

        setContent {
            OLeasingMobileTheme {
                Surface(
                    modifier = Modifier.fillMaxSize(),
                    color = MaterialTheme.colorScheme.background
                ) {
                    val navigationStack = remember { mutableStateListOf<Screen>(Screen.Dashboard) }
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
                                onSelectContrato = { id -> navigateTo(Screen.ContratoDetail(id)) }
                            )
                        }
                        is Screen.ContratosList -> {
                            ContratosListScreen(
                                onBack = { navigateBack() },
                                onSelectContrato = { id -> navigateTo(Screen.ContratoDetail(id)) },
                                onNavigateToNew = { navigateTo(Screen.NuevoContrato) }
                            )
                        }
                        is Screen.ContratoDetail -> {
                            ContratoDetailScreen(
                                idContrato = screen.idContrato,
                                onBack = { navigateBack() }
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
}
