package com.oleasing.app.ui.components

import android.content.Context
import android.content.Intent
import androidx.compose.foundation.BorderStroke
import androidx.compose.foundation.background
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material3.*
import androidx.compose.runtime.Composable
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import com.oleasing.app.ui.theme.*
import java.text.NumberFormat
import java.util.Locale

fun formatCurrency(amount: Double): String {
    val format = NumberFormat.getCurrencyInstance(Locale.US)
    return format.format(amount)
}

fun formatPercent(pct: Double): String {
    return String.format(Locale.US, "%.2f%%", pct)
}

fun shareExecutiveText(context: Context, subject: String, body: String) {
    val sendIntent = Intent().apply {
        action = Intent.ACTION_SEND
        putExtra(Intent.EXTRA_SUBJECT, subject)
        putExtra(Intent.EXTRA_TEXT, body)
        type = "text/plain"
    }
    val shareIntent = Intent.createChooser(sendIntent, subject)
    context.startActivity(shareIntent)
}

@Composable
fun MorosidadBadge(nivel: Int, modifier: Modifier = Modifier) {
    if (nivel <= 0) return
    val (bg, fg) = when (nivel) {
        1 -> Pair(Color(0xFFFEF3C7), Color(0xFFB45309))
        2 -> Pair(Color(0xFFFFEDD5), Color(0xFFC2410C))
        else -> Pair(Color(0xFFFEE2E2), Color(0xFFB91C1C))
    }
    Box(
        modifier = modifier
            .background(bg, shape = RoundedCornerShape(6.dp))
            .padding(horizontal = 8.dp, vertical = 3.dp),
        contentAlignment = Alignment.Center
    ) {
        Text(
            text = "MORA N$nivel",
            color = fg,
            fontSize = 10.sp,
            fontWeight = FontWeight.ExtraBold
        )
    }
}

@Composable
fun StatusBadge(estatus: String, modifier: Modifier = Modifier) {
    val (bgColor, textColor) = when (estatus.trim().lowercase()) {
        "activo" -> Pair(StatusActivoBg, StatusActivo)
        "concluido" -> Pair(StatusConcluidoBg, StatusConcluido)
        "baja" -> Pair(StatusBajaBg, StatusBaja)
        else -> Pair(AppSurfaceVariant, SlateSecondary)
    }

    Box(
        modifier = modifier
            .background(bgColor, shape = RoundedCornerShape(8.dp))
            .padding(horizontal = 10.dp, vertical = 4.dp),
        contentAlignment = Alignment.Center
    ) {
        Text(
            text = estatus.uppercase(),
            color = textColor,
            fontSize = 11.sp,
            fontWeight = FontWeight.ExtraBold,
            letterSpacing = 0.4.sp
        )
    }
}

@Composable
fun StatCard(
    title: String,
    value: String,
    subtitle: String? = null,
    colorAccent: Color = OrangePrimary,
    modifier: Modifier = Modifier
) {
    Card(
        modifier = modifier,
        shape = RoundedCornerShape(14.dp),
        colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
        elevation = CardDefaults.cardElevation(defaultElevation = 2.dp),
        border = BorderStroke(1.dp, AppCardBorder)
    ) {
        Column(modifier = Modifier.fillMaxWidth()) {
            Box(
                modifier = Modifier
                    .fillMaxWidth()
                    .height(4.dp)
                    .background(colorAccent)
            )
            Column(
                modifier = Modifier
                    .fillMaxWidth()
                    .padding(14.dp)
            ) {
                Text(
                    text = title.uppercase(),
                    color = MaterialTheme.colorScheme.outline,
                    fontSize = 10.sp,
                    fontWeight = FontWeight.Bold,
                    letterSpacing = 0.5.sp
                )
                Spacer(modifier = Modifier.height(6.dp))
                Text(
                    text = value,
                    color = colorAccent,
                    fontSize = 19.sp,
                    fontWeight = FontWeight.ExtraBold
                )
                if (subtitle != null) {
                    Spacer(modifier = Modifier.height(2.dp))
                    Text(
                        text = subtitle,
                        color = MaterialTheme.colorScheme.outline,
                        fontSize = 11.sp
                    )
                }
            }
        }
    }
}

@Composable
fun InfoRow(label: String, value: String, isBold: Boolean = false) {
    Row(
        modifier = Modifier
            .fillMaxWidth()
            .padding(vertical = 4.dp),
        horizontalArrangement = Arrangement.SpaceBetween,
        verticalAlignment = Alignment.CenterVertically
    ) {
        Text(
            text = label,
            color = MaterialTheme.colorScheme.outline,
            fontSize = 13.sp
        )
        Text(
            text = value,
            color = MaterialTheme.colorScheme.onSurface,
            fontSize = 14.sp,
            fontWeight = if (isBold) FontWeight.Bold else FontWeight.Medium
        )
    }
}
