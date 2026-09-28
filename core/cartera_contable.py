# ============================================================================
# PROPIEDAD INTELECTUAL Y LICENCIA COMERCIAL CERRADA
# ============================================================================
# Autor Legal y Titular de Derechos: JAVIER ILLAN GONZALEZ
# Organizacion: ORANGE CREW
# Contacto: ILLANJAVIER9@GMAIL.COM
#
# ADVERTENCIA LEGAL (MEXICO Y GLOBAL):
# Este codigo fuente y su arquitectura son propiedad intelectual exclusiva de
# JAVIER ILLAN GONZALEZ. Queda estrictamente prohibida su reproduccion,
# distribucion, modificacion, ingenieria inversa, copia o uso comercial sin la
# autorizacion expresa y por escrito del autor. Obra protegida conforme a la
# Ley Federal del Derecho de Autor y tratados internacionales aplicables.
# ============================================================================
"""
core/cartera_contable.py — Metricas de cartera para amarrar con contabilidad.

Funciones exportadas (las que usa app.py):
  - saldos_contrato_vigente
  - consolidar_cartera_vigente
  - resumen_facturacion_por_anio
  - cobertura_facturacion_periodo
  - acumulado_facturacion
  - acumulado_amortizacion_cartera
  - metricas_estilo_tabla_mensual
  - comparar_auxiliar

Misma logica que Tabla Mensual por Contrato:
  mc = (anio - fa.year)*12 + (mes - fa.month) + 1  (primer mes en firma)
  VP residual = VP_Residual de BD si existe
  redondeo a 2 decimales por celda
"""
from __future__ import annotations

from datetime import date
from typing import Any

import pandas as pd

from finanzas import calc_amort, calc_res_amort, vp_res

__all__ = [
    "saldos_contrato_vigente",
    "consolidar_cartera_vigente",
    "resumen_facturacion_por_anio",
    "cobertura_facturacion_periodo",
    "acumulado_facturacion",
    "acumulado_amortizacion_cartera",
    "tabla_intereses_mensual",
    "metricas_estilo_tabla_mensual",
    "comparar_auxiliar",
]


def _meses_transcurridos(fecha_alta: date, hoy: date) -> int:
    return max(0, (hoy.year - fecha_alta.year) * 12 + (hoy.month - fecha_alta.month))


def _anio_mes_de_mc(fecha_alta: date, mc: int) -> tuple[int, int]:
    """Calendario del mes de amortizacion mc (mc=1 = mes de firma)."""
    m0 = fecha_alta.month - 1 + (mc - 1)
    y = fecha_alta.year + m0 // 12
    m = m0 % 12 + 1
    return y, m


def _vp_residual_contrato(con: dict[str, Any], res: float, t: float, pl: int) -> float:
    """VP residual igual que Tabla Mensual: usa VP_Residual guardado si existe."""
    raw = con.get("VP_Residual")
    try:
        if raw is not None and float(raw) > 0:
            return float(raw)
    except (TypeError, ValueError):
        pass
    return float(vp_res(res, t, pl))


def _parse_fecha(val) -> date | None:
    try:
        ts = pd.to_datetime(val, errors="coerce")
        if pd.isna(ts):
            return None
        return ts.date() if hasattr(ts, "date") else date(int(ts.year), int(ts.month), int(ts.day))
    except Exception:
        return None


def saldos_contrato_vigente(con: dict[str, Any], hoy: date | None = None) -> dict[str, float]:
    """Saldos al corte, misma logica que Tabla Mensual por Contrato."""
    hoy = hoy or date.today()
    v = float(con.get("Valor_Sin_IVA") or 0)
    ant = float(con.get("Anticipo_Monto") or 0)
    inv = v - ant
    r = float(con.get("Mensualidad_Sin_IVA") or 0)
    res = float(con.get("Residual_Monto") or 0)
    com = float(con.get("Comision_Monto") or 0)
    pl = int(con.get("Plazo") or 0)
    t = round(float(con.get("Tasa_Calculada") or 0), 8)

    base = {
        "inversion_neta": round(inv, 2),
        "saldo_capital": round(inv, 2),
        "cxc_cp": 0.0,
        "cxc_lp": 0.0,
        "int_cp": 0.0,
        "int_lp": 0.0,
        "residual_lp": 0.0,
        "residual_cp": 0.0,
        "saldo_residual_activo": 0.0,
        "interes_leasing_mes": 0.0,
        "interes_residual_mes": 0.0,
        "comision_pasivo": round(com, 2),
        "renta_mensual": round(r, 2),
        "mes_actual": 0,
        "meses_restantes": pl,
    }
    if pl <= 0:
        return base

    fa = _parse_fecha(con.get("Fecha_Alta"))
    if fa is None:
        base["saldo_residual_activo"] = round(res, 2)
        return base

    mc = int((hoy.year - fa.year) * 12 + (hoy.month - fa.month) + 1)
    if mc < 1:
        mc = 0
    if mc > pl:
        mc = pl

    try:
        dfa, _, _, _, _ = calc_amort(round(inv, 4), round(r, 4), round(res, 4), pl, t)
    except Exception:
        return base
    vpr = _vp_residual_contrato(con, res, t, pl)
    try:
        dfr = calc_res_amort(round(vpr, 4), t, pl)
    except Exception:
        return base

    if mc <= 0 or dfa is None or dfr is None or len(dfa) < 1:
        saldo_cap = round(inv, 2)
        saldo_res_act = round(vpr, 2)
        int_mes = 0.0
        int_res_mes = 0.0
        int_cp = 0.0
        int_lp = 0.0
    else:
        idx = min(max(mc, 1), len(dfa)) - 1
        saldo_cap = round(float(dfa.iloc[idx]["Saldo"]), 2)
        saldo_res_act = round(float(dfr.iloc[idx]["Saldo_Fin"]), 2)
        int_mes = round(float(dfa.iloc[idx]["Interes"]), 2)
        int_res_mes = round(float(dfr.iloc[idx]["Interes"]), 2)
        int_cp = round(float(dfa.iloc[mc : min(pl, mc + 12)]["Interes"].sum()), 2) if mc < pl else 0.0
        int_lp = round(float(dfa.iloc[mc + 12 : pl]["Interes"].sum()), 2) if mc + 12 < pl else 0.0

    cxc_cp = round(r * min(12, max(0, pl - mc)), 2)
    cxc_lp = round(r * max(0, pl - mc - 12), 2)
    punto = max(0, pl - 12)
    if mc < punto:
        residual_lp, residual_cp = saldo_res_act, 0.0
    else:
        residual_lp, residual_cp = 0.0, saldo_res_act
    com_pasivo = round(com * max(0, pl - mc) / pl, 2) if pl else 0.0

    return {
        "inversion_neta": round(inv, 2),
        "saldo_capital": saldo_cap,
        "cxc_cp": cxc_cp,
        "cxc_lp": cxc_lp,
        "int_cp": int_cp,
        "int_lp": int_lp,
        "residual_lp": residual_lp,
        "residual_cp": residual_cp,
        "saldo_residual_activo": saldo_res_act,
        "interes_leasing_mes": int_mes,
        "interes_residual_mes": int_res_mes,
        "comision_pasivo": com_pasivo,
        "renta_mensual": round(r, 2),
        "mes_actual": mc,
        "meses_restantes": max(0, pl - mc),
    }


def consolidar_cartera_vigente(
    df_activos: pd.DataFrame,
    hoy: date | None = None,
) -> tuple[dict[str, float], pd.DataFrame]:
    """Agrega saldos de la cartera para amarrar con contabilidad."""
    hoy = hoy or date.today()
    vacios: dict[str, float] = {
        "n_contratos": 0,
        "inversion_neta": 0.0,
        "saldo_capital": 0.0,
        "cxc_cp": 0.0,
        "cxc_lp": 0.0,
        "cxc_total": 0.0,
        "int_cp": 0.0,
        "int_lp": 0.0,
        "residual_lp": 0.0,
        "residual_cp": 0.0,
        "residual_total": 0.0,
        "saldo_residual_activo": 0.0,
        "interes_leasing_mes": 0.0,
        "interes_residual_mes": 0.0,
        "comision_pasivo": 0.0,
        "renta_mensual": 0.0,
        "valor_sin_iva": 0.0,
        "anticipo": 0.0,
    }
    if df_activos is None or df_activos.empty:
        return vacios, pd.DataFrame()

    filas: list[dict[str, Any]] = []
    for _, row in df_activos.iterrows():
        con = row.to_dict()
        try:
            s = saldos_contrato_vigente(con, hoy)
        except Exception:
            continue
        filas.append({
            "ID_Contrato": con.get("ID_Contrato"),
            "Cliente": con.get("Cliente"),
            "Fecha_Alta": con.get("Fecha_Alta"),
            "Fecha_Vencimiento": con.get("Fecha_Vencimiento"),
            "Plazo": con.get("Plazo"),
            "Mes_Actual": s["mes_actual"],
            "Meses_Restantes": s["meses_restantes"],
            "Valor_Sin_IVA": round(float(con.get("Valor_Sin_IVA") or 0), 2),
            "Anticipo": round(float(con.get("Anticipo_Monto") or 0), 2),
            "Inversion_Neta": s["inversion_neta"],
            "Saldo_Capital": s["saldo_capital"],
            "CxC_CP": s["cxc_cp"],
            "CxC_LP": s["cxc_lp"],
            "Intereses_CP": s["int_cp"],
            "Intereses_LP": s["int_lp"],
            "Residual_LP": s["residual_lp"],
            "Residual_CP": s["residual_cp"],
            "Saldo_Residual_Activo": s["saldo_residual_activo"],
            "Interes_Leasing_Mes": s["interes_leasing_mes"],
            "Interes_Residual_Mes": s["interes_residual_mes"],
            "Comision_Pasivo": s["comision_pasivo"],
            "Renta_Mensual": s["renta_mensual"],
        })
    if not filas:
        return vacios, pd.DataFrame()

    detalle = pd.DataFrame(filas)
    totales = {
        "n_contratos": len(detalle),
        "inversion_neta": float(detalle["Inversion_Neta"].sum()),
        "saldo_capital": float(detalle["Saldo_Capital"].sum()),
        "cxc_cp": float(detalle["CxC_CP"].sum()),
        "cxc_lp": float(detalle["CxC_LP"].sum()),
        "cxc_total": float(detalle["CxC_CP"].sum() + detalle["CxC_LP"].sum()),
        "int_cp": float(detalle["Intereses_CP"].sum()),
        "int_lp": float(detalle["Intereses_LP"].sum()),
        "residual_lp": float(detalle["Residual_LP"].sum()),
        "residual_cp": float(detalle["Residual_CP"].sum()),
        "residual_total": float(detalle["Saldo_Residual_Activo"].sum()),
        "saldo_residual_activo": float(detalle["Saldo_Residual_Activo"].sum()),
        "interes_leasing_mes": float(detalle["Interes_Leasing_Mes"].sum()),
        "interes_residual_mes": float(detalle["Interes_Residual_Mes"].sum()),
        "comision_pasivo": float(detalle["Comision_Pasivo"].sum()),
        "renta_mensual": float(detalle["Renta_Mensual"].sum()),
        "valor_sin_iva": float(detalle["Valor_Sin_IVA"].sum()),
        "anticipo": float(detalle["Anticipo"].sum()),
    }
    return totales, detalle


def resumen_facturacion_por_anio(
    df_fact: pd.DataFrame,
    solo_vigentes: bool = True,
) -> pd.DataFrame:
    """Totales de facturacion agrupados por anio del periodo."""
    cols = ["anio", "n_facturas", "total", "subtotal", "mensual", "comision", "otros"]
    if df_fact is None or df_fact.empty or "periodo" not in df_fact.columns:
        return pd.DataFrame(columns=cols)
    df = df_fact.copy()
    if solo_vigentes and "cancelada" in df.columns:
        df = df[df["cancelada"].fillna(0).astype(int) == 0]
    if df.empty:
        return pd.DataFrame(columns=cols)
    df["anio"] = df["periodo"].astype(str).str[:4]
    tipo = df["tipo"].astype(str).str.upper() if "tipo" in df.columns else pd.Series([""] * len(df), index=df.index)
    rows = []
    for anio, g in df.groupby("anio"):
        t = tipo.loc[g.index] if len(tipo) else pd.Series([""] * len(g), index=g.index)
        total = float(g["total"].sum()) if "total" in g.columns else 0.0
        sub = float(g["subtotal"].sum()) if "subtotal" in g.columns else total
        rows.append({
            "anio": anio,
            "n_facturas": int(len(g)),
            "total": round(total, 2),
            "subtotal": round(sub, 2),
            "mensual": round(float(g.loc[t == "MENSUAL", "total"].sum()) if "total" in g.columns else 0.0, 2),
            "comision": round(float(g.loc[t == "COMISION", "total"].sum()) if "total" in g.columns else 0.0, 2),
            "otros": round(float(g.loc[~t.isin(["MENSUAL", "COMISION"]), "total"].sum()) if "total" in g.columns else 0.0, 2),
        })
    return pd.DataFrame(rows).sort_values("anio")


def cobertura_facturacion_periodo(esperado: float, facturado: float) -> dict[str, float]:
    """Compara renta esperada vs facturado del periodo."""
    esp = float(esperado or 0)
    fac = float(facturado or 0)
    dif = round(fac - esp, 2)
    pct = round((fac / esp * 100) if esp else 0.0, 2)
    return {"esperado": round(esp, 2), "facturado": round(fac, 2), "diferencia": dif, "cobertura_pct": pct}


def acumulado_facturacion(
    df_fact: pd.DataFrame,
    anio: int | None = None,
    hasta_periodo: str | None = None,
) -> dict[str, float]:
    """Suma facturas vigentes; opcionalmente filtra por anio y hasta periodo YYYY-MM."""
    out = {"n_facturas": 0, "total": 0.0, "subtotal": 0.0, "mensual": 0.0, "comision": 0.0, "otros": 0.0}
    if df_fact is None or df_fact.empty:
        return out
    df = df_fact.copy()
    if "cancelada" in df.columns:
        df = df[df["cancelada"].fillna(0).astype(int) == 0]
    if anio is not None and "periodo" in df.columns:
        df = df[df["periodo"].astype(str).str.startswith(str(anio))]
    if hasta_periodo and "periodo" in df.columns:
        df = df[df["periodo"].astype(str) <= str(hasta_periodo)]
    if df.empty:
        return out
    tipo = df["tipo"].astype(str).str.upper() if "tipo" in df.columns else pd.Series([""] * len(df))
    total = float(df["total"].sum()) if "total" in df.columns else 0.0
    sub = float(df["subtotal"].sum()) if "subtotal" in df.columns else total
    return {
        "n_facturas": int(len(df)),
        "total": round(total, 2),
        "subtotal": round(sub, 2),
        "mensual": round(float(df.loc[tipo == "MENSUAL", "total"].sum()) if "total" in df.columns else 0.0, 2),
        "comision": round(float(df.loc[tipo == "COMISION", "total"].sum()) if "total" in df.columns else 0.0, 2),
        "otros": round(float(df.loc[~tipo.isin(["MENSUAL", "COMISION"]), "total"].sum()) if "total" in df.columns else 0.0, 2),
    }


def acumulado_amortizacion_cartera(
    df_activos: pd.DataFrame,
    hoy: date | None = None,
    solo_ejercicio: bool = True,
) -> dict[str, float]:
    """Capital e intereses devengados (YTD del ejercicio por default)."""
    hoy = hoy or date.today()
    out = {
        "capital_amortizado": 0.0,
        "intereses_devengados": 0.0,
        "intereses_residual_devengados": 0.0,
        "rentas_esperadas_acum": 0.0,
        "capital_del_mes": 0.0,
        "intereses_del_mes": 0.0,
        "intereses_residual_del_mes": 0.0,
        "rentas_del_mes": 0.0,
        "saldo_residual_activo": 0.0,
        "saldo_capital": 0.0,
        "meses_contados": 0,
        "n_contratos": 0,
        "anio_ejercicio": hoy.year,
        "mes_corte": hoy.month,
    }
    if df_activos is None or df_activos.empty:
        return out

    # Delega a metricas_estilo_tabla_mensual (misma fuente de verdad)
    m = metricas_estilo_tabla_mensual(df_activos, hoy.year, hoy.month)
    out["capital_amortizado"] = m["capital_ytd"] if solo_ejercicio else m["capital_ytd"]
    out["intereses_devengados"] = m["interes_leasing_ytd"]
    out["intereses_residual_devengados"] = m["interes_residual_ytd"]
    out["rentas_esperadas_acum"] = m["renta_ytd"]
    out["capital_del_mes"] = m["capital_mes"]
    out["intereses_del_mes"] = m["interes_leasing_mes"]
    out["intereses_residual_del_mes"] = m["interes_residual_mes"]
    out["rentas_del_mes"] = m["renta_mes"]
    out["saldo_residual_activo"] = m["saldo_residual_mes"]
    out["saldo_capital"] = m["saldo_capital_mes"]
    out["meses_contados"] = m["n_celdas_ytd"]
    out["n_contratos"] = m["n_contratos_mes"]
    return out


def tabla_intereses_mensual(
    df_contratos: pd.DataFrame,
    anio: int,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Replica exacta de tabla_mensual_conceptos (app.py).

    Devuelve (di, dr, dc, ds) con columnas 1..12 (meses) e indice ID_Contrato.
    di = intereses leasing, dr = intereses residual, dc = comision, ds = saldo residual.
    """
    empty = pd.DataFrame()
    if df_contratos is None or df_contratos.empty:
        return empty, empty, empty, empty

    ids = df_contratos["ID_Contrato"].tolist()
    di = pd.DataFrame(index=ids, columns=range(1, 13), dtype=float)
    dr = di.copy()
    dc = di.copy()
    ds = di.copy()

    for _, row in df_contratos.iterrows():
        try:
            id_c = row["ID_Contrato"]
            fa = row["Fecha_Alta"]
            # Aceptar Timestamp / date / str como en app
            if not hasattr(fa, "year"):
                fa = pd.to_datetime(fa, errors="coerce")
            if pd.isna(fa):
                continue
            pl = int(row["Plazo"])
            t = round(float(row["Tasa_Calculada"]), 8)
            inv = float(row["Valor_Sin_IVA"]) - float(row["Anticipo_Monto"])
            dfa, _, _, _, _ = calc_amort(
                round(inv, 4),
                round(float(row["Mensualidad_Sin_IVA"]), 4),
                round(float(row["Residual_Monto"]), 4),
                pl,
                t,
            )
            try:
                vpr = float(row["VP_Residual"])
            except Exception:
                vpr = 0.0
            if vpr <= 0:
                vpr = float(vp_res(float(row["Residual_Monto"]), t, pl))
            dfr = calc_res_amort(round(vpr, 4), t, pl)
        except Exception:
            continue

        es_baja = str(row.get("Estatus", "") or "").upper() == "BAJA"
        fecha_baja = None
        if es_baja and pd.notna(row.get("Fecha_Baja")):
            fecha_baja = pd.to_datetime(row["Fecha_Baja"], errors="coerce")
            if pd.isna(fecha_baja):
                fecha_baja = None

        fv = row.get("Fecha_Vencimiento")
        try:
            if not hasattr(fv, "year") and fv is not None:
                fv = pd.to_datetime(fv, errors="coerce")
        except Exception:
            fv = None

        for mes in range(1, 13):
            fm = pd.Timestamp(int(anio), mes, 1)
            # Misma condicion que app.tabla_mensual_conceptos
            try:
                if fm < pd.Timestamp(fa.year, fa.month, 1):
                    continue
            except Exception:
                continue
            try:
                if fv is not None and not (isinstance(fv, float) and pd.isna(fv)):
                    if fm > pd.Timestamp(fv):
                        continue
            except Exception:
                pass
            if fecha_baja is not None and fm > pd.Timestamp(fecha_baja.year, fecha_baja.month, 1):
                continue
            mc = (int(anio) - int(fa.year)) * 12 + (mes - int(fa.month)) + 1
            if mc < 1 or mc > pl:
                continue
            try:
                di.at[id_c, mes] = round(float(dfa.iloc[mc - 1]["Interes"]), 2)
                dr.at[id_c, mes] = round(float(dfr.iloc[mc - 1]["Interes"]), 2)
                com = float(row.get("Comision_Monto") or 0)
                dc.at[id_c, mes] = round(com / pl, 2) if pl else 0.0
                ds.at[id_c, mes] = round(float(dfr.iloc[mc - 1]["Saldo_Fin"]), 2)
            except Exception:
                continue

    for d in (di, dr, dc, ds):
        d.dropna(how="all", inplace=True)
    return di, dr, dc, ds


def metricas_estilo_tabla_mensual(
    df_contratos: pd.DataFrame,
    anio: int,
    mes: int,
) -> dict[str, float]:
    """Totales del mes y YTD (ene..mes) leyendo las mismas celdas que Tabla Mensual.

    interes_leasing_* = suma de la tabla di (Intereses Leasing cta 208).
    interes_residual_* = suma de la tabla dr.
    capital_* se calcula con la misma amortizacion (no esta en tabla mensual UI
    pero usa el mismo dfa).
    """
    out = {
        "interes_leasing_mes": 0.0,
        "interes_residual_mes": 0.0,
        "capital_mes": 0.0,
        "renta_mes": 0.0,
        "saldo_residual_mes": 0.0,
        "saldo_capital_mes": 0.0,
        "interes_leasing_ytd": 0.0,
        "interes_residual_ytd": 0.0,
        "capital_ytd": 0.0,
        "renta_ytd": 0.0,
        "n_contratos_mes": 0,
        "n_celdas_ytd": 0,
        "anio": int(anio),
        "mes": int(mes),
    }
    if df_contratos is None or df_contratos.empty:
        return out

    di, dr, dc, ds = tabla_intereses_mensual(df_contratos, int(anio))
    mes = int(mes)
    anio = int(anio)

    # Intereses leasing / residual: suma directa de tablas (igual Excel int_208)
    if not di.empty:
        cols_ytd = [m for m in range(1, mes + 1) if m in di.columns]
        if mes in di.columns:
            col_mes = di[mes].dropna()
            out["interes_leasing_mes"] = round(float(col_mes.sum()), 2)
            out["n_contratos_mes"] = int(col_mes.ne(0).sum()) if len(col_mes) else int(col_mes.notna().sum())
        if cols_ytd:
            out["interes_leasing_ytd"] = round(float(di[cols_ytd].sum().sum()), 2)
            out["n_celdas_ytd"] = int(di[cols_ytd].notna().sum().sum())
    if not dr.empty:
        cols_ytd = [m for m in range(1, mes + 1) if m in dr.columns]
        if mes in dr.columns:
            out["interes_residual_mes"] = round(float(dr[mes].dropna().sum()), 2)
        if cols_ytd:
            out["interes_residual_ytd"] = round(float(dr[cols_ytd].sum().sum()), 2)
    if not ds.empty and mes in ds.columns:
        out["saldo_residual_mes"] = round(float(ds[mes].dropna().sum()), 2)

    # Capital y renta: misma amortizacion, solo para los contratos que tienen
    # interes en el mes/YTD (misma inclusion de tabla)
    n_mes_cap = 0
    for _, row in df_contratos.iterrows():
        try:
            id_c = row["ID_Contrato"]
            fa = row["Fecha_Alta"]
            if not hasattr(fa, "year"):
                fa = pd.to_datetime(fa, errors="coerce")
            if pd.isna(fa):
                continue
            pl = int(row["Plazo"])
            t = round(float(row["Tasa_Calculada"]), 8)
            inv = float(row["Valor_Sin_IVA"]) - float(row["Anticipo_Monto"])
            renta = float(row["Mensualidad_Sin_IVA"])
            residual = float(row["Residual_Monto"])
            dfa, _, _, _, _ = calc_amort(
                round(inv, 4), round(renta, 4), round(residual, 4), pl, t
            )
            es_baja = str(row.get("Estatus", "") or "").upper() == "BAJA"
            fecha_baja = None
            if es_baja and pd.notna(row.get("Fecha_Baja")):
                fecha_baja = pd.to_datetime(row["Fecha_Baja"], errors="coerce")
                if pd.isna(fecha_baja):
                    fecha_baja = None
            fv = row.get("Fecha_Vencimiento")
            if fv is not None and not hasattr(fv, "year"):
                fv = pd.to_datetime(fv, errors="coerce")
            uso_mes = False
            for m in range(1, mes + 1):
                fm = pd.Timestamp(anio, m, 1)
                if fm < pd.Timestamp(fa.year, fa.month, 1):
                    continue
                if fv is not None and not pd.isna(fv) and fm > fv:
                    continue
                if fecha_baja is not None and fm > pd.Timestamp(fecha_baja.year, fecha_baja.month, 1):
                    continue
                mc = (anio - int(fa.year)) * 12 + (m - int(fa.month)) + 1
                if mc < 1 or mc > pl:
                    continue
                cap = round(float(dfa.iloc[mc - 1]["Capital"]), 2)
                out["capital_ytd"] += cap
                out["renta_ytd"] += round(renta, 2)
                if m == mes:
                    out["capital_mes"] += cap
                    out["renta_mes"] += round(renta, 2)
                    out["saldo_capital_mes"] += round(float(dfa.iloc[mc - 1]["Saldo"]), 2)
                    uso_mes = True
            if uso_mes:
                n_mes_cap += 1
        except Exception:
            continue
    if out["n_contratos_mes"] == 0:
        out["n_contratos_mes"] = n_mes_cap
    for k, v in list(out.items()):
        if isinstance(v, float):
            out[k] = round(v, 2)
    return out


def comparar_auxiliar(
    df_aux: pd.DataFrame,
    totales_sistema: dict[str, float],
    mapa_columnas: dict[str, str] | None = None,
) -> pd.DataFrame:
    """Compara un auxiliar contable (Excel/CSV) contra totales del sistema."""
    if df_aux is None or df_aux.empty:
        return pd.DataFrame(columns=["Concepto_sistema", "Monto_sistema", "Monto_auxiliar", "Diferencia"])
    df = df_aux.copy()
    df.columns = [str(c).strip() for c in df.columns]
    cols_lower = {c.lower(): c for c in df.columns}
    col_cta = None
    for cand in ("cuenta", "nombre", "concepto", "descripcion", "cuenta contable", "nombre cuenta"):
        if cand in cols_lower:
            col_cta = cols_lower[cand]
            break
    col_saldo = None
    for cand in ("saldo", "saldo final", "importe", "monto", "haber", "debe"):
        if cand in cols_lower:
            col_saldo = cols_lower[cand]
            break
    if col_saldo is None:
        for c in df.columns:
            if pd.api.types.is_numeric_dtype(df[c]):
                col_saldo = c
                break
    filas: list[dict[str, Any]] = []
    busquedas = {
        "intereses_devengados": ["interes", "intereses", "ingreso financiero", "productos financieros"],
        "capital_amortizado": ["capital", "amortizacion", "amortización"],
        "rentas_esperadas_acum": ["renta", "cxc", "arrendamiento", "leasing"],
        "intereses_residual_devengados": ["residual"],
    }
    if mapa_columnas:
        for col_aux, clave in mapa_columnas.items():
            if col_aux in df.columns and clave in totales_sistema:
                try:
                    mon = float(pd.to_numeric(df[col_aux], errors="coerce").sum())
                except Exception:
                    mon = 0.0
                sis = float(totales_sistema.get(clave, 0) or 0)
                filas.append({
                    "Concepto_sistema": clave,
                    "Monto_sistema": sis,
                    "Monto_auxiliar": mon,
                    "Diferencia": round(sis - mon, 2),
                })
    elif col_cta is not None and col_saldo is not None:
        df["_saldo"] = pd.to_numeric(df[col_saldo], errors="coerce").fillna(0.0)
        for clave, palabras in busquedas.items():
            mask = df[col_cta].astype(str).str.lower().apply(
                lambda x, p=palabras: any(w in x for w in p)
            )
            mon = float(df.loc[mask, "_saldo"].sum()) if mask.any() else 0.0
            sis = float(totales_sistema.get(clave, 0) or 0)
            filas.append({
                "Concepto_sistema": clave,
                "Monto_sistema": sis,
                "Monto_auxiliar": round(mon, 2),
                "Diferencia": round(sis - mon, 2),
            })
    if not filas:
        for clave, sis in (totales_sistema or {}).items():
            if isinstance(sis, (int, float)):
                filas.append({
                    "Concepto_sistema": clave,
                    "Monto_sistema": float(sis),
                    "Monto_auxiliar": None,
                    "Diferencia": None,
                })
    return pd.DataFrame(filas)
