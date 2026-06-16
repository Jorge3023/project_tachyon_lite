"""
Parser robusto de reportes de producción SMT.
Maneja columnas duplicadas, formatos de fecha variados,
valores nulos, y distintos nombres de columnas.
"""

import pandas as pd
import numpy as np
import io
import re
from datetime import timedelta


# ── Lectura ────────────────────────────────────────────────────────────────────

def leer_excel(file_bytes: bytes) -> pd.DataFrame:
    """Lee la primera hoja del Excel sin importar su nombre."""
    xl  = pd.ExcelFile(io.BytesIO(file_bytes))
    df  = xl.parse(xl.sheet_names[0], header=0)

    # Renombrar columnas duplicadas ANTES de cualquier operación
    # pandas las llama "Col", "Col.1", "Col.2" — las dejamos únicas
    cols = []
    seen = {}
    for c in df.columns:
        c = str(c).strip()
        if c in seen:
            seen[c] += 1
            cols.append(f"{c}__{seen[c]}")
        else:
            seen[c] = 0
            cols.append(c)
    df.columns = cols
    return df


# ── Detección flexible de columnas ────────────────────────────────────────────

def _find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Devuelve el nombre real de la primera columna que coincida (case-insensitive)."""
    col_lower = {c.lower().replace(" ", "").replace("_", "").replace(".", ""): c
                 for c in df.columns}
    for cand in candidates:
        key = cand.lower().replace(" ", "").replace("_", "").replace(".", "")
        if key in col_lower:
            return col_lower[key]
    return None


# ── Parseo de fechas robusto ───────────────────────────────────────────────────

def _parse_datetime_col(series: pd.Series) -> pd.Series:
    """
    Intenta varios formatos de fecha/hora comunes en reportes SMT.
    Devuelve una Series de Timestamps (NaT para los que no parseen).
    """
    if pd.api.types.is_datetime64_any_dtype(series):
        return series

    # Formatos comunes en reportes SMT
    formatos = [
        "%d %b %Y %H:%M:%S",   # 21 May 2026 18:10:55
        "%d/%m/%Y %H:%M:%S",   # 21/05/2026 18:10:55
        "%d/%m/%Y %H:%M",      # 21/05/2026 18:10
        "%Y-%m-%d %H:%M:%S",   # 2026-05-21 18:10:55
        "%Y-%m-%dT%H:%M:%S",   # ISO
        "%m/%d/%Y %H:%M:%S",   # US format
        "%d-%m-%Y %H:%M:%S",
    ]

    s = series.astype(str).str.strip()

    # Primero intentar inferencia automática
    resultado = pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
    if resultado.notna().sum() / max(len(resultado), 1) > 0.8:
        return resultado

    # Intentar formatos uno a uno
    for fmt in formatos:
        try:
            parsed = pd.to_datetime(s, format=fmt, errors="coerce")
            ratio  = parsed.notna().sum() / max(len(parsed), 1)
            if ratio > 0.8:
                return parsed
        except Exception:
            continue

    # Fallback: inferencia sin formato
    return pd.to_datetime(s, errors="coerce")


def _parse_date_col(series: pd.Series) -> pd.Series:
    """Convierte una columna a date (sin hora), manejando duplicados y formatos varios."""
    if pd.api.types.is_datetime64_any_dtype(series):
        return series.dt.date

    s = series.astype(str).str.strip()

    formatos_fecha = [
        "%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y",
        "%d-%m-%Y", "%d %b %Y", "%Y/%m/%d",
    ]
    for fmt in formatos_fecha:
        try:
            parsed = pd.to_datetime(s, format=fmt, errors="coerce")
            if parsed.notna().sum() / max(len(parsed), 1) > 0.8:
                return parsed.dt.date
        except Exception:
            continue

    return pd.to_datetime(s, errors="coerce").dt.date


# ── Limpieza principal ────────────────────────────────────────────────────────

def limpiar_datos(df: pd.DataFrame) -> pd.DataFrame:

    # ── 1. Detectar columna DateTime ──────────────────────────────────────────
    col_dt = _find_col(df, ["DateTime", "Date Time", "Datetime", "DATETIME"])
    if col_dt is None:
        raise ValueError(
            "No se encontró la columna DateTime. "
            f"Columnas disponibles: {list(df.columns)}"
        )
    df["_datetime"] = _parse_datetime_col(df[col_dt])

    # ── 2. Detectar columna de fecha (para agrupar por día) ───────────────────
    # Preferencia: Fecha3 > Fecha 2 > Date > Date -1 > extraer de DateTime
    col_fecha = _find_col(df, ["Fecha3", "fecha3", "Fecha 3",
                                "Fecha2", "fecha2", "Fecha 2",
                                "Date", "DATE", "date",
                                "Date-1", "Date -1", "date1"])
    if col_fecha:
        df["_date"] = _parse_date_col(df[col_fecha])
    else:
        df["_date"] = df["_datetime"].dt.date

    # Si la columna de fecha tiene demasiados NaT, usar DateTime
    null_ratio = df["_date"].isna().sum() / max(len(df), 1)
    if null_ratio > 0.3:
        df["_date"] = df["_datetime"].dt.date

    # ── 3. Detectar ProductID2 ────────────────────────────────────────────────
    col_prod = _find_col(df, ["ProductID2", "productid2", "Product ID2",
                               "ProductID 2", "Productid2",
                               "ProductID",  "productid", "Product ID"])
    if col_prod is None:
        raise ValueError(
            "No se encontró la columna ProductID2 o ProductID. "
            f"Columnas disponibles: {list(df.columns)}"
        )
    df["_product"] = df[col_prod].astype(str).str.strip()

    # ── 4. Detectar Result ────────────────────────────────────────────────────
    col_res = _find_col(df, ["Result(FailedorPassed)", "Result", "RESULT",
                              "result", "ResultFailedorPassed",
                              "Result(Failed or Passed)"])
    if col_res:
        df["_result"] = (df[col_res].astype(str).str.strip()
                                    .str.capitalize()
                                    .map(lambda x: "Failed" if "fail" in x.lower() else "Passed"))
    else:
        df["_result"] = "Passed"

    # ── 5. Detectar Equipment ─────────────────────────────────────────────────
    col_eq = _find_col(df, ["Equipment", "EQUIPMENT", "equipment", "Equipo"])
    df["_equipment"] = df[col_eq].astype(str).str.strip() if col_eq else ""

    # ── 6. Detectar Shift ─────────────────────────────────────────────────────
    col_sh = _find_col(df, ["Shift", "SHIFT", "shift", "Turno"])
    df["_shift"] = df[col_sh].astype(str) if col_sh else ""

    # ── 7. Limpiar y ordenar ──────────────────────────────────────────────────
    df = df.dropna(subset=["_datetime", "_product"])
    df = df[df["_product"].str.len() > 0]
    df = df[df["_product"] != "nan"]
    df = df.sort_values("_datetime").reset_index(drop=True)

    return df


# ── Bloques ───────────────────────────────────────────────────────────────────

def asignar_bloques(df: pd.DataFrame) -> pd.DataFrame:
    """Nuevo bloque cada vez que cambia _product o _date."""
    cambia = (
        (df["_product"] != df["_product"].shift()) |
        (df["_date"]    != df["_date"].shift())
    )
    df["_bloque"] = cambia.cumsum()
    return df


# ── Cálculo por bloque ────────────────────────────────────────────────────────

def calcular_bloques(df: pd.DataFrame) -> pd.DataFrame:
    resultados = []

    for bloque_id, grupo in df.groupby("_bloque", sort=False):
        producto = grupo["_product"].iloc[0]
        fecha    = grupo["_date"].iloc[0]
        equipo   = grupo["_equipment"].iloc[0]
        turno    = grupo["_shift"].iloc[0]

        dt_min = grupo["_datetime"].min()
        dt_max = grupo["_datetime"].max()
        horas_totales = (dt_max - dt_min).total_seconds() / 3600

        passed = grupo[grupo["_result"] == "Passed"]
        failed = grupo[grupo["_result"] == "Failed"]

        qty_passed = len(passed)
        qty_failed = len(failed)

        if len(failed) >= 2:
            horas_muertas = (
                failed["_datetime"].max() - failed["_datetime"].min()
            ).total_seconds() / 3600
        else:
            horas_muertas = 0.0

        horas_reales = max(horas_totales - horas_muertas, 0.0)

        resultados.append({
            "Date":          fecha,
            "Equipment":     equipo,
            "Shift":         turno,
            "ProductID":     producto,
            "bloque_id":     bloque_id,
            "dt_inicio":     dt_min,
            "dt_fin":        dt_max,
            "horas_totales": round(horas_totales, 4),
            "horas_muertas": round(horas_muertas, 4),
            "horas_reales":  round(horas_reales,  4),
            "qty_passed":    qty_passed,
            "qty_failed":    qty_failed,
        })

    return pd.DataFrame(resultados)


# ── Consolidar ────────────────────────────────────────────────────────────────

def consolidar_reporte(bloques: pd.DataFrame) -> pd.DataFrame:
    reporte = (
        bloques
        .groupby(["Date", "Equipment", "ProductID"], as_index=False)
        .agg(
            horas_totales=("horas_totales", "sum"),
            horas_muertas=("horas_muertas", "sum"),
            horas_reales =("horas_reales",  "sum"),
            qty_passed   =("qty_passed",    "sum"),
            qty_failed   =("qty_failed",    "sum"),
            corridas     =("bloque_id",     "count"),
        )
        .sort_values(["Date", "ProductID"])
        .reset_index(drop=True)
    )
    reporte["rendimiento_%"] = (
        reporte["horas_reales"]
        / reporte["horas_totales"].replace(0, np.nan)
        * 100
    ).round(1).fillna(0)

    return reporte


# ── Exportar Excel ────────────────────────────────────────────────────────────

def exportar_excel(reporte: pd.DataFrame, bloques: pd.DataFrame) -> bytes:
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book

        fmt_header = wb.add_format({
            "bold": True, "bg_color": "#185FA5", "font_color": "#FFFFFF",
            "border": 1, "align": "center", "valign": "vcenter",
            "font_name": "Arial", "font_size": 10,
        })
        fmt_date = wb.add_format({"num_format": "dd/mm/yyyy",            "font_name": "Arial", "font_size": 10})
        fmt_dt   = wb.add_format({"num_format": "dd/mm/yyyy hh:mm:ss",   "font_name": "Arial", "font_size": 10})
        fmt_num  = wb.add_format({"num_format": "#,##0",                 "font_name": "Arial", "font_size": 10})
        fmt_dec  = wb.add_format({"num_format": "0.000",                 "font_name": "Arial", "font_size": 10})
        fmt_pct  = wb.add_format({"num_format": "0.0\"%\"",              "font_name": "Arial", "font_size": 10})
        fmt_cell = wb.add_format({"font_name": "Arial", "font_size": 10})

        def escribir_hoja(ws_name, data, col_defs):
            data.to_excel(writer, sheet_name=ws_name, index=False, startrow=1)
            ws = writer.sheets[ws_name]
            for i, (key, header, width, fmt) in enumerate(col_defs):
                ws.write(0, i, header, fmt_header)
                ws.set_column(i, i, width)
                for r, val in enumerate(data[key] if key in data.columns else [""]*len(data)):
                    ws.write(r + 2, i, val, fmt)
            ws.freeze_panes(2, 0)
            ws.autofilter(1, 0, len(data) + 1, len(col_defs) - 1)

        # Hoja Resumen
        escribir_hoja("Resumen", reporte, [
            ("Date",          "Fecha",              12, fmt_date),
            ("Equipment",     "Equipo",             10, fmt_cell),
            ("ProductID",     "Modelo (ProductID)", 30, fmt_cell),
            ("corridas",      "# Corridas",         10, fmt_num),
            ("qty_passed",    "Pzas OK",            10, fmt_num),
            ("qty_failed",    "Pzas Fallidas",      13, fmt_num),
            ("horas_totales", "Horas Totales",      14, fmt_dec),
            ("horas_muertas", "Tiempo Muerto (h)",  16, fmt_dec),
            ("horas_reales",  "Horas Reales",       14, fmt_dec),
            ("rendimiento_%", "Rendimiento %",      14, fmt_pct),
        ])

        # Hoja Detalle
        escribir_hoja("Detalle Bloques", bloques, [
            ("Date",          "Fecha",        12, fmt_date),
            ("Equipment",     "Equipo",       10, fmt_cell),
            ("ProductID",     "Modelo",       30, fmt_cell),
            ("bloque_id",     "Bloque #",      9, fmt_num),
            ("dt_inicio",     "Inicio",       20, fmt_dt),
            ("dt_fin",        "Fin",          20, fmt_dt),
            ("horas_totales", "Horas Bloque", 13, fmt_dec),
            ("horas_muertas", "T. Muerto",    12, fmt_dec),
            ("horas_reales",  "Horas Reales", 13, fmt_dec),
            ("qty_passed",    "Pzas OK",       9, fmt_num),
            ("qty_failed",    "Pzas Falla",   10, fmt_num),
        ])

    return output.getvalue()


# ── Entrada pública ───────────────────────────────────────────────────────────

def procesar_excel(file_bytes: bytes) -> dict:
    df        = leer_excel(file_bytes)
    df        = limpiar_datos(df)
    df        = asignar_bloques(df)
    bloques   = calcular_bloques(df)
    reporte   = consolidar_reporte(bloques)
    excel_out = exportar_excel(reporte, bloques)

    resumen = {
        "total_registros": len(df),
        "total_passed":    int((df["_result"] == "Passed").sum()),
        "total_failed":    int((df["_result"] == "Failed").sum()),
        "horas_totales":   round(float(reporte["horas_totales"].sum()), 3),
        "horas_muertas":   round(float(reporte["horas_muertas"].sum()), 3),
        "horas_reales":    round(float(reporte["horas_reales"].sum()),  3),
        "modelos":         int(reporte["ProductID"].nunique()),
        "fechas":          sorted([str(d) for d in reporte["Date"].unique()]),
    }

    return {
        "reporte":   reporte,
        "bloques":   bloques,
        "excel_out": excel_out,
        "resumen":   resumen,
    }
