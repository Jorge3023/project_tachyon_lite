"""
Parser de reportes de producción — Project Tachyon Lite.

Reglas de negocio:
  - El modelo se obtiene EXCLUSIVAMENTE de la columna "Job".
  - El modelo aparece una sola vez al inicio de cada bloque; las filas
    siguientes (sin valor en Job) pertenecen a ese mismo modelo hasta
    que aparece un nuevo valor en Job.
  - El valor de "Panel" en la fila donde aparece el modelo (ej. "182")
    es el conteo ACUMULADO de turnos/horas previas al archivo actual.
    NO se usa para ningún cálculo y tampoco representa una pieza nueva.
  - La PRIMERA fila de cada bloque (la que contiene el valor en Job) es
    solo metadata del acumulado anterior — NO cuenta como pieza producida.
    El conteo de piezas reales empieza desde la SEGUNDA fila del bloque.
  - Cada fila de producción (a partir de la segunda del bloque) = 1 pieza,
    sin importar el valor de Panel (no se usa último-primero+1).
  - La fecha/hora se obtiene EXCLUSIVAMENTE de "End time"
    (formato esperado: "06/06/2026 09:44:14 a. m.").
  - El "End time" de la primera fila del bloque SÍ sigue siendo válido
    como referencia temporal, aunque esa fila no cuente como pieza.
  - Columnas ignoradas por completo: Time(s), EndTo1st, EndTo1stEnd,
    EndTo2nd, EndTo2ndEnd, EndTo3rd, EndTo3rdEnd.
  - Tiempo trabajado de un modelo = desde la ÚLTIMA fila (por End time)
    del modelo ANTERIOR hasta la ÚLTIMA fila del modelo ACTUAL.
  - Para el PRIMER modelo del archivo: el tiempo trabajado inicia a las
    7:00 AM del día de su última fila (inicio de turno), no en su
    primera pieza real.
"""

import pandas as pd
import numpy as np
import io
import re
from datetime import datetime, time


# ── Columnas que deben ignorarse explícitamente ───────────────────────────────
COLUMNAS_IGNORADAS = {
    "time(s)", "endto1st", "endto1stend",
    "endto2nd", "endto2ndend", "endto3rd", "endto3rdend",
}

HORA_INICIO_TURNO = time(7, 0, 0)   # 7:00 AM


# ── Lectura del archivo (csv, xlsx, xls) ──────────────────────────────────────

def leer_archivo(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """Lee CSV o Excel, detectando el formato por extensión y por contenido."""
    nombre = (filename or "").lower()
    buffer = io.BytesIO(file_bytes)
    errores = []

    # ── CSV ──
    if nombre.endswith(".csv"):
        for sep in (",", ";", "\t"):
            for enc in ("utf-8-sig", "utf-8", "latin-1"):
                try:
                    buffer.seek(0)
                    df = pd.read_csv(buffer, sep=sep, encoding=enc, engine="python")
                    if df.shape[1] > 1:        # separador correcto si hay >1 columna
                        df.columns = [str(c).strip() for c in df.columns]
                        return df
                except Exception as exc:
                    errores.append(f"csv(sep={sep!r},enc={enc}): {exc}")
        raise ValueError(
            "No se pudo leer el archivo CSV. Verifica el delimitador y codificación. "
            f"Detalle: {' | '.join(errores[:3])}"
        )

    # ── Excel (.xlsx / .xls) ──
    for engine in (None, "openpyxl", "xlrd"):
        try:
            buffer.seek(0)
            xl = pd.ExcelFile(buffer, engine=engine) if engine else pd.ExcelFile(buffer)
            df = xl.parse(xl.sheet_names[0], header=0)

            # Deduplicar nombres de columnas repetidos
            cols, seen = [], {}
            for c in df.columns:
                c = str(c).strip()
                seen[c] = seen.get(c, 0)
                cols.append(c if seen[c] == 0 else f"{c}__dup{seen[c]}")
                seen[c] += 1
            df.columns = cols
            return df

        except Exception as exc:
            errores.append(f"{engine or 'auto'}: {exc}")

    raise ValueError(
        "No se pudo leer el archivo. Verifica que no esté dañado y sea .csv, .xlsx o .xls válido. "
        f"Detalle: {' | '.join(errores)}"
    )


# ── Detección flexible de columnas ────────────────────────────────────────────

def _find_col(df: pd.DataFrame, *candidates: str):
    normalize = lambda s: re.sub(r"[\s_\-\.\(\)]", "", s).lower()
    index = {normalize(c): c for c in df.columns}
    for cand in candidates:
        hit = index.get(normalize(cand))
        if hit:
            return hit
    return None


def _safe_series(df: pd.DataFrame, col: str) -> pd.Series:
    s = df[col]
    return s.iloc[:, 0] if isinstance(s, pd.DataFrame) else s


def _eliminar_columnas_ignoradas(df: pd.DataFrame) -> pd.DataFrame:
    normalize = lambda s: re.sub(r"[\s_\-\.\(\)]", "", str(s)).lower()
    cols_a_quitar = [c for c in df.columns if normalize(c) in COLUMNAS_IGNORADAS]
    if cols_a_quitar:
        df = df.drop(columns=cols_a_quitar)
    return df


# ── Parseo de "End time" ──────────────────────────────────────────────────────
# Formato esperado: 06/06/2026 09:44:14 a. m.  (también soporta p. m., AM/PM, etc.)

def _normalizar_ampm(texto: str) -> str:
    """Convierte variantes de a.m./p.m. en español a AM/PM estándar para strptime."""
    t = texto.strip()
    t = re.sub(r"a\.?\s*m\.?", "AM", t, flags=re.IGNORECASE)
    t = re.sub(r"p\.?\s*m\.?", "PM", t, flags=re.IGNORECASE)
    return t


def _parse_end_time(series: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(series):
        return series

    s = series.astype(str).str.strip().map(_normalizar_ampm)

    formatos = [
        "%d/%m/%Y %I:%M:%S %p",
        "%m/%d/%Y %I:%M:%S %p",
        "%d/%m/%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%d-%m-%Y %I:%M:%S %p",
    ]

    # Intento de inferencia automática primero
    resultado = pd.to_datetime(s, errors="coerce")
    if resultado.notna().mean() > 0.85:
        return resultado

    for fmt in formatos:
        try:
            parsed = pd.to_datetime(s, format=fmt, errors="coerce")
            if parsed.notna().mean() > 0.85:
                return parsed
        except Exception:
            continue

    return pd.to_datetime(s, errors="coerce")


# ── Extraer el nombre de modelo desde "Job" ───────────────────────────────────
# El número que acompaña al modelo (ej. "182") NO debe usarse para cálculos,
# solo se conserva el texto del modelo tal como aparece, sin alterarlo.

def _limpiar_job(valor) -> str | None:
    if pd.isna(valor):
        return None
    texto = str(valor).strip()
    if texto == "" or texto.lower() == "nan":
        return None
    return texto


# ── Limpieza principal ────────────────────────────────────────────────────────

def limpiar_datos(df: pd.DataFrame) -> pd.DataFrame:

    df = _eliminar_columnas_ignoradas(df)

    # ── Job (modelo) ──────────────────────────────────────────────────────────
    col_job = _find_col(df, "Job", "JOB", "job")
    if col_job is None:
        raise ValueError(
            f"No se encontró la columna 'Job'. Columnas disponibles: {list(df.columns)}"
        )
    df["_job_raw"] = _safe_series(df, col_job).map(_limpiar_job)

    # ── End time ──────────────────────────────────────────────────────────────
    col_end = _find_col(df, "End time", "EndTime", "End Time", "endtime")
    if col_end is None:
        raise ValueError(
            f"No se encontró la columna 'End time'. Columnas disponibles: {list(df.columns)}"
        )
    df["_datetime"] = _parse_end_time(_safe_series(df, col_end))

    # Descartar filas sin fecha válida (no se puede ubicar en el tiempo)
    df = df.dropna(subset=["_datetime"]).reset_index(drop=True)

    if len(df) == 0:
        raise ValueError(
            "No se encontraron registros con 'End time' válido. "
            "Verifica el formato de fecha (ej: 06/06/2026 09:44:14 a. m.)."
        )

    # ── Propagar el modelo hacia adelante ────────────────────────────────────
    # El modelo aparece una sola vez al inicio del bloque; las filas siguientes
    # (sin valor en Job) pertenecen a ese mismo modelo.
    df["_modelo"] = df["_job_raw"].ffill()

    # Si las primeras filas no tienen modelo asignado (antes del primer Job), descartarlas
    df = df.dropna(subset=["_modelo"]).reset_index(drop=True)

    if len(df) == 0:
        raise ValueError(
            "No se pudo asociar ningún registro a un modelo (columna 'Job' vacía en todo el archivo)."
        )

    return df


# ── Bloques (nuevo bloque cada vez que cambia el modelo) ──────────────────────

def asignar_bloques(df: pd.DataFrame) -> pd.DataFrame:
    cambia = df["_modelo"] != df["_modelo"].shift()
    df["_bloque"] = cambia.cumsum()
    return df


# ── Cálculo de tiempo trabajado y piezas por bloque ───────────────────────────

def calcular_bloques(df: pd.DataFrame) -> pd.DataFrame:
    """
    Para cada bloque (modelo consecutivo):
      - piezas = número de filas del bloque SIN CONTAR la primera fila
        (la que contiene el valor en Job). Esa primera fila es solo
        metadata del acumulado de turnos anteriores, no una pieza
        producida en este archivo. El conteo real inicia desde la
        segunda fila del bloque en adelante.
      - tiempo trabajado = última pieza del bloque ANTERIOR → última pieza de este bloque
      - el PRIMER bloque del archivo usa las 7:00 AM del día de su última pieza
        como punto de partida (inicio de turno), no su primera pieza.
    """
    bloques_ordenados = sorted(df["_bloque"].unique())
    resultados = []

    referencia_anterior = None   # datetime de la última pieza del bloque previo

    for idx, bloque_id in enumerate(bloques_ordenados):
        grupo   = df[df["_bloque"] == bloque_id]
        modelo  = grupo["_modelo"].iloc[0]

        # La primera fila del bloque (donde aparece Job) es metadata, no pieza
        piezas  = max(len(grupo) - 1, 0)

        ultima_pieza = grupo["_datetime"].max()

        if idx == 0:
            # Primer modelo del archivo → referencia = 7:00 AM del día de su última pieza
            dia = ultima_pieza.date()
            inicio_referencia = datetime.combine(dia, HORA_INICIO_TURNO)
        else:
            inicio_referencia = referencia_anterior

        delta = ultima_pieza - inicio_referencia
        segundos_trabajados = max(delta.total_seconds(), 0.0)

        horas_trabajadas   = int(segundos_trabajados // 3600)
        minutos_totales    = int(segundos_trabajados // 60)   # total convertido a minutos

        resultados.append({
            "Modelo":             modelo,
            "_bloque_id":         bloque_id,
            "_inicio_referencia": inicio_referencia,
            "_fin_bloque":        ultima_pieza,
            "Horas Trabajadas":   horas_trabajadas,
            "Minutos Trabajados": minutos_totales,
            "Piezas":             piezas,
            "_horas_decimal":     round(segundos_trabajados / 3600, 4),
        })

        referencia_anterior = ultima_pieza

    return pd.DataFrame(resultados)


# ── Consolidar por modelo (si el mismo modelo se repite en varios bloques) ────

def consolidar_reporte(bloques: pd.DataFrame) -> pd.DataFrame:
    """
    Si un modelo aparece en más de un bloque no consecutivo (vuelve a
    correr después de otro modelo), se suman sus piezas y tiempo trabajado.
    """
    reporte = (
        bloques
        .groupby("Modelo", as_index=False)
        .agg(
            **{
                "_horas_decimal_sum": ("_horas_decimal", "sum"),
                "Piezas":             ("Piezas", "sum"),
                "_corridas":          ("_bloque_id", "count"),
            }
        )
    )

    # Horas Trabajadas = parte entera de horas; Minutos Trabajados = TOTAL en minutos
    reporte["Horas Trabajadas"]   = reporte["_horas_decimal_sum"].astype(int)
    reporte["Minutos Trabajados"] = (reporte["_horas_decimal_sum"] * 60).round().astype(int)
    reporte = reporte.drop(columns=["_horas_decimal_sum"])

    # Mantener el orden de primera aparición del modelo en el archivo original
    orden = bloques.drop_duplicates("Modelo")["Modelo"].tolist()
    reporte["_orden"] = reporte["Modelo"].map({m: i for i, m in enumerate(orden)})
    reporte = reporte.sort_values("_orden").drop(columns=["_orden", "_corridas"]).reset_index(drop=True)

    return reporte[["Modelo", "Horas Trabajadas", "Minutos Trabajados", "Piezas"]]


# ── Exportar Excel de salida (solo 4 columnas requeridas) ─────────────────────

def exportar_excel(reporte: pd.DataFrame) -> bytes:
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book

        fmt_header = wb.add_format({
            "bold": True, "bg_color": "#185FA5", "font_color": "#FFFFFF",
            "border": 1, "align": "center", "valign": "vcenter",
            "font_name": "Arial", "font_size": 10,
        })
        fmt_cell = wb.add_format({"font_name": "Arial", "font_size": 10, "align": "center", "valign": "vcenter"})
        fmt_num  = wb.add_format({"num_format": "#,##0", "font_name": "Arial", "font_size": 10, "align": "center", "valign": "vcenter"})

        columnas = [
            ("Modelo",             "Modelo",             32, fmt_cell),
            ("Horas Trabajadas",   "Horas Trabajadas",   16, fmt_num),
            ("Minutos Trabajados", "Minutos Trabajados", 18, fmt_num),
            ("Piezas",             "Piezas",             12, fmt_num),
        ]

        reporte.to_excel(writer, sheet_name="Resultado", index=False, startrow=1, header=False)
        ws = writer.sheets["Resultado"]

        for i, (key, header, width, fmt) in enumerate(columnas):
            ws.write(0, i, header, fmt_header)
            ws.set_column(i, i, width)

        # Aplicar formato numérico a las columnas (encabezado ya escrito arriba)
        for r in range(len(reporte)):
            for i, (key, _, _, fmt) in enumerate(columnas):
                ws.write(r + 2, i, reporte.iloc[r][key], fmt)

        ws.freeze_panes(2, 0)
        ws.autofilter(1, 0, len(reporte) + 1, len(columnas) - 1)

    return output.getvalue()


# ── Entrada pública ───────────────────────────────────────────────────────────

def procesar_archivo(file_bytes: bytes, filename: str) -> dict:
    df      = leer_archivo(file_bytes, filename)
    df      = limpiar_datos(df)
    df      = asignar_bloques(df)
    bloques = calcular_bloques(df)
    reporte = consolidar_reporte(bloques)

    excel_out = exportar_excel(reporte)

    resumen = {
        "total_registros":   len(df),
        "total_modelos":     int(reporte["Modelo"].nunique()),
        "total_piezas":      int(reporte["Piezas"].sum()),
        "horas_trabajadas":  int(reporte["Horas Trabajadas"].sum()),
        "rango_fechas": sorted({
            str(d.date()) for d in [df["_datetime"].min(), df["_datetime"].max()]
        }),
    }

    return {
        "reporte":   reporte,
        "excel_out": excel_out,
        "resumen":   resumen,
    }


# Mantener compatibilidad con el nombre anterior usado por app.py
def procesar_excel(file_bytes: bytes, filename: str = "archivo.xlsx") -> dict:
    return procesar_archivo(file_bytes, filename)
