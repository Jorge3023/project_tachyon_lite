import io
import pandas as pd


# ─────────────────────────────────────────────────────────────
# Leer CSV
# ─────────────────────────────────────────────────────────────
def leer_archivo(file_bytes, nombre_archivo):

    nombre = nombre_archivo.lower()

    if nombre.endswith(".csv"):

        try:
            return pd.read_csv(
                io.BytesIO(file_bytes),
                encoding="utf-8"
            )
        except:
            return pd.read_csv(
                io.BytesIO(file_bytes),
                encoding="latin1"
            )

    elif nombre.endswith((".xlsx", ".xls")):

        return pd.read_excel(
            io.BytesIO(file_bytes)
        )

    raise ValueError("Formato no soportado")


# ─────────────────────────────────────────────────────────────
# Procesar archivo
# ─────────────────────────────────────────────────────────────

def procesar_datos(df: pd.DataFrame):

    columnas = [str(c).strip() for c in df.columns]
    df.columns = columnas

    if "Job" not in df.columns:
        raise ValueError("No se encontró la columna 'Job'")

    if "Panel" not in df.columns:
        raise ValueError("No se encontró la columna 'Panel'")

    if "End time" not in df.columns:
        raise ValueError("No se encontró la columna 'End time'")

    resultados = []

    modelo_actual = None
    paneles = []

    primera_pieza = None
    ultima_pieza = None

    fin_modelo_anterior = None

    for _, row in df.iterrows():

        job = str(row["Job"]).strip() if pd.notna(row["Job"]) else ""

        # -----------------------------------------------------
        # Cambio de modelo
        # -----------------------------------------------------
        if job != "" and job.lower() != "nan":

            if modelo_actual is not None and len(paneles) > 0:

                primer_panel = min(paneles)
                ultimo_panel = max(paneles)

                piezas = (ultimo_panel - primer_panel) + 1

                if fin_modelo_anterior is None:
                    inicio = primera_pieza
                else:
                    inicio = fin_modelo_anterior

                fin = ultima_pieza

                duracion = fin - inicio

                horas = int(duracion.total_seconds() // 3600)
                minutos = int((duracion.total_seconds() % 3600) // 60)

                resultados.append({
                    "Modelo": modelo_actual,
                    "Horas Trabajadas": horas,
                    "Minutos Trabajados": minutos,
                    "Piezas": piezas
                })

                fin_modelo_anterior = ultima_pieza

            modelo_actual = job

            paneles = []
            primera_pieza = None
            ultima_pieza = None

            continue

        # -----------------------------------------------------
        # Filas de producción
        # -----------------------------------------------------
        try:

            panel = int(float(row["Panel"]))

            fecha = pd.to_datetime(
                row["End time"],
                dayfirst=True,
                errors="coerce"
            )

            if pd.isna(fecha):
                continue

            paneles.append(panel)

            if primera_pieza is None:
                primera_pieza = fecha

            ultima_pieza = fecha

        except Exception:
            continue

    # ---------------------------------------------------------
    # Último modelo
    # ---------------------------------------------------------
    if modelo_actual is not None and len(paneles) > 0:

        primer_panel = min(paneles)
        ultimo_panel = max(paneles)

        piezas = (ultimo_panel - primer_panel) + 1

        if fin_modelo_anterior is None:
            inicio = primera_pieza
        else:
            inicio = fin_modelo_anterior

        fin = ultima_pieza

        duracion = fin - inicio

        horas = int(duracion.total_seconds() // 3600)
        minutos = int((duracion.total_seconds() % 3600) // 60)

        resultados.append({
            "Modelo": modelo_actual,
            "Horas Trabajadas": horas,
            "Minutos Trabajados": minutos,
            "Piezas": piezas
        })

    return pd.DataFrame(resultados)


# ─────────────────────────────────────────────────────────────
# Exportar Excel
# ─────────────────────────────────────────────────────────────

def exportar_excel(reporte: pd.DataFrame) -> bytes:

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

        reporte.to_excel(
            writer,
            sheet_name="Resumen",
            index=False
        )

        workbook = writer.book
        worksheet = writer.sheets["Resumen"]

        header_format = workbook.add_format({
            "bold": True,
            "bg_color": "#185FA5",
            "font_color": "#FFFFFF"
        })

        for col_num, value in enumerate(reporte.columns):
            worksheet.write(0, col_num, value, header_format)

        worksheet.set_column("A:A", 40)
        worksheet.set_column("B:D", 20)

    return output.getvalue()


# ─────────────────────────────────────────────────────────────
# Función pública
# ─────────────────────────────────────────────────────────────

def procesar_excel(file_bytes: bytes):

    df = leer_csv(file_bytes)

    reporte = procesar_datos(df)

    excel_out = exportar_excel(reporte)

    resumen = {
        "modelos": len(reporte),
        "piezas_totales": int(reporte["Piezas"].sum()),
        "horas_totales": int(reporte["Horas Trabajadas"].sum()),
        "minutos_totales": int(reporte["Minutos Trabajados"].sum())
    }

    return {
        "reporte": reporte,
        "excel_out": excel_out,
        "resumen": resumen
    }