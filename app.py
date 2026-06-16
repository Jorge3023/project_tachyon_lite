import io
import pandas as pd

def leer_csv(file_bytes):

try:
    return pd.read_csv(
        io.BytesIO(file_bytes),
        encoding="utf-8"
    )
except Exception:
    return pd.read_csv(
        io.BytesIO(file_bytes),
        encoding="latin1"
    )

def procesar_datos(df):

df.columns = [str(c).strip() for c in df.columns]

if "Job" not in df.columns:
    raise ValueError("No se encontró la columna Job")

if "Panel" not in df.columns:
    raise ValueError("No se encontró la columna Panel")

if "End time" not in df.columns:
    raise ValueError("No se encontró la columna End time")

resultados = []

modelo_actual = None

paneles = []

primera_fecha = None
ultima_fecha = None

fin_modelo_anterior = None

for _, row in df.iterrows():

    job = ""

    if pd.notna(row["Job"]):
        job = str(row["Job"]).strip()

    if job and job.lower() != "nan":

        if modelo_actual and paneles:

            primer_panel = min(paneles)
            ultimo_panel = max(paneles)

            piezas = (ultimo_panel - primer_panel) + 1

            if fin_modelo_anterior is None:
                inicio = primera_fecha
            else:
                inicio = fin_modelo_anterior

            fin = ultima_fecha

            duracion = fin - inicio

            horas = int(duracion.total_seconds() // 3600)

            minutos = int(
                (duracion.total_seconds() % 3600) // 60
            )

            resultados.append({
                "Modelo": modelo_actual,
                "Horas Trabajadas": horas,
                "Minutos Trabajados": minutos,
                "Piezas": piezas
            })

            fin_modelo_anterior = ultima_fecha

        modelo_actual = job

        paneles = []

        primera_fecha = None
        ultima_fecha = None

        continue

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

        if primera_fecha is None:
            primera_fecha = fecha

        ultima_fecha = fecha

    except Exception:
        continue

if modelo_actual and paneles:

    primer_panel = min(paneles)

    ultimo_panel = max(paneles)

    piezas = (ultimo_panel - primer_panel) + 1

    if fin_modelo_anterior is None:
        inicio = primera_fecha
    else:
        inicio = fin_modelo_anterior

    fin = ultima_fecha

    duracion = fin - inicio

    horas = int(duracion.total_seconds() // 3600)

    minutos = int(
        (duracion.total_seconds() % 3600) // 60
    )

    resultados.append({
        "Modelo": modelo_actual,
        "Horas Trabajadas": horas,
        "Minutos Trabajados": minutos,
        "Piezas": piezas
    })

return pd.DataFrame(resultados)

def exportar_excel(reporte):

output = io.BytesIO()

with pd.ExcelWriter(
    output,
    engine="xlsxwriter"
) as writer:

    reporte.to_excel(
        writer,
        sheet_name="Resumen",
        index=False
    )

    ws = writer.sheets["Resumen"]

    ws.set_column("A:A", 45)
    ws.set_column("B:D", 20)

return output.getvalue()

def procesar_excel(file_bytes):

df = leer_csv(file_bytes)

reporte = procesar_datos(df)

excel_out = exportar_excel(reporte)

resumen = {
    "modelos": int(len(reporte)),
    "piezas_totales": int(reporte["Piezas"].sum()),
    "horas_totales": int(reporte["Horas Trabajadas"].sum()),
    "minutos_totales": int(reporte["Minutos Trabajados"].sum())
}

return {
    "reporte": reporte,
    "excel_out": excel_out,
    "resumen": resumen
}