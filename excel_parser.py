import io
import pandas as pd

def leer_archivo(file_bytes, nombre_archivo):

nombre = nombre_archivo.lower()

if nombre.endswith(".csv"):

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

elif nombre.endswith((".xlsx", ".xls")):

    return pd.read_excel(
        io.BytesIO(file_bytes)
    )

raise ValueError(
    "Formato no soportado. Utiliza CSV, XLSX o XLS."
)

def procesar_datos(df):

df.columns = [
    str(c).strip()
    for c in df.columns
]

if "Job" not in df.columns:
    raise ValueError(
        "No se encontró la columna Job"
    )

if "Panel" not in df.columns:
    raise ValueError(
        "No se encontró la columna Panel"
    )

if "End time" not in df.columns:
    raise ValueError(
        "No se encontró la columna End time"
    )

resultados = []

modelo_actual = None

primera_fecha = None
ultima_fecha = None

piezas_modelo = 0

fin_modelo_anterior = None

for _, row in df.iterrows():

    job = ""

    if pd.notna(row["Job"]):
        job = str(row["Job"]).strip()

    if job and job.lower() != "nan":

        if modelo_actual is not None:

            if (
                primera_fecha is not None
                and ultima_fecha is not None
            ):

                inicio = (
                    fin_modelo_anterior
                    if fin_modelo_anterior is not None
                    else primera_fecha
                )

                duracion = (
                    ultima_fecha - inicio
                )

                total_min = int(
                    duracion.total_seconds() / 60
                )

                horas = total_min // 60
                minutos = total_min % 60

                resultados.append({
                    "Modelo": modelo_actual,
                    "Horas Trabajadas": horas,
                    "Minutos Trabajados": minutos,
                    "Piezas": piezas_modelo
                })

                fin_modelo_anterior = (
                    ultima_fecha
                )

        modelo_actual = job

        primera_fecha = None
        ultima_fecha = None
        piezas_modelo = 0

        continue

    try:

        panel = row["Panel"]

        if pd.isna(panel):
            continue

        fecha = pd.to_datetime(
            row["End time"],
            errors="coerce",
            dayfirst=True
        )

        if pd.isna(fecha):
            continue

        piezas_modelo += 1

        if primera_fecha is None:
            primera_fecha = fecha

        ultima_fecha = fecha

    except Exception:
        continue

if (
    modelo_actual is not None
    and primera_fecha is not None
    and ultima_fecha is not None
):

    inicio = (
        fin_modelo_anterior
        if fin_modelo_anterior is not None
        else primera_fecha
    )

    duracion = (
        ultima_fecha - inicio
    )

    total_min = int(
        duracion.total_seconds() / 60
    )

    horas = total_min // 60
    minutos = total_min % 60

    resultados.append({
        "Modelo": modelo_actual,
        "Horas Trabajadas": horas,
        "Minutos Trabajados": minutos,
        "Piezas": piezas_modelo
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

    ws.set_column("A:A", 40)
    ws.set_column("B:D", 20)

return output.getvalue()

def procesar_excel(
file_bytes,
nombre_archivo
):

df = leer_archivo(
    file_bytes,
    nombre_archivo
)

reporte = procesar_datos(df)

excel_out = exportar_excel(
    reporte
)

resumen = {
    "modelos": int(len(reporte)),
    "piezas_totales": int(
        reporte["Piezas"].sum()
    ) if not reporte.empty else 0,
    "horas_totales": int(
        reporte["Horas Trabajadas"].sum()
    ) if not reporte.empty else 0,
    "minutos_totales": int(
        reporte["Minutos Trabajados"].sum()
    ) if not reporte.empty else 0
}

return {
    "reporte": reporte,
    "excel_out": excel_out,
    "resumen": resumen
}
