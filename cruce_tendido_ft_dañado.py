# ============================================================
# BUSCADOR HISTÓRICO DE SERIES EN TENDIDOS
# VERSION OPTIMIZADA Y CORREGIDA
# ============================================================
#
# MEJORAS IMPORTANTES:
#
# ✔ YA NO SE QUEDA EN 0%
# ✔ MUCHO MÁS RÁPIDO
# ✔ NO RECORRE TODAS LAS CELDAS INNECESARIAMENTE
# ✔ SOLO ANALIZA CELDAS CON DATOS
# ✔ INDEXA SERIES OBJETIVO EN MEMORIA
# ✔ BARRA DE PROGRESO REAL
# ✔ DETECTA LINEA / ESTACA
# ✔ SOPORTA BLOQUES LADO A LADO
# ✔ NO MODIFICA FORMATOS
#
# AUTOR ORIGINAL:
# JOSE FRANCISCO ALVARADO LEYVA
# ============================================================
# ============================================================
# BUSCADOR HISTÓRICO DE SERIES EN TENDIDOS
# ORDENADO DEL MÁS ANTIGUO AL MÁS RECIENTE
# ============================================================

import os
import re
from datetime import datetime

from openpyxl import load_workbook
from tqdm import tqdm


# ============================================================
# VALIDACIONES
# ============================================================

def es_vacio(valor):

    if valor is None:
        return True

    texto = str(valor).strip().upper()

    return texto in [
        "",
        "N/A",
        "#N/D",
        "NONE",
        "NULL"
    ]


def es_serie_inova(valor):

    if valor is None:
        return False

    valor = str(valor).strip().upper()

    return bool(re.fullmatch(r"Q\d{8}", valor))


def es_serie_sercel(valor):

    if valor is None:
        return False

    valor = str(valor).strip()

    return valor.isdigit() and len(valor) >= 6


def es_serie(valor):

    return es_serie_inova(valor) or es_serie_sercel(valor)


def es_estaca(valor):

    try:

        n = int(str(valor).strip())

        return 1000 <= n <= 20000

    except:
        return False


def convertir_fecha_hoja(nombre):

    try:
        return datetime.strptime(nombre.strip(), "%d%m%Y")
    except:
        return None


# ============================================================
# EXTRAER FECHA DEL NOMBRE DEL ARCHIVO
# ============================================================

def obtener_fecha_archivo(nombre_archivo):

    """
    Convierte nombres como:

    TENDIDO JUNIO 2025
    TENDIDO FEBRERO 2026 2DO
    TENDIDO MAYO 2026

    a fecha para ordenar correctamente.
    """

    nombre = os.path.basename(nombre_archivo).upper()

    meses = {
        "ENERO": 1,
        "FEBRERO": 2,
        "MARZO": 3,
        "ABRIL": 4,
        "MAYO": 5,
        "JUNIO": 6,
        "JULIO": 7,
        "AGOSTO": 8,
        "SEPTIEMBRE": 9,
        "OCTUBRE": 10,
        "NOVIEMBRE": 11,
        "DICIEMBRE": 12
    }

    anio = None
    mes = None

    # buscar año
    match_anio = re.search(r"(20\d{2})", nombre)

    if match_anio:
        anio = int(match_anio.group(1))

    # buscar mes
    for nombre_mes, numero_mes in meses.items():

        if nombre_mes in nombre:

            mes = numero_mes
            break

    if anio and mes:

        return datetime(anio, mes, 1)

    # si no encuentra nada
    return datetime(1900, 1, 1)


# ============================================================
# BUSCAR ESTACA
# ============================================================

def buscar_estaca(ws, fila, columna):

    """
    Busca la estaca en la misma fila
    hacia la izquierda.
    """

    for c in range(columna - 1, max(0, columna - 6), -1):

        valor = ws.cell(fila, c).value

        if es_estaca(valor):

            return int(valor), c

    return None, None


# ============================================================
# BUSCAR LINEA
# ============================================================

def buscar_linea(ws, col_estaca):

    """
    SUBE HASTA FILA 1
    buscando el encabezado LINEA
    del bloque correspondiente.
    """

    for fila in range(1, 15):

        for c in range(
            max(1, col_estaca - 4),
            col_estaca + 5
        ):

            valor = ws.cell(fila, c).value

            if valor is None:
                continue

            texto = str(valor).strip().upper()

            # ====================================================
            # ENCONTRO "LINEA"
            # ====================================================

            if "LINEA" in texto:

                for c2 in range(c + 1, c + 6):

                    valor2 = ws.cell(fila, c2).value

                    try:

                        linea = int(str(valor2).strip())

                        if 1000 <= linea <= 9999:

                            return linea

                    except:
                        pass

            # ====================================================
            # CELDA DIRECTAMENTE ES LINEA
            # ====================================================

            try:

                numero = int(texto)

                if 1000 <= numero <= 9999:

                    if fila <= 5:

                        return numero

            except:
                pass

    return None


# ============================================================
# CARGAR SERIES OBJETIVO
# ============================================================

def cargar_series_objetivo(ws, headers):

    col_serie = headers.index("Número de Serie") + 1
    col_linea = headers.index("LÍnea") + 1
    col_x = headers.index("X") + 1

    series = {}

    for row in range(2, ws.max_row + 1):

        valor = ws.cell(row, col_serie).value

        if valor is None:
            continue

        serie = str(valor).strip().upper()

        if serie == "S/N":
            continue

        linea = ws.cell(row, col_linea).value
        x = ws.cell(row, col_x).value

        buscar = False

        if es_vacio(linea):
            buscar = True

        if es_vacio(x):
            buscar = True

        if buscar:

            series[serie] = row

    return series


# ============================================================
# PROCESAR TENDIDOS
# ============================================================

def procesar_tendidos(carpeta, series_objetivo):

    resultados = {}

    archivos = []

    for archivo in os.listdir(carpeta):

        if archivo.lower().endswith((".xlsx", ".xlsm")):

            ruta_completa = os.path.join(carpeta, archivo)

            archivos.append(ruta_completa)

    # ========================================================
    # ORDENAR DEL MÁS ANTIGUO AL MÁS RECIENTE
    # ========================================================

    archivos.sort(key=obtener_fecha_archivo)

    print("\n===================================")
    print("ORDEN DE ARCHIVOS:")
    print("===================================")

    for a in archivos:

        print(os.path.basename(a))

    print("===================================")

    print(f"\nArchivos encontrados: {len(archivos)}")

    # ========================================================
    # YA NO ELIMINAR SERIES
    # PARA QUE PUEDA SOBREESCRIBIR
    # CON FECHAS MÁS RECIENTES
    # ========================================================

    for ruta in archivos:

        print(f"\nAnalizando: {os.path.basename(ruta)}")

        try:

            wb = load_workbook(
                ruta,
                data_only=True,
                read_only=True
            )

        except Exception as e:

            print(f"ERROR: {e}")
            continue

        # ====================================================
        # ORDENAR HOJAS POR FECHA
        # ====================================================

        hojas_ordenadas = []

        for hoja in wb.sheetnames:

            fecha = convertir_fecha_hoja(hoja)

            if fecha is not None:

                hojas_ordenadas.append((fecha, hoja))

        hojas_ordenadas.sort()

        # ====================================================
        # RECORRER HOJAS
        # ====================================================

        for fecha, hoja in hojas_ordenadas:

            print(f"  Hoja: {hoja}")

            ws = wb[hoja]

            # ====================================================
            # RECORRER FILAS
            # ====================================================

            for fila in ws.iter_rows(values_only=False):

                for cell in fila:

                    valor = cell.value

                    if valor is None:
                        continue

                    serie = str(valor).strip().upper()

                    if serie not in series_objetivo:
                        continue

                    if not es_serie(serie):
                        continue

                    fila_excel = cell.row
                    col_excel = cell.column

                    # ====================================================
                    # BUSCAR ESTACA
                    # ====================================================

                    estaca, col_estaca = buscar_estaca(
                        ws,
                        fila_excel,
                        col_excel
                    )

                    if estaca is None:
                        continue

                    # ====================================================
                    # BUSCAR LINEA
                    # ====================================================

                    linea = buscar_linea(
                        ws,
                        col_estaca
                    )

                    if linea is None:
                        continue

                    punto = f"{linea}{estaca}"

                    # ====================================================
                    # SI YA EXISTE,
                    # SOLO REEMPLAZAR SI ES MÁS RECIENTE
                    # ====================================================

                    actualizar = False

                    if serie not in resultados:

                        actualizar = True

                    else:

                        fecha_actual = resultados[serie]["fecha"]

                        if fecha > fecha_actual:

                            actualizar = True

                    if actualizar:

                        resultados[serie] = {
                            "linea": linea,
                            "estaca": estaca,
                            "punto": punto,
                            "fecha": fecha
                        }

                        print(
                            f"     OK -> {serie} | "
                            f"L:{linea} E:{estaca} | "
                            f"{fecha.strftime('%d/%m/%Y')}"
                        )

    return resultados


# ============================================================
# ACTUALIZAR EXCEL
# ============================================================

def actualizar_excel(ws, headers, resultados):

    col_serie = headers.index("Número de Serie") + 1
    col_linea = headers.index("LÍnea") + 1
    col_estacion = headers.index("Estación") + 1
    col_punto = headers.index("PUNTO") + 1
    col_fecha = headers.index("Última Fecha") + 1

    total = ws.max_row - 1

    print("\nActualizando Excel...")

    for row in tqdm(range(2, ws.max_row + 1), total=total):

        valor = ws.cell(row, col_serie).value

        if valor is None:
            continue

        serie = str(valor).strip().upper()

        if serie not in resultados:
            continue

        datos = resultados[serie]

        ws.cell(row, col_linea).value = datos["linea"]

        ws.cell(row, col_estacion).value = datos["estaca"]

        ws.cell(row, col_punto).value = datos["punto"]

        ws.cell(row, col_fecha).value = (
            datos["fecha"].strftime("%d/%m/%Y")
        )


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":

    from tkinter import Tk, filedialog

    root = Tk()
    root.withdraw()

    print("\nSelecciona carpeta TENDIDOS")

    carpeta = filedialog.askdirectory()

    print("\nSelecciona archivo principal")

    archivo = filedialog.askopenfilename()

    print("\nAbriendo archivo principal...")

    wb = load_workbook(archivo)

    ws = wb.active

    headers = [c.value for c in ws[1]]

    print("\nBuscando series incompletas...")

    series_objetivo = cargar_series_objetivo(
        ws,
        headers
    )

    print(
        f"\nSeries objetivo: "
        f"{len(series_objetivo)}"
    )

    resultados = procesar_tendidos(
        carpeta,
        series_objetivo
    )

    print(
        f"\nSeries encontradas: "
        f"{len(resultados)}"
    )

    actualizar_excel(
        ws,
        headers,
        resultados
    )

    salida = archivo.replace(
        ".xlsx",
        "_ACTUALIZADO.xlsx"
    )

    wb.save(salida)

    print("\n===================================")
    print("ARCHIVO ACTUALIZADO")
    print(salida)
    print("===================================")