# ============================================================
# RELLENAR COORDENADAS DESDE CSV
# ============================================================
# Equivalente en Python de la macro:
# RellenarCoordenadasCSV
#
# Función:
# - Selecciona archivo Excel destino.
# - Selecciona archivo CSV topográfico.
# - Compara:
#       Excel destino -> PUNTO
#       CSV           -> TIPO
# - Si coincide:
#       Rellena X, Y, Z en el Excel destino.
# - Guarda SOBRE EL MISMO ARCHIVO EXCEL.
# - Respeta formato, fórmulas, colores y estructura.
# ============================================================

import csv
import os
from tkinter import Tk, filedialog, messagebox
from openpyxl import load_workbook


# ============================================================
# SELECTORES DE ARCHIVO
# ============================================================

def seleccionar_archivo(titulo, tipos):
    root = Tk()
    root.withdraw()

    archivo = filedialog.askopenfilename(
        title=titulo,
        filetypes=tipos
    )

    root.destroy()
    return archivo


# ============================================================
# NORMALIZAR ENCABEZADOS
# ============================================================

def normalizar(texto):
    if texto is None:
        return ""
    return str(texto).strip().upper()


# ============================================================
# BUSCAR COLUMNAS EN EXCEL
# ============================================================

def obtener_columnas_excel(ws):
    """
    Busca columnas necesarias en la fila 1 del Excel destino:
    PUNTO, X, Y, Z
    """

    columnas = {}

    for col in range(1, ws.max_column + 1):
        encabezado = normalizar(ws.cell(row=1, column=col).value)

        if encabezado in ["PUNTO", "X", "Y", "Z"]:
            columnas[encabezado] = col

    return columnas


# ============================================================
# LEER CSV TOPOGRÁFICO
# ============================================================

def leer_csv_topografico(ruta_csv):
    """
    Lee el CSV y lo carga en memoria como diccionario.

    Clave:
        TIPO

    Valor:
        {
            "X": valor_x,
            "Y": valor_y,
            "Z": valor_z
        }
    """

    datos_csv = {}

    with open(ruta_csv, "r", encoding="utf-8-sig", newline="") as archivo:
        lector = csv.DictReader(archivo)

        # Normalizar nombres de columnas del CSV
        columnas_originales = lector.fieldnames

        if not columnas_originales:
            raise ValueError("El CSV no tiene encabezados.")

        mapa_columnas = {
            normalizar(col): col
            for col in columnas_originales
        }

        # Validar columnas obligatorias
        for requerida in ["TIPO", "X", "Y", "Z"]:
            if requerida not in mapa_columnas:
                raise ValueError(f"No se encontró la columna {requerida} en el CSV.")

        col_tipo = mapa_columnas["TIPO"]
        col_x = mapa_columnas["X"]
        col_y = mapa_columnas["Y"]
        col_z = mapa_columnas["Z"]

        for fila in lector:
            clave = str(fila.get(col_tipo, "")).strip()

            if clave:
                datos_csv[clave] = {
                    "X": fila.get(col_x, ""),
                    "Y": fila.get(col_y, ""),
                    "Z": fila.get(col_z, "")
                }

    return datos_csv


# ============================================================
# PROCESO PRINCIPAL
# ============================================================

def rellenar_coordenadas():
    """
    Proceso principal:
    1. Selecciona Excel destino.
    2. Selecciona CSV topográfico.
    3. Busca columnas.
    4. Rellena X, Y, Z.
    5. Guarda el mismo Excel.
    """
    print("Selecciona archivo Excel destino")
    archivo_excel = seleccionar_archivo(
        "Selecciona archivo Excel destino",
        [("Excel", "*.xlsx *.xlsm")]
    )

    if not archivo_excel:
        print("No seleccionaste archivo Excel.")
        return
    print("Selecciona archivo CSV topográfico CSV")
    archivo_csv = seleccionar_archivo(
        "Selecciona archivo CSV topográfico",
        [("CSV", "*.csv")]
    )

    if not archivo_csv:
        print("No seleccionaste archivo CSV.")
        return

    try:
        print("Leyendo CSV topográfico...")
        dict_csv = leer_csv_topografico(archivo_csv)

        print(f"Registros cargados desde CSV: {len(dict_csv)}")

        print("Abriendo Excel destino...")
        wb = load_workbook(archivo_excel)
        ws = wb.active

        columnas = obtener_columnas_excel(ws)

        # Validaciones del Excel destino
        if "PUNTO" not in columnas:
            raise ValueError("No se encontró la columna PUNTO en el Excel destino.")

        if "X" not in columnas:
            raise ValueError("No se encontró la columna X en el Excel destino.")

        if "Y" not in columnas:
            raise ValueError("No se encontró la columna Y en el Excel destino.")

        if "Z" not in columnas:
            raise ValueError("No se encontró la columna Z en el Excel destino.")

        col_punto = columnas["PUNTO"]
        col_x = columnas["X"]
        col_y = columnas["Y"]
        col_z = columnas["Z"]

        encontrados = 0
        no_encontrados = 0

        print("Rellenando coordenadas...")

        for fila in range(2, ws.max_row + 1):
            punto = ws.cell(row=fila, column=col_punto).value
            clave = str(punto).strip() if punto is not None else ""

            if not clave:
                continue

            if clave in dict_csv:
                datos = dict_csv[clave]

                ws.cell(row=fila, column=col_x).value = convertir_numero(datos["X"])
                ws.cell(row=fila, column=col_y).value = convertir_numero(datos["Y"])
                ws.cell(row=fila, column=col_z).value = convertir_numero(datos["Z"])

                encontrados += 1
            else:
                no_encontrados += 1

        # Guardar sobre el mismo archivo
        wb.save(archivo_excel)

        resumen = (
            "Proceso terminado correctamente.\n\n"
            f"Coincidencias encontradas: {encontrados}\n"
            f"Sin coincidencia: {no_encontrados}\n\n"
            f"Archivo actualizado:\n{archivo_excel}"
        )

        print(resumen)
        messagebox.showinfo("Proceso terminado", resumen)

    except Exception as e:
        print("ERROR:", e)
        messagebox.showerror("Error", str(e))
def convertir_numero(valor):
    if valor is None:
        return None

    texto = str(valor).strip()

    if texto == "":
        return None

    # Por si viene con coma decimal
    texto = texto.replace(",", ".")

    try:
        numero = float(texto)

        # Si es entero, guardarlo como int
        if numero.is_integer():
            return int(numero)

        return numero

    except:
        return valor

# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    rellenar_coordenadas()