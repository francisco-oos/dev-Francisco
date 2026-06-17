# ==========================================
# IMPORTAR LIBRERÍAS
# ==========================================
"""
Script: Cruce espacial Punto vs Polígono (MEJORADO)

Mejoras:
- Soporte para polígonos irregulares
- Corrección automática de geometrías inválidas
- Detección robusta (incluye bordes)
- No modifica columnas adicionales del archivo original
- Si el archivo de puntos es Excel .xlsx/.xlsm, actualiza el MISMO archivo sin perder formato

Autor: [JOSE FRANCISCO ALVARADO LEYVA]
Mejorado por: ChatGPT
"""

import pandas as pd
from shapely.geometry import Point, Polygon
from tkinter import Tk, filedialog
import os
from openpyxl import load_workbook


# ==========================================
# SELECTOR DE ARCHIVOS
# ==========================================
def seleccionar_archivo(titulo):
    Tk().withdraw()
    return filedialog.askopenfilename(title=titulo)


def guardar_archivo(titulo, ruta_original):
    """
    Se conserva para CSV/TXT.
    Para Excel ya NO se usará, porque Excel se actualizará sobre el mismo archivo.
    """
    Tk().withdraw()
    ext = os.path.splitext(ruta_original)[1]

    return filedialog.asksaveasfilename(
        title=titulo,
        defaultextension=ext,
        filetypes=[
            ("Mismo formato", "*" + ext),
            ("Excel", "*.xlsx"),
            ("CSV", "*.csv"),
            ("TXT", "*.txt")
        ]
    )


# ==========================================
# LECTURA DE POLÍGONOS (MEJORADA)
# ==========================================
def leer_poligonos(ruta_txt):
    """
    Lee archivo TXT de polígonos.
    Mejora clave:
    - Se usa buffer(0) para reparar geometrías inválidas.
    """

    poligonos = []

    with open(ruta_txt, 'r', encoding='utf-8') as f:
        contenido = f.read()

    bloques = contenido.split("DESCRIPTION=")

    for bloque in bloques:

        if "NOMBRE=" not in bloque:
            continue

        lineas = bloque.splitlines()
        datos = {}
        coords = []

        for linea in lineas:

            if "=" in linea and not "," in linea:
                k, v = linea.split("=", 1)
                datos[k.strip()] = v.strip()

            elif "," in linea:
                try:
                    x, y, _ = linea.split(",")
                    coords.append((float(x), float(y)))
                except:
                    pass

        if len(coords) >= 3:
            try:
                poly = Polygon(coords)

                # Reparación de geometría inválida
                if not poly.is_valid:
                    poly = poly.buffer(0)

                poligonos.append({
                    "geometry": poly,
                    "PROPIETARIO": datos.get("NOMBRE", ""),
                    "DIRECCION": datos.get("UBICACION", ""),
                    "TIPO_PROP": datos.get("REGIMEN", ""),
                    "MUNICIPIO": datos.get("MUNICIPIO", ""),
                    "TELEFONO": datos.get("TELEFONO", "")
                })
            except:
                continue

    return poligonos


# ==========================================
# VALIDACIÓN DE VACÍOS
# ==========================================
def esta_vacio(valor):
    return pd.isna(valor) or str(valor).strip() == ""


# ==========================================
# ACTUALIZAR EXCEL SIN PERDER FORMATO
# ==========================================
def actualizar_excel_sin_perder_formato(ruta_excel, df_actualizado):
    """
    Esta función reemplaza el df.to_excel().
    Abre el Excel original con openpyxl, actualiza solo las celdas necesarias
    y guarda sobre el mismo archivo, respetando formato, fórmulas, estilos, anchos, etc.
    """

    wb = load_workbook(ruta_excel)
    ws = wb.active

    # Leer encabezados de la fila 1
    headers_excel = [cell.value for cell in ws[1]]

    # Columnas que este script puede actualizar
    columnas_objetivo = ["DIRECCION", "TIPO_PROP", "MUNICIPIO", "PROPIETARIO", "TELEFONO"]

    # Crear mapa: nombre columna -> número de columna en Excel
    columnas_excel = {}
    for col in columnas_objetivo:
        if col in headers_excel:
            columnas_excel[col] = headers_excel.index(col) + 1

    # Si alguna columna objetivo no existe en Excel, se crea al final
    for col in columnas_objetivo:
        if col not in columnas_excel:
            nueva_col = ws.max_column + 1
            ws.cell(row=1, column=nueva_col).value = col
            columnas_excel[col] = nueva_col

    # Recorrer dataframe actualizado y escribir solo valores de columnas objetivo
    # DataFrame empieza en índice 0, Excel empieza en fila 2 por encabezado
    for i, fila_df in df_actualizado.iterrows():
        fila_excel = i + 2

        for col in columnas_objetivo:
            valor = fila_df.get(col, "")

            # Solo escribir si hay valor
            # Esto evita meter NaN o cadenas raras
            if not esta_vacio(valor):
                ws.cell(row=fila_excel, column=columnas_excel[col]).value = valor

    # Guardar sobre el mismo archivo
    wb.save(ruta_excel)


# ==========================================
# CRUCE ESPACIAL (MEJORADO)
# ==========================================
def cruzar_puntos(ruta_puntos, poligonos, salida=None):

    ext = os.path.splitext(ruta_puntos)[1].lower()

    # ==========================================
    # LECTURA
    # ==========================================
    if ext in [".xlsx", ".xlsm"]:
        df = pd.read_excel(ruta_puntos)
    elif ext == ".xls":
        print("❌ ERROR: .xls no es recomendado para conservar formato con openpyxl.")
        print("Convierte el archivo a .xlsx y vuelve a intentarlo.")
        return
    else:
        df = pd.read_csv(ruta_puntos)

    # ==========================================
    # VALIDACIÓN
    # ==========================================
    if "X" not in df.columns or "Y" not in df.columns:
        print("❌ ERROR: El archivo debe tener columnas X y Y")
        return

    columnas_objetivo = ["DIRECCION", "TIPO_PROP", "MUNICIPIO", "PROPIETARIO", "TELEFONO"]

    # Evitar warnings de pandas al escribir texto
    for col in columnas_objetivo:
        if col in df.columns:
            df[col] = df[col].astype("object")
        else:
            # Si no existe la columna, se crea en el dataframe
            df[col] = ""

    # ==========================================
    # PROCESAMIENTO
    # ==========================================
    for i, row in df.iterrows():

        x = row["X"]
        y = row["Y"]

        if pd.isna(x) or pd.isna(y):
            continue

        try:
            punto = Point(float(x), float(y))
        except:
            continue

        # Tolerancia para incluir bordes o ligeras diferencias
        punto_tol = punto.buffer(0.00001)

        for pol in poligonos:

            # intersects incluye puntos dentro o tocando borde
            if pol["geometry"].intersects(punto_tol):

                if esta_vacio(row.get("DIRECCION")):
                    df.at[i, "DIRECCION"] = pol["DIRECCION"]

                if esta_vacio(row.get("TIPO_PROP")):
                    df.at[i, "TIPO_PROP"] = pol["TIPO_PROP"]

                if esta_vacio(row.get("MUNICIPIO")):
                    df.at[i, "MUNICIPIO"] = pol["MUNICIPIO"]

                if esta_vacio(row.get("PROPIETARIO")):
                    df.at[i, "PROPIETARIO"] = pol["PROPIETARIO"]

                if esta_vacio(row.get("TELEFONO")):
                    df.at[i, "TELEFONO"] = pol["TELEFONO"]

                break

    # ==========================================
    # EXPORTACIÓN
    # ==========================================

    # Excel: guardar sobre el MISMO archivo y conservar formato
    if ext in [".xlsx", ".xlsm"]:
        actualizar_excel_sin_perder_formato(ruta_puntos, df)
        print("✔ ARCHIVO ACTUALIZADO SIN PERDER FORMATO:", ruta_puntos)

    # CSV/TXT: conserva lógica original de generar salida
    elif ext == ".csv":
        if not salida:
            print("❌ ERROR: No se seleccionó archivo de salida.")
            return
        df.to_csv(salida, index=False)
        print("✔ ARCHIVO GENERADO:", salida)

    else:
        if not salida:
            print("❌ ERROR: No se seleccionó archivo de salida.")
            return
        df.to_csv(salida, sep="\t", index=False)
        print("✔ ARCHIVO GENERADO:", salida)


# ==========================================
# MAIN
# ==========================================
if __name__ == "__main__":

    print("Selecciona archivo de POLÍGONOS")
    ruta_poligonos = seleccionar_archivo("Polígonos TXT")

    if not ruta_poligonos:
        print("❌ No seleccionaste archivo de polígonos.")
        exit()

    print("Selecciona archivo de PUNTOS")
    ruta_puntos = seleccionar_archivo("Excel o CSV")

    if not ruta_puntos:
        print("❌ No seleccionaste archivo de puntos.")
        exit()

    ext = os.path.splitext(ruta_puntos)[1].lower()

    # Para Excel no preguntamos dónde guardar porque se guarda sobre el mismo archivo.
    # Para CSV/TXT sí mantenemos la lógica anterior.
    salida = None
    if ext not in [".xlsx", ".xlsm"]:
        salida = guardar_archivo("Guardar resultado", ruta_puntos)

        if not salida:
            print("❌ No seleccionaste archivo de salida.")
            exit()

    print("Leyendo polígonos...")
    poligonos = leer_poligonos(ruta_poligonos)
    print(f"Polígonos cargados: {len(poligonos)}")

    print("Cruzando puntos...")
    cruzar_puntos(ruta_puntos, poligonos, salida)

    print("✔ PROCESO TERMINADO")