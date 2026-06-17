
import pandas as pd
import requests
import time
from pyproj import Transformer
from tkinter import Tk, filedialog
import os

# -------------------------------
# CONFIGURACIÓN
# -------------------------------
UTM_EPSG = "EPSG:32615"  # UTM Zona 15N WGS84
LATLON_EPSG = "EPSG:4326"
COLUMNA_X = "X"
COLUMNA_Y = "Y"
COLUMNA_DIRECCION = "DIRECCION"
TIEMPO_ESPERA = 1  # segundos (obligatorio para Nominatim)

# -------------------------------
# CONVERSOR UTM → LAT/LON
# -------------------------------
transformer = Transformer.from_crs(
    UTM_EPSG,
    LATLON_EPSG,
    always_xy=True
)

# -------------------------------
# FUNCIÓN PARA OBTENER DIRECCIÓN
# -------------------------------
def obtener_direccion(lat, lon):
    url = "https://nominatim.openstreetmap.org/reverse"
    params = {
        "format": "json",
        "lat": lat,
        "lon": lon,
        "zoom": 10,
        "addressdetails": 1
    }
    headers = {
        "User-Agent": "ControlDeMaterial/1.0 (contacto@ejemplo.com)"
    }

    try:
        r = requests.get(url, params=params, headers=headers, timeout=15)
        if r.status_code == 200:
            data = r.json()
            return data.get("display_name", "")
    except Exception as e:
        return ""

    return ""

# -------------------------------
# SELECCIONAR ARCHIVO EXCEL
# -------------------------------
Tk().withdraw()
archivo = filedialog.askopenfilename(
    title="Selecciona el archivo Excel con las coordenadas",
    filetypes=[("Excel files", "*.xlsx")]
)

if not archivo:
    print("No se seleccionó ningún archivo.")
    exit()

# -------------------------------
# LEER EXCEL
# -------------------------------
df = pd.read_excel(archivo)

# Verificaciones básicas
if COLUMNA_X not in df.columns or COLUMNA_Y not in df.columns:
    print("ERROR: No se encontraron las columnas X o Y")
    exit()

if COLUMNA_DIRECCION not in df.columns:
    df[COLUMNA_DIRECCION] = ""

# -------------------------------
# PROCESO PRINCIPAL
# -------------------------------
print("Procesando coordenadas...")

for i, row in df.iterrows():

    # Si ya tiene dirección, se omite
    if pd.notna(row[COLUMNA_DIRECCION]) and str(row[COLUMNA_DIRECCION]).strip() != "":
        continue

    x = row[COLUMNA_X]
    y = row[COLUMNA_Y]

    try:
        lon, lat = transformer.transform(x, y)
        direccion = obtener_direccion(lat, lon)
        df.at[i, COLUMNA_DIRECCION] = direccion
        print(f"Fila {i+1}: OK")
        time.sleep(TIEMPO_ESPERA)

    except Exception as e:
        df.at[i, COLUMNA_DIRECCION] = "ERROR"
        print(f"Fila {i+1}: ERROR")

# -------------------------------
# GUARDAR NUEVO ARCHIVO
# -------------------------------
nombre_base = os.path.splitext(os.path.basename(archivo))[0]
salida = os.path.join(
    os.path.dirname(archivo),
    f"{nombre_base}_CON_DIRECCION.xlsx"
)

df.to_excel(salida, index=False)

print("\nPROCESO TERMINADO")
print("Archivo generado:")
print(salida)
