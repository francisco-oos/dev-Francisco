import pandas as pd
from tkinter import *
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import zipfile
import os
import tempfile
import shutil
import xml.etree.ElementTree as ET
from openpyxl import load_workbook
# ==========================================================
# ESTRUCTURA ESPERADA DE LOS ARCHIVOS EXCEL
# ==========================================================

"""
ESTRUCTURA GENERAL ESPERADA:

Este sistema fue diseñado para trabajar con archivos
Excel (.xlsx) provenientes de levantamientos o registros
que contienen evidencias fotográficas incrustadas.

----------------------------------------------------------
REQUISITOS IMPORTANTES
----------------------------------------------------------

1. LOS ARCHIVOS DEBEN SER .XLSX
   --------------------------------
   NO .xls
   NO CSV
   NO PDF

   Porque las imágenes incrustadas únicamente pueden
   extraerse desde archivos .xlsx.

----------------------------------------------------------

2. EL EXCEL PRINCIPAL DEBE TENER:
   --------------------------------

   - Una columna llamada EXACTAMENTE:

        NUM. DE SERIE ORIGINAL

   Esta columna se usa como identificador principal.

----------------------------------------------------------

3. EL ARCHIVO DE DAÑADOS DEBE TENER:
   --------------------------------

   - Una columna llamada EXACTAMENTE:

        NÚMERO DE SERIE

   Esta columna se usa para comparar coincidencias.

----------------------------------------------------------

4. LAS IMÁGENES DEBEN ESTAR:
   --------------------------------

   - INSERTADAS dentro del Excel
   - NO pegadas como comentario
   - NO vinculadas externamente

   El sistema extrae imágenes embebidas.

----------------------------------------------------------

5. COLUMNAS DE EVIDENCIA
   --------------------------------

   El sistema asume que:

   - columna 10 = EVIDENCIA 1
   - columna 11 = EVIDENCIA 2

   IMPORTANTE:
   Estas posiciones son internas de Excel.

   Si cambian las columnas donde están las imágenes,
   debe modificarse:

       if col in [10, 11]

----------------------------------------------------------

6. ORDEN DE CARGA DE ARCHIVOS
   --------------------------------

   Cuando seleccionas múltiples archivos:

   El sistema LOS CARGA EN EL ORDEN
   EN QUE WINDOWS LOS ENTREGA.

   Después:

   - Une todos los DataFrames
   - Genera una numeración global
   - Mantiene relación correcta:
       fila -> imágenes

----------------------------------------------------------

7. FUNCIONAMIENTO DEL OFFSET
   --------------------------------

   offset sirve para evitar que:

   archivo1 fila 0
   archivo2 fila 0

   se sobreescriban entre sí.

   Ejemplo:

   Archivo A:
   filas 0-100

   Archivo B:
   filas 0-50

   Con offset:
   Archivo B pasa a:
   filas 101-151

   Así las imágenes quedan correctamente asociadas.

----------------------------------------------------------

8. FLUJO GENERAL DEL SISTEMA
   --------------------------------

   PASO 1:
   Usuario carga uno o varios Excel.

   PASO 2:
   Se leen las tablas usando pandas.

   PASO 3:
   Se extraen imágenes desde XML internos.

   PASO 4:
   Las imágenes se renombran usando:

       NUMERO_SERIE_1.png
       NUMERO_SERIE_2.png

   PASO 5:
   Se guardan temporalmente en:

       temp_dir

   PASO 6:
   Se construye:

       imagenes_por_fila

   que relaciona:
       fila -> imágenes

   PASO 7:
   Se muestran los datos en Treeview.

   PASO 8:
   El usuario puede:
       - filtrar
       - buscar
       - copiar
       - abrir imágenes
       - exportar imágenes
       - agregar dañados

----------------------------------------------------------

9. RELACIÓN ENTRE FILAS E IMÁGENES
   --------------------------------

   El sistema NO guarda imágenes dentro del DataFrame.

   Guarda rutas en:

       imagenes_por_fila

   Ejemplo:

       {
           15: [
               "Q12345_1.png",
               "Q12345_2.png"
           ]
       }

   Entonces:

   fila 15
   tiene 2 evidencias.

----------------------------------------------------------

10. COLUMNAS AGREGADAS AUTOMÁTICAMENTE
    ----------------------------------

    El sistema crea:

        EVIDENCIA 1
        EVIDENCIA 2

    para mostrar el icono:

        📷

    si existen imágenes asociadas.

----------------------------------------------------------

11. ARCHIVOS TEMPORALES
    ----------------------------------

    Todas las imágenes extraídas se almacenan en:

        temp_dir

    Esta carpeta se elimina automáticamente
    al cerrar la aplicación.

----------------------------------------------------------

12. COMPARACIÓN DE DAÑADOS
    ----------------------------------

    El sistema compara:

        NUM. DE SERIE ORIGINAL
                    VS
        NÚMERO DE SERIE

    Resultado:

        ✅ Existe en dañados
        ❌ No existe

----------------------------------------------------------

13. AGREGAR A DAÑADOS
    ----------------------------------

    Cuando se agregan filas:

    - NO sobrescribe datos existentes
    - SOLO agrega series nuevas
    - Puede copiar imágenes asociadas
    - Guarda directamente en el Excel original

----------------------------------------------------------

14. FILTROS
    ----------------------------------

    Existen 2 tipos:

    A) Filtro por columna
       - selecciona columna
       - escribe valor

    B) Búsqueda global
       - busca en TODAS las columnas

----------------------------------------------------------

15. VISOR DE IMÁGENES
    ----------------------------------

    Al hacer DOBLE CLICK en:

        EVIDENCIA 1
        EVIDENCIA 2

    se abre visor donde puedes:

        - navegar
        - seleccionar
        - guardar imágenes

----------------------------------------------------------

16. MENÚ CONTEXTUAL
    ----------------------------------

    Click derecho sobre una celda:

        -> Copiar

    El valor se manda al portapapeles Windows.

----------------------------------------------------------

17. COLUMNAS OCULTAS
    ----------------------------------

    Las columnas ocultas NO se eliminan.

    Solo dejan de mostrarse visualmente
    en el Treeview.

----------------------------------------------------------

18. IMPORTANTE SOBRE LOS NOMBRES
    ----------------------------------

    El sistema convierte columnas a MAYÚSCULAS:

        df.columns = [str(c).strip().upper()]

    Por eso:

        Numero de Serie
        número de serie
        NÚMERO DE SERIE

    terminan iguales.

----------------------------------------------------------

19. LIMITACIONES IMPORTANTES
    ----------------------------------

    - Solo funciona correctamente con .xlsx
    - Imágenes deben estar incrustadas
    - Si Excel está protegido puede fallar
    - Si cambian columnas de evidencias,
      debe modificarse el código
    - Si faltan columnas obligatorias,
      se omite el archivo

----------------------------------------------------------

20. RESUMEN TÉCNICO
    ----------------------------------

    pandas:
        Manejo de tablas

    tkinter:
        Interfaz gráfica

    PIL:
        Mostrar imágenes

    zipfile:
        Abrir .xlsx internamente

    XML:
        Leer posiciones de imágenes

    openpyxl:
        Editar Excel

    shutil:
        Copiar archivos

    tempfile:
        Manejo carpeta temporal
"""
# ==========================================================
# VARIABLES GLOBALES
# ==========================================================
df_total = pd.DataFrame()           # DataFrame con todos los datos cargados
imagenes_por_fila = {}              # Diccionario {fila_global: [rutas_imagenes]}
temp_dir = os.path.join(tempfile.gettempdir(), "visor_excel_imgs")  # Carpeta temporal
df_danados = pd.DataFrame()         # DataFrame de registros "dañados"
ruta_danados_actual = None          # Ruta del archivo de dañados cargado
columnas_ocultas = set()            # Conjunto de columnas ocultas

# ==========================================================
# LIMPIEZA DE ARCHIVOS TEMPORALES
# ==========================================================
def limpiar_temp():
    """Elimina la carpeta temporal de imágenes si existe."""
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)

# ==========================================================
# EXTRACCIÓN DE IMÁGENES DESDE EXCEL (.xlsx)
# ==========================================================
def extraer_imagenes_con_posicion(archivo, temp_dir):
    """
    Extrae imágenes de un archivo Excel con su posición de fila y columna.
    Devuelve una lista de tuplas: (fila, columna, ruta_imagen)
    """
    imagenes_map = []
    with zipfile.ZipFile(archivo, 'r') as z:
        drawing_files = [f for f in z.namelist() if f.startswith("xl/drawings/drawing") and f.endswith(".xml")]
        for drawing_xml in drawing_files:
            try:
                rels_xml = drawing_xml.replace("drawings/", "drawings/_rels/") + ".rels"
                if rels_xml not in z.namelist():
                    continue
                rels = ET.fromstring(z.read(rels_xml))
                rel_dict = {rel.attrib['Id']: rel.attrib['Target'] for rel in rels}

                drawing = ET.fromstring(z.read(drawing_xml))
                ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                      'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'}

                anchors = drawing.findall("xdr:twoCellAnchor", ns) + drawing.findall("xdr:oneCellAnchor", ns)
                for anchor in anchors:
                    try:
                        row = int(anchor.find(".//xdr:from/xdr:row", ns).text)
                        col = int(anchor.find(".//xdr:from/xdr:col", ns).text)
                    except:
                        continue
                    blip = anchor.find(".//a:blip", ns)
                    if blip is None:
                        continue
                    rId = blip.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                    if rId not in rel_dict:
                        continue
                    target = rel_dict[rId].replace("../", "")
                    img_path = "xl/" + target
                    if img_path not in z.namelist():
                        continue
                    nombre = os.path.basename(img_path)
                    ruta = os.path.join(temp_dir, nombre)
                    with open(ruta, "wb") as f:
                        f.write(z.read(img_path))
                    imagenes_map.append((row, col, ruta))
            except Exception as e:
                print("Error leyendo drawing:", drawing_xml, e)
    return imagenes_map

# ==========================================================
# CARGA DE ARCHIVOS EXCEL PRINCIPAL
# ==========================================================
def cargar_archivos():
    """Carga uno o varios archivos Excel y extrae las imágenes asociadas."""
    global df_total, imagenes_por_fila
    archivos = filedialog.askopenfilenames(title="Selecciona archivos Excel", filetypes=[("Excel", "*.xlsx")])
    if not archivos:
        return

    # Limpiar DataFrames y carpeta temporal
    df_total = pd.DataFrame()
    imagenes_por_fila = {}
    limpiar_temp()
    os.makedirs(temp_dir, exist_ok=True)
    offset = 0  # Para numeración de filas global

    for archivo in archivos:
        df = pd.read_excel(archivo)
        df.columns = [str(c).strip().upper() for c in df.columns]
        col_serie = "NUM. DE SERIE ORIGINAL"
        if col_serie not in df.columns:
            messagebox.showerror("Error", f"{archivo} no tiene columna '{col_serie}'")
            continue

        # Extraer imágenes y renombrarlas según evidencia
        imagenes_map = extraer_imagenes_con_posicion(archivo, temp_dir)
        for row, col, ruta_img in imagenes_map:
            fila_df = row - 1
            fila_global = fila_df + offset
            if col in [10, 11]:  # Columnas de evidencia
                if fila_df >= len(df):
                    continue
                serie_valor = str(df.loc[fila_df, col_serie]).strip()
                if not serie_valor:
                    continue
                evidencia_num = 1 if col == 10 else 2
                nuevo_nombre = os.path.join(temp_dir, f"{serie_valor}_{evidencia_num}.png")
                try:
                    os.rename(ruta_img, nuevo_nombre)
                except:
                    continue
                imagenes_por_fila.setdefault(fila_global, []).append(nuevo_nombre)

        df_total = pd.concat([df_total, df], ignore_index=True)
        offset += len(df)

    # Añadir columnas de evidencias si no existen
    for ev in ["EVIDENCIA 1", "EVIDENCIA 2"]:
        if ev not in df_total.columns:
            df_total[ev] = ""

    # Marcar las filas con evidencia
    for i in range(len(df_total)):
        evidencias = imagenes_por_fila.get(i, [])
        df_total.at[i, "EVIDENCIA 1"] = "📷" if len(evidencias) >= 1 else ""
        df_total.at[i, "EVIDENCIA 2"] = "📷" if len(evidencias) >= 2 else ""

    # Mostrar nota informativa al usuario
    #messagebox.showinfo("Archivo cargado", f"Se cargaron {len(df_total)} filas desde {len(archivos)} archivo(s).")
    cargar_tabla(df_total)
    combo['values'] = list(df_total.columns)

# ==========================================================
# CARGAR ARCHIVO DE DAÑADOS
# ==========================================================
def cargar_danados():
    """Carga un archivo Excel con registros de dañados y actualiza la columna de estado."""
    global df_danados, df_total, ruta_danados_actual
    archivo = filedialog.askopenfilename(title="Selecciona archivo de dañados", filetypes=[("Excel", "*.xlsx")])
    if not archivo:
        return
    try:
        df_danados = pd.read_excel(archivo)
        df_danados.columns = [str(c).strip().upper() for c in df_danados.columns]
        ruta_danados_actual = archivo
        #messagebox.showinfo("Archivo de dañados", f"Se cargaron {len(df_danados)} filas desde {archivo}")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer el archivo: {e}")
        return
    refrescar_danados()

def refrescar_danados():
    """Actualiza la columna 'EN DAÑADOS' en df_total según df_danados."""
    global df_danados, df_total
    col_serie = "NÚMERO DE SERIE"
    if col_serie not in df_danados.columns:
        return
    set_danados = set(df_danados[col_serie].astype(str).str.strip())
    df_total["EN DAÑADOS"] = df_total["NUM. DE SERIE ORIGINAL"].astype(str).str.strip().apply(
        lambda x: "✅" if x in set_danados else "❌"
    )
    cargar_tabla(df_total)

# ==========================================================
# AGREGAR NUEVOS DATOS AL ARCHIVO DE DAÑADOS
# ==========================================================
def agregar_a_danados():
    """Agrega filas seleccionadas del df_total al archivo de dañados, sin sobrescribir."""
    global df_danados, df_total, ruta_danados_actual
    if df_total.empty:
        messagebox.showwarning("Aviso", "No hay datos cargados")
        return
    if df_danados.empty or ruta_danados_actual is None:
        messagebox.showwarning("Aviso", "Primero carga un archivo de dañados")
        return

    seleccion = tree.selection()
    if not seleccion:
        messagebox.showinfo("Agregar", "No hay filas seleccionadas")
        return

    col_serie_total = "NUM. DE SERIE ORIGINAL"
    col_serie_dan = "NÚMERO DE SERIE"
    set_danados = set(df_danados[col_serie_dan].astype(str).str.strip())

    carpeta = filedialog.askdirectory(title="Selecciona carpeta para guardar evidencias")
    if not carpeta:
        carpeta = None

    wb = load_workbook(ruta_danados_actual)
    ws = wb.active
    nuevas_filas = 0

    for item in seleccion:
        fila = int(item)
        serie = str(df_total.at[fila, col_serie_total]).strip()
        if serie in set_danados:
            continue  # Ya existe

        fila_vals = []
        for col in df_danados.columns:
            if col == col_serie_dan:
                fila_vals.append(serie)
            elif col in df_total.columns:
                fila_vals.append(df_total.at[fila, col])
            else:
                fila_vals.append("")
        ws.append(fila_vals)
        nuevas_filas += 1

        # Copiar imágenes asociadas
        if carpeta and fila in imagenes_por_fila:
            for ruta in imagenes_por_fila[fila]:
                nombre = os.path.basename(ruta)
                shutil.copy(ruta, os.path.join(carpeta, nombre))

    if nuevas_filas > 0:
        try:
            wb.save(ruta_danados_actual)
            messagebox.showinfo("Agregar", f"{nuevas_filas} filas agregadas al archivo de dañados.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo de dañados: {e}")

        df_danados = pd.read_excel(ruta_danados_actual)
        df_danados.columns = [str(c).strip().upper() for c in df_danados.columns]
        refrescar_danados()
    else:
        messagebox.showinfo("Agregar", "No se agregaron filas nuevas (ya existían).")

# ==========================================================
# FUNCIONES DE INTERFAZ (Treeview, filtros, visor, exportar)
# ==========================================================
def cargar_tabla(df):
    """Carga un DataFrame en el Treeview, respetando columnas ocultas."""
    tree.delete(*tree.get_children())
    columnas_visibles = [c for c in df.columns if c not in columnas_ocultas]
    tree["columns"] = columnas_visibles
    for col in columnas_visibles:
        tree.heading(col, text=col)
        tree.column(col, width=130)
    for i, row in df.iterrows():
        values = [row[c] for c in columnas_visibles]
        tree.insert("", "end", iid=i, values=values)

def aplicar_filtro():
    """Filtra el DataFrame según columna y valor ingresado."""
    if df_total.empty:
        return
    col = col_var.get()
    val = val_var.get().lower()
    if col == "" or val == "":
        return
    filtrado = df_total[df_total[col].astype(str).str.lower().str.contains(val)]
    cargar_tabla(filtrado)

def visor_imagenes(lista):
    """Abre un visor de imágenes con navegación y selección para guardar."""
    if not lista:
        messagebox.showinfo("Info", "No hay imágenes")
        return
    win = Toplevel()
    win.title("EVIDENCIAS")
    index = [0]
    seleccionadas = set()
    lbl = Label(win)
    lbl.pack()

    def mostrar():
        img = Image.open(lista[index[0]])
        img.thumbnail((900, 700))
        img_tk = ImageTk.PhotoImage(img)
        lbl.config(image=img_tk)
        lbl.image = img_tk

    def siguiente():
        if index[0] < len(lista) - 1:
            index[0] += 1
            mostrar()

    def anterior():
        if index[0] > 0:
            index[0] -= 1
            mostrar()

    def toggle_seleccion():
        if index[0] in seleccionadas:
            seleccionadas.remove(index[0])
        else:
            seleccionadas.add(index[0])

    def guardar_imagenes():
        if not lista:
            return
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta destino")
        if not carpeta:
            return
        imgs_a_guardar = [lista[i] for i in seleccionadas] if seleccionadas else lista
        for ruta in imgs_a_guardar:
            nombre = os.path.basename(ruta)
            shutil.copy(ruta, os.path.join(carpeta, nombre))
        messagebox.showinfo("Guardado", f"{len(imgs_a_guardar)} imágenes guardadas en {carpeta}")

    btn_frame = Frame(win)
    btn_frame.pack()
    Button(btn_frame, text="<<", command=anterior).pack(side="left")
    Button(btn_frame, text=">>", command=siguiente).pack(side="left")
    Button(btn_frame, text="Seleccionar/Deseleccionar", command=toggle_seleccion).pack(side="left")
    Button(btn_frame, text="Guardar imágenes", command=guardar_imagenes).pack(side="left")
    mostrar()

def exportar_imagenes_seleccionadas_automatico():
    """Exporta imágenes de las filas seleccionadas a una carpeta elegida."""
    seleccion = tree.selection()
    if not seleccion:
        messagebox.showinfo("Exportar", "No hay filas seleccionadas")
        return
    carpeta = filedialog.askdirectory(title="Seleccionar carpeta destino")
    if not carpeta:
        return
    total_guardadas = 0
    for item in seleccion:
        fila = int(item)
        if fila in imagenes_por_fila:
            for ruta in imagenes_por_fila[fila]:
                nombre = os.path.basename(ruta)
                destino = os.path.join(carpeta, nombre)
                try:
                    shutil.copy(ruta, destino)
                    total_guardadas += 1
                except Exception as e:
                    print("Error copiando:", e)
    messagebox.showinfo("Exportar", f"{total_guardadas} imágenes guardadas en {carpeta}")

def doble_click(event):
    """Al hacer doble click abre el visor de imágenes si la columna es de evidencia."""
    item = tree.identify_row(event.y)
    col_id = tree.identify_column(event.x)
    if not item or not col_id:
        return
    fila = int(item)
    col_index = int(col_id.replace("#", "")) - 1
    columnas = tree["columns"]
    if col_index >= len(columnas):
        return
    nombre_columna = columnas[col_index]
    tree.focus(item)
    tree.selection_set(item)
    if nombre_columna in ["EVIDENCIA 1", "EVIDENCIA 2"]:
        lista_imgs = imagenes_por_fila.get(fila, [])
        if not lista_imgs:
            messagebox.showinfo("Sin evidencia", "No hay imágenes")
            return
        visor_imagenes(lista_imgs)

# ==========================================================
# FUNCIONES DE COLUMNAS (Ocultar/Mostrar)
# ==========================================================
def ocultar_columna():
    """Oculta la columna seleccionada en el Treeview."""
    col = col_var.get()
    if col in df_total.columns and col not in columnas_ocultas:
        columnas_ocultas.add(col)
        cargar_tabla(df_total)

def mostrar_columnas():
    """Muestra todas las columnas ocultas."""
    columnas_ocultas.clear()
    cargar_tabla(df_total)
# ==========================================================
# MENU CONTEXTUAL PARA COPIAR (AGREGADO)
# ==========================================================
def copiar_celda(event):
    item = tree.identify_row(event.y)
    col_id = tree.identify_column(event.x)

    if not item or not col_id:
        return

    col_index = int(col_id.replace("#", "")) - 1
    columnas = tree["columns"]

    if col_index >= len(columnas):
        return

    valor = tree.item(item, "values")[col_index]

    app.clipboard_clear()
    app.clipboard_append(valor)
    app.update()

def mostrar_menu(event):
    item = tree.identify_row(event.y)
    if not item:
        return

    tree.selection_set(item)

    menu = Menu(app, tearoff=0)
    menu.add_command(label="Copiar", command=lambda: copiar_celda(event))
    menu.tk_popup(event.x_root, event.y_root)
# ==========================================================
# INTERFAZ PRINCIPAL
# ==========================================================
app = Tk()
app.title("VISOR EXCEL PRO")
app.geometry("1400x750")
app.protocol("WM_DELETE_WINDOW", lambda: (limpiar_temp(), app.destroy()))

# Panel superior con botones y filtros
frame_top = Frame(app)
frame_top.pack(fill="x")

Button(frame_top, text="📂 Cargar", command=cargar_archivos).pack(side="left")
Button(frame_top, text="📌 Cargar dañados", command=cargar_danados).pack(side="left", padx=5)
Button(frame_top, text="➕ Agregar seleccionados a dañados", command=agregar_a_danados).pack(side="left", padx=5)
Label(frame_top, text="Buscar:").pack(side="left")
search_var = StringVar()
Entry(frame_top, textvariable=search_var, width=30).pack(side="left")
Label(frame_top, text="Columna:").pack(side="left")
col_var = StringVar()
combo = ttk.Combobox(frame_top, textvariable=col_var, width=25)
combo.pack(side="left")
val_var = StringVar()
Entry(frame_top, textvariable=val_var, width=20).pack(side="left")
Button(frame_top, text="Filtrar", command=aplicar_filtro).pack(side="left")
Button(frame_top, text="Seleccionar todo", command=lambda: tree.selection_set(tree.get_children())).pack(side="left", padx=5)
Button(frame_top, text="Deseleccionar todo", command=lambda: tree.selection_remove(tree.selection())).pack(side="left", padx=5)
Button(frame_top, text="Exportar imágenes seleccionadas", command=exportar_imagenes_seleccionadas_automatico).pack(side="left", padx=5)
Button(frame_top, text="Ocultar columna", command=ocultar_columna).pack(side="left", padx=5)
Button(frame_top, text="Mostrar todas columnas", command=mostrar_columnas).pack(side="left", padx=5)

# Treeview con scroll
frame = Frame(app)
frame.pack(fill="both", expand=True)
scroll_y = Scrollbar(frame)
scroll_y.pack(side="right", fill="y")
scroll_x = Scrollbar(frame, orient="horizontal")
scroll_x.pack(side="bottom", fill="x")
tree = ttk.Treeview(frame, columns=[], show="headings", yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
scroll_y.config(command=tree.yview)
scroll_x.config(command=tree.xview)
tree.pack(fill="both", expand=True)
tree.bind("<Double-1>", doble_click)
tree.bind("<Button-3>", mostrar_menu)

# Filtro de búsqueda global
def buscar(*args):
    if df_total.empty:
        return
    txt = search_var.get().lower()
    if txt == "":
        cargar_tabla(df_total)
        return
    filtrado = df_total[df_total.apply(lambda r: r.astype(str).str.lower().str.contains(txt).any(), axis=1)]
    cargar_tabla(filtrado)

search_var.trace("w", buscar)

# ==========================================================
# INICIAR APLICACIÓN
# ==========================================================
app.mainloop()