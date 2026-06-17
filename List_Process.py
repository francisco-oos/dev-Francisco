import tkinter as tk
from tkinter import filedialog, messagebox
import cv2
import pytesseract
import pandas as pd
import numpy as np
from pdf2image import convert_from_path
import os

# ==============================
# CONFIGURACIÓN (AJUSTAR RUTAS)
# ==============================

# Ruta de Tesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Ruta de idioma (MUY IMPORTANTE)
os.environ["TESSDATA_PREFIX"] = r"C:\Program Files\Tesseract-OCR\tessdata"

# Ruta de Poppler (PDF)
POPPLER_PATH = r"C:\poppler\Library\bin"


class App:
    def __init__(self, root):
        """
        Inicializa la interfaz gráfica
        """
        self.root = root
        self.root.title("Sistema de Asistencia por Escaneo")
        self.root.geometry("750x500")

        tk.Label(root, text="Cargar lista (Imagen o PDF)", font=("Arial", 14)).pack(pady=10)

        tk.Button(root, text="Cargar Archivo", command=self.cargar_archivo).pack(pady=5)
        tk.Button(root, text="Procesar", command=self.procesar).pack(pady=5)
        tk.Button(root, text="Exportar a Excel", command=self.exportar).pack(pady=5)

        self.texto = tk.Text(root, height=15)
        self.texto.pack(pady=10, fill="both", expand=True)

        self.ruta_archivo = None
        self.datos = []

    # ==============================
    # CARGAR ARCHIVO
    # ==============================
    def cargar_archivo(self):
        self.ruta_archivo = filedialog.askopenfilename(
            filetypes=[("Archivos", "*.jpg *.png *.jpeg *.pdf")]
        )
        if self.ruta_archivo:
            messagebox.showinfo("Archivo", "Archivo cargado correctamente")

    # ==============================
    # PROCESAR ARCHIVO (PDF o Imagen)
    # ==============================
    def procesar(self):
        if not self.ruta_archivo:
            messagebox.showerror("Error", "Primero carga un archivo")
            return

        self.datos = []
        texto_total = ""

        try:
            # ===== SI ES PDF =====
            if self.ruta_archivo.lower().endswith(".pdf"):
                paginas = convert_from_path(
                    self.ruta_archivo,
                    poppler_path=POPPLER_PATH,
                    dpi=300
                )

                for i, pagina in enumerate(paginas):
                    img = np.array(pagina)
                    img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)

                    texto, datos = self.procesar_imagen(img)

                    texto_total += f"\n--- Página {i+1} ---\n{texto}\n"
                    self.datos.extend(datos)

            # ===== SI ES IMAGEN =====
            else:
                img = cv2.imread(self.ruta_archivo)
                texto, datos = self.procesar_imagen(img)

                texto_total = texto
                self.datos = datos

            # Mostrar texto detectado
            self.texto.delete(1.0, tk.END)
            self.texto.insert(tk.END, texto_total)

            messagebox.showinfo("Proceso", "Análisis completado")

        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ==============================
    # PROCESAMIENTO PRINCIPAL
    # ==============================
    def procesar_imagen(self, img):
        """
        Detecta filas de la tabla y extrae datos por cada empleado
        """
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        # Binarizar imagen (invertida para detectar líneas)
        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)

        # Detectar líneas horizontales (filas)
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
        lineas = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel)

        contours, _ = cv2.findContours(lineas, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        # Ordenar filas de arriba hacia abajo
        filas = sorted(contours, key=lambda x: cv2.boundingRect(x)[1])

        datos = []
        texto_total = ""

        for fila in filas:
            x, y, w, h = cv2.boundingRect(fila)

            # Ignorar ruido
            if h < 20:
                continue

            fila_img = img[y:y+h, x:x+w]

            # OCR por fila
            texto = pytesseract.image_to_string(fila_img, lang='spa')
            texto_total += texto + "\n"

            partes = texto.split()

            # Validación mínima
            if len(partes) >= 4:
                numero = partes[0]
                num_emp = partes[1]
                nombre = " ".join(partes[2:-1])
                puesto = partes[-1]

                # ==============================
                # DETECTAR FIRMA (lado derecho)
                # ==============================
                ancho = fila_img.shape[1]
                zona_firma = fila_img[:, int(ancho * 0.75):]

                firmado = self.detectar_firma(zona_firma)

                datos.append({
                    "numero": numero,
                    "num_emp": num_emp,
                    "nombre": nombre,
                    "puesto": puesto,
                    "firma": firmado
                })

        return texto_total, datos

    # ==============================
    # DETECCIÓN DE FIRMA
    # ==============================
    def detectar_firma(self, img):
        """
        Detecta si hay firma basado en cantidad de tinta
        """
        gris = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gris, 150, 255, cv2.THRESH_BINARY_INV)

        pixeles = cv2.countNonZero(thresh)
        area = img.shape[0] * img.shape[1]

        porcentaje = pixeles / area

        # Ajusta este valor según tus pruebas
        if porcentaje > 0.02:
            return "Firmado"
        else:
            return "No firmado"

    # ==============================
    # EXPORTAR RESULTADOS
    # ==============================
    def exportar(self):
        if not self.datos:
            messagebox.showerror("Error", "No hay datos para exportar")
            return

        df = pd.DataFrame(self.datos)

        ruta = filedialog.asksaveasfilename(defaultextension=".xlsx")

        if ruta:
            df.to_excel(ruta, index=False)
            messagebox.showinfo("Exportado", "Archivo guardado correctamente")


# ==============================
# EJECUCIÓN
# ==============================
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()