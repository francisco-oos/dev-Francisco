# ============================================================
# LANZADOR PROFESIONAL DE HERRAMIENTAS PYTHON
# ============================================================
# Detecta carpetas dentro de la carpeta principal.
# Cada carpeta se convierte en un menú desplegable.
# Cada archivo .py dentro de esa carpeta se convierte en opción ejecutable.
# ============================================================

import os
import sys
import subprocess
import threading
from tkinter import (
    Tk, Frame, Button, Text, Scrollbar, Label,
    END, BOTH, RIGHT, Y, LEFT, X, filedialog, messagebox,
    Menubutton, Menu
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
proceso_actual = None


# ============================================================
# FUNCIONES DE TEXTO
# ============================================================

def nombre_limpio(nombre):
    """
    Convierte nombres de archivo o carpeta a texto bonito.
    Ejemplo:
    cruce_gestoria_multicapa.py -> Cruce Gestoria Multicapa
    """
    nombre = os.path.splitext(nombre)[0]
    nombre = nombre.replace("_", " ").replace("-", " ")
    return nombre.title()


# ============================================================
# DETECTAR HERRAMIENTAS POR CARPETAS
# ============================================================

def obtener_herramientas_por_carpeta():
    """
    Recorre la carpeta principal.
    Cada subcarpeta encontrada será un menú.
    Cada .py dentro será una herramienta.
    """

    herramientas = {}

    for elemento in os.listdir(BASE_DIR):
        ruta_carpeta = os.path.join(BASE_DIR, elemento)

        if not os.path.isdir(ruta_carpeta):
            continue

        if elemento.startswith("__"):
            continue

        scripts = []

        for archivo in os.listdir(ruta_carpeta):
            if not archivo.lower().endswith(".py"):
                continue

            if archivo.startswith("__"):
                continue

            ruta_script = os.path.join(ruta_carpeta, archivo)
            scripts.append((nombre_limpio(archivo), ruta_script))

        if scripts:
            herramientas[nombre_limpio(elemento)] = scripts

    return herramientas


# ============================================================
# TERMINAL
# ============================================================

def escribir_terminal(texto):
    terminal.insert(END, texto)
    terminal.see(END)
    app.update_idletasks()


def limpiar_terminal():
    terminal.delete("1.0", END)


# ============================================================
# EJECUTAR SCRIPT
# ============================================================

def ejecutar_script(nombre, ruta_script):
    global proceso_actual

    if proceso_actual and proceso_actual.poll() is None:
        messagebox.showwarning(
            "Proceso en ejecución",
            "Ya hay un script ejecutándose. Espera a que termine o detenlo."
        )
        return

    if not os.path.exists(ruta_script):
        messagebox.showerror("Archivo no encontrado", ruta_script)
        return

    limpiar_terminal()
    escribir_terminal("============================================\n")
    escribir_terminal(f"Ejecutando: {nombre}\n")
    escribir_terminal(f"Ruta: {ruta_script}\n")
    escribir_terminal("============================================\n\n")

    def hilo_ejecucion():
        global proceso_actual

        try:
            env = os.environ.copy()
            env["PYTHONIOENCODING"] = "utf-8"
            env["PYTHONUTF8"] = "1"

            proceso_actual = subprocess.Popen(
                [sys.executable, "-u", ruta_script],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                cwd=os.path.dirname(ruta_script),
                env=env
            )

            for linea in proceso_actual.stdout:
                escribir_terminal(linea)

            proceso_actual.wait()

            escribir_terminal("\n============================================\n")
            escribir_terminal(f"Proceso terminado con código: {proceso_actual.returncode}\n")
            escribir_terminal("============================================\n")

        except Exception as e:
            escribir_terminal(f"\nERROR AL EJECUTAR SCRIPT:\n{e}\n")

    threading.Thread(target=hilo_ejecucion, daemon=True).start()


# ============================================================
# DETENER PROCESO
# ============================================================

def detener_proceso():
    global proceso_actual

    if proceso_actual and proceso_actual.poll() is None:
        proceso_actual.terminate()
        escribir_terminal("\n⚠ Proceso detenido por el usuario.\n")
    else:
        escribir_terminal("\nNo hay proceso activo para detener.\n")


# ============================================================
# EJECUTAR OTRO SCRIPT MANUAL
# ============================================================

def ejecutar_script_manual():
    ruta = filedialog.askopenfilename(
        title="Selecciona script Python",
        filetypes=[("Python", "*.py")]
    )

    if ruta:
        ejecutar_script(nombre_limpio(os.path.basename(ruta)), ruta)


# ============================================================
# COMPILAR SCRIPT MANUAL
# ============================================================

def compilar_script_manual():
    global proceso_actual

    ruta = filedialog.askopenfilename(
        title="Selecciona script Python para compilar",
        filetypes=[("Python", "*.py")]
    )

    if not ruta:
        return

    limpiar_terminal()
    escribir_terminal("============================================\n")
    escribir_terminal("Compilando con PyInstaller\n")
    escribir_terminal(f"Script: {ruta}\n")
    escribir_terminal("============================================\n\n")

    def hilo_compilacion():
        global proceso_actual

        try:
            env = os.environ.copy()
            env["PYTHONIOENCODING"] = "utf-8"
            env["PYTHONUTF8"] = "1"

            proceso_actual = subprocess.Popen(
                [
                    sys.executable,
                    "-m",
                    "PyInstaller",
                    "--noconfirm",
                    "--onefile",
                    ruta
                ],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                cwd=os.path.dirname(ruta),
                env=env
            )

            for linea in proceso_actual.stdout:
                escribir_terminal(linea)

            proceso_actual.wait()

            escribir_terminal("\n============================================\n")
            escribir_terminal(f"Compilación terminada con código: {proceso_actual.returncode}\n")
            escribir_terminal("Revisa la carpeta dist.\n")
            escribir_terminal("============================================\n")

        except Exception as e:
            escribir_terminal(f"\nERROR AL COMPILAR:\n{e}\n")

    threading.Thread(target=hilo_compilacion, daemon=True).start()


# ============================================================
# CREAR MENÚS DESPLEGABLES
# ============================================================

def refrescar_menus():
    """
    Borra menús actuales y vuelve a detectar carpetas/scripts.
    """

    for widget in frame_menus.winfo_children():
        widget.destroy()

    herramientas = obtener_herramientas_por_carpeta()

    if not herramientas:
        Label(
            frame_menus,
            text="No se encontraron carpetas con scripts .py",
            fg="red"
        ).pack(side=LEFT, padx=5)
        return

    for nombre_carpeta, scripts in herramientas.items():
        menu_btn = Menubutton(
            frame_menus,
            text=f"📁 {nombre_carpeta}",
            relief="raised",
            width=25
        )

        menu = Menu(menu_btn, tearoff=0)
        menu_btn.config(menu=menu)

        for nombre_script, ruta_script in scripts:
            menu.add_command(
                label=f"▶ {nombre_script}",
                command=lambda n=nombre_script, r=ruta_script: ejecutar_script(n, r)
            )

        menu_btn.pack(side=LEFT, padx=5, pady=5)

    escribir_terminal(f"Menús actualizados. Carpetas detectadas: {len(herramientas)}\n")


# ============================================================
# INTERFAZ PRINCIPAL
# ============================================================

app = Tk()
app.title("Lanzador Profesional de Herramientas Python")
app.geometry("1200x720")

frame_top = Frame(app)
frame_top.pack(fill=X, padx=10, pady=10)

Label(
    frame_top,
    text="Herramientas por categoría",
    font=("Arial", 13, "bold")
).pack(anchor="w")

Label(
    frame_top,
    text="Cada carpeta se muestra como menú desplegable. Cada .py dentro será una herramienta ejecutable.",
    font=("Arial", 9)
).pack(anchor="w")

frame_menus = Frame(app)
frame_menus.pack(fill=X, padx=10, pady=5)

frame_extra = Frame(app)
frame_extra.pack(fill=X, padx=10, pady=5)

Button(
    frame_extra,
    text="Actualizar menús",
    width=20,
    command=refrescar_menus
).pack(side=LEFT, padx=5)

Button(
    frame_extra,
    text="Ejecutar otro .py",
    width=20,
    command=ejecutar_script_manual
).pack(side=LEFT, padx=5)

Button(
    frame_extra,
    text="Compilar .py a .exe",
    width=20,
    command=compilar_script_manual
).pack(side=LEFT, padx=5)

Button(
    frame_extra,
    text="Detener proceso",
    width=20,
    command=detener_proceso
).pack(side=LEFT, padx=5)

Button(
    frame_extra,
    text="Limpiar terminal",
    width=20,
    command=limpiar_terminal
).pack(side=LEFT, padx=5)

frame_terminal = Frame(app)
frame_terminal.pack(fill=BOTH, expand=True, padx=10, pady=10)

scroll = Scrollbar(frame_terminal)
scroll.pack(side=RIGHT, fill=Y)

terminal = Text(
    frame_terminal,
    wrap="word",
    bg="black",
    fg="lime",
    insertbackground="white",
    font=("Consolas", 10),
    yscrollcommand=scroll.set
)

terminal.pack(fill=BOTH, expand=True)
scroll.config(command=terminal.yview)

escribir_terminal("Lanzador listo.\n")
escribir_terminal(f"Carpeta base: {BASE_DIR}\n\n")

refrescar_menus()

app.mainloop()