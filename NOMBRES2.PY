import os
import re
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from ttkthemes import ThemedTk
from PIL import Image, ImageTk
import threading
import pandas as pd

# Función para seleccionar el directorio de entrada
def seleccionar_directorio_entrada():
    directorio = filedialog.askdirectory()
    if directorio:
        entrada_var.set(directorio)

# Función para seleccionar el directorio de salida
def seleccionar_directorio_salida():
    directorio = filedialog.askdirectory()
    if directorio:
        salida_var.set(directorio)

# Función para abrir el directorio de salida
def abrir_directorio_salida():
    directorio = salida_var.get()
    if os.path.isdir(directorio):
        os.startfile(directorio)
    else:
        messagebox.showerror("Error", "El directorio de salida no existe.")

# Función para limpiar los campos
def limpiar_campos():
    entrada_var.set("")
    salida_var.set("")
    patron_var.set("")
    nombre_salida_var.set("")
    archivo_excel_var.set("")

# Función para cerrar la aplicación
def cerrar_aplicacion():
    root.destroy()

# Función para actualizar la barra de progreso y el porcentaje
def actualizar_progreso(progreso, max_progreso):
    porcentaje = (progreso / max_progreso) * 100
    progreso_var.set(progreso)
    porcentaje_var.set(f"{porcentaje:.2f}%")
    root.update_idletasks()

# Función para renombrar los archivos
def renombrar_archivos():
    directorio_entrada = entrada_var.get()
    directorio_salida = salida_var.get()
    patron_texto = patron_var.get()
    nombre_salida = nombre_salida_var.get()
    archivo_excel = archivo_excel_var.get()
    
    if not directorio_entrada or not directorio_salida or not patron_texto or not nombre_salida or not archivo_excel:
        messagebox.showerror("Error", "Debe llenar todos los campos.")
        return
    
    # Leer el archivo Excel
    try:
        df = pd.read_excel(archivo_excel)
        orden_impresion = df["Orden Impresión"].tolist()
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer el archivo Excel: {e}")
        return

    # Asegurarse de que haya suficientes números en el archivo Excel
    archivos_pdf = [archivo for archivo in os.listdir(directorio_entrada) if archivo.endswith('.pdf')]
    if len(archivos_pdf) != len(orden_impresion):
        messagebox.showerror("Error", "El número de archivos PDF no coincide con los números de la columna 'Orden Impresión' en el archivo Excel.")
        return

    # Crear el directorio de salida si no existe
    if not os.path.exists(directorio_salida):
        os.makedirs(directorio_salida)

    # Ordenar los archivos según su orden en el sistema
    archivos_pdf.sort(key=lambda x: os.path.getctime(os.path.join(directorio_entrada, x)))

    max_progreso = len(archivos_pdf)

    # Crear una ventana de progreso en un hilo separado
    threading.Thread(target=progreso_renombrado, args=(archivos_pdf, orden_impresion, max_progreso, directorio_entrada, directorio_salida, nombre_salida)).start()

def progreso_renombrado(archivos_pdf, orden_impresion, max_progreso, directorio_entrada, directorio_salida, nombre_salida):
    # Crear una ventana de progreso
    progreso_window = tk.Toplevel(root)
    progreso_window.title("Progreso")
    progreso_window.geometry("300x100")

    tk.Label(progreso_window, text="Renombrando archivos...", font=("Arial", 10)).pack(pady=10)
    progress_bar = ttk.Progressbar(progreso_window, variable=progreso_var, maximum=max_progreso)
    progress_bar.pack(pady=10, padx=20, fill=tk.X)

    # Etiqueta para mostrar el porcentaje
    porcentaje_label = tk.Label(progreso_window, textvariable=porcentaje_var, font=("Arial", 10))
    porcentaje_label.pack()

    for i, archivo in enumerate(archivos_pdf):
        # Usar el número de la columna "Orden Impresión" del Excel
        numero = orden_impresion[i]
        nuevo_nombre = f"{numero:05d}__{nombre_salida}_{numero}.pdf"
        
        # Ruta completa del archivo original
        ruta_original = os.path.join(directorio_entrada, archivo)
        
        # Ruta completa del nuevo archivo en el directorio de salida
        nueva_ruta = os.path.join(directorio_salida, nuevo_nombre)
        
        # Copiar el archivo al directorio de salida con el nuevo nombre
        shutil.copyfile(ruta_original, nueva_ruta)
        
        # Actualizar la barra de progreso y el porcentaje
        actualizar_progreso(i + 1, max_progreso)

    progreso_window.destroy()
    messagebox.showinfo("Éxito", "Archivos renombrados y guardados en el directorio de salida con éxito.")

# Crear la ventana principal con tema
root = ThemedTk(theme="arc")
root.title("Renombrar Cartas PDF")

# Configuración de la imagen de fondo
imagen_fondo = Image.open(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\comparar.jpg")
imagen_fondo = imagen_fondo.resize((800, 400), Image.Resampling.LANCZOS)
imagen_fondo_tk = ImageTk.PhotoImage(imagen_fondo)
label_fondo = tk.Label(root, image=imagen_fondo_tk)
label_fondo.place(relwidth=1, relheight=1)

# Variables para almacenar los valores de entrada
entrada_var = tk.StringVar()
salida_var = tk.StringVar()
patron_var = tk.StringVar()  # Asegurarse de que esté definida
nombre_salida_var = tk.StringVar()
archivo_excel_var = tk.StringVar()
progreso_var = tk.DoubleVar()
porcentaje_var = tk.StringVar()

# Crear y colocar los widgets
tk.Label(root, text="Directorio de Entrada:", bg='lightblue', font=("Arial", 12)).grid(row=0, column=0, padx=10, pady=5, sticky='e')
tk.Entry(root, textvariable=entrada_var, width=50, font=("Arial", 11)).grid(row=0, column=1, padx=10, pady=5, columnspan=2)
tk.Button(root, text="Seleccionar", command=seleccionar_directorio_entrada, font=("Arial", 12)).grid(row=0, column=3, padx=10, pady=5)

tk.Label(root, text="Directorio de Salida:", bg='lightblue', font=("Arial", 12)).grid(row=1, column=0, padx=10, pady=5, sticky='e')
tk.Entry(root, textvariable=salida_var, width=50, font=("Arial", 11)).grid(row=1, column=1, padx=10, pady=5, columnspan=2)
tk.Button(root, text="Seleccionar", command=seleccionar_directorio_salida, font=("Arial", 12)).grid(row=1, column=3, padx=10, pady=5)

tk.Label(root, text="Patrón de Búsqueda:", bg='lightblue', font=("Arial", 12)).grid(row=2, column=0, padx=10, pady=5, sticky='e')
tk.Entry(root, textvariable=patron_var, width=30, font=("Arial", 11)).grid(row=2, column=1, padx=10, pady=5, columnspan=2)

tk.Label(root, text="Nombre de Salida:", bg='lightblue', font=("Arial", 12)).grid(row=3, column=0, padx=10, pady=5, sticky='e')
tk.Entry(root, textvariable=nombre_salida_var, width=30, font=("Arial", 11)).grid(row=3, column=1, padx=10, pady=5, columnspan=2)

tk.Label(root, text="Archivo Excel (Orden Impresión):", bg='lightblue', font=("Arial", 12)).grid(row=4, column=0, padx=10, pady=5, sticky='e')
tk.Entry(root, textvariable=archivo_excel_var, width=50, font=("Arial", 11)).grid(row=4, column=1, padx=10, pady=5, columnspan=2)
tk.Button(root, text="Seleccionar", command=lambda: archivo_excel_var.set(filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])), font=("Arial", 12)).grid(row=4, column=3, padx=10, pady=5)

tk.Button(root, text="Renombrar Archivos", command=renombrar_archivos, font=("Arial", 12)).grid(row=5, column=1, padx=10, pady=20)
tk.Button(root, text="Limpiar", command=limpiar_campos, font=("Arial", 12)).grid(row=5, column=2, padx=10, pady=20)
tk.Button(root, text="Salir", command=cerrar_aplicacion, font=("Arial", 12)).grid(row=5, column=3, padx=10, pady=20)

root.mainloop()
