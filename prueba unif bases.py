import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from ttkthemes import ThemedTk
from PIL import Image, ImageTk

# Función para seleccionar el directorio de entrada
def seleccionar_directorio_entrada():
    global carpeta
    carpeta = filedialog.askdirectory()
    if carpeta:
        directorio_var.set(carpeta)

# Función para seleccionar el directorio de salida
def seleccionar_directorio_salida():
    directorio_salida = filedialog.askdirectory()
    if directorio_salida:
        directorio_salida_var.set(directorio_salida)

# Función para unificar los archivos Excel
def unificar_archivos():
    global carpeta
    global datos
    global total_registros_individuales
    global archivos_procesados
    global archivos_no_procesados

    # Limpiar las variables globales
    datos = []
    total_registros_individuales = 0
    archivos_procesados = 0
    archivos_no_procesados = []

    # Variables de salida y nombre de archivo
    directorio_salida = directorio_salida_var.get()
    nombre_archivo_salida = nombre_archivo_salida_var.get()

    # Validar la extensión del archivo de salida
    if not nombre_archivo_salida.endswith('.xlsx'):
        nombre_archivo_salida += '.xlsx'

    ruta_salida = os.path.join(directorio_salida, nombre_archivo_salida)

    # Lista todos los archivos en el directorio
    try:
        archivos_en_directorio = os.listdir(carpeta)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo acceder al directorio de entrada: {str(e)}")
        return

    for archivo in archivos_en_directorio:
        if archivo.endswith('.xlsx') or archivo.endswith('.xls') or archivo.endswith('.XLSX'):
            try:
                ruta_archivo = os.path.join(carpeta, archivo)
                # Leer el archivo Excel
                df = pd.read_excel(ruta_archivo, engine='openpyxl')  # Usar openpyxl
                datos.append(df)
                total_registros_individuales += len(df)
                archivos_procesados += 1
            except ValueError as ve:
                archivos_no_procesados.append((archivo, f"Valor inválido: {ve}"))
            except Exception as e:
                archivos_no_procesados.append((archivo, f"Error al leer: {e}"))
        else:
            archivos_no_procesados.append((archivo, "No es un archivo Excel válido"))

    if archivos_procesados == 0:
        messagebox.showwarning("Advertencia", "No se encontraron archivos de Excel en el directorio especificado.")
    else:
        # Asegurarnos de que todas las columnas sean consistentes
        columnas_comunes = set(datos[0].columns)
        for df in datos:
            if set(df.columns) != columnas_comunes:
                messagebox.showerror("Error", "Las columnas no coinciden en todos los archivos.")
                return

        # Concatena todos los DataFrames en uno solo
        try:
            datos_combinados = pd.concat(datos, ignore_index=True, sort=False)
            datos_combinados.to_excel(ruta_salida, index=False, engine='openpyxl')
            mensaje_resultados = (
                f'\nTotal de registros individuales: {total_registros_individuales}\n'
                f'Archivos procesados: {archivos_procesados}\n'
                f'Archivos no procesados: {len(archivos_no_procesados)}\n'
                f'Archivo unificado guardado en: {ruta_salida}'
            )
            messagebox.showinfo("Resultados de la Unificación", mensaje_resultados)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo: {str(e)}")

        if archivos_no_procesados:
            print('\nArchivos no procesados:')
            for archivo, error in archivos_no_procesados:
                print(f'{archivo}: {error}')

# Función para cerrar la aplicación
def cerrar_aplicacion():
    root.destroy()

# Función para aplicar el efecto de desvanecimiento (fade-in)
def fade_in(window, alpha=0):
    if alpha < 1:
        alpha += 0.05
        window.attributes("-alpha", alpha)
        root.after(50, fade_in, window, alpha)

# Función para redimensionar la imagen de fondo según el tamaño de la ventana
def resize_image(event):
    global imagen_fondo_tk
    global label_fondo

    width = event.width
    height = event.height
    imagen_fondo_resized = imagen_fondo_original.resize((width, height), Image.LANCZOS)
    imagen_fondo_tk = ImageTk.PhotoImage(imagen_fondo_resized)
    label_fondo.configure(image=imagen_fondo_tk)

# Crear la ventana principal con tema
root = ThemedTk(theme="arc")
root.title("UIFICAR BASES DE DATOS")
root.geometry("900x300")
icon_image = Image.open(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\INC.ico")
icon_photo = ImageTk.PhotoImage(icon_image)
root.iconphoto(False, icon_photo)

# Configuración de la imagen de fondo
imagen_fondo_original = Image.open(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\comparar.jpg")
imagen_fondo_original = imagen_fondo_original.resize((900, 300), Image.LANCZOS)
imagen_fondo_tk = ImageTk.PhotoImage(imagen_fondo_original)
label_fondo = tk.Label(root, image=imagen_fondo_tk)
label_fondo.place(x=0, y=0, relwidth=1, relheight=1)
label_fondo.bind('<Configure>', resize_image)

# Variables para almacenar los valores de entrada
directorio_var = tk.StringVar()
directorio_salida_var = tk.StringVar()
nombre_archivo_salida_var = tk.StringVar()

# Crear y colocar los widgets
tk.Label(root, text="Directorio de Entrada:", bg='lightblue', font=("Arial", 12)).grid(row=0, column=0, padx=10, pady=5, sticky='e')
tk.Entry(root, textvariable=directorio_var, width=50, font=("Arial", 12)).grid(row=0, column=1, padx=10, pady=5)
tk.Button(root, text="Seleccionar", command=seleccionar_directorio_entrada, font=("Arial", 12)).grid(row=0, column=2, padx=10, pady=5)

tk.Label(root, text="Directorio de Salida:", bg='lightblue', font=("Arial", 12)).grid(row=1, column=0, padx=10, pady=5, sticky='e')
tk.Entry(root, textvariable=directorio_salida_var, width=50, font=("Arial", 12)).grid(row=1, column=1, padx=10, pady=5)
tk.Button(root, text="Seleccionar", command=seleccionar_directorio_salida, font=("Arial", 12)).grid(row=1, column=2, padx=10, pady=5)

tk.Label(root, text="Nombre del Archivo de Salida:", bg='lightblue', font=("Arial", 12)).grid(row=2, column=0, padx=10, pady=5, sticky='e')
tk.Entry(root, textvariable=nombre_archivo_salida_var, width=50, font=("Arial", 12)).grid(row=2, column=1, padx=10, pady=5)

tk.Button(root, text="Unificar Archivos", command=unificar_archivos, bg="green", font=("Arial", 12)).grid(row=3, column=0, pady=10, padx=10)
tk.Button(root, text="Cerrar", command=cerrar_aplicacion, bg="red", font=("Arial", 12)).grid(row=3, column=2, pady=10, padx=10)

# Iniciar el efecto de desvanecimiento
fade_in(root)

# Ejecutar el bucle principal de la aplicación
root.mainloop()
