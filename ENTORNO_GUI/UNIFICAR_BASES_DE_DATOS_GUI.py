import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
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
                # Crea la ruta completa al archivo
                ruta_archivo = os.path.join(carpeta, archivo)
                # Lee el archivo Excel
                df = pd.read_excel(ruta_archivo)
                # Añade el DataFrame a la lista
                datos.append(df)
                # Suma la cantidad de registros del archivo individual
                total_registros_individuales += len(df)
                archivos_procesados += 1
            except Exception as e:
                archivos_no_procesados.append((archivo, str(e)))
        else:
            archivos_no_procesados.append((archivo, "No es un archivo Excel válido"))

    if archivos_procesados == 0:
        messagebox.showwarning("Advertencia", "No se encontraron archivos de Excel en el directorio especificado.")
    else:
        # Concatena todos los DataFrames en uno solo
        datos_combinados = pd.concat(datos, ignore_index=True)

        # Asignar el nombre de la primera columna a "Numero de secuencia"
        datos_combinados.columns = ["Numero de secuencia "] + list(datos_combinados.columns[1:])

        # Guarda el DataFrame combinado en un nuevo archivo Excel
        try:
            with pd.ExcelWriter(ruta_salida) as writer:
                datos_combinados.to_excel(writer, sheet_name='Hoja1', index=False)
            # Imprime los resultados
            mensaje_resultados = (
                f'\nTotal de registros individuales: {total_registros_individuales}\n'
                f'Archivos procesados: {archivos_procesados}\n'
                f'Archivos no procesados: {len(archivos_no_procesados)}\n'
                f'Archivo unificado guardado en: {ruta_salida}'
            )
            messagebox.showinfo("Resultados de la Unificación", mensaje_resultados)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo: {str(e)}")

        # Imprime la lista de archivos que no se pudieron procesar
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

    # Obtener el tamaño de la ventana
    width = event.width
    height = event.height

    # Redimensionar la imagen de fondo
    imagen_fondo_resized = imagen_fondo_original.resize((width, height), Image.LANCZOS)
    imagen_fondo_tk = ImageTk.PhotoImage(imagen_fondo_resized)

    # Actualizar la imagen de fondo en el label
    label_fondo.configure(image=imagen_fondo_tk)

# Función para centrar la ventana en la pantalla
def centrar_ventana(ventana, ancho, alto):
    screen_width = ventana.winfo_screenwidth()
    screen_height = ventana.winfo_screenheight()
    x = (screen_width // 2) - (ancho // 2)
    y = (screen_height // 2) - (alto // 2)
    ventana.geometry(f'{ancho}x{alto}+{x}+{y}')

# Crear la ventana principal con tema
root = ThemedTk(theme="arc")
root.title("UNIFICAR BASES DE DATOS")
centrar_ventana(root, 900, 300)

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
