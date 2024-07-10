import os
import re
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from ttkthemes import ThemedTk
from PIL import Image, ImageTk
import threading

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
    numero_inicial_var.set("")

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
    numero_inicial = numero_inicial_var.get()
    
    if not directorio_entrada or not directorio_salida or not patron_texto or not nombre_salida or not numero_inicial:
        messagebox.showerror("Error", "Debe llenar todos los campos.")
        return
    
    try:
        consecutivo = int(numero_inicial)
    except ValueError:
        messagebox.showerror("Error", "El número inicial debe ser un entero.")
        return
    
    # Agregar (\d+) al patrón ingresado por el usuario
    patron_completo_texto = f"{patron_texto}(\d+)"
    
    # Convertir el patrón de texto en una expresión regular
    try:
        patron = re.compile(patron_completo_texto)
    except re.error as e:
        messagebox.showerror("Error", f"Patrón de búsqueda inválido: {e}")
        return
    
    # Crear el directorio de salida si no existe
    if not os.path.exists(directorio_salida):
        os.makedirs(directorio_salida)
    
    # Obtener una lista de todos los archivos en el directorio de entrada con su tiempo de creación
    archivos_con_creacion = [(archivo, os.path.getctime(os.path.join(directorio_entrada, archivo))) for archivo in os.listdir(directorio_entrada) if archivo.endswith('.pdf')]
    
    # Ordenar los archivos por su tiempo de creación
    archivos_con_creacion.sort(key=lambda x: x[1])
    
    # Función para extraer el número del nombre del archivo
    def extraer_numero(nombre_archivo):
        match = patron.search(nombre_archivo)
        if match and match.groups():
            return int(match.group(1))
        return None
    
    # Filtrar y ordenar los archivos en función de los números extraídos
    archivos_ordenados = [archivo for archivo, _ in archivos_con_creacion if extraer_numero(archivo) is not None]
    
    max_progreso = len(archivos_ordenados)
    
    # Crear una ventana de progreso en un hilo separado
    threading.Thread(target=progreso_renombrado, args=(archivos_ordenados, consecutivo, max_progreso, directorio_entrada, directorio_salida, nombre_salida)).start()

def progreso_renombrado(archivos_ordenados, consecutivo, max_progreso, directorio_entrada, directorio_salida, nombre_salida):
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

    for i, archivo in enumerate(archivos_ordenados):
        # Nuevo nombre de archivo con el formato definido por el usuario
        nuevo_nombre = f"{consecutivo:05d}__{nombre_salida}_{consecutivo}.pdf"
        # Ruta completa del archivo original
        ruta_original = os.path.join(directorio_entrada, archivo)
        # Ruta completa del nuevo archivo en el directorio de salida
        nueva_ruta = os.path.join(directorio_salida, nuevo_nombre)
        # Copiar el archivo al directorio de salida con el nuevo nombre
        shutil.copyfile(ruta_original, nueva_ruta)
        # Incrementar el contador de consecutivo
        consecutivo += 1
        # Actualizar la barra de progreso y el porcentaje
        actualizar_progreso(i + 1, max_progreso)
    
    progreso_window.destroy()
    messagebox.showinfo("Éxito", "Archivos renombrados y guardados en el directorio de salida con éxito.")

# Función para aplicar el efecto de desvanecimiento (fade-in)
def fade_in(window, alpha=0):
    if alpha < 1:
        alpha += 0.05
        window.attributes("-alpha", alpha)
        root.after(50, fade_in, window, alpha)

# Crear la ventana principal con tema
root = ThemedTk(theme="arc")
root.title("Renombrar Cartas PDF")

# Función para centrar la ventana en la pantalla
def centrar_ventana(ventana, ancho, alto):
    screen_width = ventana.winfo_screenwidth()
    screen_height = ventana.winfo_screenheight()
    x = (screen_width // 2) - (ancho // 2)
    y = (screen_height // 2) - (alto // 2)
    ventana.geometry(f'{ancho}x{alto}+{x}+{y}')

# Llamar a la función para centrar la ventana
centrar_ventana(root, 800, 400)

icon_image = Image.open(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\INC.ico")
icon_photo = ImageTk.PhotoImage(icon_image)
root.iconphoto(False, icon_photo)

# Configuración de la imagen de fondo
imagen_fondo = Image.open(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\comparar.jpg")
imagen_fondo = imagen_fondo.resize((800, 400), Image.Resampling.LANCZOS)
imagen_fondo_tk = ImageTk.PhotoImage(imagen_fondo)
label_fondo = tk.Label(root, image=imagen_fondo_tk)
label_fondo.place(relwidth=1, relheight=1)

# Variables para almacenar los valores de entrada
entrada_var = tk.StringVar()
salida_var = tk.StringVar()
patron_var = tk.StringVar()
nombre_salida_var = tk.StringVar()
numero_inicial_var = tk.StringVar()
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

tk.Label(root, text="Número Inicial:", bg='lightblue', font=("Arial", 12)).grid(row=4, column=0, padx=10, pady=5, sticky='e')
tk.Entry(root, textvariable=numero_inicial_var, width=30, font=("Arial", 11)).grid(row=4, column=1, padx=10, pady=5, columnspan=2)

tk.Button(root, text="Renombrar", command=renombrar_archivos, font=("Arial", 12),bg='green').grid(row=5, column=0, padx=10, pady=20)
tk.Button(root, text="Limpiar", command=limpiar_campos, font=("Arial", 12), bg='orange').grid(row=5, column=1, padx=10, pady=20)
tk.Button(root, text="Abrir Carpeta de Salida", command=abrir_directorio_salida, font=("Arial", 12),bg='gray').grid(row=5, column=2, padx=10, pady=20)
tk.Button(root, text="Cerrar", command=cerrar_aplicacion, font=("Arial", 12), bg='red').grid(row=5, column=3, padx=10, pady=20)

# Iniciar el efecto de desvanecimiento
fade_in(root)

# Iniciar el bucle principal de la aplicación
root.mainloop()
