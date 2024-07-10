import os
import shutil
import re
import pandas as pd
import tkinter as tk
from tkinter import Image, filedialog, ttk, messagebox
from PIL import Image, ImageTk
import time

# Función para realizar un efecto de fade-in en la ventana.
def fade_in(window):
    for alpha in range(0, 100):
        window.attributes('-alpha', alpha / 100)
        window.update_idletasks()
        time.sleep(0.01)

# Función para realizar un efecto de fade-out en la ventana.
def fade_out(window):
    for alpha in range(100, 0, -1):
        window.attributes('-alpha', alpha / 100)
        window.update_idletasks()
        time.sleep(0.01)

# Función para verificar si el nombre de archivo tiene el formato adecuado.
def es_formato_radicado(nombre_archivo):
    patron = r"\d{4}EE\d{6}[A-Z]\d+"
    return re.match(patron, nombre_archivo) is not None

# Función principal para renombrar y mover archivos.
def renombrar_archivos():
    # Verificar si se han seleccionado todas las rutas necesarias.
    if not entry_directorio_origen.get() or not entry_directorio_destino.get() or not entry_ruta_excel.get():
        messagebox.showerror("Error", "Por favor seleccione todas las rutas antes de renombrar archivos.")
        return

    # Función para actualizar el progreso de la barra y el texto del estado.
    def actualizar_progreso():
        # Mostrar la ventana emergente.
        ventana_progreso = tk.Toplevel(root)
        ventana_progreso.title("Progreso")
        ventana_progreso.geometry("300x100")

        # Configurar la barra de progreso en la ventana emergente.
        progress_bar = ttk.Progressbar(ventana_progreso, mode='determinate', length=260)
        progress_bar.pack(pady=10)

        # Etiqueta para mostrar el estado del proceso.
        label_estado = tk.Label(ventana_progreso, text="PROCESANDO ARCHIVOS...")
        label_estado.pack()

        # Obtener los directorios de origen y destino.
        directorio_origen = entry_directorio_origen.get()
        directorio_destino = entry_directorio_destino.get()

        # Leer el archivo de Excel.
        excel_file = entry_ruta_excel.get()
        df = pd.read_excel(excel_file)
        archivos_ordenados_excel = df['SAP_ID'].tolist()

        # Obtener la lista de archivos ordenados por nombre de archivo.
        archivos = [(os.path.join(directorio_origen, archivo), archivo) for archivo in os.listdir(directorio_origen) if archivo.endswith(".pdf")]
        archivos_ordenados = sorted(archivos, key=lambda x: x[1])

        # Renombrar y mover los archivos según el orden del archivo de Excel.
        for i, (archivo_origen, nombre_archivo) in enumerate(archivos_ordenados, 1):
            if nombre_archivo in archivos_ordenados_excel:
                nuevo_nombre = f"{i:05d}__{nombre_archivo}"
            elif es_formato_radicado(nombre_archivo):
                nuevo_nombre = f"{i:05d}__{nombre_archivo}"
            else:
                nuevo_nombre = f"{i:05d}__{os.path.basename(archivo_origen)}"
            
            archivo_destino = os.path.join(directorio_destino, nuevo_nombre)
            shutil.copy(archivo_origen, archivo_destino)

            # Actualizar el valor de la barra de progreso.
            progress_value = i / len(archivos_ordenados) * 100
            progress_bar['value'] = progress_value
            ventana_progreso.update()

            # Actualizar el texto del label_estado con el porcentaje.
            label_estado.config(text=f"PROCESANDO ARCHIVOS... {int(progress_value)}%")
        
        # Mostrar mensaje de finalización.
        num_archivos_renombrados = len(archivos_ordenados)
        label_estado.config(text=f"Se han renombrado y movido {num_archivos_renombrados} archivos PDF.")
        
        # Cerrar la ventana de progreso al finalizar.
        ventana_progreso.after(2000, ventana_progreso.destroy)

    # Iniciar el proceso de actualización del progreso.
    actualizar_progreso()

# Función para seleccionar el directorio de origen.
def seleccionar_directorio_origen():
    directorio_origen = filedialog.askdirectory()
    entry_directorio_origen.delete(0, tk.END)
    entry_directorio_origen.insert(0, directorio_origen)

# Función para seleccionar el directorio de destino.
def seleccionar_directorio_destino():
    directorio_destino = filedialog.askdirectory()
    entry_directorio_destino.delete(0, tk.END)
    entry_directorio_destino.insert(0, directorio_destino)

# Función para seleccionar la ruta del archivo de Excel.
def seleccionar_ruta_excel():
    ruta_excel = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
    entry_ruta_excel.delete(0, tk.END)
    entry_ruta_excel.insert(0, ruta_excel)

# Función para limpiar los campos de entrada.
def limpiar_campos():
    entry_directorio_origen.delete(0, tk.END)
    entry_directorio_destino.delete(0, tk.END)
    entry_ruta_excel.delete(0, tk.END)

# Función para cerrar la ventana.
def cerrar_ventana():
    fade_out(root)  # Realiza una transición de fade-out antes de cerrar la ventana.
    root.after(1000, root.destroy)  # Espera un segundo antes de destruir la ventana.

# Crear la ventana principal.
root = tk.Tk()
root.title("RENOMBRAR ARCHIVOS PDF")

# Calcula el ancho y el alto de la ventana.
ancho_ventana = 600
alto_ventana = 320

# Obtiene el ancho y el alto de la pantalla.
ancho_pantalla = root.winfo_screenwidth()
alto_pantalla = root.winfo_screenheight()

# Calcula la posición x y y para centrar la ventana.
x = (ancho_pantalla // 2) - (ancho_ventana // 2)
y = (alto_pantalla // 2) - (alto_ventana // 2)

# Establece la geometría de la ventana para que aparezca centrada en la pantalla.
root.geometry(f"{ancho_ventana}x{alto_ventana}+{x}+{y}")
icon_image = Image.open(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\INC.ico")
icon_photo = ImageTk.PhotoImage(icon_image)
root.iconphoto(False, icon_photo)

# Configurar la imagen de fondo.
background_image_path = r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\fnd.png"
background_image = tk.PhotoImage(file=background_image_path)
background_label = tk.Label(root, image=background_image)
background_label.place(relx=0.5, rely=0.6, anchor=tk.CENTER)

# Mostrar la ventana gradualmente.
fade_in(root)

# Crear y colocar los elementos de la interfaz gráfica.
label_directorio_origen = tk.Label(root, text="Directorio Descargas:")
label_directorio_origen.grid(row=0, column=0, sticky="e")

entry_directorio_origen = tk.Entry(root, width=50)
entry_directorio_origen.grid(row=0, column=1)

button_seleccionar_directorio_origen = tk.Button(root, text="Seleccionar", command=seleccionar_directorio_origen,width=15)
button_seleccionar_directorio_origen.grid(row=0,column=2)

label_directorio_destino = tk.Label(root, text="Directorio Renombrados:")
label_directorio_destino.grid(row=1, column=0, sticky="e")

entry_directorio_destino = tk.Entry(root, width=50)
entry_directorio_destino.grid(row=1, column=1)

button_seleccionar_directorio_destino = tk.Button(root, text="Seleccionar", command=seleccionar_directorio_destino,width=15)
button_seleccionar_directorio_destino.grid(row=1, column=2)

label_ruta_excel = tk.Label(root, text="Ruta Base de Datos:")
label_ruta_excel.grid(row=2, column=0, sticky="e")

entry_ruta_excel = tk.Entry(root, width=50)
entry_ruta_excel.grid(row=2, column=1)

button_seleccionar_ruta_excel = tk.Button(root, text="Seleccionar", command=seleccionar_ruta_excel,width=15)
button_seleccionar_ruta_excel.grid(row=2, column=2)

button_renombrar = tk.Button(root, text="RENOMBRAR ARCHIVOS", command=renombrar_archivos, width=20)
button_renombrar.config(bg="green", fg="white", width=20)  # Ajusta el ancho a 20
button_renombrar.grid(row=3, column=1)  # Empaqueta el botón en la ventana


button_limpiar = tk.Button(root, text="LIMPIAR CAMPOS", command=limpiar_campos, bg="orange",width=15)
button_limpiar.grid(row=3, column=0, columnspan=3, padx=(0,360), pady=5)

button_cerrar = tk.Button(root, text="CERRAR", command=cerrar_ventana, bg="red",width=15 )
button_cerrar.grid(row=3, column=0, columnspan=3, padx=(443,0), pady=5)

root.mainloop()
