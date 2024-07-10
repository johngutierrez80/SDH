import os  # Módulo para interactuar con el sistema operativo
import shutil  # Módulo para operaciones de archivos de alto nivel como copiar archivos
import re  # Módulo para operaciones con expresiones regulares
import pandas as pd  # Módulo para manipulación y análisis de datos
import tkinter as tk  # Módulo para crear interfaces gráficas
from tkinter import filedialog, ttk, messagebox  # Submódulos específicos de tkinter para cuadros de diálogo, widgets y mensajes
from PIL import Image, ImageTk
import time  # Módulo para operaciones relacionadas con el tiempo

# Función para realizar un efecto de fade-in en la ventana.
def fade_in(window):
    for alpha in range(0, 100):  # Iterar de 0 a 99 para aumentar la transparencia
        window.attributes('-alpha', alpha / 100)  # Establecer la transparencia de la ventana
        window.update_idletasks()  # Actualizar la interfaz de usuario
        time.sleep(0.01)  # Esperar 0.01 segundos

# Función para realizar un efecto de fade-out en la ventana.
def fade_out(window):
    for alpha in range(100, 0, -1):  # Iterar de 100 a 1 para disminuir la transparencia
        window.attributes('-alpha', alpha / 100)  # Establecer la transparencia de la ventana
        window.update_idletasks()  # Actualizar la interfaz de usuario
        time.sleep(0.01)  # Esperar 0.01 segundos

# Función para verificar si el nombre de archivo tiene el formato adecuado.
def es_formato_radicado(nombre_archivo):
    patron = r"\d{4}EE\d{6}[A-Z]\d+"  # Expresión regular para el formato de archivo
    return re.match(patron, nombre_archivo) is not None  # Verificar si el nombre cumple con el patrón

# Función principal para renombrar y mover archivos.
def renombrar_archivos():
    # Verificar si se han seleccionado todas las rutas necesarias.
    if not entry_directorio_origen.get() or not entry_directorio_destino.get() or not entry_ruta_excel.get():
        messagebox.showerror("Error", "Por favor seleccione todas las rutas antes de renombrar archivos.")  # Mostrar mensaje de error
        return

    # Función para actualizar el progreso de la barra y el texto del estado.
    def actualizar_progreso():
        # Mostrar la ventana emergente.
        ventana_progreso = tk.Toplevel(root)  # Crear una nueva ventana emergente
        ventana_progreso.title("Progreso")  # Título de la ventana emergente
        ventana_progreso.geometry("300x100")  # Establecer tamaño de la ventana emergente

        # Configurar la barra de progreso en la ventana emergente.
        progress_bar = ttk.Progressbar(ventana_progreso, mode='determinate', length=260)  # Crear una barra de progreso
        progress_bar.pack(pady=10)  # Empaquetar la barra de progreso en la ventana

        # Etiqueta para mostrar el estado del proceso.
        label_estado = tk.Label(ventana_progreso, text="PROCESANDO ARCHIVOS...")  # Crear una etiqueta
        label_estado.pack()  # Empaquetar la etiqueta en la ventana

        # Obtener los directorios de origen y destino.
        directorio_origen = entry_directorio_origen.get()  # Obtener el directorio de origen
        directorio_destino = entry_directorio_destino.get()  # Obtener el directorio de destino

        # Leer el archivo de Excel.
        excel_file = entry_ruta_excel.get()  # Obtener la ruta del archivo de Excel
        df = pd.read_excel(excel_file)  # Leer el archivo de Excel
        archivos_ordenados_excel = df['SAP_ID'].astype(str).tolist()  # Convertir la columna SAP_ID a lista de cadenas

        # Crear un diccionario para mapear los nombres de archivos en el Excel con los archivos en el directorio.
        archivos_diccionario = {}
        for archivo in os.listdir(directorio_origen):  # Iterar sobre los archivos en el directorio de origen
            if archivo.endswith(".pdf"):  # Filtrar solo archivos PDF
                archivos_diccionario[archivo] = os.path.join(directorio_origen, archivo)  # Agregar al diccionario

        # Renombrar y mover los archivos según el orden del archivo de Excel.
        for i, nombre_archivo_excel in enumerate(archivos_ordenados_excel, 1):  # Iterar sobre la lista de archivos del Excel
            archivo_origen = None
            for nombre_archivo, ruta_archivo in archivos_diccionario.items():  # Iterar sobre el diccionario de archivos
                if nombre_archivo.startswith(nombre_archivo_excel):  # Buscar el archivo correspondiente
                    archivo_origen = ruta_archivo  # Obtener la ruta del archivo
                    break
            
            if archivo_origen:  # Si se encuentra el archivo
                nuevo_nombre = f"{i:05d}__{nombre_archivo_excel}.pdf"  # Crear el nuevo nombre del archivo
                archivo_destino = os.path.join(directorio_destino, nuevo_nombre)  # Crear la ruta de destino
                shutil.copy(archivo_origen, archivo_destino)  # Copiar el archivo al directorio de destino

                # Actualizar el valor de la barra de progreso.
                progress_value = i / len(archivos_ordenados_excel) * 100  # Calcular el progreso
                progress_bar['value'] = progress_value  # Actualizar la barra de progreso
                ventana_progreso.update()  # Actualizar la ventana emergente

                # Actualizar el texto del label_estado con el porcentaje.
                label_estado.config(text=f"PROCESANDO ARCHIVOS... {int(progress_value)}%")  # Actualizar la etiqueta de estado

        # Mostrar mensaje de finalización.
        num_archivos_renombrados = len(archivos_ordenados_excel)  # Número de archivos procesados
        label_estado.config(text=f"Se han renombrado y movido {num_archivos_renombrados} archivos PDF.")  # Mensaje de finalización
        
        # Cerrar la ventana de progreso al finalizar.
        ventana_progreso.after(2000, ventana_progreso.destroy)  # Cerrar la ventana después de 2 segundos

    # Iniciar el proceso de actualización del progreso.
    actualizar_progreso()  # Llamar a la función de actualización de progreso

# Función para seleccionar el directorio de origen.
def seleccionar_directorio_origen():
    directorio_origen = filedialog.askdirectory()  # Abrir un cuadro de diálogo para seleccionar un directorio
    entry_directorio_origen.delete(0, tk.END)  # Borrar el contenido actual de la entrada
    entry_directorio_origen.insert(0, directorio_origen)  # Insertar la ruta seleccionada en la entrada

# Función para seleccionar el directorio de destino.
def seleccionar_directorio_destino():
    directorio_destino = filedialog.askdirectory()  # Abrir un cuadro de diálogo para seleccionar un directorio
    entry_directorio_destino.delete(0, tk.END)  # Borrar el contenido actual de la entrada
    entry_directorio_destino.insert(0, directorio_destino)  # Insertar la ruta seleccionada en la entrada

# Función para seleccionar la ruta del archivo de Excel.
def seleccionar_ruta_excel():
    ruta_excel = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])  # Abrir un cuadro de diálogo para seleccionar un archivo
    entry_ruta_excel.delete(0, tk.END)  # Borrar el contenido actual de la entrada
    entry_ruta_excel.insert(0, ruta_excel)  # Insertar la ruta seleccionada en la entrada

# Función para limpiar los campos de entrada.
def limpiar_campos():
    entry_directorio_origen.delete(0, tk.END)  # Borrar el contenido de la entrada del directorio de origen
    entry_directorio_destino.delete(0, tk.END)  # Borrar el contenido de la entrada del directorio de destino
    entry_ruta_excel.delete(0, tk.END)  # Borrar el contenido de la entrada de la ruta del archivo de Excel

# Función para cerrar la ventana.
def cerrar_ventana():
    fade_out(root)  # Realiza una transición de fade-out antes de cerrar la ventana.
    root.after(1000, root.destroy)  # Espera un segundo antes de destruir la ventana.

# Crear la ventana principal.
root = tk.Tk()  # Inicializar la ventana principal de tkinter
root.title("RENOMBRAR ARCHIVOS PDF SIN ORDEN SAP_ID ")  # Establecer el título de la ventana

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
background_image_path = r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\fnd.png"  # Ruta de la imagen de fondo
background_image = tk.PhotoImage(file=background_image_path)  # Cargar la imagen de fondo
background_label = tk.Label(root, image=background_image)  # Crear una etiqueta con la imagen de fondo
background_label.place(relx=0.5, rely=0.6, anchor=tk.CENTER)  # Colocar la etiqueta en la ventana

# Mostrar la ventana gradualmente.
fade_in(root)  # Llamar a la función de fade-in para mostrar la ventana gradualmente

# Crear y colocar los elementos de la interfaz gráfica.
label_directorio_origen = tk.Label(root, text="Directorio Descargas:")  # Crear una etiqueta para el directorio de origen
label_directorio_origen.grid(row=0, column=0, sticky="e")  # Colocar la etiqueta en la ventana

entry_directorio_origen = tk.Entry(root, width=50)  # Crear una entrada de texto para el directorio de origen
entry_directorio_origen.grid(row=0, column=1)  # Colocar la entrada en la ventana

button_seleccionar_directorio_origen = tk.Button(root, text="Seleccionar", command=seleccionar_directorio_origen, width=15)  # Crear un botón para seleccionar el directorio de origen
button_seleccionar_directorio_origen.grid(row=0, column=2)  # Colocar el botón en la ventana

label_directorio_destino = tk.Label(root, text="Directorio Renombrados:")  # Crear una etiqueta para el directorio de destino
label_directorio_destino.grid(row=1, column=0, sticky="e")  # Colocar la etiqueta en la ventana

entry_directorio_destino = tk.Entry(root, width=50)  # Crear una entrada de texto para el directorio de destino
entry_directorio_destino.grid(row=1, column=1)  # Colocar la entrada en la ventana

button_seleccionar_directorio_destino = tk.Button(root, text="Seleccionar", command=seleccionar_directorio_destino, width=15)  # Crear un botón para seleccionar el directorio de destino
button_seleccionar_directorio_destino.grid(row=1, column=2)  # Colocar el botón en la ventana

label_ruta_excel = tk.Label(root, text="Ruta Base de Datos:")  # Crear una etiqueta para la ruta del archivo de Excel
label_ruta_excel.grid(row=2, column=0, sticky="e")  # Colocar la etiqueta en la ventana

entry_ruta_excel = tk.Entry(root, width=50)  # Crear una entrada de texto para la ruta del archivo de Excel
entry_ruta_excel.grid(row=2, column=1)  # Colocar la entrada en la ventana

button_seleccionar_ruta_excel = tk.Button(root, text="Seleccionar", command=seleccionar_ruta_excel, width=15)  # Crear un botón para seleccionar la ruta del archivo de Excel
button_seleccionar_ruta_excel.grid(row=2, column=2)  # Colocar el botón en la ventana

button_renombrar = tk.Button(root, text="RENOMBRAR ARCHIVOS", command=renombrar_archivos, width=20)  # Crear un botón para renombrar archivos
button_renombrar.config(bg="green", fg="white", width=20)  # Configurar el color de fondo y texto del botón
button_renombrar.grid(row=3, column=1)  # Colocar el botón en la ventana

button_limpiar = tk.Button(root, text="LIMPIAR CAMPOS", command=limpiar_campos, bg="orange", width=15)  # Crear un botón para limpiar los campos de entrada
button_limpiar.grid(row=3, column=0, columnspan=3, padx=(0, 360), pady=5)  # Colocar el botón en la ventana

button_cerrar = tk.Button(root, text="CERRAR", command=cerrar_ventana, bg="red", width=15)  # Crear un botón para cerrar la ventana
button_cerrar.grid(row=3, column=0, columnspan=3, padx=(443, 0), pady=5)  # Colocar el botón en la ventana

root.mainloop()  # Iniciar el bucle principal de la interfaz gráfica
