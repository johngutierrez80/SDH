import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk
import os
import pandas as pd

def listar_archivos_y_ubicacion_carpeta_renombrados(ruta_carpeta):
    archivos = []
    ubicaciones = []
    for nombre_archivo in os.listdir(ruta_carpeta):
        ubicacion_archivo = os.path.join(ruta_carpeta, nombre_archivo)
        archivos.append(nombre_archivo)
        ubicaciones.append(ubicacion_archivo)
    return archivos, ubicaciones

def listar_archivos_y_ubicacion_carpeta_labels(ruta_carpeta):
    archivos_labels = []
    ubicaciones_labels = []
    for nombre_archivo in os.listdir(ruta_carpeta):
        ubicacion_archivo = os.path.join(ruta_carpeta, nombre_archivo)
        # Excluir el archivo "Thumbs.db"
        if nombre_archivo != "Thumbs.db":
            archivos_labels.append(nombre_archivo)
            ubicaciones_labels.append(ubicacion_archivo)
    return archivos_labels, ubicaciones_labels

def seleccionar_carpeta_renombrados():
    ruta_carpeta_renombrados = filedialog.askdirectory()
    if ruta_carpeta_renombrados:
        ruta_carpeta_renombrados_entry.delete(0, tk.END)
        ruta_carpeta_renombrados_entry.insert(tk.END, ruta_carpeta_renombrados)
    else:
        status_label.config(text="No se ha seleccionado ninguna carpeta de archivos Renombrados.")

def seleccionar_carpeta_labels_original():
    ruta_carpeta_labels_original = filedialog.askdirectory()
    if ruta_carpeta_labels_original:
        ruta_carpeta_labels_original_entry.delete(0, tk.END)
        ruta_carpeta_labels_original_entry.insert(tk.END, ruta_carpeta_labels_original)
    else:
        status_label_original.config(text="No se ha seleccionado ninguna carpeta de archivos Labels (Original).")

def seleccionar_carpeta_labels_copia():
    ruta_carpeta_labels_copia = filedialog.askdirectory()
    if ruta_carpeta_labels_copia:
        ruta_carpeta_labels_copia_entry.delete(0, tk.END)
        ruta_carpeta_labels_copia_entry.insert(tk.END, ruta_carpeta_labels_copia)
    else:
        status_label_copia.config(text="No se ha seleccionado ninguna carpeta de archivos Labels (Copia).")

def seleccionar_ruta_excel_original():
    default_filename = "ARCHIVO_FUENTE_ORIGINAL.xlsx"  # Nombre predeterminado del archivo
    ruta_excel_original = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivos de Excel", "*.xlsx")], initialfile=default_filename)
    if ruta_excel_original:
        ruta_excel_original_entry.delete(0, tk.END)
        ruta_excel_original_entry.insert(tk.END, ruta_excel_original)

def seleccionar_ruta_excel_copia():
    default_filename = "ARCHIVO_FUENTE_COPIA.xlsx"  # Nombre predeterminado del archivo
    ruta_excel_copia = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivos de Excel", "*.xlsx")], initialfile=default_filename)
    if ruta_excel_copia:
        ruta_excel_copia_entry.delete(0, tk.END)
        ruta_excel_copia_entry.insert(tk.END, ruta_excel_copia)

def limpiar_campos_seleccion():
    ruta_carpeta_renombrados_entry.delete(0, tk.END)
    ruta_carpeta_labels_original_entry.delete(0, tk.END)
    ruta_carpeta_labels_copia_entry.delete(0, tk.END)
    ruta_excel_original_entry.delete(0, tk.END)
    ruta_excel_copia_entry.delete(0, tk.END)
    status_label.config(text="Campos de selección limpiados.")
    status_label_original.config(text="")
    status_label_copia.config(text="")

def crear_dataframes_y_escribir_excel():
    ruta_carpeta_renombrados = ruta_carpeta_renombrados_entry.get()
    ruta_carpeta_labels_original = ruta_carpeta_labels_original_entry.get()
    ruta_carpeta_labels_copia = ruta_carpeta_labels_copia_entry.get()
    ruta_excel_original = ruta_excel_original_entry.get()
    ruta_excel_copia = ruta_excel_copia_entry.get()

    if ruta_carpeta_renombrados and ruta_carpeta_labels_original and ruta_carpeta_labels_copia and ruta_excel_original and ruta_excel_copia:
        archivos_renombrados, ubicaciones_renombrados = listar_archivos_y_ubicacion_carpeta_renombrados(ruta_carpeta_renombrados)
        archivos_labels_original, ubicaciones_labels_original = listar_archivos_y_ubicacion_carpeta_labels(ruta_carpeta_labels_original)
        archivos_labels_copia, ubicaciones_labels_copia = listar_archivos_y_ubicacion_carpeta_labels(ruta_carpeta_labels_copia)

        # Crear DataFrame para el archivo fuente original
        df_fuente_original = pd.DataFrame({"Archivo": archivos_renombrados, 
                                           "Ubicación": ubicaciones_renombrados,
                                           "Label": archivos_labels_original,
                                           "Ubicación Label": ubicaciones_labels_original})

        # Crear DataFrame para el archivo fuente de copia
        df_fuente_copia = pd.DataFrame({"Archivo": archivos_renombrados, 
                                        "Ubicación": ubicaciones_renombrados,
                                        "Label": archivos_labels_copia,
                                        "Ubicación Label": ubicaciones_labels_copia})

        with pd.ExcelWriter(ruta_excel_original, engine="openpyxl", mode="w") as writer_original:
            df_fuente_original.to_excel(writer_original, index=False, header=False)

        with pd.ExcelWriter(ruta_excel_copia, engine="openpyxl", mode="w") as writer_copia:
            df_fuente_copia.to_excel(writer_copia, index=False, header=False)

        status_label_original.config(text="Datos guardados exitosamente en {}".format(ruta_excel_original))
        status_label_copia.config(text="Datos guardados exitosamente en {}".format(ruta_excel_copia))
    else:
        status_label.config(text="Por favor complete todas las rutas para generar los archivos fuente.")
        status_label_original.config(text="")
        status_label_copia.config(text="")
        
def cerrar_aplicativo():
    root.destroy()

root = tk.Tk()
root.title("ARCHIVOS FUENTE")

# Calcula el ancho y el alto de la ventana
ancho_ventana = 820
alto_ventana = 420

# Obtiene el ancho y el alto de la pantalla
ancho_pantalla = root.winfo_screenwidth()
alto_pantalla = root.winfo_screenheight()

# Calcula la posición x y y para centrar la ventana
x = (ancho_pantalla // 2) - (ancho_ventana // 2)
y = (alto_pantalla // 2) - (alto_ventana // 2)

# Establece la geometría de la ventana para que aparezca centrada en la pantalla
root.geometry(f"{ancho_ventana}x{alto_ventana}+{x}+{y}")
icon_image = Image.open(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\INC.ico")
icon_photo = ImageTk.PhotoImage(icon_image)
root.iconphoto(False, icon_photo)

# Configurar el fondo de la ventana
background_image_path = r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\fnd2.png"
background_image = Image.open(background_image_path)
background_photo = ImageTk.PhotoImage(background_image)
background_label = tk.Label(root, image=background_photo)
background_label.place(x=-5, y=30, relwidth=1, relheight=1)

ruta_carpeta_renombrados_label = tk.Label(root, text="Carpeta de Archivos Renombrados:")
ruta_carpeta_renombrados_label.grid(row=0, column=0, padx=(20, 0), sticky="w")
ruta_carpeta_renombrados_entry = tk.Entry(root, width=50)
ruta_carpeta_renombrados_entry.grid(row=0, column=1, padx=(20, 0), pady=5)

select_folder_renombrados_button = tk.Button(root, text="Seleccionar Carpeta Renombrados", command=seleccionar_carpeta_renombrados,width=30)
select_folder_renombrados_button.grid(row=0, column=2, padx=(30, 0), pady=5)

ruta_carpeta_labels_original_label = tk.Label(root, text="Carpeta de Archivos Labels (Original):")
ruta_carpeta_labels_original_label.grid(row=1, column=0, padx=(20, 0),sticky="w")
ruta_carpeta_labels_original_entry = tk.Entry(root, width=50)
ruta_carpeta_labels_original_entry.grid(row=1, column=1, padx=(20, 0), pady=5)

select_folder_labels_original_button = tk.Button(root, text="Seleccionar Carpeta Labels (Original)", command=seleccionar_carpeta_labels_original,width=30)
select_folder_labels_original_button.grid(row=1, column=2, padx=(27, 0), pady=5)

ruta_carpeta_labels_copia_label = tk.Label(root, text="Carpeta de Archivos Labels (Copia):")
ruta_carpeta_labels_copia_label.grid(row=2, column=0, padx=(20, 0),sticky="w")
ruta_carpeta_labels_copia_entry = tk.Entry(root, width=50)
ruta_carpeta_labels_copia_entry.grid(row=2, column=1, padx=(20, 0), pady=5)

select_folder_labels_copia_button = tk.Button(root, text="Seleccionar Carpeta Labels (Copia)", command=seleccionar_carpeta_labels_copia,width=30)
select_folder_labels_copia_button.grid(row=2, column=2, padx=(27, 0), pady=5)

ruta_excel_original_label = tk.Label(root, text="Ruta de Archivo Fuente (Original):")
ruta_excel_original_label.grid(row=3, column=0, padx=(20, 0), sticky="w")
ruta_excel_original_entry = tk.Entry(root, width=50)
ruta_excel_original_entry.grid(row=3, column=1, padx=(20, 0), pady=5)

select_excel_original_button = tk.Button(root, text="Salida original", command=seleccionar_ruta_excel_original,width=30)
select_excel_original_button.grid(row=3, column=2, padx=(25, 0), pady=5)

ruta_excel_copia_label = tk.Label(root, text="Ruta de Archivo Fuente (Copia):")
ruta_excel_copia_label.grid(row=4, column=0, padx=(20, 0), sticky="w")
ruta_excel_copia_entry = tk.Entry(root, width=50)
ruta_excel_copia_entry.grid(row=4, column=1, padx=(20, 0), pady=5)

select_excel_copia_button = tk.Button(root, text="Salida copia", command=seleccionar_ruta_excel_copia,width=30)
select_excel_copia_button.grid(row=4, column=2, padx=(25, 0), pady=5)

limpiar_campos_button = tk.Button(root, text="Limpiar Campos", command=limpiar_campos_seleccion, bg="orange",width=15)
limpiar_campos_button.grid(row=5, column=0, columnspan=3, padx=(0,550), pady=20)

crear_excel_button = tk.Button(root, text="Crear Excel", command=crear_dataframes_y_escribir_excel, bg="green", fg="white",width=15)
crear_excel_button.grid(row=5, column=0, columnspan=3, padx=(570,0), pady=5)

cerrar_aplicativo_button = tk.Button(root, text="CERRAR", command=cerrar_aplicativo, bg="red", fg="white",width=15)
cerrar_aplicativo_button.grid(row=6, column=0, columnspan=3, padx=(570,0), pady=5)

status_label = tk.Label(root, text="", bd=1, relief=tk.SUNKEN, anchor=tk.W)
status_label.grid(row=7, column=0, columnspan=3, padx=(20, 0), pady=15, sticky="we")

status_label_original = tk.Label(root, text="", bd=1, relief=tk.SUNKEN, anchor=tk.W)
status_label_original.grid(row=8, column=0, columnspan=3, padx=(20, 0), pady=15, sticky="we")

status_label_copia = tk.Label(root, text="", bd=1, relief=tk.SUNKEN, anchor=tk.W)
status_label_copia.grid(row=9, column=0, columnspan=3, padx=(20, 0), pady=15, sticky="we")

root.mainloop()
