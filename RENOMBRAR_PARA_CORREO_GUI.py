import os
import shutil  # Nuevo módulo para copiar archivos
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk

# Función para seleccionar el directorio de salida
def select_output_directory():
    output_directory = filedialog.askdirectory()
    if output_directory:
        output_dir_var.set(output_directory)

# Función para renombrar archivos
def rename_files():
    excel_file_path = excel_file_path_var.get()
    pdf_directory_path = pdf_directory_path_var.get()
    output_directory = output_dir_var.get()
    search_pattern = search_pattern_var.get()

    # Leer el archivo Excel
    df = pd.read_excel(excel_file_path)

    # Verificar si la columna 'nombre' existe
    if 'nombre' not in df.columns:
        messagebox.showerror("Error", "La columna 'nombre' no se encuentra en el archivo Excel.")
        return

    # Obtener la lista de nuevos nombres
    new_names = df['nombre'].astype(str).tolist()

    # Listar los archivos PDF en el directorio según el patrón especificado
    pdf_files = sorted([f for f in os.listdir(pdf_directory_path) 
                        if f.endswith('.pdf') and search_pattern in f])

    if len(pdf_files) != len(new_names):
        messagebox.showerror("Error", "El número de archivos PDF no coincide con el número de nuevos nombres en el archivo Excel.")
        return

    # Copiar y renombrar los archivos
    for pdf_file, new_name in zip(pdf_files, new_names):
        current_file = os.path.join(pdf_directory_path, pdf_file)
        new_file = os.path.join(output_directory, new_name + '.pdf')

        shutil.copy(current_file, new_file)  # Cambiado para copiar
        print(f'Copiado: {current_file} -> {new_file}')

    messagebox.showinfo("Completado", "Copiado y renombrado completado.")

# Configuración de la ventana
root = tk.Tk()
root.title("Renombrador de Archivos PDF para Correo")

# Tamaño de la ventana principal
root.geometry("800x600")

# Cargar la imagen de fondo
background_image = Image.open(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\correo.png")
background_image = background_image.resize((800, 600), Image.LANCZOS)
background_photo = ImageTk.PhotoImage(background_image)

# Crear un Label para mostrar la imagen de fondo
background_label = tk.Label(root, image=background_photo)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

# Variables para almacenar las rutas
excel_file_path_var = tk.StringVar()
pdf_directory_path_var = tk.StringVar()
output_dir_var = tk.StringVar()
search_pattern_var = tk.StringVar()

# Entradas y botones sobre la imagen con el método `place`
tk.Label(root, text="Ruta de Dicconario Excel:", bg="white").place(x=50, y=30)
tk.Entry(root, textvariable=excel_file_path_var, width=50).place(x=50, y=60)
tk.Button(root, text="Seleccionar Excel", command=lambda: excel_file_path_var.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")]))).place(x=400, y=60)

tk.Label(root, text="Ruta del directorio PDF:", bg="white").place(x=50, y=110)
tk.Entry(root, textvariable=pdf_directory_path_var, width=50).place(x=50, y=140)
tk.Button(root, text="Seleccionar PDF Directory", 
          command=lambda: pdf_directory_path_var.set(filedialog.askdirectory())).place(x=400, y=140)

tk.Label(root, text="Patrón de Renombrado:", bg="white").place(x=50, y=190)
tk.Entry(root, textvariable=search_pattern_var, width=50).place(x=50, y=220)

tk.Label(root, text="Directorio de salida:", bg="white").place(x=50, y=270)
tk.Entry(root, textvariable=output_dir_var, width=50).place(x=50, y=300)
tk.Button(root, text="Seleccionar Directorio de Salida", command=select_output_directory).place(x=400, y=300)

tk.Button(root, text="Renombrar Archivos", command=rename_files).place(x=50, y=350)

# Iniciar la aplicación
root.mainloop()
