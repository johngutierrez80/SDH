import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

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

    # Renombrar los archivos
    for pdf_file, new_name in zip(pdf_files, new_names):
        current_file = os.path.join(pdf_directory_path, pdf_file)
        new_file = os.path.join(output_directory, new_name + '.pdf')

        os.rename(current_file, new_file)
        print(f'Renombrado: {current_file} -> {new_file}')

    messagebox.showinfo("Completado", "Renombramiento completado.")

# Configuración de la ventana
root = tk.Tk()
root.title("Renombrador de Archivos PDF para Correo")

# Variables para almacenar las rutas
excel_file_path_var = tk.StringVar()
pdf_directory_path_var = tk.StringVar()
output_dir_var = tk.StringVar()
search_pattern_var = tk.StringVar()

# Entradas para la ruta del archivo Excel
tk.Label(root, text="Ruta del archivo Excel:").pack()
tk.Entry(root, textvariable=excel_file_path_var).pack()
tk.Button(root, text="Seleccionar Excel", 
          command=lambda: excel_file_path_var.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")]))).pack()

# Entradas para la ruta del directorio PDF
tk.Label(root, text="Ruta del directorio PDF:").pack()
tk.Entry(root, textvariable=pdf_directory_path_var).pack()
tk.Button(root, text="Seleccionar PDF Directory", 
          command=lambda: pdf_directory_path_var.set(filedialog.askdirectory())).pack()

# Entrada para el patrón de búsqueda
tk.Label(root, text="Patrón de Renombrado:").pack()
tk.Entry(root, textvariable=search_pattern_var).pack()

# Botón para seleccionar el directorio de salida
tk.Label(root, text="Directorio de salida:").pack()
tk.Entry(root, textvariable=output_dir_var).pack()
tk.Button(root, text="Seleccionar Directorio de Salida", command=select_output_directory).pack()

# Botón para ejecutar el renombramiento
tk.Button(root, text="Renombrar Archivos", command=rename_files).pack()

# Iniciar la aplicación
root.mainloop()
