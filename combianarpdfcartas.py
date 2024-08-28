import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter

def add_blank_page_if_odd(reader, writer):
    num_pages = len(reader.pages)
    if num_pages % 2 != 0:  # Si el número de páginas es impar
        writer.add_blank_page()  # Agregar una página en blanco

def combine_pdfs(input_dir, output_file):
    writer = PdfWriter()

    # Listar todos los archivos PDF en el directorio
    pdf_files = [f for f in os.listdir(input_dir) if f.endswith('.pdf')]
    pdf_files.sort()  # Ordenar por nombre

    for pdf_file in pdf_files:
        pdf_path = os.path.join(input_dir, pdf_file)
        reader = PdfReader(pdf_path)

        # Limpiar metadatos problemáticos (si es necesario)
        reader.metadata.clear()

        # Copiar las páginas al escritor final
        for page in reader.pages:
            writer.add_page(page)

        # Agregar una página en blanco si es necesario
        if len(reader.pages) % 2 != 0:
            writer.add_blank_page()

    # Guardar el archivo combinado en la ubicación seleccionada
    with open(output_file, "wb") as output_pdf:
        writer.write(output_pdf)

def select_input_directory():
    input_dir = filedialog.askdirectory(title="Seleccionar directorio de PDFs")
    if input_dir:
        input_directory.set(input_dir)

def select_output_file():
    output_file = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
        title="Guardar archivo combinado como"
    )
    if output_file:
        output_path.set(output_file)

def combine_pdfs_gui():
    input_dir = input_directory.get()
    output_file = output_path.get()

    if not input_dir or not output_file:
        messagebox.showwarning("Advertencia", "Por favor, selecciona un directorio y un archivo de salida.")
        return

    try:
        combine_pdfs(input_dir, output_file)
        messagebox.showinfo("Éxito", f"PDFs combinados guardados en: {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"Se produjo un error: {str(e)}")

# Crear la ventana principal
root = tk.Tk()
root.title("Combinador de PDFs")

# Variables para almacenar las rutas
input_directory = tk.StringVar()
output_path = tk.StringVar()

# Crear los elementos de la GUI
tk.Label(root, text="Directorio de PDFs:").grid(row=0, column=0, padx=10, pady=10)
tk.Entry(root, textvariable=input_directory, width=50).grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Seleccionar", command=select_input_directory).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Guardar archivo combinado como:").grid(row=1, column=0, padx=10, pady=10)
tk.Entry(root, textvariable=output_path, width=50).grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Seleccionar", command=select_output_file).grid(row=1, column=2, padx=10, pady=10)

tk.Button(root, text="Combinar PDFs", command=combine_pdfs_gui).grid(row=2, column=0, columnspan=3, pady=20)

# Iniciar la aplicación
root.mainloop()
