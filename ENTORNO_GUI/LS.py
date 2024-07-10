import os
import fitz  # Para trabajar con archivos PDF
from math import ceil
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk

def count_files_folders_and_sheets(directory, progress_var, progress_label_var):
    file_count = 0
    folder_count = 0
    total_sheets = 0
    file_names = []
    file_paths = []

    entries = os.listdir(directory)
    total_entries = len(entries)
    
    for i, entry in enumerate(entries):
        path = os.path.join(directory, entry)
        if os.path.isfile(path):
            file_count += 1
            file_names.append(entry)
            file_paths.append(path)
            if path.endswith('.pdf'):
                try:
                    pdf = fitz.open(path)
                    num_pages = pdf.page_count
                    total_sheets += ceil(num_pages / 2)
                    pdf.close()
                except Exception as e:
                    print(f'Error al procesar el archivo PDF: {path}, Error: {e}')
        elif os.path.isdir(path):
            folder_count += 1

        # Actualizar la barra de progreso
        progress_var.set((i + 1) / total_entries * 100)
        progress_label_var.set(f"Procesando: {i + 1}/{total_entries}")
        root.update_idletasks()

    return file_count, folder_count, total_sheets, file_names, file_paths

def write_count_to_file(directory, output_file, progress_var, progress_label_var, include_filenames, include_filepaths, op_selected):
    file_count, folder_count, total_sheets, file_names, file_paths = count_files_folders_and_sheets(directory, progress_var, progress_label_var)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    directory_name = os.path.basename(directory)
    with open(output_file, 'w') as f:
        f.write(f"LS - {directory_name}\n")
        f.write(f"Fecha de creación: {now}\n\n")
        f.write(f'Orden OP: {op_selected}\n')
        f.write(f'Total number of folders = {folder_count}\n')
        f.write(f'Total number of files = {file_count}\n')
        f.write(f'Total number of sheets (pages in PDF files) = {total_sheets}\n')
        if include_filenames:
            f.write('\nFile names:\n')
            for name in file_names:
                f.write(f'{name}\n')
        if include_filepaths:
            f.write('\nFile paths:\n')
            for path in file_paths:
                f.write(f'{path}\n')

def select_directory():
    directory = filedialog.askdirectory()
    if directory:
        directory_entry.delete(0, tk.END)
        directory_entry.insert(0, directory)

def select_output_file():
    file = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
    if file:
        output_entry.delete(0, tk.END)
        output_entry.insert(0, file)

def generate_report():
    directory = directory_entry.get()
    output_file = output_entry.get()
    include_filenames = include_filenames_var.get()
    include_filepaths = include_filepaths_var.get()
    op_selected = op_var.get()
    
    if not directory or not os.path.exists(directory):
        messagebox.showerror("Error", "Por favor, selecciona un directorio válido.")
        return
    
    if not output_file:
        messagebox.showerror("Error", "Por favor, selecciona un archivo de salida válido.")
        return
    
    if not op_selected:
        messagebox.showerror("Error", "Por favor, selecciona una OP válida.")
        return

    progress_var.set(0)
    progress_label_var.set("Iniciando...")
    
    progress_window = tk.Toplevel(root)
    progress_window.title("Progreso")
    
    progress_label = tk.Label(progress_window, textvariable=progress_label_var)
    progress_label.pack(padx=20, pady=10)
    
    progress_bar = ttk.Progressbar(progress_window, variable=progress_var, maximum=100)
    progress_bar.pack(padx=20, pady=10, fill=tk.X)
    
    root.update_idletasks()
    
    write_count_to_file(directory, output_file, progress_var, progress_label_var, include_filenames, include_filepaths, op_selected)
    
    progress_label_var.set("¡Completado!")
    progress_bar['value'] = 100
    
    messagebox.showinfo("Éxito", f"El archivo ha sido guardado en {output_file}.")
    progress_window.after(2000, progress_window.destroy)

def close_app():
    root.destroy()

def center_window(window, width=600, height=300):
    window.geometry(f'{width}x{height}')
    window.update_idletasks()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f'+{x}+{y}')

def resize_bg_image(event):
    new_width = event.width
    new_height = event.height
    bg_image_resized = bg_image.resize((new_width, new_height), Image.LANCZOS)
    bg_photo_resized = ImageTk.PhotoImage(bg_image_resized)
    bg_label.config(image=bg_photo_resized)
    bg_label.image = bg_photo_resized

# Crear la ventana principal
root = tk.Tk()
root.title("Generador LS Cartas")
icon_image = Image.open(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\INC.ico")
icon_photo = ImageTk.PhotoImage(icon_image)
root.iconphoto(False, icon_photo) 

# Definir tamaño de ventana fijo
window_width = 600
window_height = 300
center_window(root, window_width, window_height)

# Imagen de fondo
bg_image = Image.open(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\archivo-txt.png")
bg_image_resized = bg_image.resize((window_width, window_height), Image.LANCZOS)
bg_photo = ImageTk.PhotoImage(bg_image_resized)
bg_label = tk.Label(root, image=bg_photo)
bg_label.place(x=0, y=0, relwidth=1, relheight=1)
bg_label.bind('<Configure>', resize_bg_image)

# Variables de progreso
progress_var = tk.DoubleVar()
progress_label_var = tk.StringVar()

# Crear y colocar los widgets usando place
directory_label = tk.Label(root, text="Directorio:", bg="white", font=("Arial", 10, "bold"))
directory_label.place(x=20, y=20)
directory_entry = tk.Entry(root, width=50)
directory_entry.place(x=180, y=20)
directory_button = tk.Button(root, text="Buscar", command=select_directory, font=("Arial", 10, "bold"))
directory_button.place(x=500, y=18)

output_label = tk.Label(root, text="Archivo de salida:", bg="white", font=("Arial", 10, "bold"))
output_label.place(x=20, y=60)
output_entry = tk.Entry(root, width=50)
output_entry.place(x=180, y=60)
output_button = tk.Button(root, text="Buscar", command=select_output_file, font=("Arial", 10, "bold"))
output_button.place(x=500, y=58)

op_label = tk.Label(root, text="Orden OP:", bg="white", font=("Arial", 10, "bold"))
op_label.place(x=20, y=100)
op_var = tk.StringVar()
op_combobox = ttk.Combobox(root, textvariable=op_var, values=["307137", "307138"], font=("Arial", 10, "bold"))
op_combobox.place(x=180, y=100)

include_filenames_var = tk.BooleanVar()
include_filenames_checkbutton = tk.Checkbutton(root, text="Nombres de archivos", variable=include_filenames_var, bg="white", font=("Arial", 10, "bold"))
include_filenames_checkbutton.place(x=20, y=140)

include_filepaths_var = tk.BooleanVar()
include_filepaths_checkbutton = tk.Checkbutton(root, text="Paths de archivos", variable=include_filepaths_var, bg="white", font=("Arial", 10, "bold"))
include_filepaths_checkbutton.place(x=200, y=140)

generate_button = tk.Button(root, text="Run", command=generate_report, bg="green", font=("Arial", 10, "bold"), width=10, height=2)
generate_button.place(x=200, y=250)

close_button = tk.Button(root, text="Exit", command=close_app, bg="red", fg="white", font=("Arial", 10, "bold"), width=10, height=2)
close_button.place(x=400, y=250)

# Iniciar el bucle principal
root.mainloop()
