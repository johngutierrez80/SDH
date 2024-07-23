import os
import fitz  # Para trabajar con archivos PDF
from math import ceil
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import threading

def count_files_folders_and_sheets(directory, progress_var, progress_label_var, op_selected):
    file_count = 0
    folder_count = 0
    total_sheets = 0
    file_names = []

    entries = os.listdir(directory)
    total_entries = len(entries)
    
    for i, entry in enumerate(entries):
        path = os.path.join(directory, entry)
        if os.path.isfile(path):
            file_count += 1
            file_names.append(entry)
            if path.endswith('.pdf'):
                try:
                    pdf = fitz.open(path)
                    num_pages = pdf.page_count
                    if op_selected == "307138":
                        total_sheets += ceil(num_pages / 2) * 2
                    else:
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

    return file_count, folder_count, total_sheets, file_names

def write_count_to_file(directory, output_file, progress_var, progress_label_var, include_filenames, op_selected):
    file_count, folder_count, total_sheets, file_names = count_files_folders_and_sheets(directory, progress_var, progress_label_var, op_selected)
    
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    directory_name = "ACTOS GG/SHD" if op_selected == "307138" else os.path.basename(directory)
    
    with open(output_file, 'w') as f:
        f.write(f"LS - {directory_name}\n")
        f.write(f"Fecha de creacion: {now}\n\n")
        f.write(f'Orden OP: {op_selected}\n')
        f.write(f'Total number of folders = {folder_count}\n')
        f.write(f'Total number of files = {file_count} Archivos\n')
        f.write(f'Total number of sheets (pages in PDF files) = {total_sheets} hojas\n')
        if include_filenames:
            f.write('\nFile names:\n')
            for name in file_names:
                f.write(f'{name}\n')

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

    # Definir tamaño y centrar ventana emergente
    progress_window_width = 200
    progress_window_height = 80
    center_window(progress_window, progress_window_width, progress_window_height)
    
    progress_label = tk.Label(progress_window, textvariable=progress_label_var)
    progress_label.pack(padx=20, pady=10)
    
    progress_bar = ttk.Progressbar(progress_window, variable=progress_var, maximum=100)
    progress_bar.pack(padx=20, pady=10, fill=tk.X)
    
    root.update_idletasks()
    
    def task():
        write_count_to_file(directory, output_file, progress_var, progress_label_var, include_filenames, op_selected)
        progress_label_var.set("¡Completado!")
        progress_bar['value'] = 100
        messagebox.showinfo("Éxito", f"El archivo ha sido guardado en {output_file}.")
        progress_window.after(2000, progress_window.destroy)
    
    threading.Thread(target=task).start()

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
root.title("File List Generator")
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
bg_label.place(relx=0, rely=0, relwidth=1, relheight=1)
bg_label.bind('<Configure>', resize_bg_image)

# Variables de progreso
progress_var = tk.DoubleVar()
progress_label_var = tk.StringVar()

# Crear y colocar los widgets usando place
directory_label = tk.Label(root, text="Directory:", bg="white", font=("Arial", 10, "bold"))
directory_label.place(relx=0.2, rely=0.05)
directory_entry = tk.Entry(root, width=31)
directory_entry.place(relx=0.35, rely=0.05)
directory_button = tk.Button(root, text="Browse", command=select_directory, font=("Arial", 10, "bold"))
directory_button.place(relx=0.77, rely=0.05, width=88)

output_label = tk.Label(root, text="Output:", bg="white", font=("Arial", 10, "bold"))
output_label.place(relx=0.2, rely=0.15)
output_entry = tk.Entry(root, width=31)
output_entry.place(relx=0.35, rely=0.15)
output_button = tk.Button(root, text="Browse", command=select_output_file, font=("Arial", 10, "bold"))
output_button.place(relx=0.77, rely=0.15, width=88)

op_label = tk.Label(root, text="Order OP:", bg="white", font=("Arial", 10, "bold"))
op_label.place(relx=0.2, rely=0.3)
op_var = tk.StringVar()
op_combobox = ttk.Combobox(root, textvariable=op_var, values=["307137", "307138"], font=("Arial", 10, "bold"))
op_combobox.place(relx=0.365, rely=0.3109, width=108)

include_filenames_var = tk.BooleanVar()
include_filenames_checkbutton = tk.Checkbutton(root, text="Files", variable=include_filenames_var, bg="white", font=("Arial", 10, "bold"))
include_filenames_checkbutton.place(relx=0.2, rely=0.4)

generate_button = tk.Button(root, text="Run", command=generate_report, bg="green", font=("Arial", 10, "bold"), width=8, height=2)
generate_button.place(relx=0.54, rely=0.78)

close_button = tk.Button(root, text="Exit", command=close_app, bg="red", fg="white", font=("Arial", 10, "bold"), width=8, height=2)
close_button.place(relx=0.79, rely=0.78)

# Iniciar el bucle principal
root.mainloop()
