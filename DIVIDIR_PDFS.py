import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
from PyPDF2 import PdfReader, PdfWriter

def split_pdf(input_pdf, output_dir, pages_per_split, base_filename, start_number, progress_bar, progress_label, total_pages):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    pdf_reader = PdfReader(input_pdf)
    file_count = start_number

    for i in range(0, total_pages, pages_per_split):
        pdf_writer = PdfWriter()
        
        for j in range(i, min(i + pages_per_split, total_pages)):
            pdf_writer.add_page(pdf_reader.pages[j])
        
        output_filename = f"{str(file_count).zfill(5)}_{base_filename}_{str(file_count).zfill(5)}.pdf"
        output_filepath = os.path.join(output_dir, output_filename)
        
        with open(output_filepath, 'wb') as output_pdf:
            pdf_writer.write(output_pdf)

        file_count += 1

        # Calcular y actualizar el progreso
        progress_value = ((i + pages_per_split) / total_pages) * 100
        progress_bar['value'] = progress_value
        progress_label.config(text=f"{progress_value:.2f}%")  # Mostrar porcentaje con dos decimales
        progress_bar.update_idletasks()  # Refresca la barra de progreso y el porcentaje

    messagebox.showinfo("Completado", f"Se ha completado la división del PDF en {file_count - start_number} archivos.")

def browse_file():
    filename = filedialog.askopenfilename(
        title="Seleccionar archivo PDF",
        filetypes=[("PDF files", "*.pdf")]
    )
    pdf_path.set(filename)

def browse_directory():
    directory = filedialog.askdirectory(
        title="Seleccionar directorio de salida"
    )
    output_path.set(directory)

def start_split():
    input_pdf = pdf_path.get()
    output_dir = output_path.get()
    pages_per_split = int(pages_entry.get())
    base_filename = base_filename_entry.get()
    start_number = int(start_number_entry.get())

    if not input_pdf or not output_dir or not base_filename:
        messagebox.showwarning("Advertencia", "Por favor, completa todos los campos.")
        return

    pdf_reader = PdfReader(input_pdf)
    total_pages = len(pdf_reader.pages)

    # Crear ventana emergente con barra de progreso
    progress_window = tk.Toplevel(root)
    progress_window.title("Progreso")
    progress_window.geometry("350x150")  # Tamaño de la ventana emergente

    # Etiqueta para el mensaje de progreso
    tk.Label(progress_window, text="Dividiendo PDF...").pack(pady=10)

    # Barra de progreso
    progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=300, mode="determinate")
    progress_bar.pack(pady=10)

    # Etiqueta para mostrar el porcentaje de progreso
    progress_label = tk.Label(progress_window, text="0.00%")
    progress_label.pack()

    # Ejecuta la función de división con la barra de progreso y el porcentaje
    root.after(100, lambda: split_pdf(input_pdf, output_dir, pages_per_split, base_filename, start_number, progress_bar, progress_label, total_pages))

# Configuración de la ventana principal
root = tk.Tk()
root.title("Divisor de PDFs")

# Obtener el tamaño de la pantalla
window_width = 750
window_height = 550
root.geometry(f"{window_width}x{window_height}")

# Cargar y redimensionar la imagen de fondo usando Pillow
image = Image.open(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\pdf.png")
image = image.resize((window_width, window_height), Image.Resampling.LANCZOS)
background_image = ImageTk.PhotoImage(image)

# Configurar la imagen de fondo
background_label = tk.Label(root, image=background_image)
background_label.place(relwidth=1, relheight=1)

# Variables de entrada
pdf_path = tk.StringVar()
output_path = tk.StringVar()

# Etiquetas y campos de entrada
tk.Label(root, text="Archivo PDF:", bg='#ffffff').grid(row=0, column=0, padx=10, pady=10)
tk.Entry(root, textvariable=pdf_path, width=50).grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Seleccionar archivo", command=browse_file).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Directorio de salida:", bg='#ffffff').grid(row=1, column=0, padx=10, pady=10)
tk.Entry(root, textvariable=output_path, width=50).grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Seleccionar directorio", command=browse_directory).grid(row=1, column=2, padx=10, pady=10)

tk.Label(root, text="Número de páginas por división:", bg='#ffffff').grid(row=2, column=0, padx=10, pady=10)
pages_entry = tk.Entry(root, width=10)
pages_entry.grid(row=2, column=1, padx=10, pady=10)

tk.Label(root, text="Nombre base para los archivos de salida:", bg='#ffffff').grid(row=3, column=0, padx=10, pady=10)
base_filename_entry = tk.Entry(root, width=20)
base_filename_entry.grid(row=3, column=1, padx=10, pady=10)

tk.Label(root, text="Número inicial de la numeración:", bg='#ffffff').grid(row=4, column=0, padx=10, pady=10)
start_number_entry = tk.Entry(root, width=10)
start_number_entry.grid(row=4, column=1, padx=10, pady=10)

# Botón para iniciar la división
tk.Button(root, text="Iniciar división", command=start_split, bg="green", fg="white",width=15, height=2).grid(row=50, column=2, padx=10, pady=20)

# Botón para salir de la aplicación, con tamaño ajustado
tk.Button(root, text="Salir", command=root.quit, bg="red", fg="white", width=15, height=2).grid(row=100, column=2, padx=10, pady=20)

# Iniciar la aplicación
root.mainloop()
