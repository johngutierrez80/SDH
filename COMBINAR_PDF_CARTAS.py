import os
from tkinter import Tk, StringVar, Label, Entry, Button, filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter
from typing import List

def add_blank_page(writer: PdfWriter) -> None:
    """Añade una página en blanco al PdfWriter."""
    writer.add_blank_page()

def combine_pdfs(input_dir: str, output_file: str, preserve_order: bool = True) -> int:
    """Combina PDFs, añadiendo página en blanco si es impar, y retorna el número total de páginas esperadas."""
    writer = PdfWriter()
    
    # Listar archivos PDF
    pdf_files = [f for f in os.listdir(input_dir) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        raise ValueError("No se encontraron archivos PDF en el directorio seleccionado.")
    
    # Ordenar o mantener orden original
    if not preserve_order:
        pdf_files.sort()  # Orden alfabético
    # Si preserve_order=True, usa el orden del directorio (orden natural)

    total_original_pages = 0
    blank_pages_added = 0

    for pdf_file in pdf_files:
        pdf_path = os.path.join(input_dir, pdf_file)
        try:
            reader = PdfReader(pdf_path)
            num_pages = len(reader.pages)
            total_original_pages += num_pages
            
            # Añadir todas las páginas del PDF actual
            for page in reader.pages:
                writer.add_page(page)
            
            # Añadir página en blanco si el número de páginas es impar
            if num_pages % 2 != 0:
                add_blank_page(writer)
                blank_pages_added += 1
                
        except Exception as e:
            raise RuntimeError(f"Error procesando {pdf_file}: {str(e)}")

    # Guardar el resultado
    with open(output_file, "wb") as output_pdf:
        writer.write(output_pdf)
    
    # Retornar número esperado de páginas (originales + en blanco)
    return total_original_pages + blank_pages_added

def verify_pdf_content(input_dir: str, output_file: str, expected_pages: int) -> bool:
    """Verifica que el PDF combinado tenga el número esperado de páginas."""
    try:
        combined_reader = PdfReader(output_file)
        actual_pages = len(combined_reader.pages)
        return actual_pages == expected_pages
    except Exception as e:
        raise RuntimeError(f"Error verificando el PDF combinado: {str(e)}")

def select_input_directory() -> None:
    """Abre un diálogo para seleccionar el directorio de entrada."""
    input_dir = filedialog.askdirectory(title="Seleccionar directorio de PDFs")
    if input_dir:
        input_directory.set(input_dir)

def select_output_file() -> None:
    """Abre un diálogo para seleccionar el archivo de salida."""
    output_file = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
        title="Guardar archivo combinado como"
    )
    if output_file:
        output_path.set(output_file)

def combine_pdfs_gui() -> None:
    """Función principal para la GUI que ejecuta la combinación y verificación de PDFs."""
    input_dir = input_directory.get()
    output_file = output_path.get()

    if not input_dir or not output_file:
        messagebox.showwarning("Advertencia", "Por favor, selecciona un directorio y un archivo de salida.")
        return

    try:
        # Combinar PDFs y obtener el número esperado de páginas
        expected_pages = combine_pdfs(input_dir, output_file, preserve_order=True)
        
        # Verificar que el contenido esté completo
        if verify_pdf_content(input_dir, output_file, expected_pages):
            messagebox.showinfo("Éxito", f"PDFs combinados guardados en: {output_file}\nVerificación: Contenido completo.")
        else:
            messagebox.showwarning("Advertencia", f"PDFs combinados en: {output_file}\nPero el contenido podría estar incompleto.")
            
    except ValueError as ve:
        messagebox.showerror("Error", str(ve))
    except RuntimeError as re:
        messagebox.showerror("Error", str(re))
    except Exception as e:
        messagebox.showerror("Error", f"Error inesperado: {str(e)}")

# Configuración de la ventana principal
root = Tk()
root.title("Combinador de PDFs")
root.geometry("700x200")

# Variables para almacenar las rutas
input_directory = StringVar()
output_path = StringVar()

# Elementos de la GUI
Label(root, text="Directorio de PDFs:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
Entry(root, textvariable=input_directory, width=50).grid(row=0, column=1, padx=10, pady=10)
Button(root, text="Seleccionar", command=select_input_directory).grid(row=0, column=2, padx=10, pady=10)

Label(root, text="Guardar archivo combinado como:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
Entry(root, textvariable=output_path, width=50).grid(row=1, column=1, padx=10, pady=10)
Button(root, text="Seleccionar", command=select_output_file).grid(row=1, column=2, padx=10, pady=10)

Button(root, text="Combinar PDFs", command=combine_pdfs_gui).grid(row=2, column=1, pady=20)

# Iniciar la aplicación
root.mainloop()