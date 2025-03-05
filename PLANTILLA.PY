import pandas as pd
import barcode
from barcode.writer import ImageWriter
from reportlab.lib.pagesizes import (mm)
from reportlab.pdfgen import canvas
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def generar_codigo_barras(radicado, output_directory):
    """Genera un código de barras para un radicado específico y lo guarda en un directorio sin el texto debajo."""
    filename = os.path.join(output_directory, f"{radicado}")
     # Opciones de tamaño aproximado para el código de barras
    opciones = {
        'module_width': 0.3,    # Ajusta el ancho de cada barra en píxeles
        'module_height': 12,    # Aproximación de altura en relación al tamaño en la etiqueta
        'quiet_zone': 1,        # Espacio en blanco alrededor
        'font_size': 0          # Oculta el texto debajo del código de barras
    }
    code = barcode.get('code128', str(radicado), writer=ImageWriter())
    # Guardar el código de barras sin el texto
    code.save(filename, options={'write_text': False, 'module_width': 1, 'module_height': 25})  # Ajusta el ancho y la altura aquí
    return filename


def generar_todos_codigos_barras(csv_filename, output_directory):
    """Genera los códigos de barras para todos los registros en el CSV."""
    df = pd.read_csv(csv_filename, encoding='ISO-8859-1', delimiter=';', on_bad_lines='skip')
    for index, row in df.iterrows():
        radicado = row.get('Radicado')
        if pd.notna(radicado):
            generar_codigo_barras(str(radicado), output_directory)

def generar_pdf(datos, output_filename, barcode_filename, tipo_documento):
    """Genera un PDF basado en los datos proporcionados, incluyendo el código de barras."""
    c = canvas.Canvas(output_filename, pagesize=(85 * mm, 24 * mm))
    
    # Configuración de fuentes y tamaños
    c.setFont("Helvetica", 6)
    
    # Inserción del logo "Urban"
    logo_path = r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\logo_dark2-Photoroom.png"  # Asegúrate de colocar aquí la ruta correcta del logo
    if os.path.exists(logo_path):
        c.drawImage(logo_path, 1 * mm, 1 * mm, width=8 * mm, height=15 * mm)
    
    # Inserción de los datos en la plantilla
    c.drawString(1 * mm, 19 * mm, f"{datos.get('Nombre', '')} {datos.get('Segundo Nombre', '')} {datos.get('Apellido', '')} {datos.get('Segundo Apellido', '')}")
    c.drawString(50 * mm, 19 * mm, f"{datos.get('Dirección', '')}")
    c.drawString(3 * mm, 16 * mm, f"{datos.get('consecutivo', '')}")
    c.drawString(7 * mm, 16 * mm, f"{datos.get('Numero de Secuencia', '')}")
    c.drawString(15 * mm, 8 * mm, f"{tipo_documento}")
    c.drawString(15 * mm, 3 * mm, f"{datos.get('Prioridad', '')}")
    c.drawString(25 * mm, 11 * mm, f"{datos.get('Orden documental', '')}")
    c.drawString(40 * mm, 11 * mm, f"{datos.get('Radicado', '')}")
    c.drawString(70 * mm, 11 * mm, f"{datos.get('Fecha de Recepción', '')}")
    c.drawString(45 * mm, 8 * mm, f"{datos.get('Ciudad', '')}")
    c.drawString(65 * mm, 8 * mm, f"{datos.get('Departamento', '')}")
    c.drawString(45 * mm, 3 * mm, "130 gr.")
    c.drawString(55 * mm, 3 * mm, f"{datos.get('No acta', '')}")
    c.drawString(65 * mm, 3 * mm, f"{datos.get('Código de dependencia', '')}")
    
    # Añadir el código de barras al PDF
    if os.path.exists(barcode_filename):
        c.drawImage(barcode_filename, 65 * mm, 5 * mm, width=8.4 * mm, height=5.8 * mm)
    else:
        raise FileNotFoundError(f"No se encontró el archivo de código de barras: {barcode_filename}")

    c.save()

def generar_pdfs(csv_filename, barcode_directory, output_directory, tipo_documento):
    """Genera un PDF para cada registro en el CSV usando los códigos de barras pre-generados."""
    df = pd.read_csv(csv_filename, encoding='ISO-8859-1', delimiter=';', on_bad_lines='skip')
    for index, row in df.iterrows():
        radicado = row.get('Radicado')
        if pd.notna(radicado):
            barcode_filename = os.path.join(barcode_directory, f"{radicado}.png")
            output_filename = os.path.join(output_directory, f"documento_{tipo_documento}_{radicado}.pdf")
            generar_pdf(row, output_filename, barcode_filename, tipo_documento)

def seleccionar_archivo_csv():
    """Selecciona el archivo CSV para procesar."""
    return filedialog.askopenfilename(title="Seleccionar archivo CSV", filetypes=[("CSV files", "*.csv")])

def seleccionar_directorio():
    """Selecciona un directorio para guardar archivos."""
    return filedialog.askdirectory(title="Seleccionar directorio de salida")

def iniciar_generacion():
    """Inicia el proceso de generación de códigos de barras y PDFs."""
    csv_filename = seleccionar_archivo_csv()
    if not csv_filename:
        return

    barcode_directory = seleccionar_directorio()
    if not barcode_directory:
        return

    # Generación de todos los códigos de barras
    generar_todos_codigos_barras(csv_filename, barcode_directory)
    
    # Directorio de salida para los PDFs
    output_directory = seleccionar_directorio()
    if not output_directory:
        return

    tipo_documento = opcion_tipo.get()
    generar_pdfs(csv_filename, barcode_directory, output_directory, tipo_documento)
    messagebox.showinfo("Éxito", "Generación de PDFs completada.")

# Interfaz de usuario
root = tk.Tk()
root.title("Generador de PDF con Código de Barras")
root.geometry("400x300")

label = tk.Label(root, text="Seleccione el archivo CSV y directorios para generar códigos de barras y PDFs:")
label.pack(pady=10)

opcion_tipo = tk.StringVar(value="COPIA")
radio_copia = tk.Radiobutton(root, text="Copia", variable=opcion_tipo, value="COPIA")
radio_original = tk.Radiobutton(root, text="Original", variable=opcion_tipo, value="ORIGINAL")
radio_copia.pack()
radio_original.pack()

btn_generar = tk.Button(root, text="Generar Códigos de Barras y PDFs", command=iniciar_generacion)
btn_generar.pack(pady=20)

root.mainloop()
