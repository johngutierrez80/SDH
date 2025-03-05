import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import pandas as pd

def agregar_texto_pdf(pdf_input, pdf_output, texto_orden, x_orden, y_orden):
    temp_pdf = os.path.join(os.path.dirname(pdf_output), "temp.pdf")
    lector = PdfReader(pdf_input)
    escritor = PdfWriter()
    
    c = canvas.Canvas(temp_pdf, pagesize=letter)
    c.setFont("Helvetica-Bold", 11)
    # Texto del Orden de Impresión
    c.drawString(x_orden, y_orden, texto_orden)
    c.save()
    
    with open(temp_pdf, "rb") as temp_pdf_file:
        lector_temp = PdfReader(temp_pdf_file)
        pagina_con_texto = lector_temp.pages[0]
        pagina_original = lector.pages[0]
        
        pagina_original.merge_page(pagina_con_texto)
        escritor.add_page(pagina_original)
    
    for i in range(1, len(lector.pages)):
        escritor.add_page(lector.pages[i])
    
    with open(pdf_output, "wb") as salida:
        escritor.write(salida)
    
    os.remove(temp_pdf)

def procesar_directorio(directorio_entrada, directorio_salida, archivo_excel):
    if not os.path.exists(directorio_salida):
        os.makedirs(directorio_salida)
    
    try:
        # Leer el archivo Excel
        df = pd.read_excel(archivo_excel)
        if 'Orden Impresión' not in df.columns:
            raise ValueError("La columna 'Orde Impresión' no existe en el archivo Excel.")
        
        # Iterar sobre los archivos PDF y asignar "Orden de Impresión"
        for idx, archivo in enumerate(os.listdir(directorio_entrada)):
            if archivo.endswith(".pdf"):
                ruta_entrada = os.path.join(directorio_entrada, archivo)
                ruta_salida = os.path.join(directorio_salida, archivo)
                
                # Obtener el Orden de Impresión correspondiente
                texto_orden = str(df.loc[idx, 'Orden Impresión']) if idx < len(df) else "N/A"
                agregar_texto_pdf(
                    ruta_entrada, ruta_salida, 
                    texto_orden, 
                    75, 690  # Coordenadas para Orden de Impresión
                )
        messagebox.showinfo("Completado", f"Archivos procesados y guardados en: {directorio_salida}")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {e}")

def seleccionar_carpeta_entrada():
    ruta = filedialog.askdirectory(title="Seleccionar carpeta de entrada")
    if ruta:
        entrada_var.set(ruta)

def seleccionar_carpeta_salida():
    ruta = filedialog.askdirectory(title="Seleccionar carpeta de salida")
    if ruta:
        salida_var.set(ruta)

def seleccionar_archivo_excel():
    ruta = filedialog.askopenfilename(
        title="Seleccionar archivo Excel", 
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if ruta:
        excel_var.set(ruta)

def iniciar_proceso():
    entrada = entrada_var.get()
    salida = salida_var.get()
    archivo_excel = excel_var.get()
    
    if not entrada or not salida or not archivo_excel:
        messagebox.showerror("Error", "Por favor completa todos los campos.")
        return
    
    try:
        procesar_directorio(entrada, salida, archivo_excel)
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {e}")

# Configuración de la interfaz
root = tk.Tk()
root.title("Agregar texto a PDFs")
root.geometry("500x400")

entrada_var = tk.StringVar()
salida_var = tk.StringVar()
excel_var = tk.StringVar()

# Widgets
tk.Label(root, text="Carpeta de entrada:").pack(pady=5)
tk.Entry(root, textvariable=entrada_var, width=50).pack(pady=5)
tk.Button(root, text="Seleccionar carpeta", command=seleccionar_carpeta_entrada).pack(pady=5)

tk.Label(root, text="Carpeta de salida:").pack(pady=5)
tk.Entry(root, textvariable=salida_var, width=50).pack(pady=5)
tk.Button(root, text="Seleccionar carpeta", command=seleccionar_carpeta_salida).pack(pady=5)

tk.Label(root, text="Archivo Excel:").pack(pady=5)
tk.Entry(root, textvariable=excel_var, width=50).pack(pady=5)
tk.Button(root, text="Seleccionar archivo Excel", command=seleccionar_archivo_excel).pack(pady=5)

tk.Button(root, text="Iniciar proceso", command=iniciar_proceso, bg="green", fg="white").pack(pady=20)

# Iniciar la aplicación
root.mainloop()
