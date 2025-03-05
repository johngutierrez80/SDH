import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

def agregar_texto_pdf(pdf_input, pdf_output, texto, x, y):
    temp_pdf = os.path.join(os.path.dirname(pdf_output), "temp.pdf")
    lector = PdfReader(pdf_input)
    escritor = PdfWriter()
    
    c = canvas.Canvas(temp_pdf, pagesize=letter)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x, y, texto)
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

def procesar_directorio(directorio_entrada, directorio_salida, texto, x, y):
    if not os.path.exists(directorio_salida):
        os.makedirs(directorio_salida)
    
    for archivo in os.listdir(directorio_entrada):
        if archivo.endswith(".pdf"):
            ruta_entrada = os.path.join(directorio_entrada, archivo)
            ruta_salida = os.path.join(directorio_salida, archivo)
            agregar_texto_pdf(ruta_entrada, ruta_salida, texto, x, y)
    messagebox.showinfo("Completado", f"Archivos procesados y guardados en: {directorio_salida}")

def seleccionar_carpeta_entrada():
    ruta = filedialog.askdirectory(title="Seleccionar carpeta de entrada")
    if ruta:
        entrada_var.set(ruta)

def seleccionar_carpeta_salida():
    ruta = filedialog.askdirectory(title="Seleccionar carpeta de salida")
    if ruta:
        salida_var.set(ruta)

def iniciar_proceso():
    entrada = entrada_var.get()
    salida = salida_var.get()
    texto = texto_var.get()
    
    if not entrada or not salida or not texto:
        messagebox.showerror("Error", "Por favor completa todos los campos.")
        return
    
    try:
        procesar_directorio(entrada, salida, texto, 455, 725)
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {e}")

# Configuración de la interfaz
root = tk.Tk()
root.title("Agregar texto a PDFs")
root.geometry("500x400")

entrada_var = tk.StringVar()
salida_var = tk.StringVar()
texto_var = tk.StringVar(value="ACTA N° 80")

# Widgets
tk.Label(root, text="Carpeta de entrada:").pack(pady=5)
tk.Entry(root, textvariable=entrada_var, width=50).pack(pady=5)
tk.Button(root, text="Seleccionar carpeta", command=seleccionar_carpeta_entrada).pack(pady=5)

tk.Label(root, text="Carpeta de salida:").pack(pady=5)
tk.Entry(root, textvariable=salida_var, width=50).pack(pady=5)
tk.Button(root, text="Seleccionar carpeta", command=seleccionar_carpeta_salida).pack(pady=5)

tk.Label(root, text="Texto a agregar:").pack(pady=5)
tk.Entry(root, textvariable=texto_var, width=50).pack(pady=5)

tk.Button(root, text="Iniciar proceso", command=iniciar_proceso, bg="green", fg="white").pack(pady=20)

# Iniciar la aplicación
root.mainloop()
