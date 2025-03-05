import subprocess
from pathlib import Path
import logging
from typing import Optional
import tkinter as tk
from tkinter import filedialog, messagebox

# Configuración de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Variables globales para las rutas
GHOSTSCRIPT_PATH = Path(r"C:\Program Files\gs\gs10.04.0\bin\gswin64c.exe")
INPUT_FOLDER = None
PS_FOLDER = None
OUTPUT_PDF_FOLDER = None

def setup_environment() -> bool:
    """Verifica y configura el entorno de trabajo."""
    try:
        if not GHOSTSCRIPT_PATH.exists():
            logger.error(f"No se encontró Ghostscript en {GHOSTSCRIPT_PATH}")
            return False
        
        for folder in [PS_FOLDER, OUTPUT_PDF_FOLDER]:
            folder.mkdir(parents=True, exist_ok=True)
        
        if not INPUT_FOLDER.exists():
            logger.error(f"El directorio de entrada no existe: {INPUT_FOLDER}")
            return False
            
        return True
    except PermissionError as e:
        logger.error(f"Error de permisos al crear carpetas: {e}")
        return False
    except Exception as e:
        logger.error(f"Error al configurar el entorno: {e}")
        return False

def convert_pdf_to_ps(pdf_path: Path) -> Optional[Path]:
    """Convierte un PDF a PostScript."""
    try:
        ps_filename = pdf_path.stem + ".ps"
        ps_path = PS_FOLDER / ps_filename
        
        cmd = [
            str(GHOSTSCRIPT_PATH),
            "-dNOPAUSE",
            "-dBATCH",
            "-sDEVICE=ps2write",
            f"-sOutputFile={ps_path}",
            str(pdf_path)
        ]
        
        subprocess.run(cmd, check=True, capture_output=True, text=True)
        logger.info(f"PDF → PS: {pdf_path.name} -> {ps_filename}")
        return ps_path
    
    except subprocess.CalledProcessError as e:
        logger.error(f"Error convirtiendo PDF a PS {pdf_path.name}: {e.stderr}")
        return None

def convert_ps_to_pdf(ps_path: Path) -> Optional[Path]:
    """Convierte un PS a PDF y elimina el archivo PS."""
    try:
        pdf_filename = ps_path.stem + ".pdf"
        pdf_path = OUTPUT_PDF_FOLDER / pdf_filename
        
        cmd = [
            str(GHOSTSCRIPT_PATH),
            "-dNOPAUSE",
            "-dBATCH",
            "-sDEVICE=pdfwrite",
            f"-sOutputFile={pdf_path}",
            str(ps_path)
        ]
        
        subprocess.run(cmd, check=True, capture_output=True, text=True)
        logger.info(f"PS → PDF: {ps_path.name} -> {pdf_filename}")
        
        try:
            ps_path.unlink()
            logger.info(f"Eliminado archivo temporal: {ps_path.name}")
        except Exception as e:
            logger.warning(f"No se pudo eliminar {ps_path.name}: {e}")
            
        return pdf_path
    
    except subprocess.CalledProcessError as e:
        logger.error(f"Error convirtiendo PS a PDF {ps_path.name}: {e.stderr}")
        return None

def process_pdf_files():
    """Procesa todos los PDFs: PDF → PS → PDF y elimina PS."""
    if not INPUT_FOLDER or not PS_FOLDER or not OUTPUT_PDF_FOLDER:
        messagebox.showerror("Error", "Por favor selecciona todos los directorios antes de procesar.")
        return
    
    if not setup_environment():
        return
    
    logger.info("Iniciando proceso completo de conversión")
    status_text.delete(1.0, tk.END)
    status_text.insert(tk.END, "Iniciando proceso completo de conversión...\n")
    
    pdf_files = list(INPUT_FOLDER.glob("*.pdf"))
    total_files = len(pdf_files)
    
    if total_files == 0:
        logger.warning("No se encontraron archivos PDF en el directorio de entrada")
        status_text.insert(tk.END, "No se encontraron archivos PDF en el directorio de entrada.\n")
        return
    
    successful = 0
    for i, pdf_path in enumerate(pdf_files, 1):
        msg = f"Procesando archivo {i}/{total_files}: {pdf_path.name}"
        logger.info(msg)
        status_text.insert(tk.END, msg + "\n")
        status_text.update()
        
        ps_path = convert_pdf_to_ps(pdf_path)
        if ps_path is None:
            continue
            
        if convert_ps_to_pdf(ps_path):
            successful += 1
    
    final_msg = f"Proceso completado: {successful}/{total_files} archivos procesados exitosamente"
    logger.info(final_msg)
    status_text.insert(tk.END, final_msg + "\n")
    messagebox.showinfo("Completado", final_msg)

# Funciones para la GUI
def select_input_folder():
    global INPUT_FOLDER
    folder = filedialog.askdirectory(title="Seleccionar carpeta de entrada (PDFs originales)")
    if folder:
        INPUT_FOLDER = Path(folder)
        input_label.config(text=f"Carpeta de entrada: {INPUT_FOLDER}")

def select_ps_folder():
    global PS_FOLDER
    folder = filedialog.askdirectory(title="Seleccionar carpeta para archivos PS temporales")
    if folder:
        PS_FOLDER = Path(folder)
        ps_label.config(text=f"Carpeta PS: {PS_FOLDER}")

def select_output_folder():
    global OUTPUT_PDF_FOLDER
    folder = filedialog.askdirectory(title="Seleccionar carpeta para PDFs finales")
    if folder:
        OUTPUT_PDF_FOLDER = Path(folder)
        output_label.config(text=f"Carpeta de salida: {OUTPUT_PDF_FOLDER}")

# Configuración de la interfaz gráfica
root = tk.Tk()
root.title("Conversor PDF a PS a PDF")
root.geometry("600x400")

# Etiquetas y botones
tk.Label(root, text="Conversor de Archivos PDF").pack(pady=10)

input_button = tk.Button(root, text="Seleccionar carpeta de entrada", command=select_input_folder)
input_button.pack(pady=5)
input_label = tk.Label(root, text="Carpeta de entrada: No seleccionada")
input_label.pack()

ps_button = tk.Button(root, text="Seleccionar carpeta para PS", command=select_ps_folder)
ps_button.pack(pady=5)
ps_label = tk.Label(root, text="Carpeta PS: No seleccionada")
ps_label.pack()

output_button = tk.Button(root, text="Seleccionar carpeta de salida", command=select_output_folder)
output_button.pack(pady=5)
output_label = tk.Label(root, text="Carpeta de salida: No seleccionada")
output_label.pack()

process_button = tk.Button(root, text="Iniciar Conversión", command=process_pdf_files)
process_button.pack(pady=20)

# Área de texto para mostrar el estado
status_text = tk.Text(root, height=10, width=70)
status_text.pack(pady=10)

# Iniciar la GUI
root.mainloop()