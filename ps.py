import os
import subprocess

# Ruta de Ghostscript
GHOSTSCRIPT_PATH = r"C:\Program Files\gs\gs10.04.0\bin\gswin64c.exe"

# Verificar si Ghostscript está instalado
if not os.path.exists(GHOSTSCRIPT_PATH):
    print(f"❌ Error: No se encontró Ghostscript en {GHOSTSCRIPT_PATH}")
    exit()

# Directorios
INPUT_FOLDER = os.path.normpath(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\ACTOS_GG_DIARIOS_SHD_2025\250227-DIB+DCO-GG-lote 1\ORIGINAL FALTANTE")
PS_FOLDER = os.path.normpath(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\ACTOS_GG_DIARIOS_SHD_2025\250227-DIB+DCO-GG-lote 1\ps")

# Crear carpeta de salida si no existe
if not os.path.exists(PS_FOLDER):
    os.makedirs(PS_FOLDER)

def convert_pdf_to_ps():
    """ Convierte los PDFs del directorio de entrada a archivos .ps en el directorio PS_FOLDER """
    for filename in os.listdir(INPUT_FOLDER):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.normpath(os.path.join(INPUT_FOLDER, filename))
            ps_filename = os.path.splitext(filename)[0] + ".ps"
            ps_path = os.path.normpath(os.path.join(PS_FOLDER, ps_filename))

            try:
                subprocess.run([
                    GHOSTSCRIPT_PATH, "-dNOPAUSE", "-dBATCH", "-sDEVICE=ps2write",
                    f"-sOutputFile={ps_path}", pdf_path
                ], check=True)
                print(f"✅ Convertido a PS: {filename} -> {ps_filename}")
            except subprocess.CalledProcessError as e:
                print(f"❌ Error al convertir {filename} a PS: {e}")

if __name__ == "__main__":
    print("🔄 Iniciando conversión PDF → PS")
    convert_pdf_to_ps()
    print("\n✅ Proceso completado.")
