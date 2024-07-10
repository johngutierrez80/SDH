import os
from datetime import datetime
from pathlib import Path

def get_creation_time(file_path):
    """
    Devuelve la fecha y hora de creación del archivo, incluyendo milisegundos.
    """
    creation_time = os.path.getctime(file_path)
    return datetime.fromtimestamp(creation_time).strftime('%Y%m%d_%H%M%S_%f')[:-3]

def rename_files_with_creation_time(folder_path):
    """
    Recorre la carpeta especificada y renombra los archivos PDF agregando la fecha y hora de creación, incluyendo milisegundos.
    """
    folder = Path(folder_path)
    if not folder.is_dir():
        print(f"La ruta {folder_path} no es una carpeta válida.")
        return
    
    for pdf_file in folder.glob('*.pdf'):
        creation_time = get_creation_time(pdf_file)
        new_name = f"{pdf_file.stem}_{creation_time}{pdf_file.suffix}"
        new_path = pdf_file.with_name(new_name)
        os.rename(pdf_file, new_path)
        print(f"Archivo renombrado: {pdf_file.name} -> {new_name}")

if __name__ == "__main__":
    carpeta = r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\ACTAS_2024\240621 ACTA # 33\MUESTRAS ACTA 33"
    rename_files_with_creation_time(carpeta)
