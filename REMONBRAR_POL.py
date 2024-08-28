import os
import re

def get_last_number(filename):
    # Extraer el último número del nombre del archivo
    match = re.search(r'_(\d+)(?=\.\w+$)', filename)
    if match:
        return int(match.group(1))
    return 0

def rename_files_in_directory(directory):
    # Listar todos los archivos en el directorio
    files = os.listdir(directory)
    pdf_files = [file for file in files if file.endswith('.pdf')]

    # Crear una lista de tuplas (número, nombre del archivo)
    files_with_numbers = [(get_last_number(file), file) for file in pdf_files]

    # Ordenar los archivos por el número extraído
    files_with_numbers.sort()

    # Renombrar cada archivo
    for index, (_, filename) in enumerate(files_with_numbers, start=1):
        new_name = f"{index:06d}_{filename}"
        old_file = os.path.join(directory, filename)
        new_file = os.path.join(directory, new_name)
        os.rename(old_file, new_file)
        print(f'Renamed: {old_file} to {new_file}')

# Especificar el directorio donde se encuentran los archivos PDF
directory_path = r'\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\OP 9176-3\NUMERACION'  # Cambia esto a la ruta de tu directorio

rename_files_in_directory(directory_path)
