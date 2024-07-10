import os

def rename_pdfs(input_directory, output_directory):
    # Crear la carpeta de salida si no existe
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
    
    for filename in os.listdir(input_directory):
        if filename.endswith(".pdf"):
            # Divide el nombre del archivo en partes usando "__"
            parts = filename.split("__")
            if len(parts) > 1:
                reference_parts = parts[1].split("_")
                if len(reference_parts) > 1:
                    # Obtiene el último número de referencia después del segundo "_"
                    new_name = reference_parts[-1] 
                    # Ruta completa del archivo de entrada y salida
                    old_file_path = os.path.join(input_directory, filename)
                    new_file_path = os.path.join(output_directory, new_name)
                    # Renombra y mueve el archivo a la carpeta de salida
                    os.rename(old_file_path, new_file_path)
                    print(f"Renombrado y movido: {filename} -> {new_name}")

# Ruta de la carpeta donde están los archivos PDF
input_directory_path = r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\ACTAS_2024\240621 ACTA # 34\NUEVAS MUESTRAS ACTA 34"
# Ruta de la carpeta de salida
output_directory_path = r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\ACTAS_2024\240621 ACTA # 34\ACTA 34 FINAL"

rename_pdfs(input_directory_path, output_directory_path)
