import os
import pandas as pd

# Configuración de las rutas
excel_file_path = r'\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\ACTAS_2024\240903ACTA # 55\BASE NOMBRES.xlsx'
pdf_directory_path = r'\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\ACTAS_2024\240903ACTA # 55\MUESTRAS ACTA 55'

# Leer el archivo Excel
df = pd.read_excel(excel_file_path)

# Asegúrate de que la columna especificada existe en el DataFrame
if 'nombre' not in df.columns:
    print("La columna 'nombre' no se encuentra en el archivo Excel.")
    exit(1)

# Obtener la lista de nuevos nombres desde la columna A (columna 'nombre')
new_names = df['nombre'].astype(str).tolist()

# Listar todos los archivos en el directorio que sigan el formato especificado
pdf_files = sorted([f for f in os.listdir(pdf_directory_path) if f.endswith('.pdf') and '__ACTA_55__' in f])

# Renombrar los archivos en el directorio
if len(pdf_files) != len(new_names):
    print("El número de archivos PDF no coincide con el número de nuevos nombres en el archivo Excel.")
    exit(1)
    

for pdf_file, new_name in zip(pdf_files, new_names):
    current_file = os.path.join(pdf_directory_path, pdf_file)
    new_file = os.path.join(pdf_directory_path, new_name + '.pdf')

    os.rename(current_file, new_file)
    print(f'Renombrado: {current_file} -> {new_file}')

print("Renombramiento completado.")
