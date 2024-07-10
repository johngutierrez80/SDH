import os
import pandas as pd

# Configuración de las rutas
excel_file_path = r'\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\ACTAS_2024\240621 ACTA # 34\NUEVOS ARCHIVOS\BASE ACTA 34 PDFS FINAL.xlsx'
pdf_directory_path = r'\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\ACTAS_2024\240621 ACTA # 34\ACTA 34 FINAL'

# Leer el archivo Excel
df = pd.read_excel(excel_file_path)

# Asegúrate de que los campos específicos existen en el DataFrame
if '# COMUNICACIÓN' not in df.columns or 'LLAVE' not in df.columns:
    print("Las columnas especificadas no se encuentran en el archivo Excel.")
    exit(1)

# Obtener las listas de nombres actuales y nuevos nombres
current_names = df['# COMUNICACIÓN'].astype(str).tolist()
new_names = df['LLAVE'].astype(str).tolist()

# Renombrar los archivos en el directorio
for current_name, new_name in zip(current_names, new_names):
    current_file = os.path.join(pdf_directory_path, current_name + '.pdf')
    new_file = os.path.join(pdf_directory_path, new_name + '.pdf')

    if os.path.exists(current_file):
        os.rename(current_file, new_file)
        print(f'Renombrado: {current_file} -> {new_file}')
    else:
        print(f'Archivo no encontrado: {current_file}')

print("Renombramiento completado.")
