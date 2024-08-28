import os

# Configuración del directorio
pdf_directory_path = r'\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\OP 9176-3\SERIE 47253'

# Solicitar al usuario el número de inicio para la numeración
start_number = int(input("Introduce el número inicial para la numeración: "))

# Listar y ordenar los archivos en el directorio que sigan el formato especificado
pdf_files = sorted([f for f in os.listdir(pdf_directory_path) if f.startswith('FORM INSCRIPCION POLICIA_Records_') and f.endswith('.pdf')],
                   key=lambda x: int(x.split('_')[-1].split('.')[0]))

# Renombrar los archivos en el directorio
for idx, pdf_file in enumerate(pdf_files, start=start_number):
    new_name = f"{str(idx).zfill(6)}_{pdf_file}"
    current_file = os.path.join(pdf_directory_path, pdf_file)
    new_file = os.path.join(pdf_directory_path, new_name)

    os.rename(current_file, new_file)
    print(f'Renombrado: {current_file} -> {new_file}')

print("Renombramiento completado.")

