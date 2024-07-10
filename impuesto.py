import pandas as pd

# Función para formatear números con puntos de miles
def format_number(number):
    try:
        number = int(number)
        return f"${number:,}".replace(",", ".")
    except ValueError:
        return number

# Leer el archivo Excel
df = pd.read_excel(r'\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\ACTAS_2024\240620 ACTA # 31\ICA_COMUN_ANUAL_2023-7_FISICO_salida.xlsx')

# Aplicar el formateo a cada número en la columna 'IMPUESTO'
df['IMPUESTO'] = df['IMPUESTO'].apply(lambda x: '\n'.join([format_number(num) for num in str(x).split('\n')]))

# Guardar los cambios en un nuevo archivo Excel
df.to_excel(r'\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\ACTAS_2024\240620 ACTA # 31\tu_archivo_formateado.xlsx', index=False)

print("Formateo completado y guardado en 'tu_archivo_formateado.xlsx'")
