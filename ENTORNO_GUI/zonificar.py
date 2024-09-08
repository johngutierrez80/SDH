import pandas as pd

# Cargar las dos bases de datos
base_organizada_path = r'\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\ACTAS_2024\240903ACTA # 56\BASE ACTA 56 COMBINADA (002).xlsx'
base_sin_ordenar_path = r'\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\ACTAS_2024\240903ACTA # 56\BASE ACTA 56 COMBINADA.xlsx'

# Leer las bases de datos
df_organizada = pd.read_excel(base_organizada_path)
df_sin_ordenar = pd.read_excel(base_sin_ordenar_path)

# Fusionar ambas bases de datos usando el campo 'Numero_Identificación'
df_reorganizada = pd.merge(df_organizada[['Numero_Identificación', 'Orden Impresión']],
                           df_sin_ordenar,
                           on='Numero_Identificación',
                           how='inner')

# Ordenar los datos según la columna 'Orden Impresión' para mantener el mismo orden que la base organizada
df_reorganizada = df_reorganizada.sort_values(by='Orden Impresión')

# Asegurar que la estructura y formato sean los mismos que la base organizada
# Mantendremos las mismas columnas de la base organizada en el orden original
df_reorganizada = df_reorganizada[df_organizada.columns]

# Guardar el resultado en un nuevo archivo Excel con el formato reorganizado
output_path = r'\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\ACTAS_2024\240903ACTA # 56\reorganizada 56/base_reorganizada.xlsx'
df_reorganizada.to_excel(output_path, index=False)

print(f"Archivo reorganizado guardado en: {output_path}")
