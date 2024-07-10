import pandas as pd
import os
import re

def obtener_nombre_archivo(nombre_archivo):
    # Excluir los primeros 7 símbolos del nombre del archivo
    nombre_sin_prefijo = nombre_archivo[7:]
    # Si el nombre del archivo sigue el formato de radicado, se devuelve tal cual
    if re.match(r'\d{4}EE\d{6}O?\d?', nombre_sin_prefijo):
        return nombre_sin_prefijo
    else:
        # Si no sigue el formato de radicado, se extraen los primeros 10 caracteres
        return nombre_sin_prefijo[:10]

def comparar_archivos(excel_path, directorio_renombrados, sheet_name='Hoja1', campo_especifico='SAP_ID'):
    # Leer el archivo de Excel en un DataFrame
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    
    # Verificar si 'Numero de secuencia' está presente en el DataFrame
    if 'Numero de secuencia ' not in df.columns:
        raise ValueError("El campo 'Numero de secuencia' no está presente en el DataFrame.")
    
    # Obtener los nombres de archivo tratados sin la extensión desde el directorio de renombrados
    archivos_renombrados = os.listdir(directorio_renombrados)
    nombres_archivos_renombrados = [obtener_nombre_archivo(nombre.split('.')[0]) for nombre in archivos_renombrados]
    
    # Crear DataFrame con los nombres de archivo tratados desde el directorio de renombrados
    df_archivos_renombrados = pd.DataFrame(nombres_archivos_renombrados, columns=['Archivo_renombrado'])
    
    # Obtener el orden de los archivos según el campo específico del DataFrame
    orden_excel = df[campo_especifico].astype(str).tolist()
    
    # Verificar si los archivos en el directorio renombrado están en el mismo orden que en el archivo Excel
    if list(df_archivos_renombrados['Archivo_renombrado']) == orden_excel:
        print("Los archivos están presentes en el directorio de renombrados y en el mismo orden que en el archivo Excel.")
    else:
        print("Los archivos no están en el mismo orden que en el archivo Excel:")
        for archivo_excel, archivo_renombrado, secuencia_excel in zip(orden_excel, df_archivos_renombrados['Archivo_renombrado'], df['Numero de secuencia ']):
            if archivo_excel != archivo_renombrado:
                print(f"Archivo en el Excel (SAP_ID: {archivo_excel}, Numero de secuencia: {secuencia_excel}), Archivo en el directorio renombrado: {archivo_renombrado}")

# Rutas de entrada (usando raw string literal)
archivo_excel = r'\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\LABEL_ACTOS_GG_2024\24-03-12\Copia de Base General Imprenta 12-03-2024.xlsx'
directorio_renombrados = r'\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\Actos GG_ Diarios_SHD_2024\240312-DIB+DCO-GG-lote 1\RENOMBRADOS'  # Reemplazar con la ruta correcta

# Llamar a la función comparar_archivos
comparar_archivos(archivo_excel, directorio_renombrados)








