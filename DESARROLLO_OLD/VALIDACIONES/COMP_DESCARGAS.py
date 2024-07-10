import pandas as pd
import os
import re

def obtener_nombre_archivo(nombre_archivo):
    # Si el nombre del archivo sigue el formato de radicado, se devuelve tal cual
    if re.match(r'\d{4}EE\d{6}O?\d?', nombre_archivo):
        return nombre_archivo
    else:
        # Si no sigue el formato de radicado, se extraen los primeros 10 caracteres
        return nombre_archivo[:10]

def cargar_archivos_descargados(directorio_descargas):
    # Obtener los nombres de los archivos en el directorio de descargas
    archivos_descargados = os.listdir(directorio_descargas)

    # Obtener los nombres de archivo tratados sin la extensión
    nombres_archivos_tratados = [obtener_nombre_archivo(nombre.split('.')[0]) for nombre in archivos_descargados]

    # Crear DataFrame con los nombres de archivo tratados
    df_archivos_descargados = pd.DataFrame(nombres_archivos_tratados, columns=['Archivo_descargado'])
    #print(df_archivos_descargados)
    return df_archivos_descargados

def comparar_archivos(excel_path, directorio_descargas, sheet_name='Hoja1', campo_especifico='SAP_ID'):
    # Leer el archivo de Excel en un DataFrame
    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    # Verificar si 'Numero de secuencia' está presente en el DataFrame
    if 'Numero de secuencia ' not in df.columns:
        raise ValueError("El campo 'Numero de secuencia' no está presente en el DataFrame.")

    # Obtener los nombres de archivo tratados sin la extensión
    df_archivos_descargados = cargar_archivos_descargados(directorio_descargas)

    # Convertir los valores únicos del campo específico del DataFrame a cadenas para la comparación
    valores_excel = set(df[campo_especifico].astype(str))

    # Comparar los nombres de archivo tratados con los valores del campo específico del Excel
    archivos_no_en_excel = valores_excel - set(df_archivos_descargados['Archivo_descargado'])

    # Imprimir los archivos que no están en la carpeta de descargas
    if archivos_no_en_excel:
        print("Archivos que no están presentes en la carpeta de descargas:")
        for sap_id in archivos_no_en_excel:
            # Normalizar el SAP_ID para que coincida con el formato en el DataFrame
            sap_id_normalizado = sap_id.strip()  # Ejemplo de normalización, puede requerir más ajustes

            # Buscar el 'Numero de secuencia' correspondiente al SAP_ID normalizado en el DataFrame
            matching_rows = df[df[campo_especifico].astype(str).str.strip() == sap_id_normalizado]
            if not matching_rows.empty:
                numero_secuencia = matching_rows.iloc[0]['Numero de secuencia ']
                print(f"SAP_ID: {sap_id}, Numero de secuencia: {numero_secuencia}")
            else:
                # Buscar el 'Numero de secuencia' de los archivos no encontrados en el DataFrame
                nombre_archivo = df_archivos_descargados[df_archivos_descargados['Archivo_descargado'] == sap_id + '.pdf']
                if not nombre_archivo.empty:
                    print(f"No se encontró el SAP_ID {sap_id} en el DataFrame. Número de secuencia: {nombre_archivo.iloc[0]['Numero de secuencia ']}")
                else:
                    print(f"No se encontró el SAP_ID {sap_id} en el DataFrame.")
    else:
        print("Todos los archivos están descargados.")
# Rutas de entrada (usando raw string literal)
archivo_excel = r'\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\LABEL_ACTOS_GG_2024\24-03-27\Copia de Base General Imprenta 27-03-2024.xlsx'
directorio_descargas = r'\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\Actos GG_ Diarios_SHD_2024\240327-DIB+DCO-GG-lote 1\DESCARGAS'

# Llamar a la función comparar_archivos
comparar_archivos(archivo_excel, directorio_descargas)
