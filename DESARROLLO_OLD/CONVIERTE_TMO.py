# Diego Insuasty Mora (Convierte archivos PDF a formato Carta)

import openpyxl
import os
from PyPDF4 import PdfFileReader, PdfFileWriter

# Ruta del archivo Excel
ruta_archivo = r'\\FJCALDAS\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\tamanoarchivos24_04.xlsx'

# Ruta de la carpeta donde se guardarán los archivos convertidos
ruta_carpeta_convertidos = r'\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\Actos GG_ Diarios_SHD_2024\240304-DIB+DCO-GG-lote 1\CONVERTIDOS'

# Abre el archivo Excel
libro = openpyxl.load_workbook(ruta_archivo)
hoja = libro.active

# Recorre cada fila de la hoja Excel
for fila in hoja.iter_rows(min_row=2, values_only=True):
    # Obtiene los valores de la fila
    ruta_pdf, medida_vertical, medida_horizontal, formato = fila

    # Carga el archivo PDF
    with open(ruta_pdf, 'rb') as archivo_pdf:
        pdf_reader = PdfFileReader(archivo_pdf)

        # Obtiene el tamaño de página actual
        pagina_actual = pdf_reader.getPage(0)
        ancho_actual = pagina_actual.mediaBox.getWidth()
        alto_actual = pagina_actual.mediaBox.getHeight()

        # Obtiene el factor de conversión para ajustar al tamaño carta
        factor_ancho = 21.59 / float(ancho_actual)
        factor_alto = 27.94 / float(alto_actual)
        factor_conversion = min(factor_ancho, factor_alto)

        # Establece un factor de conversión mínimo si la página resultante es muy pequeña
        factor_minimo = 0.94
        if factor_conversion < 1:
            factor_conversion = max(factor_conversion, factor_minimo)


        # Crea un escritor de PDF para el archivo convertido
        pdf_writer = PdfFileWriter()

        # Recorre cada página del PDF y la ajusta al tamaño carta
        for num_pagina in range(pdf_reader.getNumPages()):
            pagina = pdf_reader.getPage(num_pagina)
            pagina.scale(factor_conversion, factor_conversion)
            pdf_writer.addPage(pagina)

        # Crea el nombre del archivo convertido
        nombre_archivo = os.path.basename(ruta_pdf)
        nombre_archivo_convertido = f"{nombre_archivo[:-4]}_carta.pdf"
        ruta_archivo_convertido = os.path.join(ruta_carpeta_convertidos, nombre_archivo_convertido)

        # Guarda el archivo convertido en la carpeta de destino
        with open(ruta_archivo_convertido, 'wb') as archivo_convertido:
            pdf_writer.write(archivo_convertido)

# Cierra el archivo Excel
libro.close()


