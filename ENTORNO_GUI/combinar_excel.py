import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

# Funciones para la lógica del programa
def cargar_datos(ruta_archivo):
    try:
        df = pd.read_excel(ruta_archivo)
        return df
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo cargar el archivo. Error: {e}")
        return None

def combinar_registros(df):
    # Definimos los campos que queremos mantener únicos (sin concatenar)
    campos_unicos = ['Numero_Identificación', 'Tipo_Identificación', 'Nombre_o_razón_social', 'MUNICIPIO_SAP','DEPARTAMENTO_SAP','DIRECCION_NOTIFICACION_SAP', 'Tipo de Impuesto – Indicio ']

    # Agrupamos por "Numero_Identificación" y aplicamos diferentes métodos a las columnas
    df_unificado = df.groupby('Numero_Identificación', as_index=False).agg({
        'Tipo_Identificación': 'first',    # Mantener el primer valor
        'Nombre_o_razón_social': 'first',  # Mantener el primer valor
        'MUNICIPIO_SAP': 'first',     # Mantener el primer valor
        'DEPARTAMENTO_SAP': 'first',     # Mantener el primer valor
        'DIRECCION_NOTIFICACION_SAP': 'first',     # Mantener el primer valor
        'Tipo de Impuesto – Indicio ': 'first',  # Mantener el primer valor
        
        # Concatenar las columnas que contienen múltiples valores
        'AUTOAVALUO': lambda x: '\n'.join(map(str, x)),
        'IA_CALCULADO': lambda x: '\n'.join(map(str, x)),
        'Sancion de extemporaneidad': lambda x: '\n'.join(map(str, x)),
        'Sanción minima': lambda x: '\n'.join(map(str, x)),
        'intereses 100%': lambda x: '\n'.join(map(str, x)),
        'valor total a pagar al 100%': lambda x: '\n'.join(map(str, x)),
        'valor total a pagar con beneficio (impto, sancion e interes)': lambda x: '\n'.join(map(str, x))
    })

    return df_unificado



def formatear_datos(df):
    # Definir las columnas que necesitan formato monetario
    columnas_monetarias = ["AUTOAVALUO", "IA_CALCULADO", "Sancion de extemporaneidad", 
                           "Sanción minima", "intereses 100%", 
                           "valor total a pagar al 100%", 
                           "valor total a pagar con beneficio (impto, sancion e interes)"]

    # Función para formatear los valores individuales con saltos de línea
    def formatear_monetario(valores):
        # Dividir los valores por salto de línea, aplicar formato, y luego volver a unirlos
        valores_separados = str(valores).split("\n")
        valores_formateados = [f"${float(x):,.2f}" if x.strip() != "" else "$0.00" for x in valores_separados]
        return "\n".join(valores_formateados)
    
    # Aplicar la función a las columnas monetarias
    for col in columnas_monetarias:
        df[col] = df[col].apply(formatear_monetario)
    
    return df

def exportar_datos(df, ruta_salida):
    try:
        # Exportar a Excel aplicando el formato con openpyxl
        writer = pd.ExcelWriter(ruta_salida, engine='openpyxl')
        df.to_excel(writer, index=False)
        writer.close()  # Aquí cambiamos 'save()' por 'close()'
        messagebox.showinfo("Éxito", f"Archivo exportado correctamente a {ruta_salida}")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo exportar el archivo. Error: {e}")

# Funciones para la interfaz gráfica
def seleccionar_archivo():
    archivo = filedialog.askopenfilename(title="Seleccionar archivo", filetypes=[("Archivos Excel", "*.xlsx")])
    if archivo:
        entrada_archivo.set(archivo)

def seleccionar_guardado():
    archivo_salida = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivos Excel", "*.xlsx")])
    if archivo_salida:
        salida_archivo.set(archivo_salida)

def procesar_archivo():
    ruta_archivo = entrada_archivo.get()
    ruta_salida = salida_archivo.get()

    if not ruta_archivo or not ruta_salida:
        messagebox.showwarning("Advertencia", "Debe seleccionar un archivo de entrada y una ubicación de guardado.")
        return

    df = cargar_datos(ruta_archivo)
    if df is not None:
        df_combinado = combinar_registros(df)
        df_formateado = formatear_datos(df_combinado)
        exportar_datos(df_formateado, ruta_salida)

# Configuración de la ventana principal
ventana = tk.Tk()
ventana.title("Combinación de Registros en Excel")
ventana.geometry("500x300")

# Variables para guardar las rutas de los archivos
entrada_archivo = tk.StringVar()
salida_archivo = tk.StringVar()

# Etiquetas y cuadros de texto para la selección de archivos
tk.Label(ventana, text="Archivo de Entrada:").pack(pady=10)
tk.Entry(ventana, textvariable=entrada_archivo, width=50).pack(pady=5)
tk.Button(ventana, text="Seleccionar Archivo", command=seleccionar_archivo).pack(pady=5)

tk.Label(ventana, text="Archivo de Salida:").pack(pady=10)
tk.Entry(ventana, textvariable=salida_archivo, width=50).pack(pady=5)
tk.Button(ventana, text="Seleccionar Ubicación de Guardado", command=seleccionar_guardado).pack(pady=5)

# Botón para procesar el archivo
tk.Button(ventana, text="Procesar Archivo", command=procesar_archivo).pack(pady=20)

# Iniciar la aplicación
ventana.mainloop()
