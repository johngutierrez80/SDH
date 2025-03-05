import pandas as pd
from tkinter import Tk, filedialog

# Función para cargar archivos Excel
def cargar_archivo_excel():
    root = Tk()
    root.withdraw()  # Oculta la ventana principal de Tkinter
    archivo_excel = filedialog.askopenfilename(title="Seleccione el archivo Excel", filetypes=[("Archivos Excel", "*.xlsx")])
    return archivo_excel

# Cargar archivo principal (base de datos más grande)
archivo_principal = cargar_archivo_excel()
df_principal = pd.read_excel(archivo_principal)

# Cargar archivo con NUMERO_IDENTIFICACION
archivo_filtro = cargar_archivo_excel()
df_filtro = pd.read_excel(archivo_filtro)

# Filtrar los registros de la base de datos principal según el campo "NUMERO_DOCUMENTO"
df_filtrado = df_principal[df_principal["NUMERO_DOCUMENTO"].isin(df_filtro["NUMERO_IDENTIFICACION"])]

# Guardar el DataFrame filtrado en un nuevo archivo Excel
archivo_salida = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivos Excel", "*.xlsx")], title="Guardar archivo filtrado")
df_filtrado.to_excel(archivo_salida, index=False)

print("Archivo filtrado guardado exitosamente.")
