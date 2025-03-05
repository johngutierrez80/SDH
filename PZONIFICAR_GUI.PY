import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Variables globales para almacenar cabeceras
columnas_base_organizada = []
columnas_base_sin_ordenar = []
columnas_fusionadas = []

# Función para cargar las cabeceras de un archivo seleccionado
def cargar_cabeceras(ruta_archivo, tipo_base):
    try:
        df = pd.read_excel(ruta_archivo, nrows=0)
        columnas = list(df.columns)
        
        if tipo_base == "organizada":
            global columnas_base_organizada
            columnas_base_organizada = columnas
            actualizar_opciones_menus(campo_id_organizada_menu, opciones_campo_id_organizada, columnas)
        elif tipo_base == "sin_ordenar":
            global columnas_base_sin_ordenar
            columnas_base_sin_ordenar = columnas
            actualizar_opciones_menus(campo_id_sin_ordenar_menu, opciones_campo_id_sin_ordenar, columnas)
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al cargar las cabeceras: {e}")

# Función para realizar la fusión preliminar y cargar las cabeceras fusionadas
def fusionar_bases_preliminar():
    global columnas_fusionadas
    try:
        base_organizada_path = entry_base_organizada.get()
        base_sin_ordenar_path = entry_base_sin_ordenar.get()
        campo_id_organizada = opciones_campo_id_organizada.get()
        campo_id_sin_ordenar = opciones_campo_id_sin_ordenar.get()
        
        if base_organizada_path and base_sin_ordenar_path and campo_id_organizada and campo_id_sin_ordenar:
            # Leer las bases de datos
            df_organizada = pd.read_excel(base_organizada_path)
            df_sin_ordenar = pd.read_excel(base_sin_ordenar_path)
            
            # Realizar fusión preliminar usando los campos seleccionados
            df_fusionado = pd.merge(df_organizada, df_sin_ordenar, left_on=campo_id_organizada, right_on=campo_id_sin_ordenar, how="inner")
            
            # Obtener las cabeceras del DataFrame fusionado
            columnas_fusionadas = list(df_fusionado.columns)
            
            # Actualizar el menú desplegable de campo de orden con las cabeceras fusionadas
            actualizar_opciones_menus(campo_orden_menu, opciones_campo_orden, columnas_fusionadas)
            messagebox.showinfo("Éxito", "Fusión preliminar realizada. Ahora puedes seleccionar el campo de ordenación.")
        else:
            messagebox.showwarning("Advertencia", "Selecciona ambos archivos y los campos de identificación antes de realizar la fusión preliminar.")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al fusionar las bases preliminarmente: {e}")

# Función para actualizar opciones en el menú desplegable
def actualizar_opciones_menus(menu, variable, opciones):
    menu['menu'].delete(0, 'end')
    for opcion in opciones:
        menu['menu'].add_command(label=opcion, command=tk._setit(variable, opcion))

# Función para reorganizar las bases y guardar el archivo final
def reorganizar_bases_final():
    try:
        base_organizada_path = entry_base_organizada.get()
        base_sin_ordenar_path = entry_base_sin_ordenar.get()
        output_path = entry_output_path.get()
        campo_id_organizada = opciones_campo_id_organizada.get()
        campo_id_sin_ordenar = opciones_campo_id_sin_ordenar.get()
        campo_orden_seleccionado = opciones_campo_orden.get()

        if not all([base_organizada_path, base_sin_ordenar_path, output_path, campo_id_organizada, campo_id_sin_ordenar, campo_orden_seleccionado]):
            messagebox.showerror("Error", "Por favor, selecciona todos los archivos y campos.")
            return

        # Leer las bases de datos
        df_organizada = pd.read_excel(base_organizada_path)
        df_sin_ordenar = pd.read_excel(base_sin_ordenar_path)

        # Verificar si la columna de orden existe en la base organizada
        if campo_orden_seleccionado not in df_organizada.columns:
            messagebox.showerror("Error", f"La columna '{campo_orden_seleccionado}' no se encuentra en la base organizada.")
            return

        # Fusionar ambas bases de datos usando la base sin ordenar como principal
        df_fusionado = pd.merge(
            df_organizada[[campo_id_organizada, campo_orden_seleccionado]],
            df_sin_ordenar,
            left_on=campo_id_organizada,
            right_on=campo_id_sin_ordenar,
            how="right"
        )

        # Ordenar el DataFrame fusionado según la columna de "Orden de Impresión"
        df_reorganizada = df_fusionado.sort_values(by=campo_orden_seleccionado)

        # Agregar la columna de "Orden de Impresión" y conservar las de la base sin ordenar
        columnas_finales = list(df_sin_ordenar.columns) + [campo_orden_seleccionado]
        df_reorganizada = df_reorganizada[columnas_finales]

        # Guardar el resultado en un nuevo archivo Excel
        df_reorganizada.to_excel(output_path, index=False)
        
        messagebox.showinfo("Éxito", f"Archivo reorganizado guardado en: {output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {e}")

# Funciones para seleccionar archivos
def seleccionar_base_organizada():
    archivo = filedialog.askopenfilename(title="Seleccionar Base Organizada", filetypes=[("Archivos Excel", "*.xlsx")])
    entry_base_organizada.delete(0, tk.END)
    entry_base_organizada.insert(0, archivo)
    cargar_cabeceras(archivo, "organizada")

def seleccionar_base_sin_ordenar():
    archivo = filedialog.askopenfilename(title="Seleccionar Base Sin Ordenar", filetypes=[("Archivos Excel", "*.xlsx")])
    entry_base_sin_ordenar.delete(0, tk.END)
    entry_base_sin_ordenar.insert(0, archivo)
    cargar_cabeceras(archivo, "sin_ordenar")

def seleccionar_output_path():
    archivo = filedialog.asksaveasfilename(title="Guardar Archivo Reorganizado", defaultextension=".xlsx",
                                           filetypes=[("Archivos Excel", "*.xlsx")])
    entry_output_path.delete(0, tk.END)
    entry_output_path.insert(0, archivo)

# Configuración de la interfaz gráfica
root = tk.Tk()
root.title("Reorganizador de Bases de Datos")
root.geometry("500x600")

# Widgets para seleccionar archivos
tk.Label(root, text="Base Organizada:").pack(pady=5)
entry_base_organizada = tk.Entry(root, width=50)
entry_base_organizada.pack(pady=5)
tk.Button(root, text="Seleccionar Archivo", command=seleccionar_base_organizada).pack(pady=5)

tk.Label(root, text="Base Sin Ordenar:").pack(pady=5)
entry_base_sin_ordenar = tk.Entry(root, width=50)
entry_base_sin_ordenar.pack(pady=5)
tk.Button(root, text="Seleccionar Archivo", command=seleccionar_base_sin_ordenar).pack(pady=5)

tk.Label(root, text="Guardar Archivo Reorganizado en:").pack(pady=5)
entry_output_path = tk.Entry(root, width=50)
entry_output_path.pack(pady=5)
tk.Button(root, text="Seleccionar Carpeta", command=seleccionar_output_path).pack(pady=5)

# Menús desplegables para selección de campos de identificación
tk.Label(root, text="Campo de Identificación en Base Organizada:").pack(pady=5)
opciones_campo_id_organizada = tk.StringVar(root)
campo_id_organizada_menu = tk.OptionMenu(root, opciones_campo_id_organizada, "")
campo_id_organizada_menu.pack(pady=5)

tk.Label(root, text="Campo de Identificación en Base Sin Ordenar:").pack(pady=5)
opciones_campo_id_sin_ordenar = tk.StringVar(root)
campo_id_sin_ordenar_menu = tk.OptionMenu(root, opciones_campo_id_sin_ordenar, "")
campo_id_sin_ordenar_menu.pack(pady=5)

# Botón para realizar la fusión preliminar
tk.Button(root, text="Fusión Preliminar para Selección de Columnas", command=fusionar_bases_preliminar).pack(pady=10)

# Menú desplegable para seleccionar el campo de ordenación
tk.Label(root, text="Campo para Ordenación:").pack(pady=5)
opciones_campo_orden = tk.StringVar(root)
campo_orden_menu = tk.OptionMenu(root, opciones_campo_orden, "")
campo_orden_menu.pack(pady=5)

# Botón para reorganizar y guardar el archivo
tk.Button(root, text="Reorganizar y Guardar", command=reorganizar_bases_final).pack(pady=10)

root.mainloop()
