import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Funciones para la lógica del programa
def cargar_datos(ruta_archivo):
    try:
        df = pd.read_excel(ruta_archivo)
        return df
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo cargar el archivo. Error: {e}")
        return None

def combinar_registros(df, columna_principal, columnas_unicas, columnas_concatenar):
    # Verificar si la columna principal está en el DataFrame
    if columna_principal not in df.columns:
        messagebox.showerror("Error", f"La columna '{columna_principal}' no se encuentra en los datos.")
        return None

    # Definir las columnas únicas y las que se concatenan
    agg_dict = {col: 'first' for col in columnas_unicas}
    for col in columnas_concatenar:
        agg_dict[col] = lambda x: '\n'.join(map(lambda v: str(v) if pd.notnull(v) else " ", x))

    try:
        df_unificado = df.groupby(columna_principal, as_index=False).agg(agg_dict).reset_index(drop=True)
        # Reemplazar cualquier valor restante de nan en el DataFrame
        df_unificado = df_unificado.fillna(" ")
        return df_unificado
    except Exception as e:
        messagebox.showerror("Error", f"Error al combinar registros: {e}")
        return None

def formatear_datos(df, columnas_monetarias):
    def formatear_monetario(valores):
        valores_separados = str(valores).split("\n")
        valores_formateados = [f"$ {int(float(x)):,}" if x.strip() != "" else "$ 0" for x in valores_separados]
        return "\n".join(valores_formateados)
    
    for col in columnas_monetarias:
        df[col] = df[col].apply(formatear_monetario)
    
    return df

def exportar_datos(df, ruta_salida):
    try:
        writer = pd.ExcelWriter(ruta_salida, engine='openpyxl')
        df.to_excel(writer, index=False)
        writer.close()
        messagebox.showinfo("Éxito", f"Archivo exportado correctamente a {ruta_salida}")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo exportar el archivo. Error: {e}")

# Funciones para la interfaz gráfica
def seleccionar_archivo():
    archivo = filedialog.askopenfilename(title="Seleccionar archivo", filetypes=[("Archivos Excel", "*.xlsx")])
    if archivo:
        entrada_archivo.set(archivo)
        df = cargar_datos(archivo)
        if df is not None:
            mostrar_seleccion_columnas(df)

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
        columna_principal = columna_principal_var.get()
        columnas_unicas = [col for col, val in columnas_unicas_var.items() if val.get()]
        columnas_concatenar = [col for col, val in columnas_concatenar_var.items() if val.get()]

        if not columnas_unicas or not columna_principal:
            messagebox.showwarning("Advertencia", "Debe seleccionar al menos una columna única y la columna principal.")
            return

        df_combinado = combinar_registros(df, columna_principal, columnas_unicas, columnas_concatenar)
        if df_combinado is not None:
            columnas_monetarias = [col for col, val in columnas_monetarias_var.items() if val.get()]
            df_formateado = formatear_datos(df_combinado, columnas_monetarias)
            exportar_datos(df_formateado, ruta_salida)

# Mostrar la selección de columnas únicas, concatenadas y la principal
def mostrar_seleccion_columnas(df):
    for widget in frame_columnas.winfo_children():
        widget.destroy()

    canvas = tk.Canvas(frame_columnas)
    scrollbar = tk.Scrollbar(frame_columnas, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)
    scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    tk.Label(scrollable_frame, text="Seleccionar columna principal:").pack()
    global columna_principal_var
    columna_principal_var = tk.StringVar()
    for col in df.columns:
        tk.Radiobutton(scrollable_frame, text=col, variable=columna_principal_var, value=col).pack(anchor='w')

    tk.Label(scrollable_frame, text="Seleccionar columnas únicas:").pack()
    global columnas_unicas_var
    columnas_unicas_var = {}
    for col in df.columns:
        var = tk.BooleanVar()
        chk = tk.Checkbutton(scrollable_frame, text=col, variable=var)
        chk.pack(anchor='w')
        columnas_unicas_var[col] = var

    tk.Label(scrollable_frame, text="Seleccionar columnas a concatenar:").pack()
    global columnas_concatenar_var
    columnas_concatenar_var = {}
    for col in df.columns:
        var = tk.BooleanVar()
        chk = tk.Checkbutton(scrollable_frame, text=col, variable=var)
        chk.pack(anchor='w')
        columnas_concatenar_var[col] = var

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    btn_siguiente.pack(pady=20)

# Mostrar la selección de columnas monetarias
def mostrar_seleccion_monedas(df):
    for widget in frame_columnas.winfo_children():
        widget.destroy()

    canvas = tk.Canvas(frame_columnas)
    scrollbar = tk.Scrollbar(frame_columnas, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)
    scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    tk.Label(scrollable_frame, text="Seleccionar columnas a formatear como moneda:").pack()
    global columnas_monetarias_var
    columnas_monetarias_var = {}
    for col in df.columns:
        var = tk.BooleanVar()
        chk = tk.Checkbutton(scrollable_frame, text=col, variable=var)
        chk.pack(anchor='w')
        columnas_monetarias_var[col] = var

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    btn_procesar.pack(pady=20)

# Acción del botón "Siguiente paso"
def siguiente_paso():
    ruta_archivo = entrada_archivo.get()
    df = cargar_datos(ruta_archivo)
    if df is not None:
        mostrar_seleccion_monedas(df)
        btn_siguiente.pack_forget()

# Configuración de la ventana principal
ventana = tk.Tk()
ventana.title("Combinación de Registros en Excel")
ventana.geometry("600x600")

entrada_archivo = tk.StringVar()
salida_archivo = tk.StringVar()

tk.Label(ventana, text="Archivo de Entrada:").pack(pady=10)
tk.Entry(ventana, textvariable=entrada_archivo, width=50).pack(pady=5)
tk.Button(ventana, text="Seleccionar Archivo", command=seleccionar_archivo).pack(pady=5)

tk.Label(ventana, text="Archivo de Salida:").pack(pady=10)
tk.Entry(ventana, textvariable=salida_archivo, width=50).pack(pady=5)
tk.Button(ventana, text="Seleccionar Ubicación de Guardado", command=seleccionar_guardado).pack(pady=5)

frame_columnas = tk.Frame(ventana)
frame_columnas.pack(pady=10, fill="both", expand=True)

btn_siguiente = tk.Button(ventana, text="Siguiente paso", command=siguiente_paso)
btn_siguiente.pack(pady=10)

btn_procesar = tk.Button(ventana, text="Procesar Archivo", command=procesar_archivo)

ventana.mainloop()
