import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import scrolledtext  # Para tooltips simulados si no usamos biblioteca externa
from typing import List, Optional
import time  # Para simular progreso en la barra

# Funciones de lógica (sin cambios significativos)
def cargar_datos(ruta_archivo: str) -> Optional[pd.DataFrame]:
    try:
        return pd.read_excel(ruta_archivo, dtype_backend='numpy_nullable')
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo cargar el archivo: {e}")
        return None

def combinar_registros(df: pd.DataFrame, columna_principal: str, 
                      columnas_unicas: List[str], columnas_concatenar: List[str]) -> Optional[pd.DataFrame]:
    if columna_principal not in df.columns:
        messagebox.showerror("Error", f"La columna '{columna_principal}' no existe.")
        return None
    try:
        agg_dict = {col: 'first' for col in columnas_unicas}
        agg_dict.update({col: lambda x: '\n'.join(str(v) if pd.notna(v) else " " for v in x) for col in columnas_concatenar})
        return df.groupby(columna_principal, as_index=False).agg(agg_dict).fillna(" ")
    except Exception as e:
        messagebox.showerror("Error", f"Error al combinar: {e}")
        return None

def formatear_datos(df: pd.DataFrame, columnas_monetarias: List[str], 
                    columnas_porcentajes: List[str]) -> pd.DataFrame:
    def formatear_monetario(valores):
        valores_separados = str(valores).split("\n")
        return "\n".join(f"$ {int(float(x)):,}" if x.strip() and x != " " else "$ 0" for x in valores_separados)
    
    def formatear_porcentaje(valores):
        valores_separados = str(valores).split("\n")
        return "\n".join(f"{float(x)*100:.2f}%" if x.strip() and x != " " else "0.00%" for x in valores_separados)

    df_formateado = df.copy()
    for col in columnas_monetarias:
        if col in df_formateado.columns:
            df_formateado[col] = df_formateado[col].apply(formatear_monetario)
    for col in columnas_porcentajes:
        if col in df_formateado.columns:
            df_formateado[col] = df_formateado[col].apply(formatear_porcentaje)
    return df_formateado

def exportar_datos(df: pd.DataFrame, ruta_salida: str, progress: ttk.Progressbar) -> None:
    try:
        progress['value'] = 0
        ventana.update_idletasks()
        time.sleep(0.5)  # Simulación de progreso
        progress['value'] = 50
        ventana.update_idletasks()
        with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        progress['value'] = 100
        ventana.update_idletasks()
        messagebox.showinfo("Éxito", f"Archivo guardado en: {ruta_salida}")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo exportar: {e}")
    finally:
        progress['value'] = 0

# Funciones de la interfaz gráfica
def seleccionar_archivo():
    archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
    if archivo:
        entrada_archivo.set(archivo)
        df = cargar_datos(archivo)
        if df is not None:
            mostrar_seleccion_columnas(df)
            notebook.select(tab_columnas)

def seleccionar_guardado():
    archivo_salida = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivos Excel", "*.xlsx")])
    if archivo_salida:
        salida_archivo.set(archivo_salida)

def procesar_archivo():
    ruta_archivo, ruta_salida = entrada_archivo.get(), salida_archivo.get()
    if not (ruta_archivo and ruta_salida):
        messagebox.showwarning("Advertencia", "Seleccione archivo de entrada y salida.")
        return

    df = cargar_datos(ruta_archivo)
    if df is None:
        return

    columna_principal = columna_principal_var.get()
    columnas_unicas = [col for col, var in columnas_unicas_var.items() if var.get()]
    columnas_concatenar = [col for col, var in columnas_concatenar_var.items() if var.get()]
    columnas_monetarias = [col for col, var in columnas_monetarias_var.items() if var.get()]
    columnas_porcentajes = [col for col, var in columnas_porcentajes_var.items() if var.get()]

    if not (columna_principal and columnas_unicas):
        messagebox.showwarning("Advertencia", "Seleccione columna principal y al menos una única.")
        return

    df_combinado = combinar_registros(df, columna_principal, columnas_unicas, columnas_concatenar)
    if df_combinado is not None:
        df_formateado = formatear_datos(df_combinado, columnas_monetarias, columnas_porcentajes)
        exportar_datos(df_formateado, ruta_salida, progress_bar)

def reiniciar():
    entrada_archivo.set("")
    salida_archivo.set("")
    for widget in tab_columnas.winfo_children():
        widget.destroy()
    for widget in tab_formato.winfo_children():
        widget.destroy()
    notebook.select(0)
    btn_siguiente.pack(pady=5)
    progress_bar['value'] = 0

def crear_frame_scrollable(parent: tk.Frame) -> tk.Frame:
    canvas = tk.Canvas(parent, bg="#f0f0f0")
    scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas, bg="#f0f0f0")
    scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    return scrollable_frame

# Tooltips simples (sin biblioteca externa como 'tkinter.tooltip')
def mostrar_tooltip(widget, texto):
    def entrar(event):
        x, y = widget.winfo_rootx() + 20, widget.winfo_rooty() + 20
        tooltip = tk.Toplevel(widget)
        tooltip.wm_overrideredirect(True)
        tooltip.wm_geometry(f"+{x}+{y}")
        tk.Label(tooltip, text=texto, bg="yellow", fg="black", relief="solid", borderwidth=1).pack()
        widget.tooltip = tooltip
    def salir(event):
        if hasattr(widget, 'tooltip'):
            widget.tooltip.destroy()
    widget.bind("<Enter>", entrar)
    widget.bind("<Leave>", salir)

def mostrar_seleccion_columnas(df: pd.DataFrame):
    for widget in tab_columnas.winfo_children():
        widget.destroy()

    scrollable_frame = crear_frame_scrollable(tab_columnas)
    
    tk.Label(scrollable_frame, text="Columna Principal", font=("Arial", 10, "bold"), bg="#f0f0f0").pack(pady=(10, 5))
    global columna_principal_var
    columna_principal_var = tk.StringVar()
    frame_principal = tk.Frame(scrollable_frame, bg="#f0f0f0")
    frame_principal.pack(fill="x", padx=10)
    for col in df.columns:
        radio = tk.Radiobutton(frame_principal, text=col, variable=columna_principal_var, value=col, bg="#f0f0f0")
        radio.pack(anchor="w")
        mostrar_tooltip(radio, "Columna que agrupa los registros")

    tk.Label(scrollable_frame, text="Columnas Únicas", font=("Arial", 10, "bold"), bg="#f0f0f0").pack(pady=(10, 5))
    global columnas_unicas_var
    columnas_unicas_var = {col: tk.BooleanVar() for col in df.columns}
    frame_unicas = tk.Frame(scrollable_frame, bg="#f0f0f0")
    frame_unicas.pack(fill="x", padx=10)
    for col in df.columns:
        check = tk.Checkbutton(frame_unicas, text=col, variable=columnas_unicas_var[col], bg="#f0f0f0")
        check.pack(anchor="w")
        mostrar_tooltip(check, "Mantener el primer valor de esta columna")

    tk.Label(scrollable_frame, text="Columnas a Concatenar", font=("Arial", 10, "bold"), bg="#f0f0f0").pack(pady=(10, 5))
    global columnas_concatenar_var
    columnas_concatenar_var = {col: tk.BooleanVar() for col in df.columns}
    frame_concatenar = tk.Frame(scrollable_frame, bg="#f0f0f0")
    frame_concatenar.pack(fill="x", padx=10)
    for col in df.columns:
        check = tk.Checkbutton(frame_concatenar, text=col, variable=columnas_concatenar_var[col], bg="#f0f0f0")
        check.pack(anchor="w")
        mostrar_tooltip(check, "Concatenar valores con saltos de línea")

    btn_siguiente.pack(pady=10)

def mostrar_seleccion_formato(df: pd.DataFrame):
    for widget in tab_formato.winfo_children():
        widget.destroy()

    scrollable_frame = crear_frame_scrollable(tab_formato)
    
    tk.Label(scrollable_frame, text="Columnas Monetarias", font=("Arial", 10, "bold"), bg="#f0f0f0").pack(pady=(10, 5))
    global columnas_monetarias_var
    columnas_monetarias_var = {col: tk.BooleanVar() for col in df.columns}
    frame_monetarias = tk.Frame(scrollable_frame, bg="#f0f0f0")
    frame_monetarias.pack(fill="x", padx=10)
    for col in df.columns:
        check = tk.Checkbutton(frame_monetarias, text=col, variable=columnas_monetarias_var[col], bg="#f0f0f0")
        check.pack(anchor="w")
        mostrar_tooltip(check, "Formatear como moneda ($ 1,234)")

    tk.Label(scrollable_frame, text="Columnas Porcentuales", font=("Arial", 10, "bold"), bg="#f0f0f0").pack(pady=(10, 5))
    global columnas_porcentajes_var
    columnas_porcentajes_var = {col: tk.BooleanVar() for col in df.columns}
    frame_porcentajes = tk.Frame(scrollable_frame, bg="#f0f0f0")
    frame_porcentajes.pack(fill="x", padx=10)
    for col in df.columns:
        check = tk.Checkbutton(frame_porcentajes, text=col, variable=columnas_porcentajes_var[col], bg="#f0f0f0")
        check.pack(anchor="w")
        mostrar_tooltip(check, "Formatear como porcentaje (12.34%)")

    btn_procesar.pack(pady=10)

def siguiente_paso():
    df = cargar_datos(entrada_archivo.get())
    if df is not None:
        mostrar_seleccion_formato(df)
        btn_siguiente.pack_forget()
        notebook.select(tab_formato)

# Configuración de la ventana principal
ventana = tk.Tk()
ventana.title("Combinador de Registros Excel")
ventana.geometry("700x650")
ventana.configure(bg="#e0e0e0")

# Estilo
style = ttk.Style()
style.configure("TNotebook", background="#e0e0e0")
style.configure("TButton", font=("Arial", 10))
style.configure("TProgressbar", thickness=20)

# Frame superior para entradas
frame_entradas = tk.Frame(ventana, bg="#e0e0e0")
frame_entradas.pack(pady=10, padx=10, fill="x")

tk.Label(frame_entradas, text="Archivo de Entrada:", bg="#e0e0e0", font=("Arial", 11)).grid(row=0, column=0, padx=5, pady=5, sticky="e")
entrada_archivo = tk.StringVar()
entrada_widget = tk.Entry(frame_entradas, textvariable=entrada_archivo, width=50)
entrada_widget.grid(row=0, column=1, padx=5, pady=5)
mostrar_tooltip(entrada_widget, "Ruta del archivo Excel a procesar")
ttk.Button(frame_entradas, text="Seleccionar", command=seleccionar_archivo).grid(row=0, column=2, padx=5, pady=5)

tk.Label(frame_entradas, text="Archivo de Salida:", bg="#e0e0e0", font=("Arial", 11)).grid(row=1, column=0, padx=5, pady=5, sticky="e")
salida_archivo = tk.StringVar()
salida_widget = tk.Entry(frame_entradas, textvariable=salida_archivo, width=50)
salida_widget.grid(row=1, column=1, padx=5, pady=5)
mostrar_tooltip(salida_widget, "Ruta donde se guardará el resultado")
ttk.Button(frame_entradas, text="Guardar Como", command=seleccionar_guardado).grid(row=1, column=2, padx=5, pady=5)

# Notebook para pestañas
notebook = ttk.Notebook(ventana)
notebook.pack(pady=10, padx=10, fill="both", expand=True)

tab_columnas = tk.Frame(notebook, bg="#f0f0f0")
tab_formato = tk.Frame(notebook, bg="#f0f0f0")
notebook.add(tab_columnas, text="Selección de Columnas")
notebook.add(tab_formato, text="Formato de Datos")

# Frame inferior para botones y barra de progreso
frame_botones = tk.Frame(ventana, bg="#e0e0e0")
frame_botones.pack(pady=5, fill="x")

btn_siguiente = ttk.Button(frame_botones, text="Siguiente", command=siguiente_paso)
btn_siguiente.pack(side="left", padx=5)
btn_procesar = ttk.Button(frame_botones, text="Procesar", command=procesar_archivo)
btn_reiniciar = ttk.Button(frame_botones, text="Reiniciar", command=reiniciar)
btn_reiniciar.pack(side="left", padx=5)

progress_bar = ttk.Progressbar(frame_botones, length=300, mode='determinate')
progress_bar.pack(side="left", padx=5)

ventana.mainloop()