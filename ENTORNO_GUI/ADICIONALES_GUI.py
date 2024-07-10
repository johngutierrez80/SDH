import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import pandas as pd
import os
import re

# Función para obtener el nombre del archivo de una cadena de texto.
def obtener_nombre_archivo(nombre_archivo):
    if re.match(r'\d{4}EE\d{6}O?\d?', nombre_archivo):
        return nombre_archivo
    else:
        return nombre_archivo[:10]

# Función para cargar los archivos descargados en un DataFrame.
def cargar_archivos_descargados(directorio_descargas):
    try:
        archivos_descargados = os.listdir(directorio_descargas)
        nombres_archivos_tratados = [obtener_nombre_archivo(nombre.split('.')[0]) for nombre in archivos_descargados]
        df_archivos_descargados = pd.DataFrame(nombres_archivos_tratados, columns=['Archivo_descargado'])
        return df_archivos_descargados
    except Exception as e:
        raise ValueError(f"Error al cargar archivos descargados: {str(e)}")

# Función para comparar los archivos descargados con los datos del archivo Excel.
def comparar_archivos(excel_path, directorio_descargas, sheet_name='Hoja1', campo_especifico='SAP_ID'):
    try:
        if not excel_path or not directorio_descargas:
            raise ValueError("Por favor, seleccione la ruta del archivo Excel y el directorio de descargas.")

        df_excel = pd.read_excel(excel_path, sheet_name=sheet_name)
        df_excel.columns = df_excel.columns.str.strip()
        if 'Numero de secuencia' not in df_excel.columns:
            raise ValueError("El campo 'Numero de secuencia' no está presente en el DataFrame.")

        df_descargas = cargar_archivos_descargados(directorio_descargas)
        nombres_descargas = df_descargas['Archivo_descargado']
        sap_ids_excel = df_excel[campo_especifico].astype(str)

        sap_ids_no_en_descargas = sap_ids_excel[~sap_ids_excel.isin(nombres_descargas)]
        descargas_adicionales = df_descargas[~df_descargas['Archivo_descargado'].isin(sap_ids_excel)]
        descargas_duplicadas = df_descargas[df_descargas.duplicated(subset=['Archivo_descargado'], keep=False)]

        resultados_window = tk.Toplevel(root)
        resultados_window.title("Resultados de la comparación")

        window_width = 400
        window_height = 400
        position_right = int(resultados_window.winfo_screenwidth() / 2 - window_width / 2)
        position_down = int(resultados_window.winfo_screenheight() / 2 - window_height / 2)
        resultados_window.geometry("{}x{}+{}+{}".format(window_width, window_height, position_right, position_down))

        if not sap_ids_no_en_descargas.empty:
            resultados_label = tk.Label(resultados_window, text="Archivos que no están presentes en la carpeta de descargas:")
            resultados_label.pack()

            tree = ttk.Treeview(resultados_window, columns=("SAP_ID", "Numero de secuencia"))

            tree.column("#0", width=0, stretch=tk.NO)
            tree.column("SAP_ID", anchor=tk.W, width=100)
            tree.column("Numero de secuencia", anchor=tk.W, width=150)
            tree.heading("#0", text="", anchor=tk.W)
            tree.heading("SAP_ID", text="SAP_ID", anchor=tk.W)
            tree.heading("Numero de secuencia", text="Numero de secuencia", anchor=tk.W)

            scrollbar = ttk.Scrollbar(resultados_window, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=scrollbar.set)
            scrollbar.pack(side="right", fill="y")
            tree.pack(fill="both", expand=True)

            for sap_id in sap_ids_no_en_descargas:
                matching_row = df_excel[df_excel[campo_especifico].astype(str).str.strip() == sap_id.strip()]
                if not matching_row.empty:
                    numero_secuencia = str(matching_row.iloc[0]['Numero de secuencia'])
                    tree.insert("", "end", values=(sap_id, numero_secuencia))
                else:
                    tree.insert("", "end", values=(sap_id, "No se encontró"))

            def treeview_sort_column(col):
                items = [(tree.set(k, col), k) for k in tree.get_children('')]
                items.sort()
                for index, (val, k) in enumerate(items):
                    tree.move(k, '', index)
                tree.heading(col, command=lambda: treeview_sort_column(col))

            for col in ("SAP_ID", "Numero de secuencia"):
                tree.heading(col, text=col, command=lambda c=col: treeview_sort_column(c))

        else:
            tk.Label(resultados_window, text="Todos los archivos están descargados.").pack()

        if not descargas_adicionales.empty:
            adicionales_label = tk.Label(resultados_window, text="Archivos adicionales en el directorio de descargas:")
            adicionales_label.pack()

            tree_adicionales = ttk.Treeview(resultados_window, columns=("Archivo_adicional"))

            tree_adicionales.column("#0", width=0, stretch=tk.NO)
            tree_adicionales.column("Archivo_adicional", anchor=tk.W, width=200)
            tree_adicionales.heading("#0", text="", anchor=tk.W)
            tree_adicionales.heading("Archivo_adicional", text="Archivo adicional", anchor=tk.W)

            scrollbar_adicionales = ttk.Scrollbar(resultados_window, orient="vertical", command=tree_adicionales.yview)
            tree_adicionales.configure(yscrollcommand=scrollbar_adicionales.set)
            scrollbar_adicionales.pack(side="right", fill="y")
            tree_adicionales.pack(fill="both", expand=True)

            for archivo in descargas_adicionales['Archivo_descargado']:
                tree_adicionales.insert("", "end", values=(archivo,))

        if not descargas_duplicadas.empty:
            duplicados_label = tk.Label(resultados_window, text="Archivos duplicados en el directorio de descargas:")
            duplicados_label.pack()

            tree_duplicados = ttk.Treeview(resultados_window, columns=("Archivo_duplicado"))

            tree_duplicados.column("#0", width=0, stretch=tk.NO)
            tree_duplicados.column("Archivo_duplicado", anchor=tk.W, width=200)
            tree_duplicados.heading("#0", text="", anchor=tk.W)
            tree_duplicados.heading("Archivo_duplicado", text="Archivo duplicado", anchor=tk.W)

            scrollbar_duplicados = ttk.Scrollbar(resultados_window, orient="vertical", command=tree_duplicados.yview)
            tree_duplicados.configure(yscrollcommand=scrollbar_duplicados.set)
            scrollbar_duplicados.pack(side="right", fill="y")
            tree_duplicados.pack(fill="both", expand=True)

            for archivo in descargas_duplicadas['Archivo_descargado']:
                tree_duplicados.insert("", "end", values=(archivo,))

        cerrar_button = tk.Button(resultados_window, text="Cerrar", command=resultados_window.destroy, bg="red", font=("Arial", 10), width=15)
        cerrar_button.pack(side="bottom", pady=10, anchor="center")

        resultados_window.grab_set()

    except Exception as e:
        messagebox.showerror("Error", str(e))

# Función para seleccionar el archivo Excel.
def seleccionar_archivo_excel():
    global excel_path
    excel_path = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx;*.xls")])
    entrada_archivo_excel.delete(0, tk.END)
    entrada_archivo_excel.insert(0, excel_path)

# Función para seleccionar el directorio de descargas.
def seleccionar_directorio_descargas():
    global directorio_descargas
    directorio_descargas = filedialog.askdirectory()
    entrada_directorio_descargas.delete(0, tk.END)
    entrada_directorio_descargas.insert(0, directorio_descargas)

# Función para iniciar la comparación de archivos.
def iniciar_comparacion():
    comparar_archivos(excel_path, directorio_descargas)

# Función para limpiar los campos de entrada.
def limpiar_campos():
    entrada_archivo_excel.delete(0, tk.END)
    entrada_directorio_descargas.delete(0, tk.END)

# Crear la ventana principal
root = tk.Tk()
root.title("Comparador de Archivos")

window_width = 700
window_height = 200
position_right = int(root.winfo_screenwidth() / 2 - window_width / 2)
position_down = int(root.winfo_screenheight() / 2 - window_height / 2)
root.geometry("{}x{}+{}+{}".format(window_width, window_height, position_right, position_down))


icon_image = Image.open(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\INC.ico")
icon_photo = ImageTk.PhotoImage(icon_image)
root.iconphoto(False, icon_photo)


# Frame para el formulario
frame_formulario = tk.Frame(root)
frame_formulario.pack(pady=20)

label_archivo_excel = tk.Label(frame_formulario, text="Archivo Excel:")
label_archivo_excel.grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
entrada_archivo_excel = tk.Entry(frame_formulario, width=40)
entrada_archivo_excel.grid(row=0, column=1, padx=10, pady=5)
boton_archivo_excel = tk.Button(frame_formulario, text="Seleccionar", command=seleccionar_archivo_excel, bg="lightblue", font=("Arial", 10), width=20)
boton_archivo_excel.grid(row=0, column=2, padx=10, pady=5)

label_directorio_descargas = tk.Label(frame_formulario, text="Directorio de Descargas:")
label_directorio_descargas.grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)
entrada_directorio_descargas = tk.Entry(frame_formulario, width=40)
entrada_directorio_descargas.grid(row=1, column=1, padx=10, pady=5)
boton_directorio_descargas = tk.Button(frame_formulario, text="Seleccionar", command=seleccionar_directorio_descargas, bg="lightblue", font=("Arial", 10), width=20)
boton_directorio_descargas.grid(row=1, column=2, padx=10, pady=5)

boton_comparar = tk.Button(frame_formulario, text="Comparar Archivos", command=iniciar_comparacion, bg="green", font=("Arial", 10), width=20)
boton_comparar.grid(row=2, column=1, padx=10, pady=10)

# Botón para limpiar campos
boton_limpiar = tk.Button(frame_formulario, text="Limpiar Campos", command=limpiar_campos, bg="orange", font=("Arial", 10), width=20)
boton_limpiar.grid(row=2, column=0, padx=10, pady=10)

# Botón para cerrar la ventana principal
boton_cerrar = tk.Button(frame_formulario, text="Cerrar", command=root.destroy, bg="red", font=("Arial", 10), width=20)
boton_cerrar.grid(row=2, column=2, padx=10, pady=10)

root.mainloop()
