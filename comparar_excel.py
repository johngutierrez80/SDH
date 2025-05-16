import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *

def load_excel_file():
    """Abre un diálogo para seleccionar un archivo Excel."""
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    return file_path

def get_excel_sheets(file_path):
    """Devuelve la lista de hojas de un archivo Excel o un error si falla."""
    try:
        if not file_path or not os.path.exists(file_path):
            return None, "Error: La ruta del archivo no es válida o no existe."
        excel_file = pd.ExcelFile(file_path)
        return excel_file.sheet_names, None
    except Exception as e:
        return None, f"Error al leer las hojas: {str(e)}"

def get_columns(file_path, sheet_name):
    """Devuelve la lista de columnas de una hoja específica."""
    try:
        if not file_path or not os.path.exists(file_path):
            return None, "Error: La ruta del archivo no es válida o no existe."
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        return df.columns.tolist(), None
    except Exception as e:
        return None, f"Error al leer las columnas: {str(e)}"

def compare_excel_files(file1_path, file2_path, sheet1_name, sheet2_name, compare_field):
    """Compara dos archivos Excel basándose en un campo específico."""
    try:
        # Leer los archivos Excel con las hojas especificadas
        df1 = pd.read_excel(file1_path, sheet_name=sheet1_name)
        df2 = pd.read_excel(file2_path, sheet_name=sheet2_name)

        # Verificar que el campo existe en ambos DataFrames
        if compare_field not in df1.columns or compare_field not in df2.columns:
            return None, f"Error: El campo '{compare_field}' no existe en uno de los archivos."

        # Convertir el campo de comparación a un tipo consistente (cadena)
        df1[compare_field] = df1[compare_field].astype(str)
        df2[compare_field] = df2[compare_field].astype(str)

        # Ordenar ambos DataFrames por el campo seleccionado
        df1 = df1.sort_values(by=compare_field).reset_index(drop=True)
        df2 = df2.sort_values(by=compare_field).reset_index(drop=True)

        # Alinear los DataFrames basándose en el campo (merge para encontrar coincidencias)
        merged = df1.merge(df2, on=compare_field, how="outer", suffixes=('_file1', '_file2'), indicator=True)

        differences = []
        # Identificar valores que no coinciden en el campo
        for idx, row in merged.iterrows():
            if row['_merge'] == 'left_only':
                differences.append({
                    "Fila": idx + 1,
                    "Columna": compare_field,
                    "Valor Archivo 1": str(row[compare_field]),
                    "Valor Archivo 2": "No existe",
                    "Campo": compare_field,
                    "Descripción": f"Valor '{row[compare_field]}' presente en Archivo 1 pero no en Archivo 2"
                })
            elif row['_merge'] == 'right_only':
                differences.append({
                    "Fila": idx + 1,
                    "Columna": compare_field,
                    "Valor Archivo 1": "No existe",
                    "Valor Archivo 2": str(row[compare_field]),
                    "Campo": compare_field,
                    "Descripción": f"Valor '{row[compare_field]}' presente en Archivo 2 pero no en Archivo 1"
                })

        # Filtrar solo las filas que coinciden en el campo
        matched = merged[merged['_merge'] == 'both'].reset_index(drop=True)

        # Comparar las filas coincidentes celda por celda
        for idx in range(len(matched)):
            row1 = matched.iloc[idx]
            # Obtener los valores de las columnas originales
            for col in df1.columns:
                if col == compare_field:
                    continue  # No comparar el campo de unión
                val1 = row1[f"{col}_file1"]
                val2 = row1[f"{col}_file2"]
                # Convertir ambos valores a cadenas para evitar errores de tipo
                val1_str = str(val1) if not pd.isna(val1) else "NaN"
                val2_str = str(val2) if not pd.isna(val2) else "NaN"
                if val1_str == val2_str:
                    continue
                differences.append({
                    "Fila": idx + 1,
                    "Columna": col,
                    "Valor Archivo 1": val1_str,
                    "Valor Archivo 2": val2_str,
                    "Campo": col,
                    "Descripción": f"Diferencia en el campo '{col}' para el valor '{row1[compare_field]}' en {compare_field}"
                })

        return differences, None
    except Exception as e:
        return None, f"Error al procesar los archivos: {str(e)}"

def main():
    # Crear ventana principal con tema moderno
    root = ttkb.Window(themename="flatly")
    root.title("Comparador de Archivos Excel por Campo")
    root.geometry("1000x600")
    root.resizable(True, True)

    # Frame principal
    main_frame = ttk.Frame(root, padding=20)
    main_frame.pack(fill=BOTH, expand=True)

    # Título
    title_label = ttk.Label(main_frame, text="Comparador de Archivos Excel por Campo", font=("Helvetica", 18, "bold"))
    title_label.pack(pady=10)

    # Frame para selección de archivos
    file_frame = ttk.Frame(main_frame)
    file_frame.pack(fill=X, pady=10)

    # Variables para almacenar rutas, hojas y campo
    file1_path_var = tk.StringVar()
    file2_path_var = tk.StringVar()
    sheet1_var = tk.StringVar()
    sheet2_var = tk.StringVar()
    compare_field_var = tk.StringVar()

    # Selección archivo 1
    ttk.Label(file_frame, text="Archivo 1:").grid(row=0, column=0, padx=5, sticky=W)
    ttk.Entry(file_frame, textvariable=file1_path_var, width=40).grid(row=0, column=1, padx=5)
    ttk.Button(file_frame, text="Seleccionar", command=lambda: update_file1()).grid(row=0, column=2, padx=5)

    # Selección hoja 1
    ttk.Label(file_frame, text="Hoja 1:").grid(row=0, column=3, padx=5, sticky=W)
    sheet1_combo = ttk.Combobox(file_frame, textvariable=sheet1_var, state="readonly", width=15)
    sheet1_combo.grid(row=0, column=4, padx=5)

    # Selección archivo 2
    ttk.Label(file_frame, text="Archivo 2:").grid(row=1, column=0, padx=5, sticky=W)
    ttk.Entry(file_frame, textvariable=file2_path_var, width=40).grid(row=1, column=1, padx=5)
    ttk.Button(file_frame, text="Seleccionar", command=lambda: update_file2()).grid(row=1, column=2, padx=5)

    # Selección hoja 2
    ttk.Label(file_frame, text="Hoja 2:").grid(row=1, column=3, padx=5, sticky=W)
    sheet2_combo = ttk.Combobox(file_frame, textvariable=sheet2_var, state="readonly", width=15)
    sheet2_combo.grid(row=1, column=4, padx=5)

    # Selección del campo para comparar
    ttk.Label(file_frame, text="Campo para comparar:").grid(row=2, column=0, padx=5, sticky=W)
    compare_field_combo = ttk.Combobox(file_frame, textvariable=compare_field_var, state="readonly", width=20)
    compare_field_combo.grid(row=2, column=1, padx=5, sticky=W)

    def update_file1():
        file_path = load_excel_file()
        file1_path_var.set(file_path)
        sheets, error = get_excel_sheets(file_path)
        if error:
            messagebox.showerror("Error", error)
        else:
            sheet1_combo["values"] = sheets
            sheet1_var.set(sheets[0] if sheets else "")
            # Actualizar lista de columnas
            columns, error = get_columns(file_path, sheet1_var.get())
            if error:
                messagebox.showerror("Error", error)
            else:
                compare_field_combo["values"] = columns
                compare_field_var.set(columns[0] if columns else "")

    def update_file2():
        file_path = load_excel_file()
        file2_path_var.set(file_path)
        sheets, error = get_excel_sheets(file_path)
        if error:
            messagebox.showerror("Error", error)
        else:
            sheet2_combo["values"] = sheets
            sheet2_var.set(sheets[0] if sheets else "")

    # Botón para comparar
    compare_button = ttk.Button(main_frame, text="Comparar Archivos", style="primary.TButton", command=lambda: compare_and_display())
    compare_button.pack(pady=20)

    # Frame para la tabla de diferencias
    table_frame = ttk.Frame(main_frame)
    table_frame.pack(fill=BOTH, expand=True)

    # Tabla (Treeview)
    columns = ("Fila", "Columna", "Valor Archivo 1", "Valor Archivo 2", "Campo", "Descripción")
    tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=15)
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=150, anchor="center")
    # Ajustar ancho de la columna Descripción
    tree.column("Descripción", width=300)
    tree.pack(side=LEFT, fill=BOTH, expand=True)

    # Scrollbar
    scrollbar = ttk.Scrollbar(table_frame, orient=VERTICAL, command=tree.yview)
    scrollbar.pack(side=RIGHT, fill=Y)
    tree.configure(yscrollcommand=scrollbar.set)

    def compare_and_display():
        """Ejecuta la comparación y muestra los resultados en la tabla."""
        file1_path = file1_path_var.get()
        file2_path = file2_path_var.get()
        sheet1_name = sheet1_var.get()
        sheet2_name = sheet2_var.get()
        compare_field = compare_field_var.get()

        if not file1_path or not file2_path:
            messagebox.showwarning("Advertencia", "Por favor, seleccione ambos archivos.")
            return
        if not sheet1_name or not sheet2_name:
            messagebox.showwarning("Advertencia", "Por favor, seleccione una hoja para cada archivo.")
            return
        if not compare_field:
            messagebox.showwarning("Advertencia", "Por favor, seleccione un campo para comparar.")
            return

        # Limpiar tabla
        for item in tree.get_children():
            tree.delete(item)

        # Comparar archivos
        differences, error = compare_excel_files(file1_path, file2_path, sheet1_name, sheet2_name, compare_field)
        if error:
            messagebox.showerror("Error", error)
            return

        if not differences:
            messagebox.showinfo("Resultado", "No se encontraron diferencias entre los archivos.")
            return

        # Mostrar diferencias en la tabla
        for diff in differences:
            tree.insert("", tk.END, values=(
                diff["Fila"],
                diff["Columna"],
                diff["Valor Archivo 1"],
                diff["Valor Archivo 2"],
                diff["Campo"],
                diff["Descripción"]
            ))

        messagebox.showinfo("Resultado", f"Se encontraron {len(differences)} diferencias.")

    # Iniciar el bucle principal
    root.mainloop()

if __name__ == "__main__":
    main()