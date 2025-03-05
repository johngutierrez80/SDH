import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Función para cargar y verificar duplicados en el archivo Excel
def verificar_duplicados():
    archivo = entry_archivo.get()
    
    if archivo:
        try:
            # Cargar el archivo Excel en un DataFrame
            df = pd.read_excel(archivo)
            
            # Verificar duplicados (puedes especificar columnas si es necesario)
            duplicados = df[df.duplicated(keep=False)]
            
            if not duplicados.empty:
                # Crear una nueva ventana para mostrar los duplicados
                ventana_duplicados = tk.Toplevel()
                ventana_duplicados.title("Registros Duplicados")
                
                # Crear un marco y un canvas para permitir el scroll horizontal y vertical
                frame = tk.Frame(ventana_duplicados)
                frame.pack(fill=tk.BOTH, expand=True)
                
                canvas = tk.Canvas(frame)
                canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                
                scrollbar_y = tk.Scrollbar(frame, orient=tk.VERTICAL, command=canvas.yview)
                scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
                
                scrollbar_x = tk.Scrollbar(ventana_duplicados, orient=tk.HORIZONTAL, command=canvas.xview)
                scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
                
                canvas.configure(xscrollcommand=scrollbar_x.set, yscrollcommand=scrollbar_y.set)
                
                table_frame = tk.Frame(canvas)
                canvas.create_window((0, 0), window=table_frame, anchor="nw")
                
                # Ajustar el tamaño del frame después de agregar el contenido
                def ajustar_scroll(event):
                    canvas.configure(scrollregion=canvas.bbox("all"))
                
                table_frame.bind("<Configure>", ajustar_scroll)
                
                # Añadir el número de registro como la primera columna
                duplicados.insert(0, "Registro", duplicados.index + 1)
                
                # Zebra striping - alternar colores en las filas
                colores_filas = ["#f2f2f2", "#ffffff"]
                
                # Mostrar los nombres de las columnas
                for j, col in enumerate(duplicados.columns):
                    label = tk.Label(table_frame, text=col, relief="solid", padx=5, pady=5, bg="#4CAF50", fg="white")
                    label.grid(row=0, column=j, sticky="nsew")
                
                # Mostrar los registros duplicados
                for i, row in duplicados.iterrows():
                    bg_color = colores_filas[i % 2]
                    for j, value in enumerate(row):
                        label = tk.Label(table_frame, text=value, relief="solid", padx=5, pady=5, bg=bg_color)
                        label.grid(row=i+1, column=j, sticky="nsew")
                
                # Ajuste de ancho de columnas
                for j in range(len(duplicados.columns)):
                    table_frame.grid_columnconfigure(j, weight=1)
                
            else:
                messagebox.showinfo("Información", "No se encontraron registros duplicados.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el archivo: {e}")
    else:
        messagebox.showwarning("Advertencia", "Por favor, ingrese la ruta del archivo.")

# Función para abrir el diálogo de selección de archivo
def seleccionar_archivo():
    archivo = filedialog.askopenfilename(title="Seleccione el archivo Excel", filetypes=[("Excel files", "*.xlsx *.xls")])
    entry_archivo.delete(0, tk.END)
    entry_archivo.insert(0, archivo)

# Crear la ventana principal de Tkinter
ventana = tk.Tk()
ventana.title("Verificador de Duplicados en Excel")

# Label y Entry para la ruta del archivo
label_archivo = tk.Label(ventana, text="Ruta del archivo Excel:")
label_archivo.pack(pady=5)
entry_archivo = tk.Entry(ventana, width=50)
entry_archivo.pack(pady=5)

# Botón para seleccionar el archivo
boton_seleccionar = tk.Button(ventana, text="Seleccionar Archivo", command=seleccionar_archivo)
boton_seleccionar.pack(pady=5)

# Botón para iniciar la verificación
boton_verificar = tk.Button(ventana, text="Verificar Duplicados", command=verificar_duplicados)
boton_verificar.pack(pady=20)

# Iniciar la aplicación Tkinter
ventana.mainloop()
