import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Treeview
import pandas as pd
import os
import re
from PIL import Image, ImageTk
import threading
import time
import sys

def obtener_nombre_archivo(nombre_archivo):
    # Excluir los primeros 7 símbolos del nombre del archivo
    nombre_sin_prefijo = nombre_archivo[7:]
    # Si el nombre del archivo sigue el formato de radicado, se devuelve tal cual
    if re.match(r'\d{4}EE\d{6}O?\d?', nombre_sin_prefijo):
        return nombre_sin_prefijo
    else:
        # Si no sigue el formato de radicado, se extraen los primeros 10 caracteres
        return nombre_sin_prefijo[:10]


def comparar_archivos(excel_path, directorio_renombrados, sheet_name='Hoja1', campo_especifico='SAP_ID'):
    if not excel_path or not directorio_renombrados:
        messagebox.showerror("Error", "Por favor seleccione el archivo de Excel y el directorio de archivos renombrados.")
        return

    try:
        # Leer el archivo de Excel en un DataFrame
        df = pd.read_excel(excel_path, sheet_name=sheet_name)

        # Verificar si 'Numero de secuencia' está presente en el DataFrame
        if 'Numero de secuencia ' not in df.columns:
            raise ValueError("El campo 'Numero de secuencia' no está presente en el DataFrame.")

        # Obtener los nombres de archivo tratados sin la extensión desde el directorio de renombrados
        archivos_renombrados = os.listdir(directorio_renombrados)
        nombres_archivos_renombrados = [obtener_nombre_archivo(nombre.split('.')[0]) for nombre in archivos_renombrados]

        # Crear DataFrame con los nombres de archivo tratados desde el directorio de renombrados
        df_archivos_renombrados = pd.DataFrame(nombres_archivos_renombrados, columns=['Archivo_renombrado'])

        # Obtener el orden de los archivos según el campo específico del DataFrame
        orden_excel = df[campo_especifico].astype(str).tolist()

        # Verificar si los archivos en el directorio renombrado están en el mismo orden que en el archivo Excel
        discrepancias = []
        for idx, (archivo_excel, archivo_renombrado, secuencia_excel) in enumerate(zip(orden_excel, df_archivos_renombrados['Archivo_renombrado'], df['Numero de secuencia ']), 1):
            if archivo_excel != archivo_renombrado:
                discrepancias.append((idx, archivo_renombrado))

        if discrepancias:
            mostrar_tabla_discrepancias(discrepancias, df)  # Pasar el DataFrame a la función
        else:
            messagebox.showinfo("Comparación Exitosa", "Archivos presentes en el directorio  renombrados  en el mismo orden de la Base de Datos.")

    except FileNotFoundError:
        messagebox.showerror("Error", "No se pudo encontrar el archivo especificado.")
    except pd.errors.ParserError:
        messagebox.showerror("Error", "Error al leer el archivo de Excel.")
    except ValueError as e:
        messagebox.showerror("Error", str(e))
    except Exception as e:
        messagebox.showerror("Error", f"Se produjo un error inesperado: {str(e)}")


def mostrar_tabla_discrepancias(discrepancias, df):
    ventana = tk.Toplevel()
    ventana.title("INCONSISTENCIAS RENOMBRADOS")

    # Configuración del scroll vertical
    scroll_y = tk.Scrollbar(ventana, orient="vertical")

    tree = Treeview(ventana, yscrollcommand=scroll_y.set)
    scroll_y.config(command=tree.yview)
    scroll_y.pack(side="right", fill="y")

    tree["columns"] = ("#1", "#2")
    tree.heading("#0", text="Orden Correcto")
    tree.heading("#1", text="Orden Renombrado")

    for idx, (fila, archivo) in enumerate(discrepancias, 1):
        sap_id = df.at[fila-1, "SAP_ID"]  # Ajustar el índice para acceder a la fila correcta en el DataFrame
        orden_correcto = f"{fila} - {sap_id}"  # Concatenar el número de fila con el SAP_ID
        tree.insert("", "end", text=orden_correcto, values=(archivo))

    tree.pack(expand=True, fill="both")
    # Establecer el tamaño de la ventana
    ventana.geometry("400x300")
    
    # Botón para cerrar la ventana
    cerrar_button = tk.Button(ventana, text="Cerrar", bg="red", fg="white", font=("Arial", 11), command=ventana.destroy)
    cerrar_button.pack(pady=5)
 
    # Animación de fade-in
    ventana.attributes("-alpha", 0)
    ventana.update_idletasks()
    ventana.deiconify()
    for i in range(1, 101):
        ventana.attributes("-alpha", i / 100)
        ventana.update_idletasks()
        time.sleep(0.01)


def limpiar_campos():
    archivo_excel_entry.delete(0, tk.END)
    directorio_renombrados_entry.delete(0, tk.END)

def cerrar_aplicacion():
    ventana.destroy()  # Cierra la ventana principal
    ventana.quit()     # Finaliza el bucle de eventos de tkinter
    sys.exit()         # Termina la ejecución del script

def seleccionar_archivo():
    archivo_excel_path = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
    archivo_excel_entry.delete(0, tk.END)
    archivo_excel_entry.insert(0, archivo_excel_path)

def seleccionar_directorio():
    directorio_renombrados_path = filedialog.askdirectory()
    directorio_renombrados_entry.delete(0, tk.END)
    directorio_renombrados_entry.insert(0, directorio_renombrados_path)

# Crear ventana
ventana = tk.Tk()
ventana.title("VERIFICACION DE RENOMBRADOS")


# Dimensiones de la ventana
ancho_ventana = 830
alto_ventana = 300

# Obtener dimensiones de la pantalla
ancho_pantalla = ventana.winfo_screenwidth()
alto_pantalla = ventana.winfo_screenheight()

# Calcular posición para centrar la ventana
x = (ancho_pantalla // 2) - (ancho_ventana // 2)
y = (alto_pantalla // 2) - (alto_ventana // 2)

# Definir la geometría de la ventana
ventana.geometry(f"{ancho_ventana}x{alto_ventana}+{x}+{y}")
icon_image = Image.open(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\INC.ico")
icon_photo = ImageTk.PhotoImage(icon_image)
ventana.iconphoto(False, icon_photo)



# Cargar imagen de fondo
imagen_fondo_path = r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\comparar.jpg"  # Reemplaza con la ruta de tu imagen
try:
    imagen_fondo = Image.open(imagen_fondo_path)
    imagen_fondo = imagen_fondo.resize((ancho_ventana, alto_ventana))
    imagen_fondo_tk = ImageTk.PhotoImage(imagen_fondo)
    fondo_label = tk.Label(ventana, image=imagen_fondo_tk)
    fondo_label.place(x=0, y=0, relwidth=1, relheight=1)
except Exception as e:
    print("Error al cargar la imagen de fondo:", e)

# Etiquetas
archivo_excel_label = tk.Label(ventana, text="Ruta de Base de Datos:")
archivo_excel_label.grid(row=0, column=0, padx=15, pady=25, sticky="e")

directorio_renombrados_label = tk.Label(ventana, text="Directorio renombrados:")
directorio_renombrados_label.grid(row=1, column=0, padx=10, pady=25, sticky="e")

# Entradas
archivo_excel_entry = tk.Entry(ventana, width=70)
archivo_excel_entry.grid(row=0, column=1, padx=10, pady=25)

directorio_renombrados_entry = tk.Entry(ventana, width=70)
directorio_renombrados_entry.grid(row=1, column=1, padx=10, pady=25)

# Botones
archivo_excel_button = tk.Button(ventana, text="Seleccionar Base de Datos", command=seleccionar_archivo,width=20)
archivo_excel_button.grid(row=0, column=2, padx=50, pady=5)

directorio_renombrados_button = tk.Button(ventana, text="Seleccionar directorio", command=seleccionar_directorio,width=20)
directorio_renombrados_button.grid(row=1, column=2, padx=50, pady=5)

comparar_button = tk.Button(text="Comparar Archivos", bg="green", fg="white", font=("Arial", 11), command=lambda: comparar_archivos(archivo_excel_entry.get(), directorio_renombrados_entry.get()))
comparar_button.place(relx=0.08, rely=0.55, relwidth=0.85)

limpiar_button = tk.Button(text="Limpiar Campos", command=limpiar_campos, bg="orange", font=("Arial", 11))
limpiar_button.place(relx=0.08, rely=0.7, relwidth=0.85)

cerrar_button = tk.Button(text="Cerrar", command=cerrar_aplicacion, bg="red", font=("Arial", 11))
cerrar_button.place(relx=0.08, rely=0.85, relwidth=0.85)

# Animación de fade-in
ventana.attributes("-alpha", 0)
ventana.update_idletasks()
for i in range(1, 101):
    ventana.attributes("-alpha", i / 100)
    ventana.update_idletasks()
    time.sleep(0.01)

# Ejecutar ventana
ventana.mainloop()
