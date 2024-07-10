import tkinter as tk  # Importa el módulo tkinter para crear la interfaz gráfica de usuario.
from tkinter import ttk, filedialog, messagebox  # Importa submódulos y clases específicas de tkinter.
from PIL import Image, ImageTk  # Importa clases de la biblioteca Pillow para manejar imágenes.
import pandas as pd  # Importa pandas para trabajar con datos tabulares.
import os  # Importa el módulo os para interactuar con el sistema operativo.
import re  # Importa el módulo re para el manejo de expresiones regulares.

# Función para obtener el nombre del archivo de una cadena de texto.
def obtener_nombre_archivo(nombre_archivo):
    if re.match(r'\d{4}EE\d{6}O?\d?', nombre_archivo):  # Comprueba si el nombre del archivo coincide con un patrón específico.
        return nombre_archivo  # Devuelve el nombre del archivo si coincide con el patrón.
    else:
        return nombre_archivo[:10]  # Devuelve los primeros 10 caracteres del nombre del archivo.

# Función para cargar los archivos descargados en un DataFrame.
def cargar_archivos_descargados(directorio_descargas):
    try:
        archivos_descargados = os.listdir(directorio_descargas)  # Obtiene una lista de archivos en el directorio de descargas.
        nombres_archivos_tratados = [obtener_nombre_archivo(nombre.split('.')[0]) for nombre in archivos_descargados]  # Trata los nombres de los archivos.
        df_archivos_descargados = pd.DataFrame(nombres_archivos_tratados, columns=['Archivo_descargado'])  # Crea un DataFrame con los nombres de los archivos.
        return df_archivos_descargados  # Devuelve el DataFrame con los nombres de los archivos.
    except Exception as e:
        raise ValueError(f"Error al cargar archivos descargados: {str(e)}")  # Muestra un mensaje de error si ocurre un problema al cargar los archivos.

# Función para comparar los archivos descargados con los datos del archivo Excel.
def comparar_archivos(excel_path, directorio_descargas, sheet_name='Hoja1', campo_especifico='SAP_ID'):
    try:
        if not excel_path or not directorio_descargas:  # Comprueba si se han proporcionado las rutas necesarias.
            raise ValueError("Por favor, seleccione la ruta del archivo Excel y el directorio de descargas.")

        df_excel = pd.read_excel(excel_path, sheet_name=sheet_name)  # Lee el archivo Excel y lo carga en un DataFrame.
        if 'Numero de secuencia ' not in df_excel.columns:  # Comprueba si el campo requerido está presente en el DataFrame.
            raise ValueError("El campo 'Numero de secuencia' no está presente en el DataFrame.")

        df_descargas = cargar_archivos_descargados(directorio_descargas)  # Carga los archivos descargados en un DataFrame.
        sap_ids_excel = set(df_excel[campo_especifico].astype(str))  # Obtiene los SAP_ID del archivo Excel.
        nombres_descargas = set(df_descargas['Archivo_descargado'])  # Obtiene los nombres de los archivos descargados.
        sap_ids_no_en_descargas = sap_ids_excel - nombres_descargas  # Encuentra los SAP_ID que no están en los archivos descargados.

        # Crear ventana de resultados
        resultados_window = tk.Toplevel(root)  # Crea una nueva ventana emergente.
        resultados_window.title("Resultados de la comparación")  # Establece el título de la ventana.

        # Definir tamaño y centrar ventana de resultados
        window_width = 400
        window_height = 400
        position_right = int(resultados_window.winfo_screenwidth() / 2 - window_width / 2)
        position_down = int(resultados_window.winfo_screenheight() / 2 - window_height / 2)
        resultados_window.geometry("{}x{}+{}+{}".format(window_width, window_height, position_right, position_down))

        # Si hay SAP_ID que no están en los archivos descargados
        if sap_ids_no_en_descargas:
            resultados_label = tk.Label(resultados_window, text="Archivos que no están presentes en la carpeta de descargas:")
            resultados_label.pack()

            # Crear un Treeview para mostrar los resultados
            tree = ttk.Treeview(resultados_window, columns=("SAP_ID", "Numero de secuencia"))

            # Configurar columnas y encabezados del Treeview
            tree.column("#0", width=0, stretch=tk.NO)
            tree.column("SAP_ID", anchor=tk.W, width=100)
            tree.column("Numero de secuencia", anchor=tk.W, width=100)
            tree.heading("#0", text="", anchor=tk.W)
            tree.heading("SAP_ID", text="SAP_ID", anchor=tk.W)
            tree.heading("Numero de secuencia", text="Numero de secuencia", anchor=tk.W)

            scrollbar = ttk.Scrollbar(resultados_window, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=scrollbar.set)
            scrollbar.pack(side="right", fill="y")
            tree.pack(fill="both", expand=True)

            # Iterar sobre los SAP_ID que no están en los archivos descargados
            for sap_id in sap_ids_no_en_descargas:
                matching_row = df_excel[df_excel[campo_especifico].astype(str).str.strip() == sap_id.strip()]  # Busca la fila correspondiente al SAP_ID en el archivo Excel.
                if not matching_row.empty:  # Si se encontró una fila correspondiente en el archivo Excel
                    numero_secuencia = matching_row.iloc[0]['Numero de secuencia ']  # Obtiene el número de secuencia de la fila correspondiente.
                    tree.insert("", "end", values=(sap_id, numero_secuencia))  # Inserta el SAP_ID y el número de secuencia en el Treeview.
                else:  # Si no se encontró una fila correspondiente en el archivo Excel
                    nombre_archivo = df_descargas[df_descargas['Archivo_descargado'] == sap_id + '.pdf']  # Busca el nombre del archivo en los archivos descargados.
                    if not nombre_archivo.empty:  # Si se encontró el nombre del archivo en los archivos descargados.
                        numero_secuencia = nombre_archivo.iloc[0]['Numero de secuencia ']  # Obtiene el número de secuencia del nombre del archivo.
                        tree.insert("", "end", values=(sap_id, numero_secuencia))  # Inserta el SAP_ID y el número de secuencia en el Treeview.
                    else:  # Si no se encontró el nombre del archivo en los archivos descargados.
                        tree.insert("", "end", values=(sap_id, "No se encontró"))  # Inserta el SAP_ID con un mensaje de "No se encontró" en el Treeview.

            # Función para ordenar los datos al hacer clic en el encabezado de la columna
            def treeview_sort_column(col):
                items = [(tree.set(k, col), k) for k in tree.get_children('')]
                items.sort()
                for index, (val, k) in enumerate(items):
                    tree.move(k, '', index)
                tree.heading(col, command=lambda: treeview_sort_column(col))

            # Llamar a la función de ordenación cuando se haga clic en el encabezado de la columna
            for col in ("SAP_ID", "Numero de secuencia"):
                tree.heading(col, text=col, command=lambda c=col: treeview_sort_column(c))

        else:  # Si todos los SAP_ID están en los archivos descargados.
            tk.Label(resultados_window, text="Todos los archivos están descargados.").pack()

        # Botón para cerrar la ventana de resultados
        cerrar_button = tk.Button(resultados_window, text="Cerrar", command=resultados_window.destroy, bg="red", font=("Arial", 10), width=15)
        cerrar_button.pack(side="bottom", pady=10, anchor="center")


        # Hacer que la ventana secundaria tenga el foco
        resultados_window.grab_set()

        # Mostrar la ventana principal
        root.wait_window(resultados_window)

    except ValueError as ve:  # Manejo de excepciones personalizado para errores específicos.
        messagebox.showerror("Error", str(ve))  # Muestra un mensaje de error utilizando el cuadro de diálogo de messagebox.
    except FileNotFoundError:  # Manejo de excepciones para archivos no encontrados.
        messagebox.showerror("Error", "No se han cargado las rutas completas.")  # Muestra un mensaje de error utilizando el cuadro de diálogo de messagebox.
    except pd.errors.ParserError:  # Manejo de excepciones para errores al analizar el archivo de Excel.
        messagebox.showerror("Error", "Error al analizar el archivo de Excel.")  # Muestra un mensaje de error utilizando el cuadro de diálogo de messagebox.
    except Exception as e:  # Manejo de excepciones para cualquier otro error no especificado.
        messagebox.showerror("Error", f"Error al comparar archivos: {str(e)}")  # Muestra un mensaje de error utilizando el cuadro de diálogo de messagebox.

# Función para limpiar los campos de entrada.
def limpiar_campos():
    excel_entry.delete(0, tk.END)  # Elimina el texto del campo de entrada de Excel.
    descargas_entry.delete(0, tk.END)  # Elimina el texto del campo de entrada de descargas.

# Función para cerrar la ventana principal.
def cerrar_ventana():
    root.destroy()  # Destruye la ventana principal.

# Función para seleccionar el archivo Excel.
def seleccionar_excel():
    try:
        excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])  # Abre un cuadro de diálogo para seleccionar un archivo Excel.
        if excel_path:
            excel_entry.delete(0, tk.END)  # Elimina el texto existente en el campo de entrada de Excel.
            excel_entry.insert(0, excel_path)  # Inserta la ruta del archivo seleccionado en el campo de entrada de Excel.
    except Exception as e:
        messagebox.showerror("Error", f"Error al seleccionar Base de Datos: {str(e)}")  # Muestra un mensaje de error utilizando el cuadro de diálogo de messagebox.

# Función para seleccionar el directorio de descargas.
def seleccionar_descargas():
    try:
        directorio_descargas = filedialog.askdirectory()  # Abre un cuadro de diálogo para seleccionar el directorio de descargas.
        if directorio_descargas:
            descargas_entry.delete(0, tk.END)  # Elimina el texto existente en el campo de entrada de descargas.
            descargas_entry.insert(0, directorio_descargas)  # Inserta la ruta del directorio seleccionado en el campo de entrada de descargas.
    except Exception as e:
        messagebox.showerror("Error", f"Error al seleccionar el directorio de descargas: {str(e)}")  # Muestra un mensaje de error utilizando el cuadro de diálogo de messagebox.

# Crear la ventana principal
root = tk.Tk()  # Crea una instancia de la clase Tk para la ventana principal.
root.title("VERIFICACION DESCARGAS")  # Establece el título de la ventana principal.
root.geometry("700x400")  # Establece las dimensiones de la ventana principal.
icon_image = Image.open(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\INC.ico")
icon_photo = ImageTk.PhotoImage(icon_image)
root.iconphoto(False, icon_photo)

# Cargar imagen de fondo
background_image = Image.open(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\comparar.jpg")  # Abre la imagen de fondo.
background_image = background_image.resize((700, 400))  # Redimensiona la imagen para que se ajuste a la ventana.
background_photo = ImageTk.PhotoImage(background_image)  # Convierte la imagen a un formato compatible con Tkinter.
background_label = tk.Label(root, image=background_photo)  # Crea una etiqueta para la imagen de fondo.
background_label.place(x=0, y=0, relwidth=1, relheight=1)  # Coloca la etiqueta en la ventana principal.

# Crear etiquetas y campos de entrada para la selección del archivo Excel y el directorio de descargas.
excel_label = tk.Label(root, text="Seleccionar Base Datos:")  # Crea una etiqueta para el campo de entrada de Excel.
excel_label.place(relx= 0.10, rely=0.27)  # Coloca la etiqueta en una posición específica en la ventana.
excel_entry = tk.Entry(root, width=50)  # Crea un campo de entrada para la selección del archivo Excel.
excel_entry.place(relx=0.30, rely=0.27)  # Coloca el campo de entrada en una posición específica en la ventana.

excel_button = tk.Button(root, text="Seleccionar", command=seleccionar_excel, width=15)  # Crea un botón para seleccionar el archivo Excel.
excel_button.place(relx=0.78, rely=0.27)  # Coloca el botón en una posición específica en la ventana.

descargas_label = tk.Label(root, text="Seleccionar descargas:")  # Crea una etiqueta para el campo de entrada de descargas.
descargas_label.place(relx=0.10,rely=0.47)  # Coloca la etiqueta en una posición específica en la ventana.
descargas_entry = tk.Entry(root, width=50)  # Crea un campo de entrada para la selección del directorio de descargas.
descargas_entry.place(relx=0.30, rely=0.47)  # Coloca el campo de entrada en una posición específica en la ventana.

descargas_button = tk.Button(root, text="Seleccionar", command=seleccionar_descargas, width=15)  # Crea un botón para seleccionar el directorio de descargas.
descargas_button.place(relx=0.78, rely=0.47)  # Coloca el botón en una posición específica en la ventana.

# Crea botones para comparar archivos, limpiar campos y cerrar la ventana.
comparar_button = tk.Button(root, text="Comparar Archivos", bg="green", fg="white", font=("Arial", 11), command=lambda: comparar_archivos(excel_entry.get(), descargas_entry.get()))
comparar_button.place(relx=0.1, rely=0.6, relwidth=0.85)

limpiar_button = tk.Button(root, text="Limpiar Campos", command=limpiar_campos, bg="orange", font=("Arial", 11))
limpiar_button.place(relx=0.1, rely=0.7, relwidth=0.85)

cerrar_button = tk.Button(root, text="Cerrar", command=cerrar_ventana, bg="red", font=("Arial", 11))
cerrar_button.place(relx=0.1, rely=0.8, relwidth=0.85)  # Coloca el botón en una posición específica en la ventana.

# Ajustar la geometría de la ventana para centrarla
root.update_idletasks()  # Actualiza el estado de la ventana.
window_width = root.winfo_width()  # Obtiene el ancho de la ventana.
window_height = root.winfo_height()  # Obtiene el alto de la ventana.
position_right = int(root.winfo_screenwidth() / 2 - window_width / 2)  # Calcula la posición horizontal para centrar la ventana.
position_down = int(root.winfo_screenheight() / 2 - window_height / 2)  # Calcula la posición vertical para centrar la ventana.
root.geometry("+{}+{}".format(position_right, position_down))  # Establece la posición de la ventana.

root.mainloop()  # Inicia el bucle principal de la interfaz gráfica de usuario.
