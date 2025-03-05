from PIL import Image, ImageTk
import pandas as pd
from tkinter import Tk, filedialog, Label, Button, Entry, StringVar, messagebox, Canvas

def seleccionar_archivo(var, titulo):
    """Abre un cuadro de diálogo para seleccionar un archivo de Excel y lo asigna a la variable."""
    archivo = filedialog.askopenfilename(
        title=titulo,
        filetypes=[("Archivos de Excel", "*.xlsx *.xls")]
    )
    if archivo:
        var.set(archivo)

def seleccionar_carpeta(var):
    """Abre un cuadro de diálogo para seleccionar una carpeta y la asigna a la variable."""
    carpeta = filedialog.askdirectory(title="Seleccione la carpeta para guardar el archivo")
    if carpeta:
        var.set(carpeta)

def realizar_filtrado(archivo_principal, archivo_filtro, carpeta_destino, nombre_archivo):
    """Ejecuta el filtrado basado en los archivos seleccionados y guarda el resultado."""
    try:
        # Validar que todos los campos estén llenos
        if not archivo_principal or not archivo_filtro or not carpeta_destino or not nombre_archivo:
            messagebox.showwarning("Advertencia", "Todos los campos son obligatorios.")
            return

        # Leer los archivos
        df_principal = pd.read_excel(archivo_principal, sheet_name="Hoja1")
        df_filtro = pd.read_excel(archivo_filtro, sheet_name="Hoja1")

        # Especificar las columnas
        columna_principal = "llave"  # Ajustar según los nombres de las columnas
        columna_filtro = "llave"

        # Filtrar los datos
        valores_filtro = df_filtro[columna_filtro].unique()
        df_filtrado = df_principal[df_principal[columna_principal].isin(valores_filtro)]

        # Construir la ruta final
        ruta_salida = f"{carpeta_destino}/{nombre_archivo}.xlsx"
        
        # Guardar el archivo
        df_filtrado.to_excel(ruta_salida, index=False)
        messagebox.showinfo("Éxito", f"Filtrado completado.\nArchivo guardado en:\n{ruta_salida}")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {e}")


# Configuración de la ventana principal
ventana = Tk()
ventana.title("Filtrado de Excel")
ventana.geometry("500x350")

# Cargar la imagen de fondo (usando Pillow)
try:
    ruta_imagen_fondo = r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\informes.jpg"
    imagen_fondo_original = Image.open(ruta_imagen_fondo)
    imagen_fondo_redimensionada = imagen_fondo_original.resize((500, 350), Image.Resampling.LANCZOS)
    imagen_fondo = ImageTk.PhotoImage(imagen_fondo_redimensionada)

    # Crear el canvas y colocar la imagen de fondo
    canvas = Canvas(ventana, width=500, height=350)
    canvas.pack(fill="both", expand=True)
    canvas.create_image(0, 0, image=imagen_fondo, anchor="nw")
except Exception as e:
    messagebox.showerror("Error al cargar la imagen", f"No se pudo cargar la imagen de fondo: {e}")
    canvas = Canvas(ventana, width=500, height=350)
    canvas.pack(fill="both", expand=True)

# Variables para almacenar las rutas y el nombre del archivo
archivo_principal_var = StringVar()
archivo_filtro_var = StringVar()
carpeta_destino_var = StringVar()
nombre_archivo_var = StringVar()

# Etiquetas y campos para la selección de archivos y carpetas
canvas.create_text(80, 40, text="Archivo Principal:", fill="black", font=("Arial", 10))
entry_principal = Entry(ventana, textvariable=archivo_principal_var, width=50)
entry_principal.place(x=150, y=30)
Button(ventana, text="Seleccionar", command=lambda: seleccionar_archivo(archivo_principal_var, "Seleccione el archivo principal")).place(x=400, y=25)

canvas.create_text(80, 90, text="Archivo de Filtro:", fill="black", font=("Arial", 10))
entry_filtro = Entry(ventana, textvariable=archivo_filtro_var, width=50)
entry_filtro.place(x=150, y=80)
Button(ventana, text="Seleccionar", command=lambda: seleccionar_archivo(archivo_filtro_var, "Seleccione el archivo de filtro")).place(x=400, y=75)

canvas.create_text(80, 140, text="Carpeta Destino:", fill="black", font=("Arial", 10))
entry_destino = Entry(ventana, textvariable=carpeta_destino_var, width=50)
entry_destino.place(x=150, y=130)
Button(ventana, text="Seleccionar", command=lambda: seleccionar_carpeta(carpeta_destino_var)).place(x=400, y=125)

canvas.create_text(80, 190, text="Nombre del Archivo:", fill="black", font=("Arial", 10))
entry_nombre = Entry(ventana, textvariable=nombre_archivo_var, width=50)
entry_nombre.place(x=150, y=180)

# Botón para ejecutar el filtrado
Button(ventana, text="Filtrar y Guardar", font=("Arial", 12), command=lambda: realizar_filtrado(
    archivo_principal_var.get(),
    archivo_filtro_var.get(),
    carpeta_destino_var.get(),
    nombre_archivo_var.get()
)).place(x=200, y=250)

# Ejecutar la ventana
ventana.mainloop()
