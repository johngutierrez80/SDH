from PIL import Image, ImageTk
import tkinter as tk
from tkinter import Menu, filedialog
import os

class CustomTitleBar(tk.Frame):
    def __init__(self, master=None, title=""):
        super().__init__(master, height=30)
        self.master = master
        self.pack_propagate(0)  # Evita que el marco se ajuste automáticamente al tamaño del contenido
        self.create_widgets(title)

    def create_widgets(self, title):
        self.label = tk.Label(self, text=title, font=("Arial", 25, "bold"))
        self.label.pack(fill=tk.BOTH, expand=True)
        self.fade_in()
        self.blink()

    def blink(self):
        self.label.config(fg="#3B9C9A")
        self.master.after(500, self.toggle_color)

    def toggle_color(self):
        self.label.config(fg="black")
        self.master.after(500, self.blink)    

    def fade_in(self):
        alpha = self.master.attributes("-alpha")
        alpha = min(alpha + 0.05, 1.0)
        self.master.attributes("-alpha", alpha)
        if alpha < 1.0:
            self.master.after(50, self.fade_in)

    def fade_out(self):
        alpha = self.master.attributes("-alpha")
        alpha = max(alpha - 0.05, 0.0)
        self.master.attributes("-alpha", alpha)
        if alpha > 0.0:
            self.master.after(50, self.fade_out)
        else:
            self.master.destroy()

# Funciones para ejecutar los archivos
def ejecutar_RENOMBRAR_ARCHIVOS_GUI():
    os.system(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\RENOMBRAR_ARCHIVOS_GUI.py")

def ejecutar_ARCHIVOS_FUENTE_GUI():
    os.system(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\ARCHIVOS_FUENTE_GUI.py")

def ejecutar_INSERTAR_LABELS_GUI():
    os.system(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\INSERTAR_LABELS_GUI.py")

def ejecutar_NUMERO_PAGINAS_GUI():
    os.system(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\NUMERO_PAGINAS_GUI.PY")

def ejecutar_NUMERO_PAGINAS_CARTAS_GUI():
    os.system(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\NUMERO_PAGINAS_CARTAS_GUI.PY")

def ejecutar_FILTRA_DB_GUI():
    os.system(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\FILTRA_DB_GUI.py")

def ejecutar_COMP_RENOM_GUI():
    os.system(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\COMP_RENOM_GUI.py")

def ejecutar_COM_DES_GUI():
    os.system(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\COM_DES_GUI.py")

def ejecutar_INSERT_FULL():
    os.system(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\INSERT_FULL.py")

def ejecutar_RENOMBRAR_SAP_ID_GUI():
    os.system(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\RENOMBRAR_SAP_ID_GUI.py")

def ejecutar_RENOMBRAR_ACTAS_GUI():
    os.system(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\RENOMBRAR_ACTAS_GUI.py")

def ejecutar_UNIFICAR_BASES_DE_DATOS_GUI():
    os.system(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\UNIFICAR_BASES_DE_DATOS_GUI.py")

def ejecutar_ADICIONALES_GUI():
    os.system(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\ADICIONALES_GUI.py")

def ejecutar_LS():
    os.system(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\LS.py")

root = tk.Tk()
root.title("Procesos Data Variable")  # Cambia el título de la ventana

# Obtener el tamaño del escritorio
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Establecer el tamaño de la ventana al tamaño del escritorio
root.geometry(f"{screen_width}x{screen_height}")
icon_image = Image.open(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\INC.ico")
icon_photo = ImageTk.PhotoImage(icon_image)
root.iconphoto(False, icon_photo)

root.attributes("-alpha", 0.0)  # Configurar la transparencia inicial a 0 (oculto)

# Crear una barra de título personalizada centrada
title_bar = CustomTitleBar(root, title=" GENESIS ")
title_bar.pack(fill=tk.X)

# Función para abrir un archivo de texto
def abrir_archivo():
    filepath = filedialog.askopenfilename(
        initialdir="/", title="Seleccione un archivo", filetypes=[("Archivos de texto", "*.txt")]
    )
    if filepath:
        with open(filepath, "r") as file:
            contenido = file.read()
            print(contenido)

# Cargar imagen de fondo y redimensionarla para ajustarla a la ventana
background_image = Image.open(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\INC.jpg")
background_image = background_image.resize((1200, 570), Image.LANCZOS) # Ajustar el tamaño de la imagen al tamaño de la ventana
background_photo = ImageTk.PhotoImage(background_image)

# Crear un label para la imagen de fondo y posicionarlo en la parte superior
background_label = tk.Label(root, image=background_photo)
background_label.place(x=0, y=30, relwidth=1, relheight=1)

# Cargar una segunda imagen y redimensionarla si es necesario
second_image = Image.open(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\160.jpg")
second_image = second_image.resize((800, 150), Image.LANCZOS)  # Ajustar el tamaño de la imagen según sea necesario
second_photo = ImageTk.PhotoImage(second_image)

# Crear un label para la segunda imagen y posicionarlo debajo del título y encima del fondo
second_label = tk.Label(root, image=second_photo)
second_label.place(x=screen_width//2 - 400, y=70)  # Ajustar las coordenadas según sea necesario

# Crear menú principal
menu_bar = Menu(root)
root.config(menu=menu_bar)

# Crear submenú para los procesos
apps_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="PROCESOS", menu=apps_menu, font=("Arial", 12))  # Ajustar el tamaño de la fuente del menú de aplicaciones

# Agregar las opciones de las aplicaciones al submenú
apps_menu.add_command(label="RENOMBRAR ARCHIVOS", command=ejecutar_RENOMBRAR_ARCHIVOS_GUI)
apps_menu.add_command(label="GENERAR FUENTE", command=ejecutar_ARCHIVOS_FUENTE_GUI)
apps_menu.add_command(label="INSERTAR LABELS", command=ejecutar_INSERTAR_LABELS_GUI)
apps_menu.add_command(label="INSERTAR FULL", command=ejecutar_INSERT_FULL)


# Crear submenú para los Adicionales
apps_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="ADICIONALES", menu=apps_menu, font=("Arial", 12))  # Ajustar el tamaño de la fuente del menú de Reportes

# Agregar las opciones de las reportes al submenú
apps_menu.add_command(label="RENOMBRAR SIN ORDEN SAP ID", command=ejecutar_RENOMBRAR_SAP_ID_GUI)
apps_menu.add_command(label="RENOMBRAR CARTAS ACTAS", command=ejecutar_RENOMBRAR_ACTAS_GUI)
apps_menu.add_command(label="UNIFICAR BASES DE DATOS", command=ejecutar_UNIFICAR_BASES_DE_DATOS_GUI)
apps_menu.add_command(label="COMPARAR ADICIONALES", command=ejecutar_ADICIONALES_GUI)

# Crear submenú para los Reportes
apps_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="REPORTES", menu=apps_menu, font=("Arial", 12))  # Ajustar el tamaño de la fuente del menú de Reportes

# Agregar las opciones de las reportes al submenú
apps_menu.add_command(label="REPORTE PAGINAS", command=ejecutar_NUMERO_PAGINAS_GUI)
apps_menu.add_command(label="REPORTE PAGINAS CARTAS", command=ejecutar_NUMERO_PAGINAS_CARTAS_GUI)
apps_menu.add_command(label="FILTRAR BD", command=ejecutar_FILTRA_DB_GUI)
apps_menu.add_command(label="LS CARTAS", command=ejecutar_LS)

# Crear submenú para los Validaciones
apps_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="VALIDACIONES", menu=apps_menu, font=("Arial", 12))  # Ajustar el tamaño de la fuente del menú de Validaciones

# Agregar las opciones de las Validaciones al submenú
apps_menu.add_command(label="VERIFICA RENOMBRADOS", command=ejecutar_COMP_RENOM_GUI)
apps_menu.add_command(label="VERIFICA DESCARGAS", command=ejecutar_COM_DES_GUI)

# Crear menú para salir
exit_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="SALIR", menu=exit_menu)

# Agregamos los comandos de salir al menú de salida
exit_menu.add_command(label="SALIR", command=root.quit)

root.mainloop()
