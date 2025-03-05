import os
import shutil
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from PIL import Image, ImageTk
import time
import threading

class RenombradorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("RENOMBRAR ARCHIVOS PDF SIN ORDEN SAP_ID")
        self.setup_ui()
        self.fade_in()

    def fade_in(self):
        for alpha in range(0, 100):
            self.root.attributes('-alpha', alpha / 100)
            self.root.update_idletasks()
            time.sleep(0.01)

    def fade_out(self):
        for alpha in range(100, 0, -1):
            self.root.attributes('-alpha', alpha / 100)
            self.root.update_idletasks()
            time.sleep(0.01)

    def setup_ui(self):
        # Dimensiones y centrado de la ventana
        ancho_ventana, alto_ventana = 600, 320
        ancho_pantalla = self.root.winfo_screenwidth()
        alto_pantalla = self.root.winfo_screenheight()
        x = (ancho_pantalla // 2) - (ancho_ventana // 2)
        y = (alto_pantalla // 2) - (alto_ventana // 2)
        self.root.geometry(f"{ancho_ventana}x{alto_ventana}+{x}+{y}")

        # Ícono (como en el original)
        icon_path = r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\INC.ico"
        try:
            print(f"Intentando cargar el ícono desde: {icon_path}")
            icon_image = Image.open(icon_path)
            icon_photo = ImageTk.PhotoImage(icon_image)
            self.root.iconphoto(False, icon_photo)
            print("Ícono cargado correctamente.")
        except Exception as e:
            print(f"Error al cargar el ícono: {str(e)}")

        # Configurar la imagen de fondo (como en el original)
        background_image_path = r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\fnd.png"
        try:
            print(f"Intentando cargar la imagen de fondo desde: {background_image_path}")
            self.background_image = tk.PhotoImage(file=background_image_path)
            background_label = tk.Label(self.root, image=self.background_image)
            background_label.place(relx=0.5, rely=0.6, anchor=tk.CENTER)
            print("Imagen de fondo cargada correctamente.")
        except Exception as e:
            print(f"Error al cargar la imagen de fondo: {str(e)}")

        # Elementos de la interfaz (como en el original)
        label_directorio_origen = tk.Label(self.root, text="Directorio Descargas:")
        label_directorio_origen.grid(row=0, column=0, sticky="e")

        self.entry_directorio_origen = tk.Entry(self.root, width=50)
        self.entry_directorio_origen.grid(row=0, column=1)

        button_seleccionar_directorio_origen = tk.Button(self.root, text="Seleccionar", command=self.seleccionar_directorio_origen, width=15)
        button_seleccionar_directorio_origen.grid(row=0, column=2)

        label_directorio_destino = tk.Label(self.root, text="Directorio Renombrados:")
        label_directorio_destino.grid(row=1, column=0, sticky="e")

        self.entry_directorio_destino = tk.Entry(self.root, width=50)
        self.entry_directorio_destino.grid(row=1, column=1)

        button_seleccionar_directorio_destino = tk.Button(self.root, text="Seleccionar", command=self.seleccionar_directorio_destino, width=15)
        button_seleccionar_directorio_destino.grid(row=1, column=2)

        label_ruta_excel = tk.Label(self.root, text="Ruta Base de Datos:")
        label_ruta_excel.grid(row=2, column=0, sticky="e")

        self.entry_ruta_excel = tk.Entry(self.root, width=50)
        self.entry_ruta_excel.grid(row=2, column=1)

        button_seleccionar_ruta_excel = tk.Button(self.root, text="Seleccionar", command=self.seleccionar_ruta_excel, width=15)
        button_seleccionar_ruta_excel.grid(row=2, column=2)

        button_renombrar = tk.Button(self.root, text="RENOMBRAR ARCHIVOS", command=self.renombrar_archivos, width=20)
        button_renombrar.config(bg="green", fg="white")
        button_renombrar.grid(row=3, column=1)

        button_limpiar = tk.Button(self.root, text="LIMPIAR CAMPOS", command=self.limpiar_campos, bg="orange", width=15)
        button_limpiar.grid(row=3, column=0, columnspan=3, padx=(0, 360), pady=5)

        button_cerrar = tk.Button(self.root, text="CERRAR", command=self.cerrar_ventana, bg="red", width=15)
        button_cerrar.grid(row=3, column=0, columnspan=3, padx=(443, 0), pady=5)

    def es_formato_radicado(self, nombre_archivo):
        patron = r"\d{4}EE\d{6}[A-Z]\d+"
        return re.match(patron, nombre_archivo) is not None

    def renombrar_archivos(self):
        try:
            if not all([self.entry_directorio_origen.get(), self.entry_directorio_destino.get(), self.entry_ruta_excel.get()]):
                raise ValueError("Por favor seleccione todas las rutas antes de renombrar archivos.")
            print("Iniciando proceso de renombrado...")
            thread = threading.Thread(target=self.actualizar_progreso)
            thread.start()
        except ValueError as e:
            messagebox.showerror("Error", str(e))

    def actualizar_progreso(self):
        ventana_progreso = tk.Toplevel(self.root)
        ventana_progreso.title("Progreso")
        ventana_progreso.geometry("350x150")
        ventana_progreso.transient(self.root)
        ventana_progreso.grab_set()

        progress_bar = ttk.Progressbar(ventana_progreso, mode='determinate', length=300)
        progress_bar.pack(pady=20)
        label_estado = ttk.Label(ventana_progreso, text="PROCESANDO ARCHIVOS...")
        label_estado.pack(pady=5)
        ttk.Button(ventana_progreso, text="Cancelar", command=ventana_progreso.destroy).pack(pady=10)

        try:
            directorio_origen = self.entry_directorio_origen.get()
            directorio_destino = self.entry_directorio_destino.get()
            excel_file = self.entry_ruta_excel.get()

            if not os.path.exists(directorio_origen) or not os.path.exists(directorio_destino):
                raise FileNotFoundError("Uno de los directorios no existe.")
            if not os.path.isfile(excel_file):
                raise FileNotFoundError("El archivo Excel no existe.")

            print(f"Directorio origen: {directorio_origen}")
            print(f"Directorio destino: {directorio_destino}")
            print(f"Archivo Excel: {excel_file}")

            df = pd.read_excel(excel_file)
            if 'SAP_ID' not in df.columns:
                raise ValueError("El archivo Excel no contiene la columna 'SAP_ID'.")
            archivos_ordenados_excel = df['SAP_ID'].astype(str).tolist()
            print(f"Archivos en Excel: {archivos_ordenados_excel}")

            archivos_diccionario = {archivo: os.path.join(directorio_origen, archivo)
                                    for archivo in os.listdir(directorio_origen) if archivo.endswith(".pdf")}
            print(f"Archivos en directorio origen: {list(archivos_diccionario.keys())}")

            archivos_procesados = 0
            for i, nombre_archivo_excel in enumerate(archivos_ordenados_excel, 1):
                archivo_origen = None
                for nombre_archivo, ruta in archivos_diccionario.items():
                    if nombre_archivo.startswith(nombre_archivo_excel):
                        archivo_origen = ruta
                        break

                if archivo_origen:
                    nuevo_nombre = f"{i:05d}__{nombre_archivo_excel}.pdf"
                    archivo_destino = os.path.join(directorio_destino, nuevo_nombre)
                    shutil.copy(archivo_origen, archivo_destino)
                    archivos_procesados += 1
                    print(f"Procesado: {nuevo_nombre}")

                    progress_value = i / len(archivos_ordenados_excel) * 100
                    progress_bar['value'] = progress_value
                    label_estado.config(text=f"PROCESANDO ARCHIVOS... {int(progress_value)}%")
                    ventana_progreso.update()
                else:
                    print(f"Advertencia: No se encontró archivo para {nombre_archivo_excel}")

            label_estado.config(text=f"Se han renombrado y movido {archivos_procesados} archivos PDF.")
            ventana_progreso.after(2000, ventana_progreso.destroy)
        except FileNotFoundError as e:
            messagebox.showerror("Error", str(e))
        except ValueError as e:
            messagebox.showerror("Error", str(e))
        except Exception as e:
            messagebox.showerror("Error", f"Error inesperado: {str(e)}")

    def seleccionar_directorio_origen(self):
        directorio = filedialog.askdirectory()
        self.entry_directorio_origen.delete(0, tk.END)
        self.entry_directorio_origen.insert(0, directorio)

    def seleccionar_directorio_destino(self):
        directorio = filedialog.askdirectory()
        self.entry_directorio_destino.delete(0, tk.END)
        self.entry_directorio_destino.insert(0, directorio)

    def seleccionar_ruta_excel(self):
        ruta = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
        self.entry_ruta_excel.delete(0, tk.END)
        self.entry_ruta_excel.insert(0, ruta)

    def limpiar_campos(self):
        self.entry_directorio_origen.delete(0, tk.END)
        self.entry_directorio_destino.delete(0, tk.END)
        self.entry_ruta_excel.delete(0, tk.END)

    def cerrar_ventana(self):
        if messagebox.askyesno("Confirmar", "¿Seguro que desea cerrar?"):
            self.fade_out()
            self.root.after(1000, self.root.destroy)

if __name__ == "__main__":
    root = tk.Tk()
    app = RenombradorApp(root)
    root.mainloop()