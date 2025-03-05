import tkinter as tk
from tkinter import filedialog, ttk
from PIL import Image, ImageTk
import os
import pandas as pd

class GeneradorArchivosFuente:
    def __init__(self, root):
        self.root = root
        self.root.title("ARCHIVOS FUENTE")
        self.setup_ui()
        self.center_window()

    def center_window(self):
        """Centra la ventana en la pantalla."""
        ancho_ventana = 800
        alto_ventana = 420
        ancho_pantalla = self.root.winfo_screenwidth()
        alto_pantalla = self.root.winfo_screenheight()
        x = (ancho_pantalla // 2) - (ancho_ventana // 2)
        y = (alto_pantalla // 2) - (alto_ventana // 2)
        self.root.geometry(f"{ancho_ventana}x{alto_ventana}+{x}+{y}")

    def setup_ui(self):
        """Configura la interfaz gráfica."""
        # Rutas relativas para imágenes
        script_dir = os.path.dirname(__file__)
        icon_path = os.path.join(script_dir, "Background", "INC.ico")
        background_path = os.path.join(script_dir, "Background", "fnd2.png")

        try:
            icon_image = Image.open(icon_path)
            self.root.iconphoto(False, ImageTk.PhotoImage(icon_image))
            background_image = Image.open(background_path)
            self.background_photo = ImageTk.PhotoImage(background_image)  # Guardar referencia
            tk.Label(self.root, image=self.background_photo).place(x=-5, y=30, relwidth=1, relheight=1)
        except FileNotFoundError as e:
            self.status_label.config(text=f"Error cargando imágenes: {e}")

        # Configurar columnas para diseño responsivo
        self.root.grid_columnconfigure(1, weight=1)

        # Campos y botones
        self.create_entry_field(0, "Carpeta de Archivos Renombrados:", self.seleccionar_carpeta_renombrados)
        self.create_entry_field(1, "Carpeta de Archivos Labels (Original):", self.seleccionar_carpeta_labels_original)
        self.create_entry_field(2, "Carpeta de Archivos Labels (Copia):", self.seleccionar_carpeta_labels_copia)
        self.create_entry_field(3, "Ruta de Archivo Fuente (Original):", self.seleccionar_ruta_excel_original, "Salida original")
        self.create_entry_field(4, "Ruta de Archivo Fuente (Copia):", self.seleccionar_ruta_excel_copia, "Salida copia")

        # Botones de acción
        tk.Button(self.root, text="Limpiar Campos", command=self.limpiar_campos_seleccion, bg="orange", width=15).grid(row=5, column=0, columnspan=3, padx=(0, 550), pady=20)
        tk.Button(self.root, text="Crear Excel", command=self.crear_dataframes_y_escribir_excel, bg="green", fg="white", width=15).grid(row=5, column=0, columnspan=3, padx=(570, 0), pady=5)
        tk.Button(self.root, text="CERRAR", command=self.cerrar_aplicativo, bg="red", fg="white", width=15).grid(row=6, column=0, columnspan=3, padx=(570, 0), pady=5)

        # Etiquetas de estado
        self.status_label = tk.Label(self.root, text="", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.grid(row=7, column=0, columnspan=3, padx=(20, 0), pady=15, sticky="we")
        self.status_label_original = tk.Label(self.root, text="", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_label_original.grid(row=8, column=0, columnspan=3, padx=(20, 0), pady=15, sticky="we")
        self.status_label_copia = tk.Label(self.root, text="", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_label_copia.grid(row=9, column=0, columnspan=3, padx=(20, 0), pady=15, sticky="we")

        # Barra de progreso
        self.progress = ttk.Progressbar(self.root, length=200, mode="indeterminate")
        self.progress.grid(row=10, column=0, columnspan=3, pady=10)

    def create_entry_field(self, row, label_text, command, button_text="Seleccionar Carpeta"):
        """Crea un campo de entrada con etiqueta y botón."""
        tk.Label(self.root, text=label_text).grid(row=row, column=0, padx=(20, 0), sticky="w")
        entry = tk.Entry(self.root, width=50)
        entry.grid(row=row, column=1, padx=(20, 0), pady=5, sticky="ew")
        tk.Button(self.root, text=button_text, command=lambda: command(entry), width=30).grid(row=row, column=2, padx=(30, 0), pady=5)
        return entry

    def listar_archivos(self, ruta_carpeta, excluir_thumbs=True):
        """Lista archivos y ubicaciones en una carpeta."""
        try:
            if not os.path.isdir(ruta_carpeta):
                raise ValueError("La ruta no es una carpeta válida.")
            archivos = []
            ubicaciones = []
            for nombre_archivo in os.listdir(ruta_carpeta):
                if excluir_thumbs and nombre_archivo == "Thumbs.db":
                    continue
                ubicacion_archivo = os.path.join(ruta_carpeta, nombre_archivo)
                archivos.append(nombre_archivo)
                ubicaciones.append(ubicacion_archivo)
            if not archivos:
                raise ValueError("La carpeta está vacía.")
            return archivos, ubicaciones
        except Exception as e:
            raise ValueError(f"Error al listar archivos: {e}")

    def seleccionar_carpeta(self, entry, tipo=""):
        """Selecciona una carpeta y actualiza el campo de entrada."""
        try:
            ruta = filedialog.askdirectory()
            if ruta:
                valido, mensaje = self.validar_carpeta(ruta)
                if valido:
                    entry.delete(0, tk.END)
                    entry.insert(tk.END, ruta)
                else:
                    self.status_label.config(text=mensaje)
            else:
                self.status_label.config(text=f"No se seleccionó ninguna carpeta {tipo}.")
        except Exception as e:
            self.status_label.config(text=f"Error al seleccionar carpeta: {e}")

    def seleccionar_carpeta_renombrados(self, entry):
        self.seleccionar_carpeta(entry, "de archivos Renombrados")

    def seleccionar_carpeta_labels_original(self, entry):
        self.seleccionar_carpeta(entry, "de Labels (Original)")

    def seleccionar_carpeta_labels_copia(self, entry):
        self.seleccionar_carpeta(entry, "de Labels (Copia)")

    def seleccionar_ruta_excel(self, entry, default_filename):
        """Selecciona la ruta para guardar un archivo Excel."""
        try:
            ruta = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivos de Excel", "*.xlsx")], initialfile=default_filename)
            if ruta:
                entry.delete(0, tk.END)
                entry.insert(tk.END, ruta)
        except Exception as e:
            self.status_label.config(text=f"Error al seleccionar ruta Excel: {e}")

    def seleccionar_ruta_excel_original(self, entry):
        self.seleccionar_ruta_excel(entry, "ARCHIVO_FUENTE_ORIGINAL.xlsx")

    def seleccionar_ruta_excel_copia(self, entry):
        self.seleccionar_ruta_excel(entry, "ARCHIVO_FUENTE_COPIA.xlsx")

    def validar_carpeta(self, ruta):
        """Valida si una carpeta existe y tiene contenido."""
        if not os.path.isdir(ruta):
            return False, "La ruta no es una carpeta válida."
        if not os.listdir(ruta):
            return False, "La carpeta está vacía."
        return True, ""

    def limpiar_campos_seleccion(self):
        """Limpia todos los campos de entrada."""
        for entry in [self.ruta_carpeta_renombrados_entry, self.ruta_carpeta_labels_original_entry,
                      self.ruta_carpeta_labels_copia_entry, self.ruta_excel_original_entry,
                      self.ruta_excel_copia_entry]:
            entry.delete(0, tk.END)
        self.status_label.config(text="Campos de selección limpiados.")
        self.status_label_original.config(text="")
        self.status_label_copia.config(text="")

    def crear_dataframes_y_escribir_excel(self):
        """Crea DataFrames y escribe archivos Excel."""
        try:
            self.progress.start()
            rutas = {
                "renombrados": self.ruta_carpeta_renombrados_entry.get(),
                "labels_original": self.ruta_carpeta_labels_original_entry.get(),
                "labels_copia": self.ruta_carpeta_labels_copia_entry.get(),
                "excel_original": self.ruta_excel_original_entry.get(),
                "excel_copia": self.ruta_excel_copia_entry.get()
            }
            if not all(rutas.values()):
                raise ValueError("Por favor complete todas las rutas.")

            archivos_renombrados, ubicaciones_renombrados = self.listar_archivos(rutas["renombrados"], False)
            archivos_labels_original, ubicaciones_labels_original = self.listar_archivos(rutas["labels_original"])
            archivos_labels_copia, ubicaciones_labels_copia = self.listar_archivos(rutas["labels_copia"])

            df_fuente_original = pd.DataFrame({
                "Archivo": archivos_renombrados,
                "Ubicación": ubicaciones_renombrados,
                "Label": archivos_labels_original,
                "Ubicación Label": ubicaciones_labels_original
            })
            df_fuente_copia = pd.DataFrame({
                "Archivo": archivos_renombrados,
                "Ubicación": ubicaciones_renombrados,
                "Label": archivos_labels_copia,
                "Ubicación Label": ubicaciones_labels_copia
            })

            with pd.ExcelWriter(rutas["excel_original"], engine="openpyxl", mode="w") as writer:
                df_fuente_original.to_excel(writer, index=False, header=True)
            with pd.ExcelWriter(rutas["excel_copia"], engine="openpyxl", mode="w") as writer:
                df_fuente_copia.to_excel(writer, index=False, header=True)

            self.status_label.config(text="¡Éxito! Archivos generados correctamente.")
            self.status_label_original.config(text=f"Datos guardados en {rutas['excel_original']}")
            self.status_label_copia.config(text=f"Datos guardados en {rutas['excel_copia']}")
        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}")
        finally:
            self.progress.stop()

    def cerrar_aplicativo(self):
        """Cierra la aplicación."""
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = GeneradorArchivosFuente(root)
    # Asignar entradas como atributos para usarlas en otras funciones
    app.ruta_carpeta_renombrados_entry = app.create_entry_field(0, "Carpeta de Archivos Renombrados:", app.seleccionar_carpeta_renombrados)
    app.ruta_carpeta_labels_original_entry = app.create_entry_field(1, "Carpeta de Archivos Labels (Original):", app.seleccionar_carpeta_labels_original)
    app.ruta_carpeta_labels_copia_entry = app.create_entry_field(2, "Carpeta de Archivos Labels (Copia):", app.seleccionar_carpeta_labels_copia)
    app.ruta_excel_original_entry = app.create_entry_field(3, "Ruta de Archivo Fuente (Original):", app.seleccionar_ruta_excel_original, "Salida original")
    app.ruta_excel_copia_entry = app.create_entry_field(4, "Ruta de Archivo Fuente (Copia):", app.seleccionar_ruta_excel_copia, "Salida copia")
    root.mainloop()