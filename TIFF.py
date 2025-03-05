import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pdf2image import convert_from_path, pdfinfo_from_path
from PIL import Image, ImageTk
import threading

class PDFtoTIFFConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Conversor de PDF a TIFF")
        self.root.geometry("600x400")  # Tamaño ajustado para mejor diseño
        
        # Tema moderno con ttk
        self.style = ttk.Style()
        self.style.theme_use("clam")  # Tema moderno (puedes cambiar a 'alt', 'default', etc.)

        # Fondo dinámico con fallback a color sólido
        self.root.configure(bg="#f0f0f0")  # Gris claro por defecto
        try:
            self.background_image = Image.open(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\TIFF.jpg")
            self.background_photo = ImageTk.PhotoImage(self.background_image)
            self.background_label = tk.Label(root, image=self.background_photo)
            self.background_label.place(relwidth=1, relheight=1)
        except:
            pass  # Si falla, usa el color sólido

        # Ícono de la aplicación
        try:
            icon_image = Image.open(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\Background\INC.ico")
            icon_photo = ImageTk.PhotoImage(icon_image)
            self.root.iconphoto(False, icon_photo)
        except:
            pass

        # Centrar la ventana
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width, window_height = 600, 400
        position_x = int((screen_width - window_width) / 2)
        position_y = int((screen_height - window_height) / 2)
        self.root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

        # Variables
        self.pdf_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.dpi_var = tk.IntVar(value=1200)  # DPI por defecto
        self.cancel_flag = False
        self.poppler_path = r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\poppler-24.08.0\Library\bin"
        
        # Verificar Poppler
        if not os.path.exists(self.poppler_path):
            messagebox.showerror("Error", "Poppler no encontrado. Configura una ruta válida.")
            self.root.quit()

        # Crear widgets
        self.create_widgets()

    def create_widgets(self):
        # Campo para seleccionar archivo PDF
        ttk.Label(self.root, text="Archivo PDF:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
        ttk.Entry(self.root, textvariable=self.pdf_path, width=40, state="readonly").grid(row=0, column=1, padx=10)
        ttk.Button(self.root, text="Seleccionar PDF", command=self.select_pdf_file).grid(row=0, column=2, padx=10)

        # Campo para seleccionar directorio de salida
        ttk.Label(self.root, text="Directorio de salida:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
        ttk.Entry(self.root, textvariable=self.output_dir, width=40, state="readonly").grid(row=1, column=1, padx=10)
        ttk.Button(self.root, text="Seleccionar", command=self.select_output_directory).grid(row=1, column=2, padx=10)

        # Campo para DPI
        ttk.Label(self.root, text="Resolución (DPI):").grid(row=2, column=0, padx=10, pady=10, sticky="e")
        ttk.Entry(self.root, textvariable=self.dpi_var, width=10).grid(row=2, column=1, padx=10, sticky="w")

        # Botones principales
        self.convert_button = ttk.Button(self.root, text="Convertir a TIFF", command=self.convert_to_tiff)
        self.convert_button.grid(row=3, column=1, pady=20)
        ttk.Button(self.root, text="Salir", command=self.root.quit).grid(row=3, column=2, pady=20)

    def select_pdf_file(self):
        file_path = filedialog.askopenfilename(title="Selecciona el archivo PDF", filetypes=[("PDF files", "*.pdf")])
        self.pdf_path.set(file_path)

    def select_output_directory(self):
        output_dir = filedialog.askdirectory(title="Selecciona el directorio de salida")
        self.output_dir.set(output_dir)

    def convert_to_tiff(self):
        pdf_path = self.pdf_path.get()
        output_dir = self.output_dir.get()
        dpi = self.dpi_var.get()

        if not pdf_path or not output_dir:
            messagebox.showwarning("Advertencia", "Selecciona un archivo PDF y un directorio de salida.")
            return

        # Deshabilitar botón para evitar múltiples clics
        self.convert_button.config(state="disabled")
        self.cancel_flag = False

        # Ejecutar conversión en un hilo separado
        thread = threading.Thread(target=self._convert_to_tiff_thread, args=(pdf_path, output_dir, dpi))
        thread.start()

    def _convert_to_tiff_thread(self, pdf_path, output_dir, dpi):
        # Crear ventana emergente para la barra de progreso
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Progreso de Conversión")
        progress_window.geometry("400x200")
        progress_label = ttk.Label(progress_window, text=f"Convirtiendo: {os.path.basename(pdf_path)}")
        progress_label.pack(pady=10)

        # Barra de progreso
        self.progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=300, mode="determinate")
        self.progress_bar.pack(pady=10)

        # Label para mostrar el porcentaje y páginas
        progress_percentage_label = ttk.Label(progress_window, text="0% (0/0 páginas)")
        progress_percentage_label.pack(pady=5)

        # Botón de cancelar
        cancel_button = ttk.Button(progress_window, text="Cancelar", command=self.cancel_conversion)
        cancel_button.pack(pady=10)

        try:
            # Obtener número total de páginas
            pdf_info = pdfinfo_from_path(pdf_path, poppler_path=self.poppler_path)
            total_pages = pdf_info["Pages"]
            self.progress_bar["maximum"] = total_pages
            progress_label.config(text=f"Convirtiendo: {os.path.basename(pdf_path)} ({total_pages} páginas)")

            # Extraer el nombre base del archivo PDF
            pdf_basename = os.path.splitext(os.path.basename(pdf_path))[0]

            # Convertir página por página
            for page in range(1, total_pages + 1):
                if self.cancel_flag:
                    self.root.after(0, lambda: messagebox.showinfo("Cancelado", "Conversión cancelada por el usuario."))
                    break

                images = convert_from_path(
                    pdf_path, 
                    dpi=dpi, 
                    fmt="tiff", 
                    first_page=page, 
                    last_page=page, 
                    poppler_path=self.poppler_path
                )
                img = images[0]
                output_path = os.path.join(output_dir, f"{pdf_basename}_Página_{page:03}.tiff")
                img.save(output_path, "TIFF")

                # Actualizar progreso
                porcentaje_progreso = (page / total_pages) * 100
                self.progress_bar['value'] = page
                progress_percentage_label.config(text=f"{int(porcentaje_progreso)}% ({page}/{total_pages} páginas)")
                progress_window.update_idletasks()

            if not self.cancel_flag:
                self.root.after(0, lambda: messagebox.showinfo("Éxito", f"PDF convertido exitosamente en {output_dir}"))

        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"Hubo un error al convertir el PDF: {e}"))

        finally:
            self.root.after(0, lambda: self.convert_button.config(state="normal"))
            progress_window.destroy()

    def cancel_conversion(self):
        self.cancel_flag = True

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFtoTIFFConverterApp(root)
    root.mainloop()