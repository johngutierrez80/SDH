import random
import tkinter as tk
from tkinter import messagebox

class AhorcadoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Juego del Ahorcado")
        
        # Cargar palabras desde el archivo
        self.palabras = self.cargar_palabras(r"\\fjcaldas\SDH-Secretaria_Distrital_de_Hacienda\EJECUTABLES_PROCESOS_OK\ENTORNO_GUI\JUEGOS\palabras.txt")
        
        # Selección aleatoria de palabra
        self.palabra = random.choice(self.palabras)
        self.letras_adivinadas = set()
        self.intentos_restantes = 6
        self.letras_usadas = []

        self.entrada_letra = tk.Entry(self.root)
        self.entrada_letra.pack(pady=10)

        self.boton_adivinar = tk.Button(self.root, text="Adivinar letra", command=self.adivinar_letra)
        self.boton_adivinar.pack(pady=5)

        # Configuración del estado de la palabra con fuente más grande
        self.label_estado = tk.Label(self.root, text=self.mostrar_estado(), font=("Helvetica", 24))
        self.label_estado.pack(pady=10)

        self.label_intentos = tk.Label(self.root, text=f"Intentos restantes: {self.intentos_restantes}")
        self.label_intentos.pack(pady=10)

        self.label_letras_usadas = tk.Label(self.root, text="Letras usadas: ")
        self.label_letras_usadas.pack(pady=10)

        self.label_dibujo = tk.Label(self.root, text=self.dibujo_ahorcado())
        self.label_dibujo.pack(pady=10)

    def cargar_palabras(self, archivo):
        """Carga las palabras desde un archivo de texto"""
        try:
            with open(archivo, 'r', encoding='utf-8') as file:
                palabras = [line.strip().lower() for line in file if line.strip()]
            return palabras
        except FileNotFoundError:
            messagebox.showerror("Error", f"No se encontró el archivo '{archivo}'. Asegúrate de que esté en la misma carpeta.")
            self.root.quit()
            return []

    def mostrar_estado(self):
        estado = ''.join(letra if letra in self.letras_adivinadas else ' _ ' for letra in self.palabra)
        return estado

    def adivinar_letra(self):
        letra = self.entrada_letra.get().lower()
        self.entrada_letra.delete(0, tk.END)

        if letra in self.letras_adivinadas or len(letra) != 1:
            messagebox.showinfo("Info", "Letra ya adivinada o entrada inválida.")
            return
        
        self.letras_adivinadas.add(letra)
        self.letras_usadas.append(letra)
        self.label_letras_usadas.config(text="Letras usadas: " + ', '.join(self.letras_usadas))

        if letra not in self.palabra:
            self.intentos_restantes -= 1

        self.label_estado.config(text=self.mostrar_estado())
        self.label_intentos.config(text=f"Intentos restantes: {self.intentos_restantes}")
        self.label_dibujo.config(text=self.dibujo_ahorcado())

        if all(letra in self.letras_adivinadas for letra in self.palabra):
            messagebox.showinfo("Ganaste", f"¡Felicidades! Has adivinado la palabra: {self.palabra}")
            self.root.quit()
        elif self.intentos_restantes == 0:
            messagebox.showinfo("Perdiste", f"Has perdido. La palabra era: {self.palabra}")
            self.root.quit()

    def dibujo_ahorcado(self):
        estados_dibujo = [
            "",
            " O ",
            " O\n | ",
            " O\n/| ",
            " O\n/|\\",
            " O\n/|\\\n/",
            " O\n/|\\\n/ \\"
        ]
        return estados_dibujo[6 - self.intentos_restantes]

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("400x550")
    app = AhorcadoApp(root)
    root.mainloop()
