import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d.art3d import Poly3DCollection
import tkinter as tk
from tkinter import messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class RubiksCube:
    def __init__(self, root):
        self.cube = np.zeros((3, 3, 3), dtype=int)
        self.colors = ['white', 'yellow', 'green', 'blue', 'orange', 'red']  # Colors for faces
        self.root = root
        self.init_cube()

        # Create figure for plotting the cube
        self.fig = plt.figure(figsize=(5, 5))
        self.ax = self.fig.add_subplot(111, projection='3d')
        self.ax.set_box_aspect([1, 1, 1])
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.root)
        self.canvas.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.canvas.draw()

    def init_cube(self):
        # Inicializa las caras del cubo
        for i in range(6):
            face_color = i + 1
            if i == 0:  # White face (top)
                self.cube[:, :, 0] = face_color
            elif i == 1:  # Yellow face (bottom)
                self.cube[:, :, 2] = face_color
            elif i == 2:  # Green face (front)
                self.cube[:, 0, :] = face_color
            elif i == 3:  # Blue face (back)
                self.cube[:, 2, :] = face_color
            elif i == 4:  # Orange face (left)
                self.cube[0, :, :] = face_color
            elif i == 5:  # Red face (right)
                self.cube[2, :, :] = face_color

    def plot_cube(self):
        self.ax.clear()

        r = [-0.5, 0.5]
        vertices = np.array([[x, y, z] for x in r for y in r for z in r])

        faces = [[0, 1, 3, 2], [4, 5, 7, 6], [0, 1, 5, 4], [2, 3, 7, 6], [0, 2, 6, 4], [1, 3, 7, 5]]

        for (x, y, z) in np.ndindex(self.cube.shape):
            face_colors = [self.colors[int(self.cube[x, y, z]) - 1]] * 6
            for i, face in enumerate(faces):
                self.ax.add_collection3d(Poly3DCollection([vertices[face] + [x, y, z]], color=face_colors[i], edgecolor='k'))

        self.ax.set_xlim([-0.5, 2.5])
        self.ax.set_ylim([-0.5, 2.5])
        self.ax.set_zlim([-0.5, 2.5])
        self.ax.set_xticks([])
        self.ax.set_yticks([])
        self.ax.set_zticks([])

        self.canvas.draw()

    def rotate_face(self, axis, layer, clockwise=True):
        """
        Rota una capa del cubo en el eje especificado (x, y, z) y en la dirección.
        """
        print(f"Rotando la capa {layer} en el eje {axis}, dirección: {'cw' if clockwise else 'ccw'}")
        try:
            if axis == 'x':
                self.cube[layer, :, :] = np.rot90(self.cube[layer, :, :], -1 if clockwise else 1)
            elif axis == 'y':
                self.cube[:, layer, :] = np.rot90(self.cube[:, layer, :], -1 if clockwise else 1)
            elif axis == 'z':
                self.cube[:, :, layer] = np.rot90(self.cube[:, :, layer], -1 if clockwise else 1)

            # Redibujar el cubo después de la rotación
            self.plot_cube()

        except Exception as e:
            messagebox.showerror("Error", f"Comando inválido: {e}")

    def execute_move(self, command):
        try:
            print(f"Comando recibido: {command}")  # Para verificar el comando ingresado
            
            axis, layer, direction = command.split()
            layer = int(layer)
            clockwise = direction == 'cw'

            if axis not in ['x', 'y', 'z'] or layer not in [0, 1, 2]:
                raise ValueError("Comando inválido")

            self.rotate_face(axis, layer, clockwise)
            self.plot_cube()

        except Exception as e:
            messagebox.showerror("Error", f"Comando inválido: {e}")

    def create_controls(self):
        control_frame = tk.Frame(self.root)
        control_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        instruction_label = tk.Label(control_frame, text="Introduce eje, capa y dirección (ej: z 0 cw):")
        instruction_label.pack(pady=10)

        command_entry = tk.Entry(control_frame, width=30)
        command_entry.pack(pady=10)

        def execute_command():
            command = command_entry.get()
            self.execute_move(command)

        execute_button = tk.Button(control_frame, text="Ejecutar", command=execute_command)
        execute_button.pack(pady=10)

        exit_button = tk.Button(control_frame, text="Salir", command=self.root.quit)
        exit_button.pack(pady=10)

# Main program
root = tk.Tk()
root.title("Cubo de Rubik")
root.geometry("800x600")  # Tamaño de la ventana

rubiks = RubiksCube(root)
rubiks.plot_cube()  # Mostrar el cubo inicialmente
rubiks.create_controls()  # Crear controles de entrada

root.mainloop()
