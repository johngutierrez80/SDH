import tkinter as tk
from tkinter import messagebox, simpledialog

class MissionariesCannibalsGame:
    def __init__(self, master):
        self.master = master
        self.master.title("Misioneros y Caníbales")

        # Estado inicial
        self.missionaries_left = 3
        self.cannibals_left = 3
        self.missionaries_right = 0
        self.cannibals_right = 0
        self.boat_position = 'left'  # 'left' o 'right'

        # Crear widgets
        self.left_frame = tk.Frame(self.master)
        self.left_frame.pack(side=tk.LEFT, padx=10)

        self.right_frame = tk.Frame(self.master)
        self.right_frame.pack(side=tk.RIGHT, padx=10)

        self.boat_frame = tk.Frame(self.master)
        self.boat_frame.pack(side=tk.TOP)

        self.left_label = tk.Label(self.left_frame, text="Izquierda", font=("Arial", 16))
        self.left_label.pack()

        self.missionaries_left_label = tk.Label(self.left_frame, text=f"Misioneros: {self.missionaries_left}")
        self.missionaries_left_label.pack()

        self.cannibals_left_label = tk.Label(self.left_frame, text=f"Caníbales: {self.cannibals_left}")
        self.cannibals_left_label.pack()

        self.boat_label = tk.Label(self.boat_frame, text="Bote", font=("Arial", 16))
        self.boat_label.pack()

        self.boat_position_label = tk.Label(self.boat_frame, text=self.boat_position)
        self.boat_position_label.pack()

        self.right_label = tk.Label(self.right_frame, text="Derecha", font=("Arial", 16))
        self.right_label.pack()

        self.missionaries_right_label = tk.Label(self.right_frame, text=f"Misioneros: {self.missionaries_right}")
        self.missionaries_right_label.pack()

        self.cannibals_right_label = tk.Label(self.right_frame, text=f"Caníbales: {self.cannibals_right}")
        self.cannibals_right_label.pack()

        self.move_button = tk.Button(self.master, text="Mover Bote", command=self.move_boat)
        self.move_button.pack(pady=20)

        self.update_labels()

    def move_boat(self):
        if self.boat_position == 'left':
            max_missionaries = self.missionaries_left
            max_cannibals = self.cannibals_left
        else:
            max_missionaries = self.missionaries_right
            max_cannibals = self.cannibals_right

        # Solicitar cuántos misioneros y caníbales se moverán
        missionaries_to_move = simpledialog.askinteger("Mover Misioneros", f"¿Cuántos misioneros mover (máx. {max_missionaries})?", minvalue=0, maxvalue=max_missionaries)
        cannibals_to_move = simpledialog.askinteger("Mover Caníbales", f"¿Cuántos caníbales mover (máx. {max_cannibals})?", minvalue=0, maxvalue=max_cannibals)

        if missionaries_to_move is None or cannibals_to_move is None:
            return  # Cancelar si el usuario cierra el diálogo

        # Verificar que no se muevan más de 2 personas en total
        if missionaries_to_move + cannibals_to_move > 2:
            messagebox.showwarning("Error", "¡No se pueden mover más de 2 personas a la vez!")
            return

        # Mover los misioneros y caníbales
        if self.boat_position == 'left':
            self.missionaries_left -= missionaries_to_move
            self.cannibals_left -= cannibals_to_move
            self.missionaries_right += missionaries_to_move
            self.cannibals_right += cannibals_to_move
            self.boat_position = 'right'
        else:
            self.missionaries_right -= missionaries_to_move
            self.cannibals_right -= cannibals_to_move
            self.missionaries_left += missionaries_to_move
            self.cannibals_left += cannibals_to_move
            self.boat_position = 'left'

        if self.check_validity():
            self.update_labels()
            self.check_game_status()
        else:
            messagebox.showwarning("Derrota", "¡Los caníbales han comido a los misioneros!")
            self.master.quit()  # Cierra la aplicación

    def check_validity(self):
        # Comprobar que en ninguna orilla haya más caníbales que misioneros
        left_valid = self.missionaries_left == 0 or self.missionaries_left >= self.cannibals_left
        right_valid = self.missionaries_right == 0 or self.missionaries_right >= self.cannibals_right
        return left_valid and right_valid

    def update_labels(self):
        self.missionaries_left_label.config(text=f"Misioneros: {self.missionaries_left}")
        self.cannibals_left_label.config(text=f"Caníbales: {self.cannibals_left}")
        self.boat_position_label.config(text=self.boat_position)
        self.missionaries_right_label.config(text=f"Misioneros: {self.missionaries_right}")
        self.cannibals_right_label.config(text=f"Caníbales: {self.cannibals_right}")

    def check_game_status(self):
        # Comprobar victoria
        if self.missionaries_right == 3 and self.cannibals_right == 3:
            messagebox.showinfo("Victoria!", "¡Todos han cruzado el río!")
            self.master.quit()  # Cierra la aplicación

# Crear la ventana principal
root = tk.Tk()
game = MissionariesCannibalsGame(root)
root.mainloop()
