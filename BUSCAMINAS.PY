import tkinter as tk
import random
from tkinter import messagebox

class Minesweeper:
    def __init__(self, root, size, num_bombs):
        self.size = size
        self.num_bombs = num_bombs
        self.score = 0
        self.buttons = {}
        self.bombs = set()
        self.root = root
        self.create_widgets()
        
    def create_widgets(self):
        # Crear cuadrícula de botones
        for row in range(self.size):
            for col in range(self.size):
                button = tk.Button(self.root, width=3, height=1,
                                   command=lambda r=row, c=col: self.reveal_cell(r, c))
                button.bind("<Button-3>", lambda e, r=row, c=col: self.flag_cell(r, c))
                button.grid(row=row, column=col)
                self.buttons[(row, col)] = button
        
        # Colocar bombas
        self.place_bombs()

    def place_bombs(self):
        while len(self.bombs) < self.num_bombs:
            row, col = random.randint(0, self.size - 1), random.randint(0, self.size - 1)
            self.bombs.add((row, col))

    def reveal_cell(self, row, col):
        if (row, col) in self.bombs:
            self.buttons[(row, col)].config(text="💣", bg="red")
            self.game_over("lost")
        else:
            count = self.count_adjacent_bombs(row, col)
            self.buttons[(row, col)].config(text=str(count), state="disabled", relief=tk.SUNKEN)
            self.score += 10  # Incrementar el puntaje al revelar una celda sin bomba
            if count == 0:
                self.reveal_adjacent_cells(row, col)
            if self.check_win():
                self.game_over("won")

    def reveal_adjacent_cells(self, row, col):
        for r in range(max(0, row - 1), min(self.size, row + 2)):
            for c in range(max(0, col - 1), min(self.size, col + 2)):
                if (r, c) != (row, col) and self.buttons[(r, c)]['state'] == "normal":
                    self.reveal_cell(r, c)

    def count_adjacent_bombs(self, row, col):
        count = 0
        for r in range(max(0, row - 1), min(self.size, row + 2)):
            for c in range(max(0, col - 1), min(self.size, col + 2)):
                if (r, c) in self.bombs:
                    count += 1
        return count

    def flag_cell(self, row, col):
        button = self.buttons[(row, col)]
        if button["text"] == "":
            button.config(text="🚩", fg="red")
        elif button["text"] == "🚩":
            button.config(text="")

    def game_over(self, result):
        # Revelar todas las bombas y deshabilitar los botones
        for (r, c), button in self.buttons.items():
            if (r, c) in self.bombs:
                button.config(text="💣")
            button.config(state="disabled")
        
        if result == "lost":
            messagebox.showinfo("Game Over", f"¡Perdiste! Puntaje: {self.score}")
        else:
            messagebox.showinfo("Congratulations", f"¡Ganaste! Puntaje: {self.score}")
        
        # Cerrar la ventana de juego y volver al menú
        self.root.destroy()
        show_menu()

    def check_win(self):
        non_bomb_cells = [btn for (r, c), btn in self.buttons.items() if (r, c) not in self.bombs]
        return all(btn['state'] == "disabled" for btn in non_bomb_cells)

def start_game(level):
    root = tk.Tk()
    root.title("Minesweeper")
    
    if level == "Fácil":
        size, num_bombs = 8, 10
    elif level == "Medio":
        size, num_bombs = 12, 20
    else:
        size, num_bombs = 16, 40
    
    game = Minesweeper(root, size, num_bombs)
    root.mainloop()

def show_menu():
    menu = tk.Tk()
    menu.title("Buscaminas")
    menu.geometry("300x200")
    # Etiqueta de bienvenida
    label = tk.Label(menu, text="Seleccione el nivel de dificultad", font=("Arial", 14))
    label.pack(pady=20) 

     # Botones de nivel de dificultad con espacio entre ellos
    tk.Button(menu, text="Fácil", font=("Arial", 12), width=10, command=lambda: [menu.destroy(), start_game("Fácil")]).pack(pady=5)
    tk.Button(menu, text="Medio", font=("Arial", 12), width=10, command=lambda: [menu.destroy(), start_game("Medio")]).pack(pady=5)
    tk.Button(menu, text="Difícil", font=("Arial", 12), width=10, command=lambda: [menu.destroy(), start_game("Difícil")]).pack(pady=5)

    menu.mainloop()

show_menu()
