import pygame
import random

# Inicialización de Pygame
pygame.font.init()

# Dimensiones de la ventana
WIDTH, HEIGHT = 300, 600
WIN = pygame.display.set_mode((WIDTH, HEIGHT))
pygame.display.set_caption("Tetris")

# Tamaño de los bloques y número de bloques por fila y columna
BLOCK_SIZE = 30
GRID_WIDTH = WIDTH // BLOCK_SIZE
GRID_HEIGHT = HEIGHT // BLOCK_SIZE

# Colores
BLACK = (0, 0, 0)
WHITE = (255, 255, 255)
COLORS = [
    (0, 255, 0),  # Verde
    (255, 0, 0),  # Rojo
    (0, 0, 255),  # Azul
    (255, 255, 0),  # Amarillo
    (255, 165, 0),  # Naranja
    (128, 0, 128),  # Morado
    (0, 255, 255),  # Cian
]

# Formas de las piezas de Tetris
SHAPES = [
    [[1, 1, 1, 1]],  # Línea
    [[1, 1], [1, 1]],  # Cuadro
    [[0, 1, 0], [1, 1, 1]],  # T
    [[1, 1, 0], [0, 1, 1]],  # Z
    [[0, 1, 1], [1, 1, 0]],  # S
    [[1, 1, 1], [1, 0, 0]],  # L
    [[1, 1, 1], [0, 0, 1]],  # J
]

class Piece:
    def __init__(self, x, y, shape):
        self.x = x
        self.y = y
        self.shape = shape
        self.color = random.choice(COLORS)
        self.rotation = 0

    def rotate(self):
        self.shape = [list(row) for row in zip(*self.shape[::-1])]

def create_grid(locked_positions={}):
    grid = [[0 for _ in range(GRID_WIDTH)] for _ in range(GRID_HEIGHT)]
    for y in range(GRID_HEIGHT):
        for x in range(GRID_WIDTH):
            if (x, y) in locked_positions:
                grid[y][x] = locked_positions[(x, y)]
    return grid

def draw_window(win, grid, score=0, difficulty="Easy"):
    win.fill(BLACK)

    # Dibujar la cuadrícula
    for i in range(len(grid)):
        for j in range(len(grid[i])):
            if grid[i][j] > 0:
                pygame.draw.rect(win, COLORS[grid[i][j] - 1], (j * BLOCK_SIZE, i * BLOCK_SIZE, BLOCK_SIZE, BLOCK_SIZE))

    # Mostrar el puntaje en la parte superior izquierda
    font = pygame.font.SysFont('comicsans', 10)
    label = font.render(f'Score: {score}', 1, WHITE)
    win.blit(label, (5, 5))

    # Mostrar el nivel de dificultad
    difficulty_label = font.render(f'Level: {difficulty}', 1, WHITE)
    win.blit(difficulty_label, (10, 40))

    pygame.display.update()

def valid_space(piece, grid):
    accepted_positions = [(x, y) for y in range(GRID_HEIGHT) for x in range(GRID_WIDTH) if grid[y][x] == 0]
    formatted = [(x + piece.x, y + piece.y) for y, row in enumerate(piece.shape) for x in range(len(row)) if row[x] > 0]
    
    for pos in formatted:
        if pos not in accepted_positions:
            return False
    return True

def clear_rows(grid, locked):
    increment = 0
    for i in range(len(grid)-1, -1, -1):
        if all(grid[i]):
            increment += 1
            index = i
            for j in range(len(grid[i])):
                del locked[(j, i)]

    if increment > 0:
        for i in range(index, 0, -1):
            for j in range(len(grid[i])):
                if (j, i) in locked:
                    locked[(j, i + increment)] = locked.pop((j, i))
    return increment

def draw_next_shape(piece, win):
    font = pygame.font.SysFont('comicsans', 10)
    label = font.render('Next Shape', 1, WHITE)

    sx = WIDTH - 105
    sy = 25
    shape = piece.shape

    for i, row in enumerate(shape):
        for j, col in enumerate(row):
            if col:
                pygame.draw.rect(win, piece.color, (sx + j * BLOCK_SIZE, sy + i * BLOCK_SIZE, BLOCK_SIZE, BLOCK_SIZE))

    win.blit(label, (sx, sy - 20))

def draw_text_middle(text, size, color, win, offset_y=0):
    font = pygame.font.SysFont('comicsans', size)
    label = font.render(text, 1, color)

    win.blit(label, (WIDTH / 2 - (label.get_width() / 2), HEIGHT / 2 - (label.get_height() / 2) + offset_y))
    pygame.display.update()

def main(difficulty="Easy"):
    locked_positions = {}
    grid = create_grid(locked_positions)

    change_piece = False
    run = True
    paused = False
    current_piece = Piece(4, 0, random.choice(SHAPES))
    next_piece = Piece(4, 0, random.choice(SHAPES))
    clock = pygame.time.Clock()
    fall_time = 0
    level_time = 0
    score = 0
    fall_speed = {"Easy": 0.7, "Medium": 0.5, "Hard": 0.3}.get(difficulty, 0.5)

    while run:
        grid = create_grid(locked_positions)
        fall_time += clock.get_rawtime()
        level_time += clock.get_rawtime()
        clock.tick()

        if not paused:
            if score >= 1000 and score % 1000 == 0:
                fall_speed = max(0.1, fall_speed - 0.05)

            if fall_time / 1000 >= fall_speed:
                fall_time = 0
                current_piece.y += 1
                if not valid_space(current_piece, grid) and current_piece.y > 0:
                    current_piece.y -= 1
                    change_piece = True

        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                run = False
                pygame.display.quit()

            if event.type == pygame.KEYDOWN:
                if event.key == pygame.K_p:
                    paused = not paused  # Alterna el estado de pausa

                if not paused:
                    if event.key == pygame.K_LEFT:
                        current_piece.x -= 1
                        if not valid_space(current_piece, grid):
                            current_piece.x += 1
                    elif event.key == pygame.K_RIGHT:
                        current_piece.x += 1
                        if not valid_space(current_piece, grid):
                            current_piece.x -= 1
                    elif event.key == pygame.K_DOWN:
                        current_piece.y += 1
                        if not valid_space(current_piece, grid):
                            current_piece.y -= 1
                    elif event.key == pygame.K_UP:
                        current_piece.rotate()
                        if not valid_space(current_piece, grid):
                            current_piece.rotate()

        if paused:
            draw_text_middle("Paused", 20, WHITE, WIN)
            pygame.display.update()
            continue  # Salta el resto del bucle si está en pausa

        shape_pos = [(current_piece.x + j, current_piece.y + i) for i, row in enumerate(current_piece.shape) for j, column in enumerate(row) if column > 0]

        for pos in shape_pos:
            if pos[1] > -1:
                grid[pos[1]][pos[0]] = COLORS.index(current_piece.color) + 1

        if change_piece:
            for pos in shape_pos:
                locked_positions[(pos[0], pos[1])] = COLORS.index(current_piece.color) + 1
            current_piece = next_piece
            next_piece = Piece(4, 0, random.choice(SHAPES))
            change_piece = False
            score += clear_rows(grid, locked_positions) * 10

        draw_window(WIN, grid, score, difficulty)
        draw_next_shape(next_piece, WIN)

        if any(y < 1 for x, y in locked_positions):
            WIN.fill(BLACK)
            draw_text_middle("Game Over!", 30, WHITE, WIN, offset_y=-20)
            draw_text_middle(f"Score: {score}", 20, WHITE, WIN, offset_y=40)
            pygame.display.update()
            pygame.time.delay(3000)
            run = False

def main_menu():
    run = True
    while run:
        WIN.fill(BLACK)
        draw_text_middle('Select Difficulty', 30, WHITE, WIN, offset_y=-60)
        draw_text_middle('Press 1 for Easy', 20, WHITE, WIN, offset_y=10)
        draw_text_middle('Press 2 for Medium', 20, WHITE, WIN, offset_y=40)
        draw_text_middle('Press 3 for Hard', 20, WHITE, WIN, offset_y=70)
        pygame.display.update()
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                run = False
                pygame.quit()
            if event.type == pygame.KEYDOWN:
                if event.key == pygame.K_1:
                    main("Easy")
                elif event.key == pygame.K_2:
                    main("Medium")
                elif event.key == pygame.K_3:
                    main("Hard")

if __name__ == "__main__":
    main_menu()
