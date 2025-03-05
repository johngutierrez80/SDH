import pygame
import random

# Inicializar Pygame
pygame.init()

# Configuración de pantalla
WIDTH, HEIGHT = 400, 600
screen = pygame.display.set_mode((WIDTH, HEIGHT))
pygame.display.set_caption("Flappy Bird")

# Colores
WHITE = (255, 255, 255)
GREEN = (0, 255, 0)
BLACK = (0, 0, 0)

# Cargar imágenes
bird_img = pygame.Surface((30, 30))
bird_img.fill((255, 255, 0))
pipe_img = pygame.Surface((60, HEIGHT))
pipe_img.fill(GREEN)

# Variables del juego
bird_x, bird_y = 50, HEIGHT // 2
bird_velocity = 0
gravity = 0.5
jump_strength = -10

pipe_width = 60
pipe_gap = 200
pipe_velocity = -3
pipes = []
score = 0

# Configuración de reloj y fuente
clock = pygame.time.Clock()
font = pygame.font.Font(None, 36)

# Función para generar nuevos tubos
def create_pipe():
    y = random.randint(150, HEIGHT - 150)
    top_pipe = pygame.Rect(WIDTH, y - HEIGHT - pipe_gap // 2, pipe_width, HEIGHT)
    bottom_pipe = pygame.Rect(WIDTH, y + pipe_gap // 2, pipe_width, HEIGHT)
    pipes.append((top_pipe, bottom_pipe))

# Crear los primeros tubos
create_pipe()

# Función para mostrar pantalla de inicio
def show_start_screen():
    screen.fill(WHITE)
    title_text = font.render("Flappy Bird", True, BLACK)
    start_text = font.render("Haz clic para iniciar", True, BLACK)
    screen.blit(title_text, (WIDTH // 2 - title_text.get_width() // 2, HEIGHT // 2 - 50))
    screen.blit(start_text, (WIDTH // 2 - start_text.get_width() // 2, HEIGHT // 2 + 20))
    pygame.display.flip()
    
    # Esperar a que el jugador haga clic para iniciar
    waiting = True
    while waiting:
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                pygame.quit()
                quit()
            if event.type == pygame.MOUSEBUTTONDOWN:
                waiting = False  # Sale del bucle al hacer clic

# Función para el bucle principal del juego
def game_loop():
    global bird_y, bird_velocity, pipes, score
    bird_y = HEIGHT // 2
    bird_velocity = 0
    pipes = []
    create_pipe()
    score = 0

    running = True
    while running:
        screen.fill(WHITE)
        
        # Manejar eventos
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                running = False
            if event.type == pygame.KEYDOWN:
                if event.key == pygame.K_SPACE:
                    bird_velocity = jump_strength

        # Mover pájaro
        bird_velocity += gravity
        bird_y += bird_velocity

        # Mover tubos y crear nuevos
        for top_pipe, bottom_pipe in pipes:
            top_pipe.x += pipe_velocity
            bottom_pipe.x += pipe_velocity
            pygame.draw.rect(screen, GREEN, top_pipe)
            pygame.draw.rect(screen, GREEN, bottom_pipe)
        
        # Remover tubos fuera de la pantalla
        pipes = [(top_pipe, bottom_pipe) for top_pipe, bottom_pipe in pipes if top_pipe.x > -pipe_width]
        
        # Añadir un nuevo tubo si el último está a mitad de camino
        if pipes[-1][0].x < WIDTH // 2:
            create_pipe()
        
        # Dibujar pájaro
        bird_rect = bird_img.get_rect(center=(bird_x, bird_y))
        screen.blit(bird_img, bird_rect)

        # Verificar colisiones con tubos o con el suelo
        if bird_y > HEIGHT or bird_y < 0 or any(bird_rect.colliderect(pipe) for pipe_pair in pipes for pipe in pipe_pair):
            running = False  # Terminar el juego si hay colisión

        # Actualizar y mostrar la puntuación
        score += 0.01
        score_text = font.render(f"Score: {int(score)}", True, BLACK)
        screen.blit(score_text, (10, 10))

        # Actualizar pantalla
        pygame.display.flip()
        clock.tick(60)

    # Volver a mostrar la pantalla de inicio después de terminar el juego
    show_start_screen()

# Mostrar la pantalla de inicio y luego iniciar el juego
show_start_screen()
game_loop()

pygame.quit()
