import pyautogui
import time
import os
import calendar
from datetime import datetime
import pyperclip

# Obtener tamaño de la pantalla
screen_width, screen_height = pyautogui.size()
print(f"Resolución de pantalla detectada: {screen_width}x{screen_height}")

def get_last_day_of_previous_month():
    now = datetime.now()
    year = now.year
    month = now.month - 1 if now.month > 1 else 12
    if month == 12:
        year -= 1
    last_day = calendar.monthrange(year, month)[1]
    return f"{last_day:02d}/{month:02d}/{year}"

def wait_for_color(x_ratio, y_ratio, target_color, timeout=30):
    x = int(screen_width * x_ratio)
    y = int(screen_height * y_ratio)
    start_time = time.time()

    while time.time() - start_time < timeout:
        if pyautogui.pixel(x, y) == target_color:
            print(f"Color detectado en posición ({x}, {y}): {target_color}")
            return True
        time.sleep(0.5)

    print("Tiempo de espera agotado. No se detectó el color.")
    return False

def main():
    print("Abriendo Google Chrome...")
    os.system("start chrome")
    time.sleep(3)

    print("Escribiendo URL de ASB")
    pyautogui.typewrite("https://zonasegura.asbnet.com/FOPM/servletcontroller?SUBSESSIONID=0")
    pyautogui.press("enter")

    # Esperar que la página cargue detectando un color en una posición relativa
    if wait_for_color(0.32, 0.20, (84, 86, 90), timeout=30):
        time.sleep(2)
        pyautogui.click(int(screen_width * 0.24), int(screen_height * 0.20))  # Click en Login
        pyautogui.click(int(screen_width * 0.18), int(screen_height * 0.26))  # Click en siguiente opción
        time.sleep(5)

        pyautogui.moveTo(int(screen_width * 0.32), int(screen_height * 0.45))  # Hacer focus en la página
        pyautogui.scroll(-int(screen_height * 0.4))  # Scroll dinámico
        pyautogui.click(int(screen_width * 0.18), int(screen_height * 0.30))  # Click en Reporte
        time.sleep(5)

        if wait_for_color(0.19, 0.92, (214, 113, 91), timeout=30):
            pyautogui.click(int(screen_width * 0.16), int(screen_height * 0.63))  # Click en Market Value
            time.sleep(5)
            pyautogui.click(int(screen_width * 0.56), int(screen_height * 0.35))  # Click en menú
            pyautogui.click(int(screen_width * 0.56), int(screen_height * 0.45))  # Seleccionar opción
            time.sleep(5)

            pyautogui.click(int(screen_width * 0.57), int(screen_height * 0.66))  # Click en fecha
            for _ in range(15):
                pyautogui.hotkey("backspace")

            last_day_previous_month = get_last_day_of_previous_month()
            print(f"Último día del mes anterior: {last_day_previous_month}")
            pyperclip.copy(last_day_previous_month)
            pyautogui.hotkey("ctrl", "v")
            pyautogui.press("enter")

            pyautogui.scroll(-int(screen_height * 0.6))  # Scroll dinámico

            # Captura de pantalla en región relativa
            region = (
                int(screen_width * 0.1),
                int(screen_height * 0.28),
                int(screen_width * 0.9),
                int(screen_height * 0.6)
            )
            pyautogui.screenshot(region=region).save("screenshot.png")
            print("Captura de pantalla guardada")

# Ejecutar script principal
main()
