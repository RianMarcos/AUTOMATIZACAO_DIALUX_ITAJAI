from screeninfo import get_monitors

def calculate_dpi():
    monitors = get_monitors()
    for monitor in monitors:
        width_px = monitor.width
        height_px = monitor.height
        width_mm = monitor.width_mm
        height_mm = monitor.height_mm

        # Converta milímetros para polegadas
        width_in = width_mm / 25.4
        height_in = height_mm / 25.4

        # Calcule DPI
        dpi_x = width_px / width_in
        dpi_y = height_px / height_in

        print(f'Monitor: {monitor.name}')
        print(f'Resolution: {width_px}x{height_px} pixels')
        print(f'Size: {width_mm}x{height_mm} mm')
        print(f'DPI: {dpi_x:.2f} (horizontal), {dpi_y:.2f} (vertical)\n')

# Chamar a função para calcular a DPI
calculate_dpi()
