import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment

ruta_archivo = "C:/Users/Alessandro/Desktop/Ejemplo ASB.csv"
output_excel_path = "C:/Users/Alessandro/Desktop/Ejemplo_ASB.xlsx"

# Leer el archivo CSV y procesarlo como ya lo tienes
csv_rows = []

with open(ruta_archivo, "r") as file:
    for row in file:
        csv_rows.append(row)

# Step 2: Isolate the header and data lines
header_line = csv_rows[2].strip()
data_lines = csv_rows[3:]

headers = []

for header in header_line.split(","):
    headers.append(header)

list_of_lists = []

for line in data_lines:
    info = []
    for data in line.split(","):
        info.append(data.replace('"', ''))
    while True:
        if len(info) > 24:
            info.pop()
        else:
            break
    list_of_lists.append(info)

# Crear un DataFrame con los datos
df = pd.DataFrame(list_of_lists, columns=headers)

# Guardar el DataFrame en un archivo Excel
df.to_excel(output_excel_path, index=False)

# Cargar el archivo Excel guardado con openpyxl para modificarlo
wb = load_workbook(output_excel_path)
ws = wb.active

for col in ws.columns:
    max_length = 0
    column = col[0].column_letter
    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)  # Se ajusta el tamaño para que no quede demasiado justo
    ws.column_dimensions[column].width = adjusted_width

# Centrar todo el contenido de las celdas

for row in ws.iter_rows():
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")

# Guardar los cambios en el archivo Excel
wb.save(output_excel_path)

print(f"Datos guardados exitosamente en {output_excel_path}")
