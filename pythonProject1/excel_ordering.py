import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment

ruta_archivo = "C:/Users/USER/Downloads/Ejemplo ASB.csv"
output_excel_path = "C:/Users/USER/Desktop/Ejemplo_ASB_ordenado.xlsx"

# Leer el archivo CSV y procesarlo como ya lo tienes
csv_rows = []

with open(ruta_archivo, "r") as file:
    for row in file:
        csv_rows.append(row)

# Step 2: Isolate the header and data lines
header_line = csv_rows[2].strip()
data_lines = csv_rows[3:]

headers = [header.strip() for header in header_line.split(",")]

list_of_lists = []
for line in data_lines:
    info = [data.replace('"', '').strip() for data in line.split(",")]
    while len(info) > 24:
        info.pop()
    list_of_lists.append(info)

# Crear un DataFrame
df = pd.DataFrame(list_of_lists, columns=headers)

# Verificar si las columnas a ordenar existe antes de continuar
if "Trade date" in df.columns and "Nature" in df.columns and "Currency" in df.columns:
    # Renombrar las columnas
    df.rename(columns={"Trade date": "Fecha", "Nature": "Operación", "Currency":"Moneda"}, inplace=True)

    # Mover "Fecha" y "Operación" a las primeras posiciones
    cols = ["Fecha", "Operación", "Moneda"] + [col for col in df.columns if col not in ["Fecha", "Operación", "Moneda"]]

    # Reordenar el DataFrame
    df = df[cols]

    # Reemplazar "Income" por "Ingresos" en la columna "Operación"
    df["Operación"] = df["Operación"].replace("Income", "Ingresos")

# Guardar el DataFrame en un archivo Excel
df.to_excel(output_excel_path, index=False)

# Modificar el formato con openpyxl
wb = load_workbook(output_excel_path)
ws = wb.active

# Ajustar ancho de las columnas
for col in ws.columns:
    max_length = 0
    column = col[0].column_letter
    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        except:
            pass
    ws.column_dimensions[column].width = max_length + 2

# Centrar todo el contenido de las celdas
for row in ws.iter_rows():
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")

# Guardar los cambios en el nuevo archivo Excel
wb.save(output_excel_path)

print(f"Datos guardados exitosamente en {output_excel_path}")
