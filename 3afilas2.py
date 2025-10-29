import os
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

# Obtener la fecha actual
current_date = datetime.now().strftime("%Y%m%d")

# Nombre del archivo de la lista
lista_file = f"lista_{current_date}.txt"

# Comprobar si el archivo existe
if not os.path.exists(lista_file):
    print(f"⚠️ El archivo '{lista_file}' no existe en el directorio actual.")
else:
    # Leer las rutas de los archivos desde el archivo de texto
    with open(lista_file, 'r') as file:
        file_paths = [line.strip() for line in file if line.strip()]

    # Procesar cada archivo Excel
    for excel_file in file_paths:
        try:
            # Leer el archivo Excel SIN interpretar encabezado
            df = pd.read_excel(excel_file, sheet_name='Sheet1', header=None)

            print(f"🔎 Primeras filas de '{excel_file}' antes de eliminar:\n", df.head(10))

            # Asignar la primera fila como encabezado
            df.columns = df.iloc[0]  # Primera fila como encabezado
            df = df.drop(index=0)    # Eliminar la fila del encabezado antiguo

            # Eliminar filas 2, 3 y 4 (sin contar encabezado)
            df = df.drop(index=[1, 2, 3], errors='ignore').reset_index(drop=True)

            print(f"✅ Primeras filas después de la limpieza:\n", df.head(10))

            # Guardar el DataFrame modificado en el archivo Excel
            with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name='Sheet1', index=False)

            # Usar openpyxl para abrir el archivo y escribir "Date" en la celda A1
            workbook = load_workbook(excel_file)
            sheet = workbook['Sheet1']
            sheet['A1'] = 'Date'  # Escribir 'Date' en la celda A1

            # Guardar los cambios en el archivo con openpyxl
            workbook.save(excel_file)

            print(f"✅ Filas 2, 3 y 4 eliminadas correctamente y 'Date' añadido en A1 de '{excel_file}'.")

        except Exception as e:
            print(f"❌ Error al procesar '{excel_file}': {e}")
