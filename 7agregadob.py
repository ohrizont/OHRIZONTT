import os
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook

def process_files(file_paths):
    """
    Procesa los archivos listados y extrae las columnas de la hoja 'kkddb2b' de cada uno,
    agregándolas horizontalmente en un nuevo archivo Excel, con un prefijo basado en el nombre del archivo.
    """
    # DataFrame vacío para almacenar los datos combinados
    combined_df = pd.DataFrame()

    # Procesar cada archivo
    for file_path in file_paths:
        print(f"Procesando archivo: {file_path}")
        
        try:
            # Leer el archivo Excel
            xl = pd.ExcelFile(file_path)

            # Verificar si la hoja "kkddb2" existe
            if 'kkddb2' not in xl.sheet_names:
                print(f"El archivo {file_path} no contiene una hoja 'kkddb2'. Ignorando.")
                continue

            # Leer la hoja "kkddb2"
            df = xl.parse('kkddb2')

            # Imprimir las columnas y las primeras filas para depurar
            print(f"Columnas en {file_path}: {df.columns.tolist()}")
            print(f"Primeras filas de {file_path}:\n{df.head()}")

            # Limpiar los nombres de las columnas (eliminar espacios en blanco)
            df.columns = df.columns.str.strip()

            # Verificar si las columnas necesarias están presentes, de forma flexible
            required_columns = ["compra2", "Compra", "Venta", "ventap", "Stop_Loss_Compra", "Take_profit_Compra", "Close", "Precio_Compra", "cta", "bolsa", "Valor"
]
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                print(f"El archivo {file_path} falta las siguientes columnas: {', '.join(missing_columns)}. Ignorando.")
                continue

            # Extraer el prefijo del nombre del archivo (antes del primer punto)
            file_name = os.path.basename(file_path)
            prefix = file_name.split('.')[0]

            # Renombrar las columnas agregando el prefijo
            rename_dict = {col: f"{prefix}_{col}" for col in required_columns if col in df.columns}
            df.rename(columns=rename_dict, inplace=True)

            # Agregar las columnas renombradas a las del DataFrame combinado
            combined_df = pd.concat([combined_df, df[list(rename_dict.values())]], axis=1)

        except Exception as e:
            print(f"Error al procesar el archivo {file_path}: {e}")

    # Si se han recopilado datos, escribirlos en un nuevo archivo
    if not combined_df.empty:
        # Nombre del archivo de salida con la fecha de hoy
        today = datetime.now().strftime("%Y%m%d")
        output_file = f"agregado_{today}.xlsx"

        if os.path.exists(output_file):
            # Cargar el archivo existente
            with pd.ExcelWriter(output_file, engine='openpyxl', mode='a') as writer:
                libro = writer.book
                existing_sheet_names = libro.sheetnames
                
                # Verificar si ya existe una hoja llamada 'kkddb2'
                base_sheet_name = "kkddb2"
                sheet_name = base_sheet_name
                counter = 1
                while sheet_name in existing_sheet_names:
                    sheet_name = f"{base_sheet_name}_{counter}"
                    counter += 1

                # Guardar los datos en una nueva hoja con nombre correlativo si ya existe
                combined_df.to_excel(writer, index=False, sheet_name=sheet_name)
        else:
            # Si el archivo no existe, crearlo y guardar los datos
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                combined_df.to_excel(writer, index=False, sheet_name="kkddb2")

        print(f"kkddb2 guardados en: {output_file}")
    else:
        print("No se encontraron datos para guardar.")

def main():
    # Obtener la fecha de hoy en el formato YYYYMMDD
    today = datetime.now().strftime("%Y%m%d")

    # Ruta del fichero de texto con las rutas de los archivos Excel
    file_list_path = f"lista_indicadores_{today}.txt"

    # Verificar si el archivo de lista existe
    if not os.path.exists(file_list_path):
        print(f"El fichero {file_list_path} no existe.")
        return

    # Leer las rutas de los archivos desde el fichero de texto
    with open(file_list_path, "r") as file:
        file_paths = [line.strip() for line in file.readlines() if line.strip()]

    # Verificar que hay archivos en la lista
    if not file_paths:
        print("La lista de archivos está vacía.")
        return

    # Procesar los archivos
    process_files(file_paths)

if __name__ == "__main__":
    main()