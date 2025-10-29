import pandas as pd
from datetime import datetime
from openpyxl import load_workbook

def create_summary_sheet(file_path):
    """
    Lee el archivo Excel, extrae las últimas filas de cada columna para cada variable,
    incluyendo el primer valor de las columnas con "close" en el nombre,
    y las organiza en una nueva hoja llamada 'Rkkddb2' en el mismo archivo.
    """
    # Leer el archivo Excel
    xl = pd.ExcelFile(file_path)
    
    # Verificar si la hoja "kkddb2" existe
    if 'kkddb2' not in xl.sheet_names:
        print(f"El archivo {file_path} no contiene una hoja 'kkddb2'.")
        return
    
    # Leer la hoja "kkddb2"
    df = xl.parse('kkddb2')
    
    # Limpiar los nombres de las columnas (eliminar espacios en blanco)
    df.columns = df.columns.str.strip()

    # Identificar todas las variables (por ejemplo, ACS_Compra, ACS_Venta, etc.)
    variables = [col.split('_')[0] for col in df.columns if len(col.split('_')) > 1]
    unique_variables = list(set(variables))  # Filtramos las variables únicas
    
    # Crear un DataFrame vacío para almacenar el kkddb2
    summary_df = pd.DataFrame(columns=['Variable', 'compra2', 'Compra', 'Venta', 'ventap', 'Stop_Loss_Compra', 'Take_profit_Compra', 'Close', 'Precio_Compra', 'cta', 'bolsa', 'Valor', 'First_Close'])
    
    # Procesar cada variable
    for variable in unique_variables:
        row = {'Variable': variable}
        
        # Leer los últimos valores de cada columna de datos
        for category in ['compra2', 'Compra', 'Venta', 'ventap', 'Stop_Loss_Compra', 'Take_profit_Compra', 'Close', 'Precio_Compra', 'cta', 'bolsa', 'Valor']:
            # Construir el nombre de la columna, como ACS_Compra
            column_name = f"{variable}_{category}"
            
            # Verificar si la columna existe en el DataFrame
            if column_name in df.columns:
                # Tomar el último valor no nulo de la columna
                last_value = df[column_name].dropna().iloc[-1]
                row[category] = last_value
            else:
                row[category] = None  # Si no existe la columna, poner None

        # Buscar y guardar el primer valor de las columnas que contienen "Close"
        close_column_name = f"{variable}_Close"
        if close_column_name in df.columns:
            first_close_value = df[close_column_name].dropna().iloc[0]
            row['First_Close'] = first_close_value
        else:
            row['First_Close'] = None
        
        # Añadir la fila con la variable y sus valores a summary_df
        summary_df = pd.concat([summary_df, pd.DataFrame([row])], ignore_index=True)
    
    # Si el archivo de salida ya existe, cargarlo y agregar la nueva hoja sin borrar las existentes
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
        libro = writer.book
        existing_sheet_names = libro.sheetnames
        
        # Verificar si ya existe una hoja llamada 'Rkkddb2'
        base_sheet_name = "Rkkddb2"
        sheet_name = base_sheet_name
        counter = 1
        while sheet_name in existing_sheet_names:
            sheet_name = f"{base_sheet_name}_{counter}"
            counter += 1

        # Guardar los datos en una nueva hoja con nombre correlativo si ya existe
        summary_df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"kkddb2 guardado en: {file_path}")

# Ruta del archivo que quieres procesar
today = datetime.now().strftime("%Y%m%d")
file_path = f"agregado_{today}.xlsx"

# Crear la hoja kkddb2 y guardarla en el mismo archivo
create_summary_sheet(file_path)