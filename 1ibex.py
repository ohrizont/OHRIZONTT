import yfinance as yf
import pandas as pd
from datetime import datetime
import os

# Lista de símbolos de las empresas del IBEX 35
symbols_ibex = [
    'ACX.MC', 'ACS.MC', 'AENA.MC', 'ALM.MC', 'AMS.MC', 'MT.AS', 'BBVA.MC', 'SAB.MC', 'SAN.MC',
    'BKT.MC', 'CABK.MC', 'CLNX.MC', 'CIE.MC', 'COL.MC', 'ENG.MC', 'ELE.MC', 'FER.MC', 'FDR.MC',
    'GRF.MC', 'IAG.MC', 'IBE.MC', 'ITX.MC', 'IDR.MC', 'MAP.MC', 'MEL.MC', 'MRL.MC', 'NTGY.MC',
    'REP.MC', 'TEF.MC', 'VIS.MC'
]

# Obtener datos desde enero de 2022
start_date = '2022-01-01'
end_date = pd.Timestamp.now()

# Carpeta para guardar los archivos
current_date = datetime.now().strftime("%Y%m%d")
output_dir = os.path.join(os.getcwd(), current_date)
os.makedirs(output_dir, exist_ok=True)

# Iterar sobre cada símbolo del IBEX 35
for ibex_symbol in symbols_ibex:
    try:
        # Descargar datos del IBEX 35
        ibex_data = yf.download(ibex_symbol, start=start_date, end=end_date)
    except Exception as e:
        print(f"Error downloading {ibex_symbol}: {e}")
        continue

    # Calcular la volatilidad para el símbolo del IBEX 35
    ibex_data['Volatility'] = ibex_data['Close'].rolling(window=20).std()

    # Obtener datos adicionales para cada símbolo
    additional_symbols = ['EURUSD=X', 'CNY=X', '^IRX', '^FVX', '^TNX', '^TYX']
    for symbol in additional_symbols:
        try:
            additional_data = yf.download(symbol, start=start_date, end=end_date)
            ibex_data[symbol] = additional_data['Close']
        except Exception as e:
            print(f"Error downloading {symbol}: {e}")

    # Nombre del archivo de salida
    output_file = os.path.join(output_dir, f"{ibex_symbol}_{current_date}.xlsx")

    # Guardar los datos en un archivo Excel
    ibex_data.to_excel(output_file)

    print(f"Datos de '{ibex_symbol}' descargados y guardados en '{output_file}'")
