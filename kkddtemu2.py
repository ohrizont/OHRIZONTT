from datetime import datetime
import pandas as pd
import openpyxl
import numpy as np

def calculate_ichimoku(df):
    """Calcula los componentes del Ichimoku Kinko Hyo"""
    
    # Función auxiliar para calcular líneas de conversión
    def calc_midpoint(high, low, period):
        high_val = high.rolling(window=period).max()
        low_val = low.rolling(window=period).min()
        return (high_val + low_val) / 2
    
    # Tenkan-sen (Línea de Conversión): (máximo 9 períodos + mínimo 9 períodos) / 2
    df['Tenkan_sen'] = calc_midpoint(df['High'], df['Low'], 9)
    
    # Kijun-sen (Línea Base): (máximo 26 períodos + mínimo 26 períodos) / 2
    df['Kijun_sen'] = calc_midpoint(df['High'], df['Low'], 26)
    
    # Senkou Span A (Primera Línea Adelantada): (Tenkan-sen + Kijun-sen) / 2
    df['Senkou_Span_A'] = ((df['Tenkan_sen'] + df['Kijun_sen']) / 2)
    
    # Senkou Span B (Segunda Línea Adelantada): (máximo 52 períodos + mínimo 52 períodos) / 2
    df['Senkou_Span_B'] = calc_midpoint(df['High'], df['Low'], 52)
    
    return df

def calculate_stochastic(df, period=14):
    """Calcula los componentes del Stochastic Oscillator"""
    low_min = df['Low'].rolling(window=period).min()
    high_max = df['High'].rolling(window=period).max()
    df['Stochastic_%K'] = (df['Close'] - low_min) / (high_max - low_min) * 100
    df['Stochastic_%D'] = df['Stochastic_%K'].rolling(window=3).mean()
    return df

def calculate_temu(df, period=20):
    """Calcula el indicador Temu_20"""
    df['Temu_20'] = df['Close'].rolling(window=period).mean()
    return df

def initialize_dataframe(df):
    """Inicializa el DataFrame con las columnas necesarias y valores por defecto"""
    # Seleccionar solo las columnas necesarias
    kkddb2_df = df[['Date', 'Open', 'High', 'Low', 'Close', 'Volume',
                    'Stochastic_K', 'Stochastic_D', 'ADX', 'SMA', 'Average_True_Range']].copy()
    
    # Calcular el close_26 y los indicadores Ichimoku
    kkddb2_df = calculate_ichimoku(kkddb2_df)
    kkddb2_df['close_26'] = kkddb2_df['Close'].shift(26)
    
    # Calcular Temu_20, Stochastic_%K y Stochastic_%D
    kkddb2_df = calculate_temu(kkddb2_df)
    kkddb2_df = calculate_stochastic(kkddb2_df)
    
    # Calcular señales
    kkddb2_df['Compra'] = ((kkddb2_df['Stochastic_%K'] > kkddb2_df['Stochastic_%D']) & 
                          (kkddb2_df['Temu_20'] > kkddb2_df['Close'])).astype(int)
    kkddb2_df['Venta'] = ((kkddb2_df['Stochastic_%K'] < kkddb2_df['Stochastic_%D']) & 
                         (kkddb2_df['Tenkan_sen'] > kkddb2_df['Kijun_sen']) & 
                         (kkddb2_df['Temu_20'] < kkddb2_df['Close'])).astype(int)
    
    # Calcular niveles de Stop Loss y Take Profit
    kkddb2_df['Stop_Loss'] = 0.85 * kkddb2_df['Senkou_Span_B']
    kkddb2_df['Take_profit'] = 1.6 * kkddb2_df['Senkou_Span_A']
    
    # Inicializar columnas de trading
    initial_columns = {
        'cta': 100.0,
        'bolsa': 0.0,
        'Stop_Loss_Compra': None,
        'Take_profit_Compra': None,
        'Precio_Compra': None,
        'Rentabilidad': 0.0,
        'compra2': 0,
        'ventap': 0.0,
        'Valor': 100.0
    }
    
    for col, value in initial_columns.items():
        kkddb2_df[col] = value
    
    return kkddb2_df

def process_trading_logic(kkddb2_df):
    """Implementa la lógica de trading con gestión mejorada de stop loss y take profit"""
    en_bolsa = False
    p = 0.5  # Porcentaje de la posición que se venderá en take profit
    
    for i in range(1, len(kkddb2_df)):
        # Actualizar cuenta y bolsa
        kkddb2_df.loc[i, 'cta'] = kkddb2_df.loc[i - 1, 'cta']
        kkddb2_df.loc[i, 'bolsa'] = kkddb2_df.loc[i - 1, 'bolsa'] * kkddb2_df.loc[i, 'Close'] / kkddb2_df.loc[i - 1, 'Close']
        
        # Copiar valores de stop loss y take profit del período anterior si estamos en bolsa
        if en_bolsa:
            kkddb2_df.loc[i, 'Stop_Loss_Compra'] = kkddb2_df.loc[i - 1, 'Stop_Loss_Compra']
            kkddb2_df.loc[i, 'Take_profit_Compra'] = kkddb2_df.loc[i - 1, 'Take_profit_Compra']
        
        # Lógica de compra
        if kkddb2_df.loc[i, 'Compra'] == 1:
            if kkddb2_df.loc[i, 'cta'] > 0:
                # Ejecutar compra
                kkddb2_df.loc[i, 'bolsa'] = kkddb2_df.loc[i, 'cta'] + kkddb2_df.loc[i - 1, 'bolsa']
                kkddb2_df.loc[i, 'cta'] = 0
                kkddb2_df.loc[i, 'Precio_Compra'] = kkddb2_df.loc[i, 'Close']
                # Memorizar Stop Loss y Take Profit de entrada
                kkddb2_df.loc[i, 'Stop_Loss_Compra'] = kkddb2_df.loc[i, 'Stop_Loss']
                kkddb2_df.loc[i, 'Take_profit_Compra'] = kkddb2_df.loc[i, 'Take_profit']
                if not en_bolsa:
                    kkddb2_df.loc[i, 'compra2'] = 1
                    en_bolsa = True
            else:
                # Actualizar Stop Loss solo si mejora la posición actual
                if (kkddb2_df.loc[i, 'Stop_Loss_Compra'] is None or 
                    kkddb2_df.loc[i, 'Stop_Loss'] > kkddb2_df.loc[i, 'Stop_Loss_Compra']):
                    kkddb2_df.loc[i, 'Stop_Loss_Compra'] = kkddb2_df.loc[i, 'Stop_Loss']
        
        # Lógica de take profit (venta parcial)
        elif (en_bolsa and 
              pd.notnull(kkddb2_df.loc[i, 'High']) and 
              pd.notnull(kkddb2_df.loc[i, 'Take_profit_Compra']) and 
              kkddb2_df.loc[i, 'High'] >= kkddb2_df.loc[i, 'Take_profit_Compra']):
            # Ejecutar venta parcial
            venta_parcial = p * kkddb2_df.loc[i, 'bolsa']
            kkddb2_df.loc[i, 'cta'] = kkddb2_df.loc[i - 1, 'cta'] + venta_parcial
            kkddb2_df.loc[i, 'bolsa'] = (1 - p) * kkddb2_df.loc[i, 'bolsa']
            kkddb2_df.loc[i, 'ventap'] = p
            
            # Actualizar Stop Loss y Take Profit para la posición restante
            if kkddb2_df.loc[i, 'Stop_Loss'] > kkddb2_df.loc[i, 'Stop_Loss_Compra']:
                kkddb2_df.loc[i, 'Stop_Loss_Compra'] = kkddb2_df.loc[i, 'Stop_Loss']
            kkddb2_df.loc[i, 'Take_profit_Compra'] = kkddb2_df.loc[i, 'Take_profit']
        
        # Lógica de venta total (por señal de venta o stop loss)
        elif (en_bolsa and 
              (kkddb2_df.loc[i, 'Venta'] == 1 or 
               (pd.notnull(kkddb2_df.loc[i, 'Low']) and 
                pd.notnull(kkddb2_df.loc[i, 'Stop_Loss_Compra']) and 
                kkddb2_df.loc[i, 'Low'] <= kkddb2_df.loc[i, 'Stop_Loss_Compra']))):
            # Ejecutar venta total
            kkddb2_df.loc[i, 'cta'] += kkddb2_df.loc[i, 'bolsa']
            kkddb2_df.loc[i, 'bolsa'] = 0
            # Reiniciar valores de Stop Loss y Take Profit
            kkddb2_df.loc[i, 'Stop_Loss_Compra'] = None
            kkddb2_df.loc[i, 'Take_profit_Compra'] = None
            kkddb2_df.loc[i, 'Precio_Compra'] = None
            en_bolsa = False
        
        # Calcular rentabilidad si estamos en posición
        if en_bolsa and pd.notnull(kkddb2_df.loc[i, 'Precio_Compra']):
            kkddb2_df.loc[i, 'Rentabilidad'] = ((kkddb2_df.loc[i, 'Close'] / 
                                                kkddb2_df.loc[i, 'Precio_Compra']) - 1) * 100
        else:
            kkddb2_df.loc[i, 'Rentabilidad'] = 0
        
        # Actualizar valor total
        kkddb2_df.loc[i, 'Valor'] = kkddb2_df.loc[i, 'cta'] + kkddb2_df.loc[i, 'bolsa']
    
    return kkddb2_df

def main():
    # Obtener la fecha actual
    hoy = datetime.now()
    fecha_str = hoy.strftime("%Y%m%d")
    nombre_archivo_lista = f"lista_indicadores_{fecha_str}.txt"
    
    # Leer las rutas de los archivos Excel
    try:
        with open(nombre_archivo_lista, 'r') as file:
            rutas_archivos = file.read().splitlines()
    except Exception as e:
        print(f"Error al leer el archivo de texto: {e}")
        return
    
    # Procesar cada archivo Excel
    for archivo in rutas_archivos:
        try:
            # Cargar datos
            df = pd.read_excel(archivo, sheet_name='Sheet1')
            
            # Verificar columnas necesarias
            columnas_necesarias = ['Date', 'Open', 'High', 'Low', 'Close', 'Volume',
                                 'Stochastic_K', 'Stochastic_D', 'ADX', 'SMA', 'Average_True_Range']
            
            faltan_columnas = [col for col in columnas_necesarias if col not in df.columns]
            if faltan_columnas:
                print(f"Columnas faltantes en {archivo}: {', '.join(faltan_columnas)}")
                continue
            
            # Preparar datos
            df.fillna(0, inplace=True)
            
            # Inicializar y procesar el DataFrame
            kkddb2_df = initialize_dataframe(df)
            kkddb2_df = process_trading_logic(kkddb2_df)
            
            # Redondear valores
            columns_to_round = ['cta', 'bolsa', 'Stop_Loss', 'Take_profit', 'Valor', 
                              'compra2', 'ventap', 'Rentabilidad',
                              'Tenkan_sen', 'Kijun_sen', 'Senkou_Span_A', 'Senkou_Span_B']
            kkddb2_df[columns_to_round] = kkddb2_df[columns_to_round].round(1)
            
            # Guardar resultados
            libro = openpyxl.load_workbook(archivo)
            hojas_existentes = libro.sheetnames
            hoja_base = 'kkddb2'
            contador = 1
            
            while f"{hoja_base}_{contador}" in hojas_existentes:
                contador += 1
            nombre_nueva_hoja = f"{hoja_base}_{contador}" if hoja_base in hojas_existentes else hoja_base
            
            with pd.ExcelWriter(archivo, engine='openpyxl', mode='a') as writer:
                kkddb2_df.to_excel(writer, sheet_name=nombre_nueva_hoja, index=False)
            
            print(f"Resultados guardados en '{nombre_nueva_hoja}' del archivo {archivo}.")
            
        except Exception as e:
            print(f"Error al procesar {archivo}: {e}")

if __name__ == "__main__":
    main()