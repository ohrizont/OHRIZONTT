import pandas as pd
import talib as ta
import numpy as np
import os
from datetime import datetime

# -----------------------------------
# Script para calcular indicadores técnicos de Momentum, Tendencia, Volumen, Osciladores, Volatilidad y Ichimoku
# Incluye hoja resumen con ecuaciones y datos utilizados.
# -----------------------------------

# Obtener la fecha actual en formato AAAAMMDD
hoy = datetime.today().strftime('%Y%m%d')

# Archivo de entrada con la lista de rutas
file_list_path = f'lista_{hoy}.txt'
if not os.path.exists(file_list_path):
    raise FileNotFoundError(f"No se encontró el archivo: {file_list_path}")

with open(file_list_path, 'r') as f:
    file_paths = [line.strip() for line in f]

# -----------------------------------
# FUNCIONES PERSONALIZADAS CON DOCSTRINGS
# -----------------------------------

def calculate_dpo(close, period):
    """ Desfase del precio (Detrended Price Oscillator - DPO). """
    return close - close.shift(period)

def calculate_awesome_oscillator(high, low):
    """ Oscilador Awesome (AO). """
    sma5 = ta.SMA((high + low) / 2, timeperiod=5)
    sma34 = ta.SMA((high + low) / 2, timeperiod=34)
    return sma5 - sma34

def calculate_nvi(close, volume):
    """ Índice de Volumen Negativo (NVI). """
    nvi = np.zeros(len(close))
    nvi[0] = 1000
    for i in range(1, len(nvi)):
        if volume[i] < volume[i - 1]:
            nvi[i] = nvi[i - 1] + ((close[i] - close[i - 1]) / close[i - 1]) * nvi[i - 1]
        else:
            nvi[i] = nvi[i - 1]
    return nvi

def calculate_pvt(close, volume):
    """ Tendencia de los Precios Volumen (PVT). """
    return (volume * ((close - close.shift(1)) / close.shift(1))).cumsum()

def calculate_smi(high, low, close, period=14):
    """ Índice de Momento Estocástico (SMI). """
    smoothed_k = ta.EMA(close, timeperiod=period)
    high_max = high.rolling(window=period).max()
    low_min = low.rolling(window=period).min()
    return ((smoothed_k - low_min) / (high_max - low_min) * 100).rolling(window=3).mean()

def calculate_historical_volatility(close, window=30):
    """ Volatilidad Histórica. """
    return close.pct_change().rolling(window=window).std() * np.sqrt(window)

def calculate_mass_index(high, low, period=9):
    """ Índice de Masa. """
    hl_range = high - low
    ema1 = hl_range.ewm(span=period, adjust=False).mean()
    ema2 = ema1.ewm(span=period, adjust=False).mean()
    return (ema1 / ema2).rolling(window=25).sum()

# Función para calcular umbrales de la zona dinámica de Stochastic
def calculate_sdz_thresholds(stochastic_k):
    upper_threshold = stochastic_k.rolling(window=14).mean() + 2 * stochastic_k.rolling(window=14).std()
    lower_threshold = stochastic_k.rolling(window=14).mean() - 2 * stochastic_k.rolling(window=14).std()
    return upper_threshold, lower_threshold

# Cálculo de las divergencias MACD
def calculate_macd_divergence(close, macd, macd_signal):
    """
    Detecta las divergencias de la MACD comparando la dirección del precio con la MACD.
    Devuelve una serie con las señales de divergencia: 'Bullish' (alcista), 'Bearish' (bajista), o 'None'.
    """
    # Divergencia alcista: El precio hace nuevos mínimos, pero la MACD no.
    # Divergencia bajista: El precio hace nuevos máximos, pero la MACD no.
    divergence = []
    for i in range(2, len(close)):
        if close[i] > close[i-1] and macd[i] < macd[i-1]:  # Divergencia bajista
            divergence.append('Bearish')
        elif close[i] < close[i-1] and macd[i] > macd[i-1]:  # Divergencia alcista
            divergence.append('Bullish')
        else:
            divergence.append('None')
    # Añadir 'None' para los primeros dos valores ya que no tenemos suficientes datos para comparar
    divergence = ['None', 'None'] + divergence
    return pd.Series(divergence, index=close.index)

# Funciones para el cálculo de HMA
def WMA(series, period):
    return ta.WMA(series, timeperiod=period)

def HMA(series, period):
    half_length = int(period / 2)
    sqrt_length = int(period ** 0.5)
    wma_half = WMA(series, half_length)
    wma_full = WMA(series, period)
    hma_series = WMA(2 * wma_half - wma_full, sqrt_length)
    return hma_series

# -----------------------------------
# PROCESAMIENTO DE CADA ARCHIVO
# -----------------------------------
for file_path in file_paths:
    data = pd.read_excel(file_path, sheet_name='Sheet1')  # Leer Sheet1 directamente

    # Asegurarse de que los datos sean numéricos
    data['High'] = pd.to_numeric(data['High'], errors='coerce')
    data['Low'] = pd.to_numeric(data['Low'], errors='coerce')
    data['Close'] = pd.to_numeric(data['Close'], errors='coerce')
    data['Volume'] = pd.to_numeric(data['Volume'], errors='coerce')

    # Calcular MACD antes de crear el DataFrame de osciladores
    macd, macd_signal, macd_hist = ta.MACD(data['Close'], fastperiod=12, slowperiod=26, signalperiod=9)

    # Calcular las divergencias de la MACD
    macd_divergence = macd - macd_signal

    # Indicadores de Momentum
    momentum = pd.DataFrame({
        'ADL': ta.AD(data['High'], data['Low'], data['Close'], data['Volume']),
        'Chande_MO': ta.CMO(data['Close'], timeperiod=14),
        'Force_Index': (data['Close'] - data['Close'].shift(1)) * data['Volume'],
        'Relative_Strength': ta.RSI(data['Close'], timeperiod=14),
        'Momentum': ta.MOM(data['Close'], timeperiod=10),
        'ROC': ta.ROC(data['Close'], timeperiod=10),
        'Ultimate_Osc': ta.ULTOSC(data['High'], data['Low'], data['Close'])
    })

    # Indicadores de Tendencia
    trend = pd.DataFrame({
        'ADX': ta.ADX(data['High'], data['Low'], data['Close'], timeperiod=14),
        'Aroon_Up': ta.AROON(data['High'], data['Low'], timeperiod=14)[0],
        'Aroon_Down': ta.AROON(data['High'], data['Low'], timeperiod=14)[1],
        'Bollinger_High': ta.BBANDS(data['Close'])[0],
        'Bollinger_Mid': ta.BBANDS(data['Close'])[1],
        'Bollinger_Low': ta.BBANDS(data['Close'])[2],
        'SMA': ta.SMA(data['Close'], timeperiod=14),
        'SAR': ta.SAR(data['High'], data['Low'], acceleration=0.02, maximum=0.2),
        'EMA_10': ta.EMA(data['Close'], timeperiod=10),
        'EMA_20': ta.EMA(data['Close'], timeperiod=20),
        'EMA_50': ta.EMA(data['Close'], timeperiod=50),
        'EMA_100': ta.EMA(data['Close'], timeperiod=100),
        'EMA_200': ta.EMA(data['Close'], timeperiod=200),
        'DEMA_20': ta.DEMA(data['Close'], timeperiod=20),
        'TEMA_20': ta.TEMA(data['Close'], timeperiod=20),
        'Golden_Cross': (ta.EMA(data['Close'], timeperiod=50) > ta.EMA(data['Close'], timeperiod=200)).astype(int),
        'Death_Cross': (ta.EMA(data['Close'], timeperiod=50) < ta.EMA(data['Close'], timeperiod=200)).astype(int),
        'Keltner_High': ta.EMA(data['Close'], timeperiod=20) + (2 * ta.ATR(data['High'], data['Low'], data['Close'], timeperiod=10)),
        'Keltner_Low': ta.EMA(data['Close'], timeperiod=20) - (2 * ta.ATR(data['High'], data['Low'], data['Close'], timeperiod=10)),
        'KAMA_10': ta.KAMA(data['Close'], timeperiod=10),
        'HMA_20': HMA(data['Close'], 20),
        'VWAP': (data['Volume'] * (data['High'] + data['Low'] + data['Close']) / 3).cumsum() / data['Volume'].cumsum(),
        'WMA_20': ta.WMA(data['Close'], timeperiod=20),
        'FRAMA_20': ta.KAMA(data['Close'], timeperiod=20),
        'Donchian_High': data['High'].rolling(window=20).max(),
        'Donchian_Low': data['Low'].rolling(window=20).min(),
        'SuperTrend': data['Close'] - ta.ATR(data['High'], data['Low'], data['Close'], timeperiod=10)
    })

    # Indicadores de Volumen
    volumen = pd.DataFrame({
        'Accumulation_Distribution': ta.AD(data['High'], data['Low'], data['Close'], data['Volume']),
        'Chaikin_Oscillator': ta.ADOSC(data['High'], data['Low'], data['Close'], data['Volume'], fastperiod=3, slowperiod=10),
        'Ease_of_Movement': (data['High'] - data['Low']) / data['Volume'],
        'VWMA': (data['Close'] * data['Volume']).cumsum() / data['Volume'].cumsum(),
        'Money_Flow_Index': ta.MFI(data['High'], data['Low'], data['Close'], data['Volume'], timeperiod=14),
        'Negative_Volume_Index': calculate_nvi(data['Close'], data['Volume']),
        'On_Balance_Volume': ta.OBV(data['Close'], data['Volume']),
        'Positive_Volume_Index': data['Volume'].where(data['Close'] > data['Close'].shift(1)).cumsum(),
        'Price_Volume_Trend': calculate_pvt(data['Close'], data['Volume']),
        'Volume': data['Volume'],
        'Media_Volumen_20d': data['Volume'].rolling(window=20).mean(),
        'CVI': data['Volume'].cumsum(),
        'PPO': ta.PPO(data['Close'], fastperiod=12, slowperiod=26, matype=0),
        'VFI': np.log(data['Close'] / data['Close'].shift(1)).fillna(0) * data['Volume'],
        'TMF': ((data['Close'] - data['Low']) - (data['High'] - data['Close'])).rolling(window=21).sum() / data['Volume'].rolling(window=21).sum()
    })

    # Cálculo de Bandas de Bollinger y Stochastic una sola vez
    bollinger_high, bollinger_mid, bollinger_low = ta.BBANDS(data['Close'])
    stochastic_k, stochastic_d = ta.STOCH(data['High'], data['Low'], data['Close'])
    sdz_upper, sdz_lower = calculate_sdz_thresholds(stochastic_k)

    # Indicadores de Osciladores
    osciladores = pd.DataFrame({
       'Awesome_Oscillator': calculate_awesome_oscillator(data['High'], data['Low']),
       'CCI': ta.CCI(data['High'], data['Low'], data['Close'], timeperiod=14),
       'DPO': calculate_dpo(data['Close'], 20),
       'RSI': ta.RSI(data['Close'], timeperiod=14),
       'SMI': calculate_smi(data['High'], data['Low'], data['Close']),
       'Stochastic_K': stochastic_k,
       'Stochastic_D': stochastic_d,
       'SDZ_Upper': sdz_upper,
       'SDZ_Lower': sdz_lower,
       'SDZ': 100 * (stochastic_k - stochastic_d) / (stochastic_k.max() - stochastic_d.min()),  # Corregido
       'TRIX': ta.TRIX(data['Close'], timeperiod=14),
       'Williams_%R': ta.WILLR(data['High'], data['Low'], data['Close'], timeperiod=14),
       'Bollinger_%b': (data['Close'] - bollinger_low) / (bollinger_high - bollinger_low) * 100,
       'MACD': macd,
       'MACD_Signal': macd_signal,
       'MACD_Histogram': macd_hist,
       'MACD_Divergence': macd_divergence,
       'Stochastic_RSI': ta.STOCHRSI(data['Close'], timeperiod=14, fastk_period=3, fastd_period=3, fastd_matype=0)[0],
       'Connors_RSI': (ta.RSI(data['Close'], timeperiod=3) + ta.RSI(data['Close'].diff().fillna(0), timeperiod=2) + data['Close'].pct_change(100).rank(pct=True) * 100) / 3,
       'Elder_Ray_Bull': data['High'] - ta.EMA(data['Close'], timeperiod=13),
       'Elder_Ray_Bear': data['Low'] - ta.EMA(data['Close'], timeperiod=13),
       'Vortex_Positive': ta.PLUS_DI(data['High'], data['Low'], data['Close'], timeperiod=14),
       'Vortex_Negative': ta.MINUS_DI(data['High'], data['Low'], data['Close'], timeperiod=14),
       'RVI': ((data['Close'] - data['Low']) - (data['High'] - data['Close'])) / (data['High'] - data['Low']),
       'Schaff_Trend_Cycle': ta.MACD(data['Close'], fastperiod=23, slowperiod=50, signalperiod=10)[2]
    })

    # Indicadores de Volatilidad
    volatilidad = pd.DataFrame({
        'Average_True_Range': ta.ATR(data['High'], data['Low'], data['Close'], timeperiod=14),
        'Desviación_Típica': data['Close'].rolling(window=14).std(),
        'Indice_de_Masa': calculate_mass_index(data['High'], data['Low']),
        'Volatilidad_de_Chaikin': ta.ADOSC(data['High'], data['Low'], data['Close'], data['Volume']),
        'Bollinger_Bandwidth': (bollinger_high - bollinger_low) / bollinger_mid,
        'Volatilidad_Histórica': calculate_historical_volatility(data['Close']),
        'Donchian_Width': data['High'].rolling(window=20).max() - data['Low'].rolling(window=20).min(),
        'Chandelier_Exit': data['High'].rolling(window=22).max() - ta.ATR(data['High'], data['Low'], data['Close'], timeperiod=22) * 3,
        'Ulcer_Index': ((data['Close'] / data['Close'].cummax() - 1) ** 2).rolling(window=14).mean() ** 0.5,
        'TSI': ta.TEMA(data['Close'] - data['Close'].shift(1), timeperiod=25)
    })

    # Indicadores de Ichimoku
    ichimoku = pd.DataFrame({
        'Tenkan_sen': (data['High'].rolling(window=9).max() + data['Low'].rolling(window=9).min()) / 2,
        'Kijun_sen': (data['High'].rolling(window=26).max() + data['Low'].rolling(window=26).min()) / 2,
        'Senkou_Span_A': ((data['High'].rolling(window=9).max() + data['Low'].rolling(window=9).min()) / 2 +
                      (data['High'].rolling(window=26).max() + data['Low'].rolling(window=26).min()) / 2) / 2,
        'Senkou_Span_B': ((data['High'].rolling(window=52).max() + data['Low'].rolling(window=52).min()) / 2).shift(26),
        'Chikou_Span': data['Close'].shift(-26)
    })

    # Indicadores de Ciclos y Patrones
    ciclos_y_patrones = pd.DataFrame({
        'Fisher_Transform': 0.5 * np.log((1 + data['Close'].pct_change().fillna(0)) / (1 - data['Close'].pct_change().fillna(0))),
        'ZigZag': data['Close'].where(abs(data['Close'].pct_change()) > 0.05),
        'Parabolic_SAR_AF': ta.SAR(data['High'], data['Low'], acceleration=0.02, maximum=0.2)
    })

    # Añadir los nuevos indicadores a Sheet1 (al final de las columnas)
    sheet1_with_indicators = pd.concat([data, momentum, trend, volumen, osciladores, volatilidad, ichimoku, ciclos_y_patrones], axis=1)

    # Crear el nuevo directorio si no existe
    nuevo_directorio = os.path.join(os.getcwd(), hoy)
    if not os.path.exists(nuevo_directorio):
        os.makedirs(nuevo_directorio)

    # Guardar el archivo actualizado
    output_file = os.path.join(nuevo_directorio, os.path.basename(file_path).replace('.xlsx', f'_indicadores_{hoy}.xlsx'))
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        sheet1_with_indicators.to_excel(writer, sheet_name='Sheet1', index=False)

print("Procesamiento completado y hojas con los indicadores añadidos a Sheet1.")