import pandas as pd
import numpy as np
from math import radians, sin, cos, sqrt, atan2

# -------------------------------------------------
# Distancia Haversine
# -------------------------------------------------
def haversine(lat1, lon1, lat2, lon2):
    R = 6371
    lat1, lon1, lat2, lon2 = map(radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = sin(dlat/2)**2 + cos(lat1)*cos(lat2)*sin(dlon/2)**2
    return 2 * R * atan2(sqrt(a), sqrt(1 - a))

# -------------------------------------------------
# PREPROCESADO BASE (sin eliminar aún)
# -------------------------------------------------
def preparar_features(df):
    df = df.copy()

    # Fechas
    df["trans_date_trans_time"] = pd.to_datetime(df["trans_date_trans_time"])
    df["dob"] = pd.to_datetime(df["dob"])

    # ---- tiempo ----
    df["hour"] = df["trans_date_trans_time"].dt.hour
    df["dayofweek"] = df["trans_date_trans_time"].dt.dayofweek
    df["is_weekend"] = (df["dayofweek"] >= 5).astype(int)

    df["hour_sin"] = np.sin(2 * np.pi * df["hour"] / 24)
    df["hour_cos"] = np.cos(2 * np.pi * df["hour"] / 24)

    df["day_index"] = (df["unix_time"] // 86400).astype(int)

    # ---- edad ----
    df["age"] = (df["trans_date_trans_time"] - df["dob"]).dt.days / 365.25

    # ---- distancia absoluta ----
    df["distance"] = df.apply(
        lambda r: haversine(
            r["lat"], r["long"],
            r["merch_lat"], r["merch_long"]
        ),
        axis=1
    )

    # ---- transformaciones ----
    df["amt"] = np.log1p(df["amt"])
    df["city_pop"] = np.log1p(df["city_pop"])
    df["gender"] = (df["gender"] == "M").astype(int)

    return df

# -------------------------------------------------
# DISTANCE_DELTA (solo con TRAIN)
# -------------------------------------------------
def add_distance_delta(df_train, df_test):
    mean_by_cc = df_train.groupby("cc_num")["distance"].mean()
    global_mean = df_train["distance"].mean()

    def apply(df):
        df = df.copy()
        df["mean_distance_cc"] = df["cc_num"].map(mean_by_cc)
        df["mean_distance_cc"].fillna(global_mean, inplace=True)
        df["distance_delta"] = df["distance"] - df["mean_distance_cc"]
        return df

    return apply(df_train), apply(df_test)

# -------------------------------------------------
# ELIMINACIÓN FINAL DE VARIABLES
# -------------------------------------------------
def eliminar_columnas(df):
    drop_cols = [
        "Unnamed: 0",
        "first", "last",
        "street",
        "trans_num",
        "lat", "long",
        "merch_lat", "merch_long",
        "dob",
        "trans_date_trans_time",
        "unix_time",
        "hour"
    ]
    return df.drop(columns=drop_cols, errors="ignore")

# -------------------------------------------------
# MAIN
# -------------------------------------------------
if __name__ == "__main__":
    df_train = pd.read_csv("fraudtrain.csv")
    df_test = pd.read_csv("fraudtest.csv")

    df_train = preparar_features(df_train)
    df_test = preparar_features(df_test)

    df_train, df_test = add_distance_delta(df_train, df_test)

    df_train = eliminar_columnas(df_train)
    df_test = eliminar_columnas(df_test)

    df_train.to_csv("pre_fraudtrain.csv", index=False)
    df_test.to_csv("pre_fraudtest.csv", index=False)

    print("✔ Preprocesado completado correctamente")
    print("Columnas finales:", sorted(df_train.columns))
