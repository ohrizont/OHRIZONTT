import pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import MiniBatchKMeans
from pathlib import Path

# ----------------------------
# Archivos
# ----------------------------
TRAIN = "pre_fraudtrain.csv"
TEST = "pre_fraudtest.csv"
EXCEL = "fraude.xlsx"
HOJA = "k-means"

# ----------------------------
# Variables numéricas
# ----------------------------
FEATURES = [
    "amt", "gender", "zip",
    "lat", "long", "city_pop",
    "unix_time",
    "merch_lat", "merch_long",
    "dayofweek", "is_weekend",
    "hour_sin", "hour_cos",
    "age", "distance"
]

# ----------------------------
# Leer datos
# ----------------------------
df_train = pd.read_csv(TRAIN)
df_test = pd.read_csv(TEST)

X_train = df_train[FEATURES]
X_test = df_test[FEATURES]

# ----------------------------
# Escalado
# ----------------------------
scaler = StandardScaler()
X_train_scaled = scaler.fit_transform(X_train)
X_test_scaled = scaler.transform(X_test)

# ----------------------------
# K-Means
# ----------------------------
kmeans = MiniBatchKMeans(
    n_clusters=10,
    random_state=42,
    batch_size=20000
)

df_train["cluster"] = kmeans.fit_predict(X_train_scaled)
df_test["cluster"] = kmeans.predict(X_test_scaled)

# ----------------------------
# Métricas por cluster (TEST)
# ----------------------------
fraude_medio = df_test["is_fraud"].mean()

resumen = (
    df_test
    .groupby("cluster")["is_fraud"]
    .agg(
        transacciones="count",
        fraudes="sum",
        ratio_fraude="mean"
    )
    .reset_index()
)

resumen["lift"] = resumen["ratio_fraude"] / fraude_medio
resumen = resumen.sort_values("lift", ascending=False)

# ----------------------------
# Guardar en Excel
# ----------------------------
mode = "a" if Path(EXCEL).exists() else "w"
with pd.ExcelWriter(EXCEL, engine="openpyxl", mode=mode) as writer:
    resumen.to_excel(writer, sheet_name=HOJA, index=False)

print("K-Means + LIFT completado. Resultados en fraude.xlsx (hoja 'k-means').")
