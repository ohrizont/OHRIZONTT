import pandas as pd
import numpy as np
import lightgbm as lgb

# -------------------------------
# Configuración
# -------------------------------
TRAIN = "pre_fraudtrain.csv"
TEST  = "pre_fraudtest.csv"

FEATURES = [
    "amt",
    "hour_cos",
    "dayofweek",
    "is_weekend",
    "age",
    "gender"
]
TARGET = "is_fraud"

TH_LOW  = 0.66   # amt <= media
TH_HIGH = 0.51   # amt > media

# -------------------------------
# Cargar datos
# -------------------------------
df_train = pd.read_csv(TRAIN)
df_test  = pd.read_csv(TEST)

X_train = df_train[FEATURES]
y_train = df_train[TARGET]

X_test = df_test[FEATURES]
y_test = df_test[TARGET]

# -------------------------------
# Media de amt (solo train)
# -------------------------------
media_amt = X_train["amt"].mean()
print(f"Media de amt (train): {media_amt:.4f}")

# -------------------------------
# Entrenar LightGBM
# -------------------------------
train_set = lgb.Dataset(X_train, label=y_train)

params = {
    "objective": "binary",
    "metric": "auc",
    "learning_rate": 0.05,
    "num_leaves": 31,
    "min_data_in_leaf": 200,
    "feature_fraction": 0.8,
    "bagging_fraction": 0.8,
    "bagging_freq": 5,
    "verbosity": -1,
    "scale_pos_weight": (len(y_train) - y_train.sum()) / y_train.sum()
}

model = lgb.train(params, train_set, num_boost_round=500)

# -------------------------------
# Predicciones
# -------------------------------
df_test = df_test.copy()
df_test["score"] = model.predict(X_test)

# Umbral dinámico
df_test["threshold"] = np.where(
    df_test["amt"] <= media_amt,
    TH_LOW,
    TH_HIGH
)

df_test["marked"] = (df_test["score"] >= df_test["threshold"]).astype(int)

# -------------------------------
# Función métricas por tramo
# -------------------------------
def metrics_block(df, name):
    total_ops = len(df)
    total_amt = df["amt"].sum()

    fraud_df = df[df[TARGET] == 1]
    marked_df = df[df["marked"] == 1]
    fraud_marked_df = df[(df[TARGET] == 1) & (df["marked"] == 1)]

    recall_ops = len(fraud_marked_df) / max(len(fraud_df), 1)
    precision = len(fraud_marked_df) / max(len(marked_df), 1)

    pct_ops_marked = len(marked_df) / total_ops
    pct_amt_marked = marked_df["amt"].sum() / total_amt if total_amt > 0 else 0

    recall_amt = fraud_marked_df["amt"].sum() / fraud_df["amt"].sum() if fraud_df["amt"].sum() > 0 else 0

    print(f"\n=== {name.upper()} ===")
    print(f"Recall fraude (ops)        : {recall_ops:.4f}")
    print(f"Recall fraude (importe)    : {recall_amt:.4f}")
    print(f"Precisión                 : {precision:.4f}")
    print(f"% operaciones marcadas     : {pct_ops_marked*100:.2f}%")
    print(f"% importe total marcado    : {pct_amt_marked*100:.2f}%")
    print(f"Importe marcado total     : {marked_df['amt'].sum():.2f}")

# -------------------------------
# Métricas globales
# -------------------------------
metrics_block(df_test, "global")

# -------------------------------
# Métricas por tramo
# -------------------------------
metrics_block(
    df_test[df_test["amt"] <= media_amt],
    "amt <= media"
)

metrics_block(
    df_test[df_test["amt"] > media_amt],
    "amt > media"
)
