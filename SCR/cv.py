import lightgbm as lgb
import json
import pandas as pd
import numpy as np

# -------------------------------
# CONFIGURACIÓN
# -------------------------------
MODEL_FILE = "lightgbm_fraude.txt"
RULES_FILE = "fraud_rules.json"
TEST_FILE  = "pre_fraudtest.csv"
TARGET = "is_fraud"

# -------------------------------
# CARGAR MODELO
# -------------------------------
model = lgb.Booster(model_file=MODEL_FILE)

# -------------------------------
# CARGAR REGLAS
# -------------------------------
with open(RULES_FILE) as f:
    rules = json.load(f)

media_amt = rules["amt_mean_train"]
TH_LOW = rules["threshold_low"]
TH_HIGH = rules["threshold_high"]
FEATURES = rules["features"]

# -------------------------------
# CARGAR TEST
# -------------------------------
df = pd.read_csv(TEST_FILE)

X_test = df[FEATURES]

# -------------------------------
# SCORING
# -------------------------------
df = df.copy()
df["score"] = model.predict(X_test)

# Umbral dinámico
df["threshold"] = np.where(
    df["amt"] <= media_amt,
    TH_LOW,
    TH_HIGH
)

df["marked"] = (df["score"] >= df["threshold"]).astype(int)

# -------------------------------
# FUNCIÓN MÉTRICAS
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

    recall_amt = (
        fraud_marked_df["amt"].sum() / fraud_df["amt"].sum()
        if fraud_df["amt"].sum() > 0 else 0
    )

    print(f"\n=== {name.upper()} ===")
    print(f"Recall fraude (ops)        : {recall_ops:.4f}")
    print(f"Recall fraude (importe)    : {recall_amt:.4f}")
    print(f"Precisión                 : {precision:.4f}")
    print(f"% operaciones marcadas     : {pct_ops_marked*100:.2f}%")
    print(f"% importe total marcado    : {pct_amt_marked*100:.2f}%")
    print(f"Importe marcado total     : {marked_df['amt'].sum():.2f}")

# -------------------------------
# RESULTADOS
# -------------------------------
print("VALIDACIÓN FINAL — MODELO Y REGLAS CONGELADAS")
print(f"Media amt (train): {media_amt:.4f}")
print(f"Umbrales: <= media → {TH_LOW} | > media → {TH_HIGH}")

metrics_block(df, "global")
metrics_block(df[df["amt"] <= media_amt], "amt <= media")
metrics_block(df[df["amt"] > media_amt], "amt > media")