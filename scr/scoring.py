import pandas as pd
import numpy as np
from sklearn.ensemble import GradientBoostingClassifier

# =========================
# CONFIG
# =========================
FILE = "BankChurners1.xlsx"
SHEET = "BankChurners"
OUT_SHEET = "scoring_final_gb"
TOP_PERCENT = 0.10   # 10% clientes a contactar
RANDOM_STATE = 42

# =========================
# CARGA
# =========================
df = pd.read_excel(FILE, sheet_name=SHEET)

# Target solo para entrenamiento
y = (df["Attrition_Flag"] == "Attrited Customer").astype(int)

# Features
X = df[
    [
        "Months_Inactive_12_mon",
        "Contacts_Count_12_mon",
        "Total_Ct_Chng_Q4_Q1",
        "Total_Trans_Ct",
        "Total_Amt_Chng_Q4_Q1",
    ]
].copy()

# =========================
# TRANSFORMACIONES
# =========================
for v in ["Total_Ct_Chng_Q4_Q1", "Total_Amt_Chng_Q4_Q1"]:
    X[v] = np.log1p(X[v])

# =========================
# ENTRENAR MODELO FINAL
# =========================
gb = GradientBoostingClassifier(
    n_estimators=300,
    learning_rate=0.05,
    max_depth=3,
    subsample=0.8,
    min_samples_leaf=100,
    random_state=RANDOM_STATE
)

gb.fit(X, y)

# =========================
# SCORING
# =========================
df["score_churn"] = gb.predict_proba(X)[:, 1]

# =========================
# RANKING
# =========================
df = df.sort_values("score_churn", ascending=False).reset_index(drop=True)

df["rank_churn"] = df.index + 1
df["percentil_churn"] = df["rank_churn"] / len(df)
df["decil_churn"] = pd.qcut(df["score_churn"], 10, labels=False, duplicates="drop") + 1

# =========================
# FLAG OPERATIVO
# =========================
cutoff = int(len(df) * TOP_PERCENT)
df["flag_contactar"] = (df["rank_churn"] <= cutoff).astype(int)

# =========================
# COLUMNAS FINALES
# =========================
cols_out = [
    "CLIENTNUM",
    "score_churn",
    "rank_churn",
    "percentil_churn",
    "decil_churn",
    "flag_contactar",
    "Attrition_Flag"
]

df_out = df[cols_out]

# =========================
# GUARDAR EN EXCEL
# =========================
with pd.ExcelWriter(
    FILE,
    engine="openpyxl",
    mode="a",
    if_sheet_exists="new"
) as writer:
    df_out.to_excel(writer, sheet_name=OUT_SHEET, index=False)

print("✔ Scoring y ranking final generados")
print(f"Clientes a contactar (top {int(TOP_PERCENT*100)}%): {cutoff}")
