import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import roc_auc_score, precision_score, recall_score

# =========================
# CONFIG
# =========================
FILE = "BankChurners1.xlsx"
SHEET = "BankChurners"
OUT_SHEET = "logistica_conductual_cluster"
RANDOM_STATE = 42

# =========================
# CARGA
# =========================
df = pd.read_excel(FILE, sheet_name=SHEET)

y = (df["Attrition_Flag"] == "Attrited Customer").astype(int)

X = df[
    [
        "Months_Inactive_12_mon",
        "Contacts_Count_12_mon",
        "Total_Ct_Chng_Q4_Q1",
        "Total_Trans_Ct",
        "Total_Amt_Chng_Q4_Q1",
        "cluster_hdbscan"
    ]
].copy()

# Log en ratios
for v in ["Total_Ct_Chng_Q4_Q1", "Total_Amt_Chng_Q4_Q1"]:
    X[v] = np.log1p(X[v])

# Escalado SOLO para variables continuas
scaler = StandardScaler()
X_scaled = X.copy()
X_scaled[
    [
        "Months_Inactive_12_mon",
        "Contacts_Count_12_mon",
        "Total_Ct_Chng_Q4_Q1",
        "Total_Trans_Ct",
        "Total_Amt_Chng_Q4_Q1"
    ]
] = scaler.fit_transform(
    X[
        [
            "Months_Inactive_12_mon",
            "Contacts_Count_12_mon",
            "Total_Ct_Chng_Q4_Q1",
            "Total_Trans_Ct",
            "Total_Amt_Chng_Q4_Q1"
        ]
    ]
)

# =========================
# TRAIN / TEST
# =========================
Xtr, Xte, ytr, yte = train_test_split(
    X_scaled, y, test_size=0.3, random_state=RANDOM_STATE, stratify=y
)

# =========================
# LOGÍSTICA
# =========================
logit = LogisticRegression(
    penalty="l2",
    C=1.0,
    max_iter=1000,
    random_state=RANDOM_STATE
)

logit.fit(Xtr, ytr)

proba = logit.predict_proba(Xte)[:, 1]
auc = roc_auc_score(yte, proba)

# =========================
# COEFICIENTES
# =========================
coef_df = pd.DataFrame({
    "variable": X_scaled.columns,
    "coef_log_odds": logit.coef_[0]
}).sort_values("coef_log_odds", ascending=False)

# =========================
# EVALUACIÓN UMBRALES
# =========================
thresholds = np.arange(0.05, 0.55, 0.05)
rows_thr = []

for t in thresholds:
    preds = (proba >= t).astype(int)
    rows_thr.append({
        "threshold": t,
        "precision": precision_score(yte, preds),
        "recall": recall_score(yte, preds)
    })

threshold_df = pd.DataFrame(rows_thr)

# =========================
# LIFT POR TOP %
# =========================
df_eval = pd.DataFrame({
    "y_true": yte.values,
    "proba": proba
}).sort_values("proba", ascending=False)

base_rate = df_eval["y_true"].mean()

rows_lift = []
for p in [0.1, 0.2, 0.3]:
    n = int(len(df_eval) * p)
    top = df_eval.head(n)

    churn_rate = top["y_true"].mean()
    lift = churn_rate / base_rate

    rows_lift.append({
        "top_percent": p,
        "churn_rate": churn_rate,
        "lift": lift
    })

lift_df = pd.DataFrame(rows_lift)

# =========================
# RESUMEN
# =========================
summary_df = pd.DataFrame({
    "metric": ["AUC", "base_churn_rate"],
    "value": [auc, base_rate]
})

# =========================
# GUARDAR EN EXCEL
# =========================
with pd.ExcelWriter(
    FILE,
    engine="openpyxl",
    mode="a",
    if_sheet_exists="new"
) as writer:
    summary_df.to_excel(writer, sheet_name=OUT_SHEET, index=False)
    coef_df.to_excel(writer, sheet_name="coeficientes_cluster", index=False)
    threshold_df.to_excel(writer, sheet_name="umbrales_cluster", index=False)
    lift_df.to_excel(writer, sheet_name="lift_cluster", index=False)

print("✔ Logística conductual + cluster_hdbscan completada")
