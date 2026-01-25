import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.tree import DecisionTreeClassifier, export_text
from sklearn.metrics import roc_auc_score

# =========================
# CONFIG
# =========================
FILE = "BankChurners1.xlsx"
SHEET = "BankChurners"
OUT_SHEET = "arbol_decision"
RANDOM_STATE = 42

# =========================
# CARGA
# =========================
df = pd.read_excel(FILE, sheet_name=SHEET)

y = (df["Attrition_Flag"] == "Attrited Customer").astype(int)

# Variable HDBSCAN (alto riesgo)
df["hdbscan_high_risk"] = df["cluster_hdbscan"].isin([3, 6, 7]).astype(int)

X = df[
    [
        "Months_Inactive_12_mon",
        "Contacts_Count_12_mon",
        "Total_Ct_Chng_Q4_Q1",
        "Total_Trans_Ct",
        "Total_Amt_Chng_Q4_Q1",
        "hdbscan_high_risk",
    ]
].copy()

# =========================
# TRANSFORMACIONES
# =========================
for v in ["Total_Ct_Chng_Q4_Q1", "Total_Amt_Chng_Q4_Q1"]:
    X[v] = np.log1p(X[v])

# =========================
# TRAIN / TEST
# =========================
Xtr, Xte, ytr, yte = train_test_split(
    X, y,
    test_size=0.3,
    random_state=RANDOM_STATE,
    stratify=y
)

# =========================
# ÁRBOL DE DECISIÓN
# =========================
tree = DecisionTreeClassifier(
    max_depth=4,
    min_samples_leaf=200,
    random_state=RANDOM_STATE
)

tree.fit(Xtr, ytr)

# =========================
# EVALUACIÓN
# =========================
proba = tree.predict_proba(Xte)[:, 1]
auc = roc_auc_score(yte, proba)

# =========================
# REGLAS DEL ÁRBOL
# =========================
rules = export_text(tree, feature_names=list(X.columns))

rules_df = pd.DataFrame(
    {"rule": rules.split("\n")}
)

# =========================
# LIFT TOP %
# =========================
df_eval = pd.DataFrame({
    "y_true": yte.values,
    "proba": proba
}).sort_values("proba", ascending=False)

base_rate = df_eval["y_true"].mean()

rows_lift = []
for p in [0.1, 0.2]:
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
    lift_df.to_excel(writer, sheet_name="lift_arbol", index=False)
    rules_df.to_excel(writer, sheet_name="reglas_arbol", index=False)

print("✔ Árbol de decisión entrenado y guardado en Excel")
print(f"AUC Árbol: {auc:.4f}")