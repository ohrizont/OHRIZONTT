import pandas as pd
import numpy as np
import hdbscan
from sklearn.preprocessing import StandardScaler

# =========================
# CONFIG
# =========================
FILE = "BankChurners1.xlsx"
SHEET = "BankChurners"
OUT_SHEET = "hdbscan_conductual"

# =========================
# CARGA
# =========================
df = pd.read_excel(FILE, sheet_name=SHEET)
churn = df["Attrition_Flag"]

# =========================
# VARIABLES CONDUCTUALES
# =========================
vars_cond = [
    "Months_Inactive_12_mon",
    "Contacts_Count_12_mon",
    "Total_Ct_Chng_Q4_Q1",
    "Total_Trans_Ct",
    "Total_Amt_Chng_Q4_Q1"
]

X = df[vars_cond].copy()

# Log en ratios
for v in ["Total_Ct_Chng_Q4_Q1", "Total_Amt_Chng_Q4_Q1"]:
    X[v] = np.log1p(X[v])

# Escalado
X_scaled = StandardScaler().fit_transform(X)

# =========================
# HDBSCAN
# =========================
clusterer = hdbscan.HDBSCAN(
    min_cluster_size=300,
    min_samples=50,
    metric="euclidean"
)

labels = clusterer.fit_predict(X_scaled)

# =========================
# RESULTADOS
# =========================
out = df[vars_cond].copy()
out["cluster_hdbscan"] = labels
out["churn"] = churn

summary = (
    out.groupby("cluster_hdbscan")["churn"]
    .apply(lambda x: (x == "Attrited Customer").mean())
)

sizes = out["cluster_hdbscan"].value_counts().sort_index()

results = pd.DataFrame({
    "size": sizes,
    "churn_rate": summary
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
    results.to_excel(writer, sheet_name=OUT_SHEET)
    out.to_excel(writer, sheet_name="conductual_hdbscan", index=False)

print("✔ HDBSCAN conductual completado")
