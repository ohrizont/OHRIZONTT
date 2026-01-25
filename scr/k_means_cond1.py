import pandas as pd
import numpy as np
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score

# =========================
# CONFIG
# =========================
FILE = "BankChurners1.xlsx"
SHEET_ORIG = "BankChurners"
OUT_SHEET = "kmeans_conductual"
RANDOM_STATE = 42

# =========================
# CARGA
# =========================
df = pd.read_excel(FILE, sheet_name=SHEET_ORIG)

# Guardamos churn para análisis posterior
churn = df["Attrition_Flag"]

# =========================
# VARIABLES CONDUCTUALES
# =========================
behavioral_vars = [
    "Months_Inactive_12_mon",
    "Contacts_Count_12_mon",
    "Total_Ct_Chng_Q4_Q1",
    "Total_Trans_Ct",
    "Total_Amt_Chng_Q4_Q1"
]

X = df[behavioral_vars].copy()

# =========================
# TRANSFORMACIONES
# =========================
# Log en ratios
for v in ["Total_Ct_Chng_Q4_Q1", "Total_Amt_Chng_Q4_Q1"]:
    X[v] = np.log1p(X[v])

# Escalado
scaler = StandardScaler()
X_scaled = scaler.fit_transform(X)

# =========================
# K-MEANS k = 2..5
# =========================
rows = []

for k in range(2, 6):
    km = KMeans(n_clusters=k, n_init=30, random_state=RANDOM_STATE)
    labels = km.fit_predict(X_scaled)

    sil = silhouette_score(X_scaled, labels)

    tmp = pd.DataFrame({
        "cluster": labels,
        "churn": churn
    })

    churn_rate = (
        tmp.groupby("cluster")["churn"]
        .apply(lambda x: (x == "Attrited Customer").mean())
        .to_dict()
    )

    sizes = tmp["cluster"].value_counts().to_dict()

    rows.append({
        "k": k,
        "silhouette": sil,
        "sizes": sizes,
        "churn_rate": churn_rate
    })

results = pd.DataFrame(rows)

# =========================
# MODELO FINAL (k=2)
# =========================
km_final = KMeans(n_clusters=2, n_init=30, random_state=RANDOM_STATE)
labels_final = km_final.fit_predict(X_scaled)

df_out = df[behavioral_vars].copy()
df_out["cluster_conductual"] = labels_final
df_out["churn"] = churn

# =========================
# CENTROIDES (ESCALADOS)
# =========================
centroids = pd.DataFrame(
    km_final.cluster_centers_,
    columns=behavioral_vars
)

# =========================
# GUARDAR EN EXCEL
# =========================
with pd.ExcelWriter(
    FILE,
    engine="openpyxl",
    mode="a",
    if_sheet_exists="new"
) as writer:
    results.to_excel(writer, sheet_name=OUT_SHEET, index=False)
    centroids.to_excel(writer, sheet_name="centroides_conductual")
    df_out.to_excel(writer, sheet_name="conductual_con_cluster", index=False)

print("✔ K-means conductual completado")
