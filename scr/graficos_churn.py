import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# =========================
# CONFIG
# =========================
FILE = "BankChurners1.xlsx"
SHEET = "scoring_final_gb"

# =========================
# CARGA
# =========================
df = pd.read_excel(FILE, sheet_name=SHEET)

df["y_true"] = (df["Attrition_Flag"] == "Attrited Customer").astype(int)

# =========================
# 1️⃣ CURVA DE GANANCIAS
# =========================
df_sorted = df.sort_values("score_churn", ascending=False).reset_index(drop=True)

df_sorted["cum_churn"] = df_sorted["y_true"].cumsum()
df_sorted["pct_clients"] = np.arange(1, len(df_sorted)+1) / len(df_sorted)
df_sorted["pct_churn"] = df_sorted["cum_churn"] / df_sorted["y_true"].sum()

plt.figure()
plt.plot(df_sorted["pct_clients"], df_sorted["pct_churn"], label="Modelo")
plt.plot([0,1], [0,1], "--", label="Aleatorio")
plt.xlabel("% Clientes contactados")
plt.ylabel("% Churn capturado")
plt.title("Curva de Ganancias - Churn")
plt.legend()
plt.tight_layout()
plt.savefig("curva_ganancias.png")
plt.close()

# =========================
# 2️⃣ LIFT POR DECIL
# =========================
lift_decil = df.groupby("decil_churn")["y_true"].mean()
lift = lift_decil / df["y_true"].mean()

plt.figure()
lift.plot(kind="bar")
plt.axhline(1, linestyle="--")
plt.ylabel("Lift")
plt.title("Lift por decil de riesgo")
plt.tight_layout()
plt.savefig("lift_por_decil.png")
plt.close()

# =========================
# 3️⃣ DISTRIBUCIÓN SCORE
# =========================
plt.figure()
plt.hist(df[df["y_true"] == 0]["score_churn"], bins=30, alpha=0.6, label="No churn")
plt.hist(df[df["y_true"] == 1]["score_churn"], bins=30, alpha=0.6, label="Churn")
plt.xlabel("Score churn")
plt.ylabel("Clientes")
plt.title("Distribución del score de churn")
plt.legend()
plt.tight_layout()
plt.savefig("distribucion_score.png")
plt.close()

# =========================
# 4️⃣ IMPORTANCIAS (manual)
# =========================
importance = {
    "Total_Trans_Ct": 0.49,
    "Total_Ct_Chng_Q4_Q1": 0.26,
    "Total_Amt_Chng_Q4_Q1": 0.12,
    "Months_Inactive_12_mon": 0.09,
    "Contacts_Count_12_mon": 0.05
}

plt.figure()
plt.barh(list(importance.keys()), list(importance.values()))
plt.xlabel("Importancia relativa")
plt.title("Importancia de variables")
plt.tight_layout()
plt.savefig("importancia_variables.png")
plt.close()

print("✔ Gráficos generados y guardados")
