import pandas as pd
import numpy as np
from sklearn.preprocessing import OneHotEncoder, StandardScaler
from sklearn.compose import ColumnTransformer

# =========================
# CARGA DE DATOS
# =========================
FILE = "BankChurners1.xlsx"
SHEET = "BankChurners"

df = pd.read_excel(FILE, sheet_name=SHEET)

# Guardar churn aparte (NO se usa aquí)
churn = df["Attrition_Flag"]

# Eliminar ID y target
df = df.drop(columns=["CLIENTNUM", "Attrition_Flag"])

# =========================
# DEFINICIÓN DE VARIABLES
# =========================
categorical_cols = [
    "Gender",
    "Education_Level",
    "Marital_Status",
    "Income_Category",
    "Card_Category"
]

numeric_cols = [
    "Customer_Age",
    "Dependent_count",
    "Months_on_book",
    "Total_Relationship_Count",
    "Months_Inactive_12_mon",
    "Contacts_Count_12_mon",
    "Credit_Limit",
    "Total_Revolving_Bal",
    "Avg_Open_To_Buy",
    "Total_Amt_Chng_Q4_Q1",
    "Total_Trans_Amt",
    "Total_Trans_Ct",
    "Total_Ct_Chng_Q4_Q1",
    "Avg_Utilization_Ratio"
]

# =========================
# TRANSFORMACIONES PREVIAS
# =========================
log_vars = [
    "Total_Revolving_Bal",
    "Total_Amt_Chng_Q4_Q1",
    "Total_Ct_Chng_Q4_Q1",
    "Total_Trans_Amt"
]

for v in log_vars:
    df[v] = np.log1p(df[v])

# Forzar categóricas a texto
for c in categorical_cols:
    df[c] = df[c].astype(str)

# =========================
# PREPROCESADOR
# =========================
preprocess = ColumnTransformer(
    transformers=[
        ("num", StandardScaler(), numeric_cols),
        ("cat", OneHotEncoder(handle_unknown="ignore", sparse_output=False), categorical_cols)
    ]
)

X = preprocess.fit_transform(df)

feature_names = preprocess.get_feature_names_out()

# =========================
# GUARDAR EN EXCEL (OPCIÓN 2)
# =========================
X_df = pd.DataFrame(X, columns=feature_names, index=df.index)

with pd.ExcelWriter(
    FILE,
    engine="openpyxl",
    mode="a",
    if_sheet_exists="new"
) as writer:
    X_df.to_excel(writer, sheet_name="preprocesado", index=False)

print("✔ Preprocesamiento completado y guardado en Excel")
