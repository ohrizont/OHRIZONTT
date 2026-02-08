import pandas as pd

file = "train.xlsx"
cols = [
    "Company_grp","TypeName","Inches",
    "ScreenResolution","Ram","OpSys","Weight","Price_euros", "Flash_GB",	"SSD_GB", "HDD_GB", "gpu_brand", "gpu_class", "gpu_tier", "cpu_brand", "cpu_class", "cpu_tier"

]

# Leer datos
df = pd.read_excel(file, sheet_name="train", usecols=cols)

# Crear dataframe de variables únicas
max_len = max(df[c].nunique() for c in cols)

data = {}
for c in cols:
    values = df[c].dropna().unique()
    data[c] = list(values) + [None] * (max_len - len(values))

df_vars = pd.DataFrame(data)

# Escribir en el mismo fichero
with pd.ExcelWriter(file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df_vars.to_excel(writer, sheet_name="variables", index=False)