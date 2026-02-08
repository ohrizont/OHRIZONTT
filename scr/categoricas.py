import pandas as pd
from sklearn.preprocessing import OneHotEncoder
from sklearn.compose import ColumnTransformer

file = "train.xlsx"

cat_cols = [
    "Company","TypeName","OpSys",
    "gpu_brand","gpu_class","gpu_tier",
    "cpu_brand","cpu_class","cpu_tier"
]


num_cols = [
    "Inches","screen_pixels","Ram","Weight",
    "Flash_GB","SSD_GB","HDD_GB"
]

# Leer datos
df_train = pd.read_excel(file, sheet_name="train")
df_test  = pd.read_excel(file, sheet_name="test")

y_train = df_train["Price_euros"]
X_train = df_train[cat_cols + num_cols]
X_test  = df_test[cat_cols + num_cols]

# Transformer
preprocess = ColumnTransformer(
    transformers=[
        ("cat", OneHotEncoder(handle_unknown="ignore", sparse_output=False), cat_cols),
        ("num", "passthrough", num_cols)
    ]
)

# Fit SOLO en train
X_train_enc = preprocess.fit_transform(X_train)
X_test_enc  = preprocess.transform(X_test)

# Nombres de columnas
cat_names = preprocess.named_transformers_["cat"].get_feature_names_out(cat_cols)

all_columns = list(cat_names) + num_cols

# DataFrames finales
train_model = pd.DataFrame(X_train_enc, columns=all_columns)
train_model["Price_euros"] = y_train.values

test_model = pd.DataFrame(X_test_enc, columns=all_columns)

# Escribir nuevas hojas
with pd.ExcelWriter(file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    train_model.to_excel(writer, sheet_name="train_model", index=False)
    test_model.to_excel(writer, sheet_name="test_model", index=False)
