import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor
from sklearn.preprocessing import OneHotEncoder
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
from sklearn.metrics import mean_absolute_error, r2_score
from sklearn.inspection import permutation_importance
from sklearn.impute import SimpleImputer

file = "train.xlsx"

# Columnas
num_cols = [
    "Inches", "Ram", "cpu_rating", "Weight", "SSD_GB"
]

cat_cols = [
    "Company_grp", "OpSys", "gpu_tier"
]

# Leer datos
df = pd.read_excel(file, sheet_name="train")

X = df[num_cols + cat_cols]
y = df["Price_euros"]

X_train, X_val, y_train, y_val = train_test_split(
    X, y, test_size=0.2, random_state=42
)

# =========================
# Preprocesado con imputación
# =========================
num_pipe = Pipeline([
    ("imputer", SimpleImputer(strategy="median"))
])

cat_pipe = Pipeline([
    ("imputer", SimpleImputer(strategy="most_frequent")),
    ("onehot", OneHotEncoder(handle_unknown="ignore"))
])

preprocess = ColumnTransformer(
    transformers=[
        ("cat", cat_pipe, cat_cols),
        ("num", num_pipe, num_cols)
    ]
)

# Modelo
rf = RandomForestRegressor(
    n_estimators=300,
    max_depth=None,
    min_samples_leaf=3,
    random_state=42,
    n_jobs=4
)

pipe = Pipeline([
    ("prep", preprocess),
    ("rf", rf)
])

# Entrenar
pipe.fit(X_train, y_train)

# Evaluar
preds = pipe.predict(X_val)
print("MAE:", mean_absolute_error(y_val, preds))
print("R2 :", r2_score(y_val, preds))

# =========================
# Importancia por permutación
# =========================
r = permutation_importance(
    pipe,
    X_val,
    y_val,
    n_repeats=10,
    random_state=42,
    scoring="neg_mean_absolute_error"
)

perm_imp = pd.DataFrame({
    "variable": X_val.columns,
    "importance": r.importances_mean,
    "std": r.importances_std
}).sort_values("importance", ascending=False)

print("\nImportancia por permutación:")
print(perm_imp)
