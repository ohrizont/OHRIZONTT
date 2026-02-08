import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor
from sklearn.preprocessing import OneHotEncoder
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
from sklearn.metrics import mean_absolute_error, r2_score

# =========================
# Configuración
# =========================
file = "train.xlsx"
RANDOM_STATE = 42

target = "Price_euros"

num_cols = ["Inches", "Ram", "Weight", "cpu_rating", "SSD_GB"]
cat_cols = ["Company_grp", "TypeName", "OpSys"]  # cpu_tier eliminado

# =========================
# Cargar datos
# =========================
df = pd.read_excel(file, sheet_name="train")

# Imputar posibles NaN en cpu_rating
df["cpu_rating"] = df["cpu_rating"].fillna(df["cpu_rating"].median())

X = df[num_cols + cat_cols]
y = df[target]

# =========================
# Split aleatorio 80 / 20
# =========================
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.2, shuffle=True, random_state=RANDOM_STATE
)

# =========================
# Preprocesado
# =========================
preprocess = ColumnTransformer(
    transformers=[
        ("cat", OneHotEncoder(handle_unknown="ignore"), cat_cols),
        ("num", "passthrough", num_cols)
    ]
)

# =========================
# Modelo Random Forest
# =========================
rf = RandomForestRegressor(
    n_estimators=300,
    max_depth=None,
    min_samples_leaf=2,
    random_state=RANDOM_STATE,
    n_jobs=4
)

pipe = Pipeline([
    ("prep", preprocess),
    ("rf", rf)
])

# =========================
# Entrenar
# =========================
pipe.fit(X_train, y_train)

# =========================
# Predicciones
# =========================
pred_train = pipe.predict(X_train)
pred_test = pipe.predict(X_test)

# =========================
# Métricas
# =========================
def metrics(y_true, y_pred):
    return (
        mean_absolute_error(y_true, y_pred),
        r2_score(y_true, y_pred)
    )

mae_tr, r2_tr = metrics(y_train, pred_train)
mae_te, r2_te = metrics(y_test, pred_test)

print("===== RESULTADOS =====")
print(f"TRAIN (80%) -> MAE: {mae_tr:.3f} | R2: {r2_tr:.4f}")
print(f"TEST  (20%) -> MAE: {mae_te:.3f} | R2: {r2_te:.4f}")

# =========================
# Estadísticos adicionales
# =========================
rf_fitted = pipe.named_steps["rf"]

n_features = pipe.named_steps["prep"].transform(X_train).shape[1]

print("\n===== ESTADÍSTICOS RF =====")
print("Nº árboles:", rf_fitted.n_estimators)
print("Profundidad media de árboles:",
      np.mean([est.tree_.max_depth for est in rf_fitted.estimators_]))
print("Nº hojas medio:",
      np.mean([est.tree_.n_leaves for est in rf_fitted.estimators_]))
print("Nº features usadas:", n_features)
