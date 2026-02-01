import pandas as pd
import lightgbm as lgb
from sklearn.metrics import roc_auc_score, average_precision_score

# -------------------------------
# Cargar datos
# -------------------------------
df_train = pd.read_csv("pre_fraudtrain.csv")
df_test  = pd.read_csv("pre_fraudtest.csv")

FEATURES = [
    "amt",
    "hour_cos",
    "dayofweek",
    "is_weekend",
    "age",
    "gender"
]
TARGET = "is_fraud"

X_train = df_train[FEATURES]
y_train = df_train[TARGET]

X_test = df_test[FEATURES]
y_test = df_test[TARGET]

# -------------------------------
# Dataset LightGBM
# -------------------------------
train_set = lgb.Dataset(X_train, label=y_train)
test_set  = lgb.Dataset(X_test, label=y_test, reference=train_set)

# -------------------------------
# Parámetros baseline
# -------------------------------
params = {
    "objective": "binary",
    "metric": ["auc", "average_precision"],
    "boosting_type": "gbdt",
    "learning_rate": 0.05,
    "num_leaves": 31,
    "max_depth": -1,
    "min_data_in_leaf": 200,
    "feature_fraction": 0.8,
    "bagging_fraction": 0.8,
    "bagging_freq": 5,
    "verbosity": -1,
    "scale_pos_weight": (len(y_train) - y_train.sum()) / y_train.sum()
}

# -------------------------------
# Entrenamiento
# -------------------------------
model = lgb.train(
    params,
    train_set,
    num_boost_round=500,
    valid_sets=[test_set]
)

# -------------------------------
# Evaluación
# -------------------------------
y_prob = model.predict(X_test)

roc_auc = roc_auc_score(y_test, y_prob)
pr_auc  = average_precision_score(y_test, y_prob)

print("\n=== LIGHTGBM BASELINE ===")
print(f"ROC AUC : {roc_auc:.4f}")
print(f"PR  AUC : {pr_auc:.4f}")
