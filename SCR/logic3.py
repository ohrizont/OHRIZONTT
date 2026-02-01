import pandas as pd
import numpy as np
from pathlib import Path

from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (
    roc_auc_score,
    average_precision_score,
    log_loss,
    brier_score_loss,
    confusion_matrix,
    classification_report
)

import statsmodels.api as sm
from statsmodels.stats.outliers_influence import variance_inflation_factor

# =================================================
# CONFIG
# =================================================
TRAIN = "pre_fraudtrain.csv"
TEST  = "pre_fraudtest.csv"
EXCEL = "fraude.xlsx"
SHEET = "logic"

FEATURES = [
    "amt", "gender", "age",
    "is_weekend", "dayofweek",
     "hour_cos"
]
TARGET = "is_fraud"

# =================================================
# LOAD DATA
# =================================================
df_train = pd.read_csv(TRAIN)
df_test  = pd.read_csv(TEST)

X_train = df_train[FEATURES]
y_train = df_train[TARGET]
X_test  = df_test[FEATURES]
y_test  = df_test[TARGET]

# =================================================
# SCALE
# =================================================
scaler = StandardScaler()
X_train_s = scaler.fit_transform(X_train)
X_test_s  = scaler.transform(X_test)

# =================================================
# LOGISTIC REGRESSION (sklearn)
# =================================================
clf = LogisticRegression(max_iter=1000, class_weight="balanced", n_jobs=-1)
clf.fit(X_train_s, y_train)

y_prob = clf.predict_proba(X_test_s)[:, 1]
y_pred = (y_prob >= 0.5).astype(int)

# =================================================
# METRICS
# =================================================
roc_auc = roc_auc_score(y_test, y_prob)
pr_auc  = average_precision_score(y_test, y_prob)
ll      = log_loss(y_test, y_prob)
brier   = brier_score_loss(y_test, y_prob)

print("\n=== MÉTRICAS GLOBALES ===")
print(f"ROC AUC : {roc_auc:.4f}")
print(f"PR  AUC : {pr_auc:.4f}")
print(f"LogLoss : {ll:.4f}")
print(f"Brier   : {brier:.4f}")

metrics_df = pd.DataFrame({
    "metric": ["ROC_AUC", "PR_AUC", "LogLoss", "Brier"],
    "value":  [roc_auc, pr_auc, ll, brier]
})

# =================================================
# CONFUSION MATRIX
# =================================================
cm = confusion_matrix(y_test, y_pred)
cm_df = pd.DataFrame(
    cm,
    index=["Actual_NoFraud", "Actual_Fraud"],
    columns=["Pred_NoFraud", "Pred_Fraud"]
)

print("\n=== CONFUSION MATRIX ===")
print(cm_df)

# =================================================
# CLASSIFICATION REPORT
# =================================================
report_df = pd.DataFrame(
    classification_report(y_test, y_pred, output_dict=True)
).transpose()

print("\n=== CLASSIFICATION REPORT ===")
print(report_df)

# =================================================
# STATISTICAL LOGIT (statsmodels)
# =================================================
X_sm = sm.add_constant(X_train_s)
logit = sm.Logit(y_train, X_sm)
res = logit.fit(disp=False)

coef_df = pd.DataFrame({
    "variable": ["const"] + FEATURES,
    "coef": res.params,
    "std_err": res.bse,
    "z_value": res.tvalues,
    "p_value": res.pvalues,
    "odds_ratio": np.exp(res.params)
})

print("\n=== COEFICIENTES CON SIGNIFICACIÓN ===")
print(coef_df)

# =================================================
# VIF
# =================================================
vif_df = pd.DataFrame({
    "variable": FEATURES,
    "VIF": [
        variance_inflation_factor(X_train_s, i)
        for i in range(X_train_s.shape[1])
    ]
})

print("\n=== VIF (COLINEALIDAD) ===")
print(vif_df)

# =================================================
# SAVE TO EXCEL (ORDERED)
# =================================================
if Path(EXCEL).exists():
    writer = pd.ExcelWriter(EXCEL, engine="openpyxl", mode="a", if_sheet_exists="replace")
else:
    writer = pd.ExcelWriter(EXCEL, engine="openpyxl", mode="w")

with writer:
    row = 0
    metrics_df.to_excel(writer, sheet_name=SHEET, index=False, startrow=row)
    row += len(metrics_df) + 3

    cm_df.to_excel(writer, sheet_name=SHEET, startrow=row)
    row += len(cm_df) + 3

    report_df.to_excel(writer, sheet_name=SHEET, startrow=row)
    row += len(report_df) + 3

    coef_df.to_excel(writer, sheet_name=SHEET, index=False, startrow=row)
    row += len(coef_df) + 3

    vif_df.to_excel(writer, sheet_name=SHEET, index=False, startrow=row)

print("\n✔ TODO calculado, mostrado y guardado correctamente")
