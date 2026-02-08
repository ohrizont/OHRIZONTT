import pandas as pd

file = "train.xlsx"
threshold = 10

# Usamos train para calcular frecuencias
df_train = pd.read_excel(file, sheet_name="train")
counts = df_train["Company"].value_counts()

major_brands = counts[counts >= threshold].index

with pd.ExcelWriter(file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    for sheet in ["train", "test"]:
        df = pd.read_excel(file, sheet_name=sheet)

        df["Company_grp"] = df["Company"].where(
            df["Company"].isin(major_brands), "Other"
        )

        df.to_excel(writer, sheet_name=sheet, index=False)
