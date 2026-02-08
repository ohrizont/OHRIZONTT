import pandas as pd

file = "train.xlsx"

def clean_ram(x):
    if isinstance(x, str):
        return float(x.replace("GB", "").strip())
    return x

def clean_weight(x):
    if isinstance(x, str):
        return float(
            x.lower()
             .replace("kg", "")
             .replace(",", ".")
             .strip()
        )
    return x

with pd.ExcelWriter(file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    for sheet in ["train", "test"]:
        df = pd.read_excel(file, sheet_name=sheet)

        df["Ram"] = df["Ram"].apply(clean_ram)
        df["Weight"] = df["Weight"].apply(clean_weight)

        df.to_excel(writer, sheet_name=sheet, index=False)
