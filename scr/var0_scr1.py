import pandas as pd
import re

file = "train.xlsx"

df = pd.read_excel(file, sheet_name="test")

def extract_resolution(x):
    if isinstance(x, str):
        m = re.search(r"\b\d{4}x\d{4}\b", x)
        return m.group(0) if m else None
    return None

df["ScreenResolution"] = df["ScreenResolution"].apply(extract_resolution)

with pd.ExcelWriter(file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df.to_excel(writer, sheet_name="variables", index=False)