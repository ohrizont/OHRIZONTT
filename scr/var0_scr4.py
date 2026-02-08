import pandas as pd
import re

file = "train.xlsx"

def resolution_to_pixels(x):
    if isinstance(x, str):
        m = re.search(r"(\d{3,4})x(\d{3,4})", x)
        if m:
            return int(m.group(1)) * int(m.group(2))
    return None

with pd.ExcelWriter(file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    for sheet in ["train", "test"]:
        df = pd.read_excel(file, sheet_name=sheet)

        df["screen_pixels"] = df["ScreenResolution"].apply(resolution_to_pixels)

        df.to_excel(writer, sheet_name=sheet, index=False)