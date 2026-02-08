import pandas as pd
import re

file = "train.xlsx"

def to_gb(value, unit):
    value = float(value)
    return value * 1024 if unit.lower() == "tb" else value

def parse_memory(mem):
    flash = ssd = hdd = 0

    if not isinstance(mem, str):
        return flash, ssd, hdd

    parts = re.split(r"\s*\+\s*", mem)

    for p in parts:
        m = re.search(r"([\d\.]+)\s*(TB|GB)\s*(SSD|HDD|Flash Storage|Hybrid)", p, re.I)
        if not m:
            continue

        value, unit, kind = m.groups()
        gb = to_gb(value, unit)

        kind = kind.lower()
        if "flash" in kind:
            flash += gb
        elif "ssd" in kind or "hybrid" in kind:
            ssd += gb
        elif "hdd" in kind:
            hdd += gb

    return flash, ssd, hdd


with pd.ExcelWriter(file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    for sheet in ["train", "test"]:
        df = pd.read_excel(file, sheet_name=sheet)

        df[["Flash_GB", "SSD_GB", "HDD_GB"]] = df["Memory"].apply(
            lambda x: pd.Series(parse_memory(x))
        )

        df.to_excel(writer, sheet_name=sheet, index=False)