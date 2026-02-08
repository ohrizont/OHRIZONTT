import pandas as pd

file = "train.xlsx"

def cpu_features(cpu):
    if not isinstance(cpu, str):
        return "Unknown", "Consumer", "Low"

    c = cpu.lower()

    # BRAND
    brand = "AMD" if "amd" in c or "ryzen" in c else "Intel"

    # CLASS
    if "xeon" in c:
        cpu_class = "Pro"
    elif any(x in c for x in ["atom", "celer", "pentium", "core m", " y"]):
        cpu_class = "LowPower"
    elif any(x in c for x in ["i7", "hq", "hk", "ryzen 7", "ryzen 5"]):
        cpu_class = "Performance"
    else:
        cpu_class = "Consumer"

    # TIER
    if any(x in c for x in ["xeon", "i7", "hq", "hk", "ryzen 7"]):
        tier = "High"
    elif any(x in c for x in ["i5", "ryzen 5", "i3", "a10", "a12"]):
        tier = "Mid"
    else:
        tier = "Low"

    return brand, cpu_class, tier


with pd.ExcelWriter(file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    for sheet in ["train", "test"]:
        df = pd.read_excel(file, sheet_name=sheet)

        df[["cpu_brand", "cpu_class", "cpu_tier"]] = df["Cpu"].apply(
            lambda x: pd.Series(cpu_features(x))
        )

        df.to_excel(writer, sheet_name=sheet, index=False)
