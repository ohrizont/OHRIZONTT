import pandas as pd

file = "train.xlsx"

def gpu_features(gpu):
    if not isinstance(gpu, str):
        return "Unknown", "Unknown", "Low"

    g = gpu.lower()

    # BRAND
    if "intel" in g:
        brand = "Intel"
    elif "amd" in g or "radeon" in g or "firepro" in g:
        brand = "AMD"
    elif "nvidia" in g or "geforce" in g or "quadro" in g:
        brand = "Nvidia"
    else:
        brand = "Other"

    # CLASS
    if "intel" in g or "graphics" in g and "iris" not in g:
        gpu_class = "Integrated"
    elif "quadro" in g or "firepro" in g:
        gpu_class = "Pro"
    elif "gtx" in g or "rx" in g:
        gpu_class = "Gaming"
    else:
        gpu_class = "Consumer"

    # TIER
    if any(x in g for x in ["1080", "1070", "1060", "980", "970", "rx 580", "rx 560"]):
        tier = "High"
    elif any(x in g for x in ["1050", "960", "950", "940mx", "mx150", "mx130", "iris"]):
        tier = "Mid"
    else:
        tier = "Low"

    return brand, gpu_class, tier


with pd.ExcelWriter(file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    for sheet in ["train", "test"]:
        df = pd.read_excel(file, sheet_name=sheet)

        df[["gpu_brand", "gpu_class", "gpu_tier"]] = df["Gpu"].apply(
            lambda x: pd.Series(gpu_features(x))
        )

        df.to_excel(writer, sheet_name=sheet, index=False)
