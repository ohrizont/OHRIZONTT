import pandas as pd

file = "train.xlsx"

def transform(df):
    out = pd.DataFrame()

    # Target
    if "Price_euros" in df.columns:
        out["Price_euros"] = df["Price_euros"]

    # Numéricas
    out["Ram"] = df["Ram"]
    out["screen_pixels"] = df["screen_pixels"]
    out["SSD_GB"] = df["SSD_GB"]
    out["Flash_GB"] = df["Flash_GB"]
    out["HDD_GB"] = df["HDD_GB"]

    # GPU brand → dummies
    out["gpu_amd"] = (df["gpu_brand"] == "AMD").astype(int)
    out["gpu_intel"] = (df["gpu_brand"] == "Intel").astype(int)
    out["gpu_nvidia"] = (df["gpu_brand"] == "Nvidia").astype(int)

    # GPU tier → dummies
    out["gpu_tier_low"] = (df["gpu_tier"] == "Low").astype(int)
    out["gpu_tier_mid"] = (df["gpu_tier"] == "Mid").astype(int)
    out["gpu_tier_high"] = (df["gpu_tier"] == "High").astype(int)

    # GPU / CPU flags
    out["gpu_pro"] = (df["gpu_class"] == "Pro").astype(int)
    out["cpu_high"] = (df["cpu_tier"] == "High").astype(int)
    out["cpu_pro"] = (df["cpu_class"] == "Pro").astype(int)

    # Marca premium
    out["is_premium_brand"] = df["Company"].isin(["Apple", "Razer"]).astype(int)

    # Peso dicotómico
    out["is_ultralight"] = (df["Weight"] < 1.3).astype(int)

    return out


df_train = pd.read_excel(file, sheet_name="train")
df_test  = pd.read_excel(file, sheet_name="test")

trans2_train = transform(df_train)
trans2_test  = transform(df_test)

with pd.ExcelWriter(file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    trans2_train.to_excel(writer, sheet_name="trans2_train", index=False)
    trans2_test.to_excel(writer, sheet_name="trans2_test", index=False)
