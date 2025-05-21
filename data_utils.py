import pandas as pd

def load_excel_data(path, sheet_name):
    return pd.read_excel(path, sheet_name=sheet_name)

def clean_transactions(df):
    df_clean = df.copy()
    df_clean = df_clean[df_clean["Quantity"] > 0]
    df_clean = df_clean[~df_clean["Invoice"].astype(str).str.startswith("C")]
    df_clean = df_clean[df_clean["Customer ID"].notna()]
    return df_clean
