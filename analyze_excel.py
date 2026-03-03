import pandas as pd
import sys

def analyze(filepath):
    print(f"--- Analyzing {filepath} ---")
    try:
        xls = pd.ExcelFile(filepath)
        print(f"Sheet names: {xls.sheet_names}")
        for sheet in xls.sheet_names:
            df = pd.read_excel(filepath, sheet_name=sheet, nrows=20)
            print(f"\nSheet: {sheet}")
            print("Columns:", list(df.columns))
            print("Head:")
            print(df.head(10).to_string())
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    analyze("c:/Users/tenni/OneDrive/デスクトップ/Amazon納品梱包情報変換ツール/Excel①（Amazonの梱包グループのシート）.xlsx")
    analyze("c:/Users/tenni/OneDrive/デスクトップ/Amazon納品梱包情報変換ツール/Excel②（ラクマートの輸送箱などの情報シート）.xlsx")
