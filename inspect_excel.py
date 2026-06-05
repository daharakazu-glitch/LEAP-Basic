import pandas as pd
excel_file = "見出語・用例リスト(Part 1～5，＋α，外来語).xlsx"
try:
    xls = pd.ExcelFile(excel_file)
    print("Sheets:", xls.sheet_names)
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, nrows=5)
        print(f"\nSheet: {sheet}")
        print("Columns:", df.columns.tolist())
        print(df.head(2))
except Exception as e:
    print(f"Error: {e}")
