import urllib.request
import pandas as pd
import io
import os

url = "https://docs.google.com/spreadsheets/d/1fQhkaRqvd97x7712aK2y3UDgV8sf72R7RxTFRSzGKuY/export?format=xlsx"
dest_path = "見出語・用例リスト_改訂版.xlsx"

print("Downloading LEAP Revision spreadsheet...")
try:
    req = urllib.request.Request(
        url, 
        headers={'User-Agent': 'Mozilla/5.0'}
    )
    with urllib.request.urlopen(req) as response:
        content = response.read()
    
    with open(dest_path, "wb") as f:
        f.write(content)
        
    print(f"Downloaded and saved to {dest_path} successfully!")
    
    xls = pd.ExcelFile(dest_path)
    print("Sheets available:", xls.sheet_names)
    
    for name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=name)
        print(f"\nSheet [{name}]:")
        print("Columns:", df.columns.tolist())
        print("Shape:", df.shape)
        # Check unique weeks
        week_col = [c for c in df.columns if 'Week' in c]
        if week_col:
            col = week_col[0]
            # Fill forward to get accurate counts
            weeks = df[col].ffill().unique()
            print(f"Unique values in {col} ({len(weeks)}):", list(weeks)[:10], "...")
        else:
            print("No Week column found.")
except Exception as e:
    print(f"Error: {e}")
