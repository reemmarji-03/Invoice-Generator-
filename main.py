import pandas as pd
import glob


paths = glob.glob("Invoices/*.xlsx")

for path in paths:
    df = pd.read_excel(path, sheet_name="Sheet 1")
    print(df)