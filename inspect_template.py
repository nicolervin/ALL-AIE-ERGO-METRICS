import pandas as pd
from pathlib import Path
p = Path('AIE_Ergo_Dashboard_Template NICOLES VERSION.xlsx')
print('exists', p.exists())
xl = pd.ExcelFile(p)
print('SHEETS', xl.sheet_names)
for s in xl.sheet_names:
    df = pd.read_excel(p, sheet_name=s)
    print('\nSHEET', s)
    print(df.head().to_string())
    print('COLUMNS', list(df.columns))
