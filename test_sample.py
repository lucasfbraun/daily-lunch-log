import pandas as pd
from datetime import datetime

data = [
    {'data': '20/11/2025', 'codigo': '123', 'valor': 7.11},
    {'data': '21/11/2025', 'codigo': '123', 'valor': 5.00},
    {'data': '19/11/2025', 'codigo': '456', 'valor': 6.78},
    {'data': '20/11/2025', 'codigo': '456', 'valor': 1.22},
    {'data': '21/11/2025', 'codigo': '789', 'valor': 10.00},
]

df = pd.DataFrame(data)
# Write to Excel
path = 'sample_test.xlsx'
df.to_excel(path, index=False)
print('Wrote', path)
