import pandas as pd

df = pd.DataFrame({'data':['2025-11-20'],'codigo':[125.0],'Valor Refeição':[7.11]})
df.to_excel('tmp_code_float.xlsx', index=False)
print('wrote tmp_code_float.xlsx')
