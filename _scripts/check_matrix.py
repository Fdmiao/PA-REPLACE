import pandas as pd

BASE = r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求'
df = pd.read_excel(BASE + r'\防火墙功能对比矩阵_空表.xlsx', sheet_name='对比矩阵', header=None)
print('shape:', df.shape)
print('non-empty rows:', df.dropna(how='all').shape)
tail = df.tail(10).to_string()
with open(BASE + r'\_scripts\matrix_tail.txt', 'w', encoding='utf-8') as f:
    f.write(tail)
