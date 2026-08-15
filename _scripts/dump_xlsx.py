import pandas as pd
import json

base = r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求'

out = {}

f1 = base + r'\PA替代需求列表.xlsx'
df = pd.read_excel(f1, header=None)
rows = []
for i, r in df.iterrows():
    vals = []
    for v in r.tolist():
        if pd.isna(v):
            vals.append('')
        else:
            vals.append(str(v).replace('\n', ' / '))
    rows.append([i + 1] + vals)
out['PA替代需求列表'] = {'sheets': ['工作表1'], 'rows': rows}

f2 = base + r'\防火墙功能对比矩阵_空表.xlsx'
xl = pd.ExcelFile(f2)
info = []
for s in xl.sheet_names:
    d = pd.read_excel(f2, sheet_name=s, header=None)
    srows = []
    for i, r in d.iterrows():
        vals = []
        for v in r.tolist():
            if pd.isna(v):
                vals.append('')
            else:
                vals.append(str(v).replace('\n', ' / '))
        srows.append([i + 1] + vals)
    info.append({'name': s, 'shape': list(d.shape), 'rows': srows})
out['防火墙功能对比矩阵_空表'] = info

with open(base + r'\_scripts\xlsx_dump.json', 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, indent=1)
print('done')
