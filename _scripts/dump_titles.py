import json

idx = json.load(open(r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_scripts\manual_index.json', encoding='utf-8'))
with open(r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_scripts\manual_titles.txt', 'w', encoding='utf-8') as f:
    for x in idx:
        f.write(f"{x['file']} | {x['title']}\n")
print('written', len(idx))
