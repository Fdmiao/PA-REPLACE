import re

h = open(r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\fw-comparison-workflow.html', encoding='utf-8').read()
text = re.sub(r'<[^>]+>', ' ', h)
text = re.sub(r'\s+', ' ', text)

out = []
i = text.find('Phase 6')
while i != -1:
    out.append(f'===== Phase 6 @ {i} =====\n' + text[i:i + 2500])
    i = text.find('Phase 6', i + 2500)

# 红线相关
j = text.find('红线')
seen = []
while j != -1 and len(seen) < 6:
    seg = text[max(0, j - 150):j + 600]
    if not any(abs(j - s) < 300 for s in seen):
        seen.append(j)
        out.append(f'===== 红线 @ {j} =====\n' + seg)
    j = text.find('红线', j + 600)

with open(r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_scripts\workflow_p6.txt', 'w', encoding='utf-8') as f:
    f.write('\n\n'.join(out))
print('segments:', len(out))
