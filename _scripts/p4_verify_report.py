import os
import re

b = r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\fw-comparison-report'
html = open(os.path.join(b, 'fw-comparison-report.html'), encoding='utf-8').read()
refs = re.findall(r'(?:src|href)="\./?((?:_shared|assets)[^"]+)"', html)
print('引用资源检查:')
missing = 0
for r in sorted(set(refs)):
    p = os.path.join(b, r)
    ok = os.path.exists(p)
    missing += (not ok)
    print(('  OK      ' if ok else '  MISSING'), r)
print('HTML 大小:', len(html), '字节 | 缺失:', missing)
# mermaid 与图表容器核对
print('mermaid块:', html.count('class="mermaid"'), '| 图表容器:', len(re.findall(r'id="chart-', html)))
