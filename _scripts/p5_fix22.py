import re

BASE = r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求'
html = open(BASE + r'\fw-comparison-report\fw-comparison-report.html', encoding='utf-8').read()
plan = open(BASE + r'\对比执行计划.md', encoding='utf-8').read()

for name, text in [('报告HTML', html), ('计划MD', plan)]:
    print(f'=== {name} ===')
    for m in re.finditer(r'[^>。；\n]{0,30}22\s*条[^<。；\n]{0,25}', text):
        print('  …' + m.group().strip() + '…')
