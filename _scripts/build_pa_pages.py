import json, sys
import pypdf
import os

BASE = r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_scripts'

# PA PDF 页文本索引：{doc, page_no, text, bookmark}——bookmark 取该页所属书签路径（书签级别<=2）
docs = [
    ('webhelp', r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_sources\docs\pa\pan-os-web-interface-help-11-1-zh-cn.pdf'),
    ('admin', r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_sources\docs\pa\pan-os-admin-11-1-zh-cn.pdf'),
]

out_path = os.path.join(BASE, 'pa_pages.json')
result = {}

for name, path in docs:
    r = pypdf.PdfReader(path)
    n = len(r.pages)

    # 书签→页码映射（含层级路径）
    marks = []

    def walk(items, prefix=''):
        for it in items:
            if isinstance(it, list):
                walk(it, prefix)
            else:
                try:
                    pno = r.get_destination_page_number(it)
                    marks.append((pno, prefix + ' ' + (it.title or '').strip()))
                except Exception:
                    pass

    walk(r.outline)
    marks.sort()
    # 为每页找最近的前置书签（level 任意，取路径最长者）
    page_mark = {}
    idx = 0
    cur = ''
    for pno in range(n):
        while idx < len(marks) and marks[idx][0] <= pno:
            cur = marks[idx][1]
            idx += 1
        page_mark[pno] = cur

    texts = {}
    for pno in range(n):
        try:
            t = (r.pages[pno].extract_text() or '').strip()
        except Exception:
            t = ''
        if t:
            texts[pno] = {'m': page_mark[pno], 't': t}
    result[name] = {'n': n, 'pages': texts}
    print(name, 'pages:', n, 'with text:', len(texts))

with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False)
print('saved:', out_path)
