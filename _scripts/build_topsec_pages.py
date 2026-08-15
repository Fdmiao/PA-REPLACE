import os, re, json, html as h

SRC = r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_sources\topsec_manual'
OUT = r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_scripts\topsec_pages.json'


def read_gb(path):
    raw = open(path, 'rb').read()
    for enc in ('gb18030', 'utf-8'):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode('gb18030', errors='replace')


def strip_tags(s):
    s = re.sub(r'(?s)<head.*?</head>', '', s, flags=re.I)
    s = re.sub(r'(?s)<[^>]+>', ' ', s)
    s = h.unescape(s)
    return re.sub(r'\s+', ' ', s).strip()


pages = {}
for fn in sorted(os.listdir(SRC)):
    if not fn.lower().endswith('.html'):
        continue
    body = strip_tags(read_gb(os.path.join(SRC, fn)))
    if body:
        pages[fn] = body

with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(pages, f, ensure_ascii=False)
print('topsec pages indexed:', len(pages))
