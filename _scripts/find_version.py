import os, re, json

SRC = r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_sources\topsec_manual'

def read_gb(path):
    with open(path, 'rb') as f:
        raw = f.read()
    for enc in ('gb18030', 'utf-8'):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode('gb18030', errors='replace')

def strip_tags(s):
    s = re.sub(r'(?s)<[^>]+>', ' ', s)
    import html as h
    s = h.unescape(s)
    return re.sub(r'\s+', ' ', s).strip()

patterns = [
    r'(?:版本|version|Version|VERSION)\s*[:：]?\s*([A-Z]*\s*[\d]+(?:\.[\d]+){1,3}[A-Za-z0-9\-]*)',
    r'(?:NGFW|TFOS|TopOS|OS)\s*[\s:：]?\s*([Vv]?\d+(?:\.\d+){1,3})',
    r'([Vv]\d{1,2}(?:\.\d+){1,3})\s*(?:版本|version)',
    r'TFOS-[\w\.]+',
    r'NGFW-?OS\s*[\d\.]+',
]

hits = {}
for fn in sorted(os.listdir(SRC)):
    if not fn.lower().endswith('.html'):
        continue
    doc = read_gb(os.path.join(SRC, fn))
    body = strip_tags(re.sub(r'(?s)<head.*?</head>', '', doc, flags=re.I))
    for pat in patterns:
        for m in re.finditer(pat, body):
            v = m.group(0)[:60]
            ctx = body[max(0, m.start() - 60):m.end() + 60]
            hits.setdefault(v, []).append({'file': fn, 'ctx': ctx})

out = r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_scripts\version_hits.json'
with open(out, 'w', encoding='utf-8') as f:
    json.dump({k: v[:3] for k, v in sorted(hits.items(), key=lambda x: -len(x[1]))}, f, ensure_ascii=False, indent=1)
print('variants:', len(hits))
