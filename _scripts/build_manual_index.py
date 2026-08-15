import os, re, json, html as htmllib

SRC = r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_sources\topsec_manual'
OUT = r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_scripts\manual_index.json'

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
    s = htmllib.unescape(s)
    return re.sub(r'\s+', ' ', s).strip()

index = []
for fn in sorted(os.listdir(SRC)):
    if not fn.lower().endswith('.html'):
        continue
    p = os.path.join(SRC, fn)
    try:
        doc = read_gb(p)
    except Exception as e:
        continue
    m = re.search(r'<title>(.*?)</title>', doc, re.S | re.I)
    title = strip_tags(m.group(1)) if m else ''
    md = re.search(r'name="description"\s+content="(.*?)"', doc, re.S | re.I)
    desc = strip_tags(htmllib.unescape(md.group(1)))[:200] if md else ''
    body = strip_tags(re.sub(r'(?s)<head.*?</head>', '', doc, flags=re.I))
    index.append({'file': fn, 'title': title, 'desc': desc, 'len': len(body)})

with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(index, f, ensure_ascii=False, indent=1)

print('pages:', len(index))
print('with title:', sum(1 for x in index if x['title']))
