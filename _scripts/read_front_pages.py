import os, re, html as h

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
    s = re.sub(r'(?s)<head.*?</head>', '', s, flags=re.I)
    s = re.sub(r'(?s)<[^>]+>', ' ', s)
    s = h.unescape(s)
    return re.sub(r'\s+', ' ', s).strip()

out = []
for fn in ['ref487725245.html', 'topic125.html', 'toc490041559.html', 'toc490041566.html', 'faq.html', 'ngfw.html', 'toc462736546.html']:
    p = os.path.join(SRC, fn)
    if os.path.exists(p):
        out.append(f'===== {fn} =====')
        out.append(strip_tags(read_gb(p))[:2500])
        out.append('')

with open(r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_scripts\manual_front_pages.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(out))
print('done')
