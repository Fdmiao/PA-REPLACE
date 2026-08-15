import re, html as h

SRC = r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_sources\topsec_manual\NGFW3.7一本通.hhc'
OUT = r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_scripts\topsec_catalog.txt'

raw = open(SRC, 'rb').read()
for enc in ('gb18030', 'utf-8'):
    try:
        doc = raw.decode(enc)
        break
    except UnicodeDecodeError:
        continue

items = re.findall(r'(?is)<param\s+name="(Name|Local)"\s+value="(.*?)"', doc)
lines = []
depth = 0
pending_name = None
seq = []
for kind, val in items:
    val = h.unescape(val).strip()
    if kind == 'Name':
        seq.append(('N', val))
    else:
        seq.append(('L', val))

# hhc structure: [ul] opens level, [/ul] closes. Build tree by tracking ul tags order.
tree_lines = []
level = 0
i = 0
last_name = None
# simpler: walk raw tokens
tokens = re.findall(r'(?is)<param\s+name="Name"\s+value="(.*?)"|<param\s+name="Local"\s+value="(.*?)"|<(ul|/ul)>', doc)
for name, local, tag in tokens:
    if tag == 'ul':
        level += 1
    elif tag == '/ul':
        level -= 1
    elif name:
        last_name = h.unescape(name).strip()
    elif local:
        fn = h.unescape(local).strip().split('/')[-1]
        if last_name:
            tree_lines.append('  ' * max(level - 1, 0) + f'{last_name}  [{fn}]')
            last_name = None

with open(OUT, 'w', encoding='utf-8') as f:
    f.write('\n'.join(tree_lines))
print('catalog entries:', len(tree_lines))
