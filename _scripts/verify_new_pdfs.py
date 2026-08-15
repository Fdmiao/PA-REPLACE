import pypdf, os

BASE = r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_sources\docs\pa'

for fn in ['pan-os-admin-11-1-zh-cn.pdf', 'pan-os-web-interface-help-11-1-zh-cn.pdf']:
    p = os.path.join(BASE, fn)
    raw = open(p, 'rb').read(2048)
    print(f'===== {fn} | {os.path.getsize(p):,}B =====')
    try:
        r = pypdf.PdfReader(p)
        print('pages:', len(r.pages))
        t1 = (r.pages[0].extract_text() or '').strip().replace('\n', ' ')[:120]
        print('p1:', t1)
        outline = ''
        try:
            o = r.outline
            def walk(items, d=0, out=[]):
                for it in items:
                    if isinstance(it, list):
                        walk(it, d + 1)
                    else:
                        out.append('  ' * d + (it.title or ''))
                return out
            titles = walk(o)[:40]
            outline = '\n'.join(titles)
        except Exception as e:
            outline = f'(outline fail: {type(e).__name__})'
        print('outline head:')
        print(outline[:1800])
    except Exception as e:
        print('BROKEN:', type(e).__name__, str(e)[:150])
    print()
