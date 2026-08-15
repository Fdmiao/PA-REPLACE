import pypdf, os

BASE = r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_sources\docs\pa'

def walk(items, d=0, out=None):
    if out is None:
        out = []
    for it in items:
        if isinstance(it, list):
            walk(it, d + 1, out)
        else:
            try:
                pg = None
                try:
                    pg = it.page.indirect_reference.idnum
                except Exception:
                    pass
                out.append('  ' * d + (it.title or '').strip())
            except Exception:
                pass
    return out

for fn, outname in [('pan-os-web-interface-help-11-1-zh-cn.pdf', 'pa_catalog_webhelp.txt'),
                    ('pan-os-admin-11-1-zh-cn.pdf', 'pa_catalog_admin.txt')]:
    p = os.path.join(BASE, fn)
    r = pypdf.PdfReader(p)
    lines = walk(r.outline)
    with open(r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_scripts' + '\\' + outname, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print(outname, 'entries:', len(lines))
