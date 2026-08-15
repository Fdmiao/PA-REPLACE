import json, re, sys, os

BASE = r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_scripts'

topsec = json.load(open(os.path.join(BASE, 'topsec_pages.json'), encoding='utf-8'))
ts_titles = {x['file']: x['title'] for x in json.load(open(os.path.join(BASE, 'manual_index.json'), encoding='utf-8'))}
pa = json.load(open(os.path.join(BASE, 'pa_pages.json'), encoding='utf-8'))


def ctx(text, kw, w=90):
    i = text.find(kw)
    if i < 0:
        return ''
    return text[max(0, i - w):i + len(kw) + w].replace('\n', ' ')


def search_topsec(kws, limit=6):
    hits = []
    for fn, body in topsec.items():
        score = sum(body.count(k) for k in kws if k in body)
        if score:
            hits.append((score, fn, body))
    hits.sort(key=lambda x: -x[0])
    out = []
    for score, fn, body in hits[:limit]:
        frag = ''
        for k in kws:
            if k in body:
                frag = ctx(body, k)
                break
        out.append(f'  [{score}] {ts_titles.get(fn, "")} ({fn}): {frag[:200]}')
    return out


def search_pa(kws, limit=6, docs=None):
    out = []
    for doc in (docs or ['webhelp', 'admin']):
        d = pa[doc]
        hits = []
        for pno, info in d['pages'].items():
            t = info['t']
            score = sum(t.count(k) for k in kws if k in t)
            if score:
                hits.append((score, int(pno), info))
        hits.sort(key=lambda x: -x[0])
        for score, pno, info in hits[:limit]:
            frag = ''
            for k in kws:
                if k in info['t']:
                    frag = ctx(info['t'], k)
                    break
            out.append(f'  [{doc} p{pno + 1} score{score}] {info["m"][:60]}: {frag[:180]}')
    return out


def run(query_file, out_file):
    q = json.load(open(query_file, encoding='utf-8'))
    lines = []
    for item in q:
        no, fp, tk, pk = item['no'], item['fp'], item['topsec'], item['pa']
        lines.append(f'### {no} {fp}')
        lines.append('-- 天融信:')
        lines.extend(search_topsec(tk) or ['  (无命中)'])
        lines.append('-- PA:')
        lines.extend(search_pa(pk) or ['  (无命中)'])
        lines.append('')
    with open(out_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print('report written:', out_file, f'({len(q)} items)')


if __name__ == '__main__':
    run(sys.argv[1], sys.argv[2])
