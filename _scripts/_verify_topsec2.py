# -*- coding: utf-8 -*-
import json, os

base="/Users/个人资料/trx/2026/PA替代/_scripts"
topsec=json.load(open(os.path.join(base,'topsec_pages.json'), encoding='utf-8'))
ts_titles={x['file']: x['title'] for x in json.load(open(os.path.join(base,'manual_index.json'), encoding='utf-8'))}

def ctx(text, kw, w=90):
    i=text.find(kw)
    if i<0: return ''
    return text[max(0,i-w):i+len(kw)+w].replace('\n',' ')

def search(kws, limit=6):
    hits=[]
    for fn, body in topsec.items():
        score=sum(body.count(k) for k in kws if k in body)
        if score:
            hits.append((score, fn, body))
    hits.sort(key=lambda x:-x[0])
    out=[]
    for score, fn, body in hits[:limit]:
        frag=''
        for k in kws:
            if k in body:
                frag=ctx(body,k)
                break
        out.append(f'  [{score}] {ts_titles.get(fn,"")} ({fn}): {frag[:240]}')
    return out

queries={
 "3b_规则命中统计": ["显示统计","匹配次数","命中次数","规则命中","策略命中"],
 "4b_DSCP": ["DSCP","差分服务","IP优先级","ToS"],
 "8b_链路探测": ["链路探测","探测规则","丢包","时延","抖动","链路质量"],
 "2b_策略备注": ["备注","描述","策略描述","规则描述"],
}
for name, kws in queries.items():
    print(f'===== {name} =====')
    res=search(kws)
    if res:
        for r in res: print(r)
    else:
        print('  (无命中)')
    print()
