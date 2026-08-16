# -*- coding: utf-8 -*-
import json, os

base="/Users/个人资料/trx/2026/PA替代/_scripts"
topsec=json.load(open(os.path.join(base,'topsec_pages.json'), encoding='utf-8'))
ts_titles={x['file']: x['title'] for x in json.load(open(os.path.join(base,'manual_index.json'), encoding='utf-8'))}

def ctx(text, kw, w=90):
    i=text.find(kw)
    if i<0: return ''
    return text[max(0,i-w):i+len(kw)+w].replace('\n',' ')

def search(kws, limit=5):
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
        out.append(f'  [{score}] {ts_titles.get(fn,"")} ({fn}): {frag[:220]}')
    return out

queries={
 "1_策略移动克隆": ["策略移动","移动策略","克隆","复制策略","策略复制"],
 "2_审核注释": ["审核注释","注释","策略备注","备注"],
 "3_规则使用点击数": ["命中次数","命中数","规则使用","使用情况","未使用规则"],
 "4_QoS_DSCP": ["DSCP","服务类型","ToS","差分服务","优先级匹配"],
 "5_解密流量镜像": ["解密镜像","端口镜像","镜像","解密流量"],
 "6_NPBroker": ["网络数据包代理","数据包代理","第三方安全","安全链"],
 "7_max_encap": ["封装级数","最大封装","隧道嵌套","嵌套隧道","封装层数"],
 "8_SDWAN_链路质量": ["SD-WAN","SDWAN","链路质量","抖动","丢包率","延迟"],
}

for name, kws in queries.items():
    print(f'===== {name} =====')
    res=search(kws)
    if res:
        for r in res: print(r)
    else:
        print('  (无命中)')
    print()
