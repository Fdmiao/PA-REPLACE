# -*- coding: utf-8 -*-
import json, os

base="/Users/个人资料/trx/2026/PA替代/_scripts"
topsec=json.load(open(os.path.join(base,'topsec_pages.json'), encoding='utf-8'))
ts_titles={x['file']: x['title'] for x in json.load(open(os.path.join(base,'manual_index.json'), encoding='utf-8'))}
pa=json.load(open(os.path.join(base,'pa_pages.json'), encoding='utf-8'))

def ctx(text, kw, w=80):
    i=text.find(kw)
    if i<0: return ''
    return text[max(0,i-w):i+len(kw)+w].replace('\n',' ')

def search_topsec(kws, limit=3):
    hits=[]
    for fn, body in topsec.items():
        score=sum(body.count(k) for k in kws if k in body)
        if score: hits.append((score, fn, body))
    hits.sort(key=lambda x:-x[0])
    out=[]
    for score, fn, body in hits[:limit]:
        frag=''
        for k in kws:
            if k in body:
                frag=ctx(body,k); break
        out.append(f'    [{score}] {ts_titles.get(fn,"")}({fn}): {frag[:150]}')
    return out

def search_pa(kws, limit=2):
    out=[]
    for doc in ['webhelp','admin']:
        d=pa[doc]
        hits=[]
        for pno, info in d['pages'].items():
            t=info['t']
            score=sum(t.count(k) for k in kws if k in t)
            if score: hits.append((score, int(pno), info))
        hits.sort(key=lambda x:-x[0])
        for score, pno, info in hits[:limit]:
            frag=''
            for k in kws:
                if k in info['t']:
                    frag=ctx(info['t'],k); break
            out.append(f'    [{doc} p{pno+1} s{score}] {info["m"][:40]}: {frag[:140]}')
    return out

# 每项: (行号, 需求, 天融信关键词, PA关键词)
items=[
 ("R2 邮件告警SMTPS", ["SMTPS","邮件告警"], ["SMTPS","邮件"]),
 ("R3 FireMon对接", ["FireMon"], []),
 ("R4 Ansible对接", ["Ansible"], []),
 ("R5 restful api", ["restful","REST API"], []),
 ("R6 日志源IP过滤", ["日志","源IP"], []),
 ("R10 会话开始时间", ["会话","开始时间"], []),
 ("R13 授权平滑升级", ["授权","升级"], []),
 ("R14 Facebook识别", ["Facebook","应用识别"], []),
 ("R23 双因子认证", ["双因子","双因素"], []),
 ("R24 地址组嵌套", ["地址组","嵌套"], []),
 ("R25 代理升级规则库", ["代理","规则库","升级"], []),
 ("R26 配置锁", ["配置锁","锁定"], []),
 ("R27 带外管理syslog", ["带外","管理路由","syslog"], []),
 ("R29 fortinet转换", ["fortinet","配置转换"], []),
 ("R30 PA转换", ["Palo Alto","配置转换"], ["配置转换","转换"]),
 ("R32 RED", ["RED","随机早期检测"], []),
 ("R34 IGMP", ["IGMP","组播"], []),
 ("R35 MLD", ["MLD"], []),
 ("R36 PIM-SM", ["PIM-SM"], []),
 ("R37 PIM-DM", ["PIM-DM"], []),
 ("R38 源目的域名访问控制", ["域名","访问控制"], []),
 ("R39 组播IPS", ["组播","IPS"], []),
 ("R40 域名分类", ["域名分类","URL分类"], []),
 ("R41 HTTP3/QUIC", ["HTTP3","QUIC","WAF"], []),
 ("R42 XFF", ["X-Forwarded","XFF"], []),
 ("R43 云沙箱", ["云沙箱","沙箱"], []),
 ("R44 英文界面", ["英文","English"], []),
 ("R49 PA网络加速", [], ["硬件加速","加速"]),
 ("R50 PA IPS加速", [], ["IPS","加速"]),
 ("R52 PA ipsec加速", [], ["IPSec","加速","硬件"]),
 ("R53 HA配置同步", ["配置同步","HA"], []),
 ("R54 虚系统HA", ["虚系统","虚拟系统"], []),
 ("R55 NAT策略组512", ["NAT","策略组"], []),
 ("R56 25G/40G接口", ["25G","40G"], []),
 ("R57 接口光衰", ["光衰"], []),
 ("R58 IPSEC链路切换", ["IPSEC","链路","切换"], []),
 ("R59 PA解密镜像", [], ["解密","镜像"]),
]

for label, tk, pk in items:
    print(f'===== {label} =====')
    if tk:
        res=search_topsec(tk)
        print('  天融信:', res if res else '(无命中)')
    if pk:
        res=search_pa(pk)
        print('  PA:', res if res else '(无命中)')
    print()
