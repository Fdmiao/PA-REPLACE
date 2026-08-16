# -*- coding: utf-8 -*-
from pypdf import PdfReader
import re
base="/Users/个人资料/trx/2026/PA替代"
r=PdfReader(base+"/_sources/docs/pa/pan-os-web-interface-help-11-1-zh-cn.pdf")
# 策略章节 113-195 页（PDF页索引 112-194）
print("总页数:", len(r.pages))
# 抽取 112-195 页文本，找目录标题行（带页码）
for pno in range(112, 195):
    t=r.pages[pno].extract_text() or ""
    t=re.sub(r'\s+',' ',t).strip()
    if not t: continue
    # 打印每页首行作为标题线索
    first=t[:90]
    print(f"PDF{pno+1}: {first}")
