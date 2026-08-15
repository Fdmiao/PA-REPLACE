import requests, os

BASE = r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_sources'
UA = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36',
      'Accept': 'text/html,application/pdf,*/*'}

targets = [
    # PDF 归档
    ('https://www.commoncriteriaportal.org/files/epfiles/st_vid11482-vr.pdf', BASE + r'\docs\pa\cc-st_vid11482-vr.pdf'),
    ('https://www.paloaltonetworks.com/apps/pan/public/downloadResource?pagePath=/content/pan/en_US/resources/datasheets/pa-3400-series', BASE + r'\docs\pa\pa-3400-series-datasheet.pdf'),
    ('https://docs.paloaltonetworks.com/content/dam/techdocs/en_US/pdf/pan-os/12-1/pan-os-admin-12-1-en.pdf', BASE + r'\docs\pa\pan-os-admin-12-1-en.pdf'),
    # HTML 快照：PA 侧
    ('https://docs.paloaltonetworks.com/ngfw/new-features/by-version/panos/12-1', BASE + r'\web\pa\panos-12-1-new-features.html'),
    ('https://docs.paloaltonetworks.com/ngfw/administration/certifications', BASE + r'\web\pa\panos-certifications.html'),
    ('https://www.paloaltonetworks.com/network-security/next-generation-firewall-hardware', BASE + r'\web\pa\pa-ngfw-hardware.html'),
    # HTML 快照：天融信侧
    ('https://www.topsec.com.cn/products/TopNGFW.html', BASE + r'\web\topsec\topsec-topngfw-product.html'),
    ('https://www.topsec.com.cn/newsx/4441.html', BASE + r'\web\topsec\topsec-news-4441-cert.html'),
]

results = []
for url, path in targets:
    name = os.path.basename(path)
    try:
        r = requests.get(url, headers=UA, timeout=90, allow_redirects=True, verify=True)
        ct = r.headers.get('Content-Type', '')
        ok = r.status_code == 200 and len(r.content) > 1000
        if ok:
            with open(path, 'wb') as f:
                f.write(r.content)
        results.append(f"{'OK ' if ok else 'FAIL'} | {name} | {r.status_code} | {len(r.content):,}B | {ct[:40]}")
    except Exception as e:
        results.append(f"FAIL | {name} | EXC | {type(e).__name__}: {str(e)[:80]}")

with open(BASE + r'\..\_scripts\dl_results.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(results))
print('\n'.join(results))
