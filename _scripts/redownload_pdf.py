import requests, os, pypdf

BASE = r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_sources'
UA = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36'}


def download_check(url, path):
    name = os.path.basename(path)
    try:
        with requests.get(url, headers=UA, timeout=(30, 120), stream=True) as r:
            r.raise_for_status()
            clen = int(r.headers.get('Content-Length', 0))
            chunks = []
            total = 0
            for chunk in r.iter_content(chunk_size=1 << 16):
                chunks.append(chunk)
                total += len(chunk)
        data = b''.join(chunks)
        tail_ok = b'%%EOF' in data[-2048:]
        with open(path, 'wb') as f:
            f.write(data)
        try:
            rd = pypdf.PdfReader(path)
            pages = len(rd.pages)
            parse = 'parse-ok'
        except Exception as e:
            pages = 0
            parse = f'parse-FAIL:{type(e).__name__}'
        print(f'{name}: {total:,}B (CL={clen:,}) tail={tail_ok} pages={pages} {parse}')
    except Exception as e:
        print(f'{name}: DOWNLOAD-FAIL {type(e).__name__}: {str(e)[:100]}')


download_check('https://docs.paloaltonetworks.com/content/dam/techdocs/zh_CN/pdf/pan-os/11-0/pan-os-admin-11-0-zh-cn.pdf',
               BASE + r'\docs\pa\pan-os-admin-11-0-zh-cn.pdf')

# 复验其余三份 PDF
for f in ['docs/pa/cc-st_vid11482-vr.pdf', 'docs/pa/pa-3400-series-datasheet.pdf', 'docs/topsec_ngfw_product_doc.pdf']:
    p = os.path.join(BASE, f)
    raw = open(p, 'rb').read()
    tail_ok = b'%%EOF' in raw[-2048:]
    try:
        rd = pypdf.PdfReader(p)
        print(f'{os.path.basename(f)}: {len(raw):,}B tail={tail_ok} pages={len(rd.pages)} parse-ok')
    except Exception as e:
        print(f'{os.path.basename(f)}: {len(raw):,}B tail={tail_ok} parse-FAIL {type(e).__name__}')
