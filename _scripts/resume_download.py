import requests, os, time, pypdf

URL = 'https://docs.paloaltonetworks.com/content/dam/techdocs/zh_CN/pdf/pan-os/11-0/pan-os-admin-11-0-zh-cn.pdf'
PATH = r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_sources\docs\pa\pan-os-admin-11-0-zh-cn.pdf'
UA = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36'}

tmp = PATH + '.part'
max_rounds = 30
for round_no in range(1, max_rounds + 1):
    have = os.path.getsize(tmp) if os.path.exists(tmp) else 0
    headers = dict(UA)
    if have:
        headers['Range'] = f'bytes={have}-'
    try:
        with requests.get(URL, headers=headers, timeout=(20, 60), stream=True) as r:
            if r.status_code == 416:
                break
            r.raise_for_status()
            clen = int(r.headers.get('Content-Length', 0))
            if r.status_code == 200 and have:
                have = 0
                mode = 'wb'
            else:
                mode = 'ab'
            with open(tmp, mode) as f:
                for chunk in r.iter_content(chunk_size=1 << 15):
                    f.write(chunk)
        size = os.path.getsize(tmp)
        print(f'round {round_no}: now {size:,}B (this CL={clen:,})')
        if clen and size >= have + clen:
            break
        if not clen and size == have:
            break
    except Exception as e:
        print(f'round {round_no}: {type(e).__name__}, retrying, have={os.path.getsize(tmp) if os.path.exists(tmp) else 0:,}B')
        time.sleep(2)

raw = open(tmp, 'rb').read()
ok = b'%%EOF' in raw[-2048:]
print('size:', f'{len(raw):,}B', '| tail EOF:', ok)
if ok:
    try:
        rd = pypdf.PdfReader(tmp)
        print('pages:', len(rd.pages))
        t = (rd.pages[0].extract_text() or '')[:80].replace('\n', ' ')
        print('page1:', t)
        os.replace(tmp, PATH)
        print('SAVED ->', PATH)
    except Exception as e:
        print('still broken:', type(e).__name__, str(e)[:150])
else:
    print('incomplete, keep .part for next retry')
