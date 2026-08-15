from pypdf import PdfReader

P = r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_sources\docs\topsec_ngfw_product_doc.pdf'
r = PdfReader(P)
out = [f'pages: {len(r.pages)}', f'metadata: {r.metadata}', '']
for i in range(min(6, len(r.pages))):
    t = (r.pages[i].extract_text() or '').strip()
    out.append(f'===== page {i + 1} =====')
    out.append(t[:800])
    out.append('')

with open(r'e:\zt\ggz\研发部\3.7项目\需求分析\PA-替代需求\_scripts\topsec_pdf_head.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(out))
print('done')
