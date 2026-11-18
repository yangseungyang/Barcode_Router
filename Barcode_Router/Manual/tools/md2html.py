import sys, re, html, os
from PIL import Image

src, out = sys.argv[1], sys.argv[2]
base = os.path.dirname(os.path.abspath(src))
lines = open(src, encoding='utf-8').read().split('\n')
MAXW = 620

def inline(t):
    t = html.escape(t, quote=False)
    t = re.sub(r'`([^`]+)`', r'<code>\1</code>', t)
    t = re.sub(r'\*\*([^*]+)\*\*', r'<b>\1</b>', t)
    return t

o = []
i = 0
in_list = False
def close_list():
    global in_list
    if in_list:
        o.append('</ul>'); in_list = False

while i < len(lines):
    ln = lines[i]
    s = ln.strip()
    if s.startswith('```'):
        close_list()
        i += 1; buf = []
        while i < len(lines) and not lines[i].strip().startswith('```'):
            buf.append(html.escape(lines[i])); i += 1
        o.append('<pre>' + '\n'.join(buf) + '</pre>')
        i += 1; continue
    m = re.match(r'^(#{1,4})\s+(.*)', s)
    if m:
        close_list()
        n = len(m.group(1)); o.append(f'<h{n}>{inline(m.group(2))}</h{n}>'); i += 1; continue
    if s == '---':
        close_list(); i += 1; continue
    m = re.match(r'^!\[([^\]]*)\]\(([^)]+)\)', s)
    if m:
        close_list()
        p = os.path.join(base, m.group(2))
        w, h = Image.open(p).size
        if w > MAXW: h = int(h * MAXW / w); w = MAXW
        o.append(f'<p class=img><img src="{os.path.abspath(p)}" width={w} height={h}></p><p class=cap>{inline(m.group(1))}</p>')
        i += 1; continue
    if s.startswith('|'):
        close_list()
        rows = []
        while i < len(lines) and lines[i].strip().startswith('|'):
            rows.append(lines[i].strip()); i += 1
        cells = [[c.strip() for c in r.strip('|').split('|')] for r in rows]
        t = ['<table>']
        for k, r in enumerate(cells):
            if k == 1 and all(re.fullmatch(r':?-+:?', c) for c in r): continue
            tag = 'th' if k == 0 else 'td'
            t.append('<tr>' + ''.join(f'<{tag}>{inline(c)}</{tag}>' for c in r) + '</tr>')
        t.append('</table>')
        o.append(''.join(t))
        if not any('[[TOC]]' in x for x in o):
            o.append('<p class=tochead>목차</p><p>[[TOC]]</p>')
        continue
    m = re.match(r'^(\s*)[-*]\s+(.*)', ln)
    if m:
        lvl = len(m.group(1)) // 2
        if not in_list:
            o.append('<ul>'); in_list = True
        o.append(f'<li class=l{lvl}>{inline(m.group(2))}</li>'); i += 1; continue
    m = re.match(r'^(\s*)(\d+)\.\s+(.*)', ln)
    if m:
        close_list()
        lvl = len(m.group(1)) // 2
        o.append(f'<p class=num style="margin-left:{12 + lvl*14}pt"><b>{m.group(2)}.</b> {inline(m.group(3))}</p>'); i += 1; continue
    if s == '':
        close_list(); i += 1; continue
    close_list()
    o.append(f'<p>{inline(s)}</p>'); i += 1
close_list()

css = '''
body { font-family: "Malgun Gothic"; font-size: 10pt; }
h1 { font-size: 20pt; color: #1F3864; border-bottom: 2px solid #1F3864; margin: 0 0 8pt 0; }
h2 { page-break-before: always; font-size: 14pt; color: #1F3864; border-bottom: 1px solid #1F3864; margin: 16pt 0 6pt 0; page-break-after: avoid; }
h3 { font-size: 11.5pt; color: #2F5496; margin: 10pt 0 4pt 0; page-break-after: avoid; }
h4 { font-size: 10.5pt; color: #2F5496; margin: 8pt 0 3pt 0; page-break-after: avoid; }
p { margin: 3pt 0; line-height: 140%; }
table { border-collapse: collapse; width: 100%; margin: 4pt 0 6pt 0; }
th { background: #D9E2F3; border: 1px solid #A6A6A6; padding: 2pt 4pt; font-size: 9pt; text-align: left; }
td { border: 1px solid #A6A6A6; padding: 2pt 4pt; font-size: 9pt; vertical-align: top; }
code { font-family: Consolas; font-size: 9pt; background: #F2F2F2; }
pre { font-family: Consolas; font-size: 8.5pt; background: #F2F2F2; border: 1px solid #D0D0D0; padding: 4pt; margin: 4pt 0; }
ul { margin: 2pt 0 2pt 18pt; }
li { margin: 1pt 0; }
li.l1 { margin-left: 18pt; }
p.img { text-align: center; margin: 6pt 0 0 0; }
p.tochead { font-size: 13pt; font-weight: bold; color: #1F3864; margin: 14pt 0 4pt 0; }
p.cap { text-align: center; font-size: 9pt; color: #595959; margin: 2pt 0 8pt 0; }
'''
doc = '<html><head><meta charset="utf-8"><style>' + css + '</style></head><body>' + '\n'.join(o) + '</body></html>'
open(out, 'w', encoding='utf-8').write(doc)
print('ok', len(o))
