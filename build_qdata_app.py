import re, json, sys, glob, os
from docx import Document

SRC = "src_docx"
OUT = "qdata_customs.json"
CIRC = {'①':1,'②':2,'③':3,'④':4,'⑤':5}
CSET = set(CIRC)

# 회차별 A형 정답(41~80). 복수정답은 리스트. 클로드 검증분 — 수정 금지.
EXPECTED = {
 28:[1,3,2,3,4,1,3,1,4,2,2,5,1,5,2,3,3,1,5,3,5,4,5,5,4,2,4,5,2,4,4,1,4,5,2,1,4,3,2,4],
 29:[2,5,2,4,4,3,4,3,3,2,2,4,3,1,1,1,1,2,5,5,3,4,1,3,4,5,5,5,5,1,4,5,2,4,1,5,3,1,2,3],
 30:[3,3,4,5,1,5,3,5,3,5,1,1,1,4,[1,2,5],4,2,2,1,2,5,2,4,4,3,4,1,5,3,2,4,5,4,1,2,2,3,5,3,4],
 31:[4,1,2,3,2,4,2,4,1,5,5,2,3,4,3,4,2,4,1,1,2,5,2,1,5,4,4,1,3,3,5,5,3,2,3,4,3,1,5,1],
 32:[1,3,1,4,1,4,2,5,1,5,5,3,4,2,2,3,3,2,2,1,5,1,1,4,2,3,5,4,3,4,4,5,5,2,4,4,2,3,3,5],
 33:[3,2,1,4,2,1,3,4,4,5,5,4,5,5,3,3,1,5,5,2,1,3,3,2,3,4,1,4,4,2,2,2,1,5,3,4,5,4,1,2],
 34:[2,4,1,2,1,4,2,4,1,5,3,1,3,4,5,4,2,1,5,5,4,2,3,1,3,5,4,4,2,5,5,5,4,2,3,1,3,1,3,3],
 35:[5,1,1,4,1,5,1,5,4,2,3,3,3,2,2,2,5,4,2,2,4,2,5,5,5,3,1,5,1,1,3,4,4,1,2,5,1,3,2,4],
 36:[3,2,2,5,1,2,2,2,1,4,5,2,4,5,1,3,5,4,2,5,4,3,1,1,4,3,5,1,2,5,5,3,5,4,4,3,4,1,3,5],
 37:[3,4,2,4,4,4,5,2,4,3,5,4,2,1,5,4,4,4,4,5,1,3,3,2,1,5,3,4,3,3,1,5,3,3,5,4,5,4,4,2],
 38:[2,1,1,3,3,2,3,1,1,3,4,1,4,1,4,2,4,4,3,5,2,5,5,4,3,5,1,5,2,5,4,5,1,3,2,4,1,3,5,2],
 39:[2,1,2,5,3,4,1,3,5,5,1,3,5,4,2,3,4,3,2,5,4,4,5,1,3,4,4,2,2,5,4,3,5,1,1,2,4,3,2,1],
 40:[3,5,4,1,3,2,1,5,3,4,1,2,4,5,2,4,1,3,5,4,2,3,4,4,3,3,4,5,1,1,1,5,2,5,5,2,2,3,1,2],
 41:[2,3,4,3,1,3,1,5,2,2,1,1,3,4,4,1,5,3,5,5,3,1,5,1,2,4,5,5,4,5,5,3,4,4,2,4,2,3,2,2],
 42:[1,5,5,5,1,4,3,2,1,5,3,1,4,5,5,3,2,3,4,4,2,5,4,5,2,3,2,2,3,4,1,2,4,3,1,5,2,5,1,1],
 43:[4,5,1,3,5,1,3,2,5,1,4,3,1,4,1,2,5,5,1,2,3,4,2,4,2,1,3,2,2,3,3,5,3,4,3,4,5,5,2,5],
}

def paras(path):
    return [p.text.strip() for p in Document(path).paragraphs if p.text.strip()]

def core_from(lines):
    hdr = re.compile(r'^(?:문\s*)?(\d{2})\\?\.\s*(.*)$')
    idx = [(i, int(m.group(1)), m.group(2).strip())
           for i, l in enumerate(lines)
           for m in [hdr.match(l)] if m and 41 <= int(m.group(1)) <= 80]
    idx.append((len(lines), None, None))
    core = {}
    for k in range(len(idx)-1):
        s, num, stem = idx[k]; e = idx[k+1][0]
        opts, pas, ans = [], [], None
        for l in lines[s+1:e]:
            t = l.strip()
            if t.replace('▸', '').strip().startswith('정답'):
                got = [CIRC[c] for c in t if c in CIRC]
                if got:
                    ans = got[0] if len(got) == 1 else sorted(got)
                continue
            if t and t[0] in CSET:
                opts.append(t)
            elif not opts:
                pas.append(t)
        core[num] = {'q': stem, 'passage': '\n'.join(pas), 'opts': opts, 'ans': ans}
    return core

def why_from(path):
    lines = paras(path)
    hdr = re.compile(r'^(?:문\s*)?(\d{2})\\?\.\s')
    idx = [(i, int(hdr.match(l).group(1))) for i, l in enumerate(lines)
           if hdr.match(l) and 41 <= int(hdr.match(l).group(1)) <= 80]
    idx.append((len(lines), None))
    why = {}
    for k in range(len(idx)-1):
        s, num = idx[k]; e = idx[k+1][0]; buf, on = [], False
        for l in lines[s:e]:
            t = l.strip()
            if ('오답' in t) and ('이유' in t):
                on = True
                r = re.sub(r'^[▸\s]*오답\s*이유[:\s]*', '', t).strip()
                if r:
                    buf.append(r)
                continue
            if on:
                buf.append(t)
        why[num] = '\n'.join(buf).strip()
    return why

def find(rnd, ohap):
    for f in sorted(glob.glob(os.path.join(SRC, "*.docx"))):
        n = os.path.basename(f)
        is_o = ('오답이유' in n) or ('HAESUL' in n)
        if is_o != ohap:
            continue
        if ('%d회' % rnd) in n or ('HAESUL-%d' % rnd) in n:
            return f
    return None

qdata = {}; rows = []; allok = True
for rnd in range(28, 44):
    ft = find(rnd, False)
    if not ft:
        continue                      # docx 없으면 건너뜀(31·38 등)
    core = core_from(paras(ft))
    op = find(rnd, True)
    why = why_from(op) if op else {}
    got = [core.get(n, {}).get('ans') for n in range(41, 81)]
    opts5 = all(len(core.get(n, {}).get('opts', [])) == 5 for n in range(41, 81))
    amatch = (got == EXPECTED[rnd])
    if not (amatch and opts5 and len(core) == 40):
        allok = False
    wc = sum(1 for n in range(41, 81) if why.get(n))
    for n in range(41, 81):
        c = core.get(n, {})
        qdata['%d-%d' % (rnd, n)] = {'q': c.get('q', ''), 'passage': c.get('passage', ''),
                                     'opts': c.get('opts', []), 'ans': c.get('ans'), 'why': why.get(n, '')}
    rows.append((rnd, len(core), opts5, amatch, wc, got))

print("회차 문항 선지5 정답assert why채움")
for rnd, cnt, o5, amatch, wc, got in rows:
    print(rnd, cnt, ("OK" if o5 else "X"), ("일치" if amatch else "불일치"), "why=%d/40" % wc)
    if not amatch:
        print("  불일치:", [(41+i, got[i], EXPECTED[rnd][i]) for i in range(40) if got[i] != EXPECTED[rnd][i]])

if allok and rows:
    json.dump(qdata, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print("\nOK -> %s 저장, %d개 회차 %d문항" % (OUT, len(rows), len(qdata)))
else:
    print("\n[중단] assert 실패 또는 입력없음 -> 저장 안 함. index.html 건드리지 말 것.")
    sys.exit(1)
