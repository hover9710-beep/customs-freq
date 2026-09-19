import re, json, sys, glob, os
from docx import Document

SRC = "src_docx"            # 여기에 docx 7개
OUT = "qdata_customs.json"
CIRC = {'①':1,'②':2,'③':3,'④':4,'⑤':5}
CSET = set(CIRC)

# 클로드 컨테이너 검증 완료 정답(41~80) — assert 기준. 절대 수정 금지.
EXPECTED = {
 40:[3,5,4,1,3,2,1,5,3,4,1,2,4,5,2,4,1,3,5,4,2,3,4,4,3,3,4,5,1,1,1,5,2,5,5,2,2,3,1,2],
 41:[2,3,4,3,1,3,1,5,2,2,1,1,3,4,4,1,5,3,5,5,3,1,5,1,2,4,5,5,4,5,5,3,4,4,2,4,2,3,2,2],
 42:[1,5,5,5,1,4,3,2,1,5,3,1,4,5,5,3,2,3,4,4,2,5,4,5,2,3,2,2,3,4,1,2,4,3,1,5,2,5,1,1],
 43:[4,5,1,3,5,1,3,2,5,1,4,3,1,4,1,2,5,5,1,2,3,4,2,4,2,1,3,2,2,3,3,5,3,4,3,4,5,5,2,5],
}

def paras(path):
    return [p.text.strip() for p in Document(path).paragraphs if p.text.strip()]

def find(rnd, ohap):
    pats = ["%d회" % rnd, "HAESUL-%d" % rnd]
    for f in sorted(glob.glob(os.path.join(SRC, "*.docx"))):
        n = os.path.basename(f)
        is_ohap = ('오답이유' in n) or ('HAESUL' in n)
        if is_ohap != ohap:
            continue
        if any(p in n for p in pats):
            return f
    return None

def parse_core(lines):
    hdr = re.compile(r'^(\d{2})\\?\.\s*(.*)$')
    idx = [(i, int(m.group(1)), m.group(2).strip())
           for i, l in enumerate(lines)
           for m in [hdr.match(l)] if m and 41 <= int(m.group(1)) <= 80]
    idx.append((len(lines), None, None))
    core = {}
    for k in range(len(idx) - 1):
        s, num, stem = idx[k]; e = idx[k+1][0]
        opts, pas, ans = [], [], None
        for l in lines[s+1:e]:
            t = l.strip()
            if t.replace('▸', '').strip().startswith('정답'):
                for c in t:
                    if c in CIRC:
                        ans = CIRC[c]; break
                continue
            if t and t[0] in CSET:
                opts.append(t)
            elif not opts:
                pas.append(t)
        core[num] = {'q': stem, 'passage': '\n'.join(pas), 'opts': opts, 'ans': ans}
    return core

def parse_why(path):
    lines = paras(path)
    hdr = re.compile(r'^문?(\d{2})\\?\.\s')   # 40회 오답이유판은 '문41. …' 형식
    idx = [(i, int(hdr.match(l).group(1))) for i, l in enumerate(lines)
           if hdr.match(l) and 41 <= int(hdr.match(l).group(1)) <= 80]
    idx.append((len(lines), None))
    why = {}
    for k in range(len(idx) - 1):
        s, num = idx[k]; e = idx[k+1][0]
        buf, on = [], False
        for l in lines[s:e]:
            t = l.strip()
            if ('오답' in t) and ('이유' in t):
                on = True
                rest = re.sub(r'^[\u25b8\s]*오답\s*이유[:\s]*', '', t).strip()
                if rest:
                    buf.append(rest)
                continue
            if on:
                buf.append(t)
        why[num] = '\n'.join(buf).strip()
    return why

qdata = {}
rows = []
allok = True
for rnd in [40, 41, 42, 43]:
    ft = find(rnd, False)
    if not ft:
        print("FULLTEXT %d 없음" % rnd); sys.exit(1)
    core = parse_core(paras(ft))
    op = find(rnd, True)
    why = parse_why(op) if op else {}
    got = [core.get(n, {}).get('ans') for n in range(41, 81)]
    opts5 = all(len(core.get(n, {}).get('opts', [])) == 5 for n in range(41, 81))
    ok = (got == EXPECTED[rnd]) and opts5 and len(core) == 40
    if not ok:
        allok = False
    wcount = sum(1 for n in range(41, 81) if why.get(n))
    for n in range(41, 81):
        c = core.get(n, {})
        qdata["%d-%d" % (rnd, n)] = {
            'q': c.get('q', ''), 'passage': c.get('passage', ''),
            'opts': c.get('opts', []), 'ans': c.get('ans'), 'why': why.get(n, '')
        }
    rows.append((rnd, len(core), opts5, ok, wcount, got))

print("회차 문항 선지5 정답assert why채움")
for rnd, cnt, o5, ok, wc, got in rows:
    print(rnd, cnt, ("OK" if o5 else "X"),
          ("일치" if got == EXPECTED[rnd] else "불일치"), "why=%d/40" % wc)
    if got != EXPECTED[rnd]:
        print("  불일치:", [(41+i, got[i], EXPECTED[rnd][i]) for i in range(40) if got[i] != EXPECTED[rnd][i]])

if allok:
    json.dump(qdata, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print("\nOK -> %s 저장, 총 %d 문항" % (OUT, len(qdata)))
else:
    print("\n[중단] 정답 assert 실패 -> qdata_customs.json 저장 안 함. index.html 건드리지 말 것.")
    sys.exit(1)
