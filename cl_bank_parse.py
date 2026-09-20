#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""CL-BANK-{회차}.tsv -> 관세법 QDATA(회차-문번 키). 여러 회차 병합.
열: 자산ID|회차|교시|문항|단위유형|기호|문두|지문|문두방향|정답(최종)|정오판정|원문행
조합형: ㄱㄴㄷ 보기 -> passage / 선택지 ①~⑤ -> opts.  ox=선지별 O/X(선지·빈칸형).
사용: py cl_bank_parse.py   (src_cl 및 하위폴더 전부 병합)
"""
import sys, re, json, csv, glob, os

CIRC = {'①':1,'②':2,'③':3,'④':4,'⑤':5}
SRC = "src_cl"
OUT = "cl_qdata.json"

def parse_file(path, rnd):
    q = {}
    with open(path, encoding='utf-8') as f:
        rd = csv.reader(f, delimiter='\t')
        next(rd, None)  # 헤더
        for r in rd:
            if len(r) < 11:
                continue
            num  = r[3].strip()
            unit = r[4].strip()
            sym  = r[5].strip()
            moon = r[6].strip()
            jimun= r[7].strip()
            ans_raw = r[9].strip()
            ox = r[10].strip()
            key = f"{rnd}-{num}"
            e = q.setdefault(key, {'q':'', 'bogi':[], 'opts':{}, 'ans':None, 'ox':{}})
            if moon and not e['q']:
                e['q'] = moon
            if ans_raw and e['ans'] is None:
                e['ans'] = ans_raw
            if sym in CIRC:
                # 선지형·빈칸형의 ①~⑤, 조합(선택지)의 ①~⑤ 모두 여기로 -> opts
                e['opts'][CIRC[sym]] = jimun
                if ox in ('O', 'X'):
                    e['ox'][CIRC[sym]] = ox
            elif re.match(r'^[ㄱ-ㅎ]$', sym):
                # 조합형 보기 ㄱㄴㄷ -> passage
                e['bogi'].append(f"{sym}. {jimun}")
    out = {}
    for key, e in q.items():
        opts = [e['opts'].get(i, '') for i in range(1, 6)]
        ans = None
        if e['ans']:
            m = [CIRC[c] for c in e['ans'] if c in CIRC]
            if m:
                ans = m[0] if len(m) == 1 else sorted(set(m))
        out[key] = {'q': e['q'], 'passage': '\n'.join(e['bogi']),
                    'opts': opts, 'ans': ans, 'ox': e['ox']}
    return out

def rnd_from_name(fn):
    m = re.search(r'CL-BANK-(\d+)', fn)
    return m.group(1) if m else None

if __name__ == '__main__':
    files = sorted(set(glob.glob(os.path.join(SRC, "**", "*.tsv"), recursive=True)))
    merged, rows = {}, []
    for path in files:
        rnd = rnd_from_name(os.path.basename(path))
        if not rnd:
            continue
        d = parse_file(path, rnd)
        merged.update(d)
        bad   = [k for k in d if sum(1 for o in d[k]['opts'] if o) != 5]
        noans = [k for k in d if d[k]['ans'] is None]
        rows.append((int(rnd), len(d), bad, noans))
    print("회차  문항  선지5X            정답없음")
    for rnd, n, bad, noans in sorted(rows):
        print(f"{rnd}   {n}   {('OK' if not bad else bad)}   {('OK' if not noans else noans)}")
    json.dump(merged, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f"\n저장: {OUT}  총 {len(merged)} 문항 ({len(rows)}개 회차)")
