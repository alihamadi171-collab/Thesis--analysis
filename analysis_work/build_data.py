# Assemble final unit-level data (source text + revised codes + verified/corrected translations + themes)
import json,csv,sys
sys.path.insert(0,'.')
from check import parse
d=json.load(open('source_extract.json'))
cb={r['ID']:r for r in csv.DictReader(open('codebook.tsv'),delimiter='\t')}
th=json.load(open('themes.json'))
c2s={}
for t in th['themes']:
    for s in t['subs']:
        for c in s['codes']: c2s[c]=s
fx={r['id']:r for r in csv.DictReader(open('translation_fixes.tsv'),delimiter='\t')}
M={f'P{i:02d}':parse(f'P{i:02d}') for i in range(1,11)}
out=[]
for r in d['prev']:
    pid,n=r[0].split('-C'); n=int(n); e='E'+str(int(pid[1:]))
    codes=M[pid][n]
    f=fx.get(r[0])
    subs=[]
    for c in codes:
        s=c2s.get(c)
        if s and s['id'] not in [x['id'] for x in subs]: subs.append(s)
    subs.sort(key=lambda s:[int(x) for x in s['id'].split('.')])
    themes=[f"{s['id']} {s['en']} / {s['ar']}" for s in subs]
    if not themes:
        nt=[k for k in ['K6','CONTEXT','PROC'] if k in codes]
        themes=[th['nonthematic'][nt[0]]] if nt else []
    if 'REFLEX' in codes: themes.append(th['nonthematic']['REFLEX'])
    out.append(dict(uid=f"{e}-{n:03d}",old=r[0],interview=e,n=n,ar=r[1],codes=codes,
        coding=[f"{c} — {cb[c]['Arabic label']}" for c in codes],
        en=f['english'] if f else r[3], back=f['back_translation'] if f else r[4],
        corrected=bool(f), note=f['note'] if f else '', themes=themes))
json.dump(out,open('final_units.json','w'),ensure_ascii=False)
print(len(out),'units;',sum(o['corrected'] for o in out),'corrected translations;',
      sum(1 for o in out if set(o['codes'])-{'PROC','K6','REFLEX','CONTEXT'}),'substantive')
