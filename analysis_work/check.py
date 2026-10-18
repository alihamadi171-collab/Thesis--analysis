import json,csv,sys,re,collections
d=json.load(open('source_extract.json'))
codes={r['ID'] for r in csv.DictReader(open('codebook.tsv'),delimiter='\t')}
n_by={}
for r in d['prev']:
    p,n=r[0].split('-C'); n_by[p]=max(n_by.get(p,0),int(n))
def parse(p):
    m={}
    for ln in open(f'coding/{p}.txt'):
        ln=ln.strip()
        if not ln: continue
        rng,*cs=ln.split()
        a,b=(rng.split('-')+[rng])[:2]
        for i in range(int(a),int(b)+1):
            if i in m: print(p,'dup',i)
            m[i]=cs
            for c in cs:
                if c not in codes: print(p,'bad code',c,i)
    miss=[i for i in range(1,n_by[p]+1) if i not in m]
    if miss: print(p,'missing',miss[:20],len(miss))
    extra=[i for i in m if i>n_by[p]]
    if extra: print(p,'extra',extra)
    return m
if __name__=='__main__':
    for p in sys.argv[1:]:
        m=parse(p); c=collections.Counter(x for v in m.values() for x in v)
        print(p,len(m),'units; substantive',sum(1 for v in m.values() if set(v)-{'PROC','K6','REFLEX'}))
    fx=list(csv.reader(open('translation_fixes.tsv'),delimiter='\t'))
    ids={r[0] for r in d['prev']}
    for r in fx[1:]:
        if len(r)!=4 or r[0] not in ids: print('bad fix row',r[:2],len(r))
