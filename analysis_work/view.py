import json,sys
d=json.load(open('analysis_work/source_extract.json'))
p=sys.argv[1]; a=int(sys.argv[2]) if len(sys.argv)>2 else 1; b=int(sys.argv[3]) if len(sys.argv)>3 else 9999
for r in d['prev']:
    pid,n=r[0].split('-C'); n=int(n)
    if pid==p and a<=n<=b: print(f"{n}|{r[1]}|{r[3]}")
