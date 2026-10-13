# Rebuild data.js from the "BG Fund - Historical data" Excel file.
# Usage: python3 scripts/update_data.py "<file>.xlsx"   (needs openpyxl)
import sys, json, datetime, openpyxl
xl = sys.argv[1]
out = sys.argv[2] if len(sys.argv) > 2 else "data.js"
wb = openpyxl.load_workbook(xl, data_only=True, read_only=True)
def num(v):
    return float(v) if isinstance(v,(int,float)) and not isinstance(v,bool) else None
def rows(ws, start, datecol):
    for r in list(ws.iter_rows(values_only=True))[start:]:
        if len(r)>datecol and isinstance(r[datecol], datetime.datetime): yield r
def rec(d, pairs):
    o = {"d": d.strftime("%Y-%m-%d")}
    for k, v in pairs:
        v = num(v)
        if v is not None: o[k] = v
    return o
perf=[]
for r in rows(wb["Performance attribution"], 4, 1):
    keys=["vs","mc","cb","vt","wa","es","ra","ls","cs","cl","ca","cx","ts","t"]
    o=rec(r[1], [("yr", r[0])]+list(zip(keys, r[2:16])))
    perf.append(o)
exp=[]
groups=["cb","vt","wa","eq","ra","ls","cl","cs","cx","tr","qe","sy","ot"]
for r in rows(wb["Exposure (%)"], 5, 0):
    pairs=[]; c=1
    for i,g in enumerate(groups):
        s_ = (num(r[c+2]) or 0)+(num(r[c+3]) or 0) if i==0 else r[c+2]
        pairs += [(g,r[c]),(g+"_l",r[c+1]),(g+"_s",r[c+2])]
        c += 4 if i==0 else 3
    pairs += [("g",r[42]),("n",r[43]),("l",r[44]),("s",r[45]),("am",r[47])]
    exp.append(rec(r[0], pairs))
ear=[]
for r in rows(wb["Equity at risk"], 4, 0):
    ear.append(rec(r[0], [("te",r[16]),("vt",r[18]),("et",r[19]),("ct",r[20]),("tt",r[21])]))
open(out,"w").write("const DATA = "+json.dumps({"perf":perf,"exp":exp,"ear":ear},separators=(",",":"))+";")
