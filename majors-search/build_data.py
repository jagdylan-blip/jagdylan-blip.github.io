"""Convert the UNL 'Majors by College' workbook to data.json.
Usage: python3 build_data.py path/to/workbook.xlsx   (needs openpyxl)
Sheets must be in chronological order; first sheet year is set by FIRST_YEAR."""
import sys, json, re, collections, openpyxl
FIRST_YEAR = 2016
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
years = list(range(FIRST_YEAR, FIRST_YEAR + len(wb.sheetnames)))
COLLEGE = {"Business Administration": "Business"}
SMALL = {"and","of","in","for","the","to","a","an","at","on","&","or"}
ACR = {"ii","iii","iv","uno","unl","it","hr","gis","ba","bs","cfa","stem","tesol","k-12","pk-12"}
def title(s):
    s=re.sub(r"\s*&\s*"," & ",s)
    out=[]
    for i,w in enumerate(re.split(r"(\s+)", s.strip().lower())):
        if not w.strip(): out.append(w); continue
        core=w.strip("()/,-")
        if core in ACR: out.append(w.upper())
        elif i and w in SMALL: out.append(w)
        else: out.append(re.sub(r"(^|[(/\-])([a-z])", lambda m:m.group(1)+m.group(2).upper(), w))
    return "".join(out)
def deg(d):
    d=(d or "").strip()
    return "NA" if d in ("","NDEG","NA") else d
progs = collections.OrderedDict()
for y, ws in zip(years, wb):
    for r in list(ws.iter_rows(values_only=True))[1:]:
        if not r[1] or not str(r[1]).strip(): continue
        try: n = int(float(r[4]))
        except (TypeError, ValueError): continue
        name = title(str(r[1])); d = deg(r[3])
        college = COLLEGE.get(str(r[0]).strip(), str(r[0]).strip())
        p = progs.setdefault((name, d), {"name":name,"deg":d,"college":college,"counts":{}})
        p["counts"][y] = p["counts"].get(y, 0) + n
        p["college"] = college  # latest year wins
out = []
for p in progs.values():
    c = p["college"]
    p["level"] = ("Graduate" if c.startswith("Grad Studies") else "Professional" if c == "Law"
                  else "Non-degree" if p["deg"] == "NA" else "Undergraduate")
    p["counts"] = [p["counts"].get(y) for y in years]
    out.append(p)
out.sort(key=lambda p:(p["name"], p["deg"]))
json.dump({"years":years,"programs":out}, open("data.json","w"), separators=(",",":"))
print(len(out), "programs", years[0], "-", years[-1])
