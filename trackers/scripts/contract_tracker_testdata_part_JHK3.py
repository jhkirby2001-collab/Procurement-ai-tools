import datetime as _d
_pt = wb["Portfolio Tracker"]; _to = wb["Task Order Tracker"]
_rows = {str(_pt[f"A{r}"].value): r for r in range(2, 88)}
for po, t, w in [("271670", "Master Service", 500000), ("132457", "Standard", None), ("254233", "Construction (JOC)", 200000), ("180514", "Master Service", 300000)]:
    _pt[f"AR{_rows[po]}"] = t; _pt[f"W{_rows[po]}"] = w
D = lambda *a: _d.datetime(*a)
_rel = [("271670", "Task Order", 300000, 50000, "Issued", D(2026,1,10), D(2027,1,1)),
        ("271670", "Task Order", 150000, None, "In progress", D(2026,6,1), D(2026,11,20)),
        ("271670", "Task Order", 100000, None, "Pending approval", None, None),
        ("271670", "Task Order", 80000, None, "Issued", D(2026,9,1), D(2027,3,1)),
        ("254233", "Job Order Release", 50000, None, "Issued", D(2026,2,1), D(2031,1,1)),
        ("132457", "Task Order", 1000, None, "Issued", D(2026,2,1), D(2027,1,1)),
        ("254233", "Job Order Release", 20000, 25000, "Complete", D(2026,3,1), D(2026,9,1)),
        ("254233", "Release Change Order", -10000, None, "Issued", D(2026,4,1), D(2027,9,1)),
        ("271670", "Task Order", 99999, None, "Cancelled", D(2026,5,1), D(2027,1,1)),
        ("180514", "Task Order", 120000, 60000, "In progress", D(2026,7,15), D(2027,6,30))]
for i, (po, ty, amt, paid, st, iss, end) in enumerate(_rel):
    r = 5 + i
    _to[f"A{r}"] = f"R-{i+1:03}"; _to[f"B{r}"] = _pt[f"A{_rows[po]}"].value; _to[f"F{r}"] = ty; _to[f"G{r}"] = "Sample scope " + str(i+1)
    _to[f"L{r}"] = amt; _to[f"M{r}"] = paid; _to[f"N{r}"] = st; _to[f"I{r}"] = iss; _to[f"K{r}"] = end
