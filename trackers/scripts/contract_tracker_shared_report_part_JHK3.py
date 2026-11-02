# Shared contracts on the Dashboard and Summary. Exec'd inside build_contract_tracker_JHK3.py after the
# dashboard and task order parts (uses their globals: db, sm, tile, sh, T0, TOPMAX, SOON, CLOSED_CAT, PT ...).
NSTAT_SH = len(EXP)

# ---------------- Dashboard: section after the Top 10
S0 = T0 + 16
sh(f"B{S0}", "Shared contracts (owned by another department)")
SHR = f'{PT("AT")},"Shared"'
tile("B", S0 + 1, "SHARED CONTRACTS", f'=COUNTIF({PT("AT")},"Shared")', "Owning Dept Code is not yours")
tile("E", S0 + 1, "SHARED - NEED ACTION", f'=COUNTIFS({SHR},{PT("AL")},"<="&{TOPMAX})', "Priority 1-2", alert=("FBE3E3", "C00000"))
tile("H", S0 + 1, "SHARED - LAPSED", f'=COUNTIFS({SHR},{PT("AD")},"<0",{PT("AH")},"<>"&{CLOSED_CAT})', "Past end date, not closed", alert=("FBE3E3", "C00000"))
tile("K", S0 + 1, "SHARED - EXPIRING SOON", f'=COUNTIFS({SHR},{PT("AD")},">=0",{PT("AD")},"<="&{SOON},{PT("AH")},"<>"&{CLOSED_CAT})',
     '="Ending within "&' + SOON + '&" days"', alert=("FFF4D6", "8A6100"))
tile("N", S0 + 1, "OWNED CONTRACTS", f'=COUNTIF({PT("AT")},"Owned")', "Your department code")
tile("Q", S0 + 1, "CODE NOT SET", f'=COUNTIF({PT("AT")},"Not set")', "Owning Dept Code is blank", alert=("FFF4D6", "8A6100"))

hr4 = S0 + 6
sh(f"B{hr4 - 1}", "Contracts by Owning Dept Code and expiration status")
c = db[f"B{hr4}"]; c.value = "Dept Code"
c2 = db[f"C{hr4}"]; c2.value = "Contracts"
for cc in (c, c2):
    cc.font = font(bold=True, color="FFFFFF"); cc.fill = fill(NAVY); cc.border = BOX; cc.alignment = Alignment(horizontal="center")
for j, s in enumerate(range(E0, E0 + NSTAT_SH)):
    col = L(4 + j)
    h = db[f"{col}{hr4}"]; h.value = f'=IF({AC}!$C${s}="","",{AC}!$C${s})'
    h.font = font(bold=True, color="FFFFFF", size=9); h.fill = fill(NAVY); h.border = BOX
    h.alignment = Alignment(horizontal="center", wrap_text=True)
NCODES = CODE1 - CODE0 + 1
for i in range(NCODES):
    r = hr4 + 1 + i; src = f"{AC}!$C${CODE0 + i}"
    db[f"B{r}"] = f'=IF({src}="","",{src})'
    db[f"C{r}"] = f'=IF(B{r}="","",COUNTIF({PT("AS")},B{r}))'
    for j in range(NSTAT_SH):
        col = L(4 + j)
        db[f"{col}{r}"] = f'=IF(OR($B{r}="",{col}${hr4}=""),"",COUNTIFS({PT("AS")},$B{r},{PT("AI")},{col}${hr4}))'
    for j in range(2 + NSTAT_SH):
        x = db.cell(r, 2 + j); x.border = BOX; x.font = font(size=9, bold=(j < 2)); x.alignment = Alignment(horizontal="center")
db.conditional_formatting.add(f"B{hr4+1}:{L(3+NSTAT_SH)}{hr4+NCODES}",
                              FormulaRule(formula=[f'$B{hr4+1}={MY_CODE}'], fill=fill("D6EEF9")))
db[f"B{hr4 + NCODES + 1}"] = "Your department's row is shaded blue. Codes come from Alert Controls section 15."
db[f"B{hr4 + NCODES + 1}"].font = font(size=9, italic=True, color=GREY_HDR)

cb = BarChart(); cb.type = "bar"; cb.grouping = "stacked"; cb.overlap = 100; cb.title = "Contracts by Owning Dept Code"
cb.add_data(Reference(db, min_col=4, max_col=3 + NSTAT_SH, min_row=hr4, max_row=hr4 + NCODES), titles_from_data=True)
cb.set_categories(Reference(db, min_col=2, min_row=hr4 + 1, max_row=hr4 + NCODES))
for srs, cname in zip(cb.series, [e[4] for e in EXP]):
    srs.graphicalProperties.solidFill = CHART_HEX[cname]; srs.graphicalProperties.line.solidFill = CHART_HEX[cname]
cb.x_axis.scaling.orientation = "maxMin"; cb.y_axis.majorGridlines = None
cb.height, cb.width = 8, 14
db.add_chart(cb, f"K{hr4}")

# ---------------- Summary: Owned vs Shared x Expiration Status
sr0 = sm.max_row + 2
shead(f"A{sr0}", "Owned vs. shared contracts")
th(f"A{sr0+1}", "Ownership"); th(f"B{sr0+1}", "Contracts")
for j, s in enumerate(range(E0, E1 + 1)):
    col = L(3 + j)
    cc = sm[f"{col}{sr0+1}"]; cc.value = f'=IF({AC}!$C${s}="","",{AC}!$C${s})'
    cc.font = font(bold=True, color="FFFFFF"); cc.fill = fill(BLUE_HDR); cc.border = BOX
    cc.alignment = Alignment(horizontal="center", wrap_text=True)
for i, grp in enumerate(["Owned", "Shared", "Not set"]):
    r = sr0 + 2 + i
    td(f"A{r}", grp, "@", bold=True)
    td(f"B{r}", f'=COUNTIF({PT("AT")},A{r})', bold=True)
    for j in range(E1 - E0 + 1):
        col = L(3 + j)
        td(f"{col}{r}", f'=IF({col}${sr0+1}="","",COUNTIFS({PT("AT")},$A{r},{PT("AI")},{col}${sr0+1}))')
