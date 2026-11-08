"""Insert the working copy's "Reviewed" column at A of the Portfolio Tracker, the way Excel does.

The build parks the Reviewed values in column ZZ (REVIEWED_COL=1). This runs LibreOffice headless to insert a
column at A - every formula, conditional format, data validation and filter that points at the Portfolio Tracker
shifts one column right, exactly as when the column is inserted in Excel - moves the parked values into A,
recalculates and saves in place.

Usage: python insert_reviewed_column_JHK3.py <workbook.xlsx> [freeze cell, default D2] [scroll-to cell, default same]
(D2 = columns A-C and row 1 stay visible, as in the working copy.)
Needs LibreOffice (soffice) on PATH. In sandboxes that block AF_UNIX sockets, put the xlsx skill's scripts folder
on PYTHONPATH so office.soffice can supply its socket shim.
"""
import os, sys, subprocess, tempfile, re
from pathlib import Path
from openpyxl.utils import column_index_from_string

path = Path(sys.argv[1]).resolve()
freeze = sys.argv[2] if len(sys.argv) > 2 else "D2"
scroll = sys.argv[3] if len(sys.argv) > 3 else freeze
m = re.fullmatch(r"([A-Z]+)(\d+)", freeze)
fcol, frow = column_index_from_string(m.group(1)) - 1, int(m.group(2)) - 1
try:
    from office.soffice import get_soffice_env
    env = get_soffice_env()
except ImportError:
    env = dict(os.environ)

MACRO = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
    Sub InsertReviewed()
      Dim doc As Object, sh As Object, src As Object, dst As Object
      doc = ThisComponent
      sh = doc.Sheets.getByName("Portfolio Tracker")
      sh.Columns.insertByIndex(0, 1)
      src = sh.getCellRangeByPosition(702, 0, 702, 299)
      dst = sh.getCellRangeByPosition(0, 0, 0, 299)
      dst.setDataArray(src.getDataArray())
      src.clearContents(1023)
      sh.Columns.getByIndex(0).Width = 2600
      dst.VertJustify = 1
      doc.calculateAll()
      doc.store()
      doc.close(True)
    End Sub
</script:module>"""

with tempfile.TemporaryDirectory(prefix="lo_profile_") as prof:
    url = Path(prof).as_uri()
    subprocess.run(["soffice", "--headless", "--terminate_after_init", f"-env:UserInstallation={url}"],
                   env=env, capture_output=True, timeout=120)
    mdir = Path(prof) / "user" / "basic" / "Standard"
    (mdir / "Module1.xba").write_text(MACRO)
    before = path.stat().st_mtime
    r = subprocess.run(["soffice", "--headless", "--norestore", f"-env:UserInstallation={url}",
                        "vnd.sun.star.script:Standard.Module1.InsertReviewed?language=Basic&location=application", str(path)],
                       env=env, capture_output=True, text=True, timeout=600)
    if r.returncode != 0 or path.stat().st_mtime == before:
        sys.exit(f"LibreOffice did not update the file: {r.stderr.strip() or r.returncode}")
# Freeze panes: LibreOffice headless does not keep a frozen split, so write it into the sheet XML directly.
import zipfile, shutil
with zipfile.ZipFile(path) as z:
    wbx = z.read("xl/workbook.xml").decode()
    rid = re.search(r'<sheet [^>]*name="Portfolio Tracker"[^>]*r:id="([^"]+)"', wbx).group(1)
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    tgt = (re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or re.search(r'Target="([^"]+)"[^>]*Id="%s"' % rid, rels)).group(1)
    sheet = "xl/" + tgt.split("xl/")[-1].lstrip("/")
    items = [(i, z.read(i.filename)) for i in z.infolist()]
pane = (f'<pane xSplit="{fcol}" ySplit="{frow}" topLeftCell="{scroll}" activePane="bottomRight" state="frozen"/>'
        f'<selection pane="topRight"/><selection pane="bottomLeft"/><selection pane="bottomRight" activeCell="{scroll}" sqref="{scroll}"/>')
tmp = path.with_suffix(".tmp.xlsx")
with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
    for info, data in items:
        if info.filename == sheet:
            x = data.decode()
            x = re.sub(r'(<sheetView\b[^>]*>).*?(</sheetView>)', lambda m: m.group(1) + pane + m.group(2), x, count=1, flags=re.S)
            data = x.encode()
        out.writestr(info, data)
shutil.move(tmp, path)
print("Reviewed column inserted at A of the Portfolio Tracker; frozen at", freeze, "scrolled to", scroll)
