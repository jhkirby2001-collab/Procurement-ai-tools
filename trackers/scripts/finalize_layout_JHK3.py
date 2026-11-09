"""Put the Portfolio Tracker columns into their final layout, the way Excel would.

Runs LibreOffice headless on a built workbook: moves every column into contract_tracker_layout_JHK3.FINAL_ORDER
(cut-and-insert, so every formula, conditional format, validation and filter on every tab follows the column),
deletes the REMOVED columns, folds the HELPERS into a collapsed +/- group, recalculates and saves in place. Then writes
the freeze panes (columns A-C and row 1) and the whole-row highlight / filter ranges into the sheet XML.

Usage: python finalize_layout_JHK3.py <workbook.xlsx>
Needs LibreOffice (soffice). In sandboxes that block AF_UNIX sockets, put the xlsx skill's scripts folder on
PYTHONPATH so office.soffice can supply its socket shim. Do not run recalc.py afterwards: LibreOffice drops the freeze.
"""
import os, re, sys, shutil, subprocess, tempfile, zipfile
from pathlib import Path
from openpyxl.utils import get_column_letter
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from contract_tracker_layout_JHK3 import FINAL_ORDER, REMOVED, HELPERS

path = Path(sys.argv[1]).resolve()
try:
    from office.soffice import get_soffice_env
    env = get_soffice_env()
except ImportError:
    env = dict(os.environ)

def basic_str(s):
    return '"' + s.replace('"', '""').replace("\n", '" & Chr(10) & "') + '"'

order = ", ".join(basic_str(h) for h in FINAL_ORDER)
removed = ", ".join(basic_str(h) for h in REMOVED)
h0, h1 = FINAL_ORDER.index(HELPERS[0]), FINAL_ORDER.index(HELPERS[-1])

MACRO = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
    Function FindHeader(sh As Object, h As String) As Long
      Dim c As Long
      For c = 0 To 900
        If sh.getCellByPosition(c, 0).getString() = h Then
          FindHeader = c
          Exit Function
        End If
      Next c
      FindHeader = -1
    End Function

    Sub Finalize()
      Dim doc As Object, sh As Object, src As Object, cols As Object
      Dim order As Variant, removed As Variant, i As Long, j As Long, w As Long, vis As Boolean
      doc = ThisComponent
      sh = doc.Sheets.getByName("Portfolio Tracker")
      cols = sh.Columns
      order = Array({order})
      removed = Array({removed})
      For i = 0 To UBound(order)
        j = FindHeader(sh, order(i))
        If j &gt;= 0 And j &lt;&gt; i Then
          w = cols.getByIndex(j).Width
          vis = cols.getByIndex(j).IsVisible
          cols.insertByIndex(i, 1)
          If j &gt;= i Then j = j + 1
          src = sh.getCellRangeByPosition(j, 0, j, 1048575).getRangeAddress()
          sh.moveRange(sh.getCellByPosition(i, 0).getCellAddress(), src)
          cols.removeByIndex(j, 1)
          cols.getByIndex(i).Width = w
          cols.getByIndex(i).IsVisible = vis
        End If
      Next i
      For i = 0 To UBound(removed)
        j = FindHeader(sh, removed(i))
        If j &gt;= 0 Then cols.removeByIndex(j, 1)
      Next i
      Dim grp As Object
      grp = sh.getCellRangeByPosition({h0}, 0, {h1}, 299).getRangeAddress()
      sh.group(grp, com.sun.star.table.TableOrientation.COLUMNS)
      sh.hideDetail(grp)
      doc.calculateAll()
      doc.store()
      doc.close(True)
    End Sub
</script:module>"""

with tempfile.TemporaryDirectory(prefix="lo_profile_") as prof:
    url = Path(prof).as_uri()
    subprocess.run(["soffice", "--headless", "--terminate_after_init", f"-env:UserInstallation={url}"],
                   env=env, capture_output=True, timeout=120)
    (Path(prof) / "user" / "basic" / "Standard" / "Module1.xba").write_text(MACRO)
    before = path.stat().st_mtime
    r = subprocess.run(["soffice", "--headless", "--norestore", f"-env:UserInstallation={url}",
                        "vnd.sun.star.script:Standard.Module1.Finalize?language=Basic&location=application", str(path)],
                       env=env, capture_output=True, text=True, timeout=1500)
    if r.returncode != 0 or path.stat().st_mtime == before:
        sys.exit(f"LibreOffice did not update the file: {r.stderr.strip() or r.returncode}")

# Sheet XML: freeze A-C + row 1; whole-row highlight and filter span every column.
last = get_column_letter(len(FINAL_ORDER))
with zipfile.ZipFile(path) as z:
    wbx = z.read("xl/workbook.xml").decode()
    rid = re.search(r'<sheet [^>]*name="Portfolio Tracker"[^>]*r:id="([^"]+)"', wbx).group(1)
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    tgt = (re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or re.search(r'Target="([^"]+)"[^>]*Id="%s"' % rid, rels)).group(1)
    sheet = "xl/" + tgt.split("xl/")[-1].lstrip("/")
    items = [(i, z.read(i.filename)) for i in z.infolist()]
pane = ('<pane xSplit="3" ySplit="1" topLeftCell="D2" activePane="bottomRight" state="frozen"/>'
        '<selection pane="topRight"/><selection pane="bottomLeft"/><selection pane="bottomRight" activeCell="D2" sqref="D2"/>')
tmp = path.with_suffix(".tmp.xlsx")
with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
    for info, data in items:
        if info.filename == sheet:
            x = data.decode()
            x = re.sub(r'(<sheetView\b[^>]*>).*?(</sheetView>)', lambda m: m.group(1) + pane + m.group(2), x, count=1, flags=re.S)
            # whole-row "Act now" / "Plan" highlight: the conditionalFormatting whose rule tests the Action Needed text
            def widen(m):
                block = m.group(0)
                rng = re.search(r'sqref="([^"]+)"', block).group(1)
                if ('&quot;Act now&quot;' in block or '&quot;Plan&quot;' in block) and rng.startswith(("A2:", "B2:")):
                    block = block.replace(f'sqref="{rng}"', f'sqref="A2:{last}300"', 1)
                return block
            x = re.sub(r'<conditionalFormatting\b.*?</conditionalFormatting>', widen, x, flags=re.S)
            x = re.sub(r'<autoFilter ref="[^"]+"', f'<autoFilter ref="A1:{last}300"', x, count=1)
            data = x.encode()
        out.writestr(info, data)
shutil.move(tmp, path)
print(f"Portfolio Tracker laid out: {len(FINAL_ORDER)} columns (A-{last}), {len(REMOVED)} removed, helpers grouped; frozen at D2")
