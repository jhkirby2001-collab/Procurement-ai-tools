// Build the one-page Contract Tracker Staff Quick Start.
// Usage: node build_contract_quick_start_JHK3.js ../outputs/Contract_Tracker_Staff_Quick_Start_JHK3.docx
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
        AlignmentType, BorderStyle, LevelFormat, VerticalAlign } = require("docx");

const NAVY = "1F3864", TEAL = "0E5E5A", GREY = "595959";
const FONT = "Arial";
const W = 12240 - 2 * 864;            // content width at 0.6" margins
const border = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const borders = { top: border, bottom: border, left: border, right: border };
const r = (text, o = {}) => new TextRun({ text, font: FONT, size: o.size || 19, bold: o.bold, color: o.color, italics: o.italics });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], spacing: { before: o.before || 0, after: o.after ?? 60 }, alignment: o.align, numbering: o.numbering });
const h = (text) => new Paragraph({ children: [r(text, { bold: true, size: 22, color: TEAL })], spacing: { before: 140, after: 60 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: TEAL, space: 2 } } });

function cell(content, w, o = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA }, borders, verticalAlign: VerticalAlign.CENTER,
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 40, bottom: 40, left: 90, right: 90 },
    children: [p(typeof content === "string" ? [r(content, { bold: o.bold, color: o.color, size: o.size })] : content, { after: 0, align: o.align })]
  });
}
function table(widths, rows) {
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: widths, rows: rows.map(cells => new TableRow({ children: cells })) });
}
const head = (labels, widths) => labels.map((t, i) => cell(t, widths[i], { fill: NAVY, color: "FFFFFF", bold: true, size: 18 }));

// ---- 4 columns table
const cw = [2300, 4100, W - 2300 - 4100];
const cols = table(cw, [
  head(["Column", "What to enter", "Example"], cw),
  [cell("Action Status", cw[0], { bold: true }), cell("Pick from the dropdown: where the renewal work stands.", cw[1]), cell("Submitted to DPS", cw[2])],
  [cell("Next Action", cw[0], { bold: true }), cell("The one next step, in a few words.", cw[1]), cell("Send usage estimate to bureau", cw[2])],
  [cell("Next Action Due Date", cw[0], { bold: true }), cell("When that step is due (mm/dd/yyyy).", cw[1]), cell("11/15/2026", cw[2])],
  [cell("Next Action Done?", cw[0], { bold: true }), cell("Pick Yes when finished, then enter the next action and date.", cw[1]), cell("Yes", cw[2])],
]);

// ---- colours table
const pill = (label, fill, color, w) => cell(label, w, { fill, color, bold: true, align: AlignmentType.CENTER, size: 17 });
const kw = [1700, (W - 3400) / 2, 1700, (W - 3400) / 2];
const colours = table(kw, [
  head(["Expiration", "Meaning", "Next Action", "Meaning"], kw),
  [pill("LAPSED", "C00000", "FFFFFF", kw[0]), cell("Past end date. Top priority.", kw[1]), pill("OVERDUE", "F4A6A6", "7A0000", kw[2]), cell("Past its due date.", kw[3])],
  [pill("RED", "F4A6A6", "7A0000", kw[0]), cell("Act now: little time left or past the Start-By date.", kw[1]), pill("DUE SOON", "FFD966", "5C4400", kw[2]), cell("Due within 14 days.", kw[3])],
  [pill("YELLOW", "FFD966", "5C4400", kw[0]), cell("Start preparing.", kw[1]), pill("UPCOMING", "FFF2CC", "5C4400", kw[2]), cell("Due within 30 days.", kw[3])],
  [pill("IN PROCESS", "BDD7EE", "1F3864", kw[0]), cell("Submitted and moving. Keep watching.", kw[1]), cell("Priority", kw[2], { bold: true, align: AlignmentType.CENTER }), cell("1 = most urgent. Work top down.", kw[3])],
  [pill("GREEN", "C6EFCE", "1E5631", kw[0]), cell("On track.", kw[1]), cell("", kw[2]), cell("", kw[3])],
  [pill("CLOSED", "D9D9D9", "404040", kw[0]), cell("Replaced, voided or let expire.", kw[1]), cell("", kw[2]), cell("", kw[3])],
]);

const bullets = { reference: "b", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
  style: { paragraph: { indent: { left: 360, hanging: 240 } } } }] };
const nums = { reference: "n", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
  style: { paragraph: { indent: { left: 360, hanging: 300 } } } }] };
const B = (runs) => p(runs, { numbering: { reference: "b", level: 0 }, after: 30 });
const N = (runs) => p(runs, { numbering: { reference: "n", level: 0 }, after: 30 });

const doc = new Document({
  numbering: { config: [bullets, nums] },
  styles: { default: { document: { run: { font: FONT, size: 19 } } } },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 720, bottom: 640, left: 864, right: 864 } } },
    children: [
      p(r("CONTRACT TRACKER", { bold: true, size: 17, color: TEAL }), { after: 0 }),
      p(r("Staff Quick Start", { bold: true, size: 40, color: NAVY }), { after: 40 }),
      p(r("About 10 minutes a week. You update four columns; the tracker does the rest.", { italics: true, color: GREY }), { after: 80 }),

      h("Your job"),
      p([r("Keep the "), r("four columns below", { bold: true }), r(" current for every contract you own. Everything else on the "), r("Portfolio Tracker", { bold: true }), r(" tab calculates itself.")]),

      h("The 4 columns you update"),
      cols,

      h("Weekly routine"),
      N([r("Open the tracker from the team's "), r("OneDrive Trackers", { bold: true }), r(" folder and go to the "), r("Portfolio Tracker", { bold: true }), r(" tab.")]),
      N([r("Filter "), r("Proposed Owner", { bold: true }), r(" to your name. Sort "), r("Priority", { bold: true }), r(" smallest to largest.")]),
      N([r("Work top down. For each contract you touched this week, update the 4 columns. Add detail in "), r("Status / Notes", { bold: true }), r(" if needed.")]),
      N([r("Clear anything showing "), r("OVERDUE", { bold: true }), r(" first, then "), r("LAPSED", { bold: true }), r(" and "), r("RED", { bold: true }), r(".")]),
      N([r("Save. Done.")]),

      h("What the colours mean"),
      colours,
      p(r("The status word is always written in the cell, so you never have to rely on colour alone.", { italics: true, color: GREY, size: 17 }), { before: 40 }),

      h("Please don't"),
      B([r("Type in columns marked "), r("(auto)", { bold: true }), r(" with grey headers. They are formulas.")]),
      B([r("Type your own values in dropdown columns. Pick from the list; ask for a new choice if you need one.")]),
      B([r("Delete a contract row. Set its Action Status to a Closed status instead.")]),
      B([r("Change the "), r("Alert Controls", { bold: true }), r(" tab. Rules and thresholds are set by the portfolio lead.")]),
      B([r("Save a copy somewhere else. The OneDrive Trackers folder is the only master copy.")]),

      h("Need help?"),
      p([r("See the "), r("How To Use", { bold: true }), r(" tab in the tracker for the full guide, or ask your portfolio lead.")]),
    ]
  }]
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(process.argv[2], b); console.log("ok"); });
