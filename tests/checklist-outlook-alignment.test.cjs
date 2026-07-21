const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");

assert.match(
  html,
  /<td class="procedure-no center">\$\{escapeHtml\(procedureGroup\.no\)\}<\/td>/,
  "checklist number 11 should carry the center class into Outlook export markup"
);
assert.match(
  html,
  /<td class="sub-row-number">\$\{escapeHtml\(row\.no\)\}<\/td>/,
  "checklist sub-row numbers should keep their independent alignment"
);
assert.match(
  html,
  /if \(cell\.classList\.contains\("sub-row-number"\)\) \{[\s\S]*?appendInlineStyle\(cell, "text-align:right;padding-right:14px;"\);/,
  "Outlook export should inline right alignment for checklist sub-row numbers"
);
assert.match(
  html,
  /if \(cell\.classList\.contains\("center"\)\) \{[\s\S]*?appendInlineStyle\(cell, "text-align:center;"\);/,
  "Outlook export should inline the center alignment"
);
assert.match(
  html,
  /<table class="doc-table table2">[\s\S]*?<th style="width: 80px">No\.<\/th>[\s\S]*?<th style="width: 330px">Aktivitas<\/th>[\s\S]*?<th style="width: 140px">Status<\/th>[\s\S]*?<th style="width: 170px">PIC<\/th>[\s\S]*?<th style="width: 170px">Target<\/th>[\s\S]*?<th style="width: 310px">Keterangan<\/th>/,
  "checklist preview should use spacious column widths"
);
assert.match(
  html,
  /if \(table\.classList\.contains\("table2"\)\) \{\s*return \["80px", "330px", "140px", "170px", "170px", "310px"\];/,
  "Outlook export should preserve the spacious checklist column widths"
);
assert.match(
  html,
  /if \(table\.classList\.contains\("table2"\)\) \{\s*return "1200px";/,
  "Outlook checklist table should be wide enough to prevent narrow wrapping"
);

console.log("checklist Outlook alignment tests passed");
