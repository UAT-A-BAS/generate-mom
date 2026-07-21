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
  /<table class="doc-table table2">[\s\S]*?<th style="width: 80px">No\.<\/th>[\s\S]*?<th style="width: 360px">Aktivitas<\/th>[\s\S]*?<th style="width: 140px">Status<\/th>[\s\S]*?<th style="width: 170px">PIC<\/th>[\s\S]*?<th style="width: 170px">Target<\/th>[\s\S]*?<th style="width: 280px">Keterangan<\/th>/,
  "web checklist preview should keep its existing column widths"
);
assert.match(
  html,
  /if \(table\.classList\.contains\("table2"\)\) \{\s*return \["80px", "400px", "140px", "190px", "170px", "220px"\];/,
  "Outlook export should preserve the spacious checklist column widths"
);
assert.match(
  html,
  /if \(table\.classList\.contains\("table2"\)\) \{\s*return "1200px";/,
  "Outlook checklist table should be wide enough to prevent narrow wrapping"
);
assert.match(
  html,
  /<table class="doc-table table1">[\s\S]*?<th style="width: 80px">No\.<\/th>[\s\S]*?<th style="width: 200px">Nomor &amp; Nama BPRO<\/th>[\s\S]*?<th style="width: 220px">Changes ID &amp; Changes Name<\/th>[\s\S]*?<th style="width: 220px">Release ID &amp; Release Name<\/th>[\s\S]*?<th style="width: 280px">Link Blueprint<\/th>[\s\S]*?<th style="width: 200px">Apakah diperlukan SK\/SE\/Service News\/Memo\?<\/th>[\s\S]*?<th style="width: 200px">Pelaku UAT by User<\/th>/,
  "web certification preview should keep its existing column widths"
);
assert.match(
  html,
  /if \(table\.classList\.contains\("table1"\)\) \{\s*return \["80px", "240px", "260px", "220px", "280px", "200px", "200px"\];/,
  "Outlook export should preserve the spacious certification column widths"
);
assert.match(
  html,
  /if \(table\.classList\.contains\("table1"\)\) \{\s*return "1480px";/,
  "Outlook certification table should be wide enough to avoid narrow wrapping"
);
assert.match(
  html,
  /border:1px solid #111;padding:10px 12px;vertical-align:middle;white-space:normal;word-break:normal;overflow-wrap:break-word;line-break:auto;/,
  "Outlook tables should use roomier padding and natural wrapping"
);
assert.match(
  html,
  /table\.querySelectorAll\(":scope > thead > tr:first-child > th"\)/,
  "Outlook widths should be applied only to table headers"
);
assert.match(
  html,
  /cell\.style\.width = width;/,
  "Outlook header widths should replace the web preview widths"
);
assert.match(
  html,
  /width:\$\{tableWidth\};min-width:0;max-width:none;border-collapse:collapse;table-layout:auto;/,
  "Outlook tables should remain adjustable instead of using fixed layout"
);
assert.doesNotMatch(
  html,
  /document\.createElement\("colgroup"\)/,
  "Outlook export should not add a colgroup that fights manual resizing"
);

console.log("checklist Outlook alignment tests passed");
