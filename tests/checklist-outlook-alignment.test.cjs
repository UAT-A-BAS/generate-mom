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
  /if \(cell\.classList\.contains\("center"\)\) \{[\s\S]*?appendInlineStyle\(cell, "text-align:center;"\);/,
  "Outlook export should inline the center alignment"
);

console.log("checklist Outlook alignment tests passed");
