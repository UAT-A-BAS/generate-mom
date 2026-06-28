const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");

assert.match(
  html,
  /\.lampiran-date-card\s+\.field-body\s*>\s*textarea:not\(\.date-display\)\s*\{(?=[^}]*height:\s*64px;)(?=[^}]*resize:\s*vertical;)(?=[^}]*overflow:\s*auto\s*!important;)[^}]*\}/s,
  "all free-text Lampiran fields should be vertically resizable"
);
assert.match(
  html,
  /\.lampiran-date-card\s+\.field,[^{]*\.lampiran-feature-card\s+\.field,[^{]*\.lampiran-scenario-card\s+\.field\s*\{[^}]*align-content:\s*start;/s,
  "Lampiran fields should remain top-aligned when one field grows"
);
assert.match(
  html,
  /function autoGrowFeatureTextarea\(textarea\)[\s\S]*?isVerticallyResizableTextarea\(textarea\)[\s\S]*?getAutoGrowHeight\(/,
  "vertically resizable Lampiran textareas should use the generic auto-grow helper"
);
assert.match(
  html,
  /querySelectorAll\("textarea"\)/,
  "all textareas should be checked after render so resizable Lampiran fields auto-grow"
);
assert.match(
  html,
  /elements\.lampiranRows\.addEventListener\("input",[\s\S]*?refreshTextareaScrollState\(event\.target\);[\s\S]*?\}\);/,
  "Lampiran textareas should auto-grow while typing"
);

console.log("lampiran input tests passed");
