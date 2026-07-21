const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const macro = fs.readFileSync(path.join(__dirname, "..", "ExportMOMToDraft.bas"), "utf8");

assert.match(macro, /\.HTMLBody = htmlContent/);
assert.match(macro, /\.Save\s*\r?\n\s*\.Display/);

[
  /FixTable2HeaderForOutlook/i,
  /FixDisplayedTableHeaders/i,
  /WordEditor/i,
  /AllowAutoFit/i,
  /AutoFitBehavior/i,
  /PreferredWidth/i,
  /SetWidth/i,
].forEach((forbiddenPattern) => {
  assert.doesNotMatch(
    macro,
    forbiddenPattern,
    `Outlook macro must not mutate Word tables after HTML import: ${forbiddenPattern}`
  );
});

assert.equal(
  (macro.match(/\.Save\b/g) || []).length,
  1,
  "Outlook draft should be saved once and never re-saved after table rendering"
);

console.log("Outlook VBA performance tests passed");
