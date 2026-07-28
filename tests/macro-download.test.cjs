const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const projectRoot = path.join(__dirname, "..");
const html = fs.readFileSync(path.join(projectRoot, "index.html"), "utf8");
const macroPath = path.join(projectRoot, "ExportMOMToDraft.bas");

assert.equal(fs.existsSync(macroPath), true, "Outlook macro download file should exist");
assert.match(
  html,
  /<a[\s\S]*?class="hero-download"[\s\S]*?href="\.\/ExportMOMToDraft\.bas"[\s\S]*?download="ExportMOMToDraft\.bas"[\s\S]*?>[\s\S]*?Macro Outlook[\s\S]*?<\/a>/,
  "header should include the Macro Outlook download button"
);
assert.match(html, /\.hero-download svg\s*\{/, "download button should style its icon");

console.log("Macro Outlook download tests passed");
