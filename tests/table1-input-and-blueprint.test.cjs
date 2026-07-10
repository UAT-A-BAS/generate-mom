const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");

function extractFunctionSource(source, name) {
  const start = source.indexOf(`function ${name}`);
  assert.notEqual(start, -1, `${name} should exist`);
  const bodyStart = source.indexOf("{", start);
  let depth = 0;

  for (let index = bodyStart; index < source.length; index += 1) {
    if (source[index] === "{") {
      depth += 1;
    } else if (source[index] === "}") {
      depth -= 1;
      if (depth === 0) {
        return source.slice(start, index + 1);
      }
    }
  }

  throw new Error(`${name} body should close`);
}

assert.match(
  html,
  /\.table1-detail-row\s+\.field-body\s*>\s*\.free-textarea\s*\{(?=[^}]*height:\s*64px;)(?=[^}]*resize:\s*vertical;)[^}]*\}/s,
  "all Table 1 feature textareas should be vertically resizable"
);
assert.match(
  html,
  /\.table1-detail-grid\s+\.field\s*\{[^}]*align-content:\s*start;/s,
  "Table 1 feature fields should remain top-aligned when one textarea grows"
);
assert.match(
  html,
  /#table2Rows\s+textarea\[data-field="note"\]\s*\{[^}]*resize:\s*vertical;/s,
  "Table 2 note textareas should be vertically resizable"
);
assert.equal(
  (html.match(/<textarea\b[^>]*class="free-textarea blueprint-link-textarea"[^>]*data-field="blueprintLink"[^>]*rows="1"[^>]*>/g) || []).length,
  2,
  "Release and BPRO blueprint links should both render as multiline textareas"
);
assert.doesNotMatch(
  html,
  /<input\b[^>]*type="url"[^>]*data-field="blueprintLink"/,
  "blueprint links should not render as single-line URL inputs"
);
assert.match(
  html,
  /\.field-body\s*>\s*\.blueprint-link-textarea\s*\{(?=[^}]*min-height:\s*64px;)(?=[^}]*resize:\s*vertical;)[^}]*\}/s,
  "all blueprint link textareas should auto-grow and remain manually resizable"
);
assert.match(
  html,
  /upgradeFreeTextInputsToTextarea\(elements\.table1Projects\);\s*refreshTextareasScrollState\(elements\.table1Projects\);/,
  "blueprint textareas should resize immediately when Table 1 is rendered"
);
assert.match(
  html,
  /blueprintValue !== "-"\s*\? formatMultilineWithLinks\(blueprintValue\)/,
  "Table 1 preview should render every Enter-separated blueprint URL independently"
);

const context = {
  table1BlueprintLevel: "bpro",
  escapeHtml(value) {
    return `${value ?? ""}`
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  },
};
vm.createContext(context);
for (const functionName of [
  "getAutoGrowHeight",
  "formatJoined",
  "formatMultilineWithLinks",
  "joinDocumentNeeds",
  "buildGroupMeta",
  "normalizeBlueprintMergeKey",
  "buildMergeMeta",
]) {
  vm.runInContext(extractFunctionSource(html, functionName), context);
}

assert.equal(context.getAutoGrowHeight(64, 138), 138);
assert.equal(context.getAutoGrowHeight(200, 138), 200);
assert.equal(context.getAutoGrowHeight(32, 40), 64);

const multilineBlueprintMarkup = context.formatMultilineWithLinks(
  "https://example.com/blueprint-a\nhttps://example.com/blueprint-b"
);
assert.equal(
  (multilineBlueprintMarkup.match(/<a class="doc-link"/g) || []).length,
  2,
  "each Enter-separated blueprint URL should render as its own link"
);
assert.match(
  multilineBlueprintMarkup,
  /blueprint-a<\/a><br>\s*<a class="doc-link"[^>]*>https:\/\/example\.com\/blueprint-b<\/a>/,
  "Enter-separated blueprint links should keep their line break"
);

const sharedBlueprintRows = [
  {
    bproNumber: "BPRO-1",
    bproName: "Fitur A",
    releaseId: "REL-1",
    releaseName: "Release",
    blueprintLink: "https://example.com/blueprint",
    documentNeeds: [],
    uatByUser: "User A",
  },
  {
    bproNumber: "BPRO-2",
    bproName: "Fitur B",
    releaseId: "REL-1",
    releaseName: "Release",
    blueprintLink: "https://example.com/blueprint",
    documentNeeds: [],
    uatByUser: "User B",
  },
];
const sharedMeta = context.buildMergeMeta(sharedBlueprintRows);
assert.deepEqual(JSON.parse(JSON.stringify(sharedMeta.map((entry) => entry.blueprint))), [
  { rowSpan: 2, show: true },
  { rowSpan: 0, show: false },
]);

const blankMeta = context.buildMergeMeta(
  sharedBlueprintRows.map((row) => ({ ...row, blueprintLink: "-" }))
);
assert.deepEqual(JSON.parse(JSON.stringify(blankMeta.map((entry) => entry.blueprint))), [
  { rowSpan: 1, show: true },
  { rowSpan: 1, show: true },
]);

console.log("table1 input and blueprint tests passed");
