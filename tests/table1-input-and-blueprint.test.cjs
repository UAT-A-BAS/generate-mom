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
  /\.table1-detail-row\s+\.field-body\s*>\s*\.free-textarea\s*\{[^}]*resize:\s*vertical;/s,
  "all Table 1 feature textareas should be vertically resizable"
);
assert.match(
  html,
  /#table2Rows\s+textarea\[data-field="note"\]\s*\{[^}]*resize:\s*vertical;/s,
  "Table 2 note textareas should be vertically resizable"
);

const context = { table1BlueprintLevel: "bpro" };
vm.createContext(context);
for (const functionName of [
  "formatJoined",
  "joinDocumentNeeds",
  "buildGroupMeta",
  "normalizeBlueprintMergeKey",
  "buildMergeMeta",
]) {
  vm.runInContext(extractFunctionSource(html, functionName), context);
}

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
