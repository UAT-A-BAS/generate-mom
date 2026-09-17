const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");

// Regression guard for a real production bug: a structural patch that targets a *nested*
// array (for example `table1ProjectsState/0/relationPackages`) used to update the model
// without re-rendering, so the other editor's added row was invisible even though both
// sides agreed on the data.
assert.match(
  html,
  /function applyStructuralArrayPatch\(path, value\)/,
  "structural array patches must have a dedicated apply path"
);

assert.match(
  html,
  /const isStructuralArray = Array\.isArray\(value\) && !path\.endsWith\("\/documentNeeds"\);/,
  "array values are structural, except the documentNeeds checkbox list"
);
assert.match(
  html,
  /if \(isStructuralArray\) \{\s*applyStructuralArrayPatch\(path, value\);\s*return;\s*\}/,
  "applyFieldPatchValue must route array values to the structural apply path"
);

// Every array-bearing section must re-render, at both the root path and nested paths.
for (const [root, renderer] of [
  ["table1ProjectsState", "renderTable1Projects"],
  ["table3State", "renderTable3Rows"],
  ["lampiranState", "renderLampiranSection"],
]) {
  assert.match(
    html,
    new RegExp(`parts\\[0\\] === "${root}"[\\s\\S]{0,400}?${renderer}\\(\\);`),
    `${root} structural patches must re-render via ${renderer}()`
  );
}

// A brand new client must publish a baseline for every array as soon as it learns whether
// the session is empty. Without that baseline the first structural click re-sends every
// array and overwrites work another editor already committed.
assert.match(
  html,
  /refreshCollabArraySignatures\(collectDraftPayload\(\)\)/,
  "the client must prime array signatures against the current draft"
);
assert.match(
  html,
  /if \(ackedFullDraft\) \{\s*refreshCollabArraySignatures\(collectDraftPayload\(\)\);/,
  "a confirmed full write must re-prime the array signatures"
);

// Editing a field that lives inside an array changes that array's contents, so its
// signature has to be refreshed when the field patch is sent and acknowledged.
assert.match(
  html,
  /refreshSignaturesForTouchedPaths\(ops\.map\(\(op\) => op\.path\)\);/,
  "sending a field patch must refresh the enclosing array signatures"
);
assert.match(
  html,
  /refreshSignaturesForTouchedPaths\(\s*ops\.filter\(\(op\) => op\?\.applied !== false\)\.map\(\(op\) => op\.path\)\s*\);/,
  "acknowledged patches must refresh the enclosing array signatures"
);

console.log("collab structural render tests passed");
