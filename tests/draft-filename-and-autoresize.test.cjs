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
    if (source[index] === "{") depth += 1;
    if (source[index] === "}") depth -= 1;
    if (depth === 0) return source.slice(start, index + 1);
  }

  throw new Error(`${name} body should close`);
}

const filenameContext = {};
vm.createContext(filenameContext);
for (const functionName of ["sanitizeDraftFilePart", "buildDraftFileName"]) {
  vm.runInContext(extractFunctionSource(html, functionName), filenameContext);
}

assert.equal(
  filenameContext.buildDraftFileName({
    table1ProjectsState: [{ projectName: " BERTA Release 4/2026 " }],
  }),
  "BERTA Release 4 2026_MOM.json"
);
assert.equal(
  filenameContext.buildDraftFileName({ table1ProjectsState: [{ projectName: "" }] }),
  "Project_MOM.json"
);
assert.match(html, /Export to Outlook/);
assert.match(html, /function exportOutlookResult\(\)/);
assert.match(html, /exportOutlookBtn/);
assert.match(html, /previewDrawerExportOutlookBtn/);
assert.match(html, /outlook: exportOutlookResult/);
assert.match(html, /elements\.exportOutlookBtn\.addEventListener\("click", exportOutlookResult\)/);
assert.match(html, /sanitizeDraftFilePart\(getPrimaryProjectName\(\)\)/);
assert.doesNotMatch(html, /slugifyDraftFilePart/);

const resizeContext = {
  window: {
    getComputedStyle(textarea) {
      return textarea.computedStyle;
    },
  },
};
vm.createContext(resizeContext);
for (const functionName of [
  "getAutoGrowHeight",
  "isVerticallyResizableTextarea",
  "autoGrowFeatureTextarea",
]) {
  vm.runInContext(extractFunctionSource(html, functionName), resizeContext);
}

function makeTextarea(resize = "vertical") {
  return {
    computedStyle: { resize, minHeight: "42px" },
    matches: (selector) => selector === "textarea",
    getBoundingClientRect: () => ({ height: 80 }),
    offsetHeight: 80,
    clientHeight: 70,
    scrollHeight: 120,
    style: {},
  };
}

const resizableTextarea = makeTextarea();
resizeContext.autoGrowFeatureTextarea(resizableTextarea);
assert.equal(resizableTextarea.style.height, "130px");

const fixedTextarea = makeTextarea("none");
resizeContext.autoGrowFeatureTextarea(fixedTextarea);
assert.equal(fixedTextarea.style.height, undefined);

assert.match(
  html,
  /function refreshTextareasScrollState\(root = document\)[\s\S]*?querySelectorAll\("textarea"\)/,
  "all textareas should be checked so every vertically draggable field can auto-resize"
);

console.log("draft filename and auto-resize tests passed");
