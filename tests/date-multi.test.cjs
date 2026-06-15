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

const functionNames = [
  "formatDateForDisplay",
  "parseDateRangeValue",
  "normalizeDateSegment",
  "parseDateMultiValue",
  "formatDateSegmentForDisplay",
  "formatDateRangeForDisplay",
  "formatDateForNative",
  "nextIsoDate",
  "expandDateSegmentsToIsoValues",
  "collapseIsoValuesToDateSegments",
  "sortDateSegments",
  "normalizeDateSegments",
  "isDateSelectedInSegments",
  "addDateRangeToSegments",
  "removeDateFromSegments",
];
const context = {};
vm.createContext(context);
vm.runInContext(functionNames.map((name) => extractFunctionSource(html, name)).join("\n"), context);

assert.equal(
  context.formatDateRangeForDisplay(
    "26-05-2026, 21-05-2026 - 22-05-2026, 21-05-2026"
  ),
  "21-05-2026 - 22-05-2026, 26-05-2026"
);

assert.equal(
  context.formatDateRangeForDisplay(
    "21-05-2026 - 22-05-2026, 23-05-2026 - 24-05-2026"
  ),
  "21-05-2026 - 22-05-2026, 23-05-2026 - 24-05-2026"
);

assert.deepEqual(
  JSON.parse(
    JSON.stringify(
      context.normalizeDateSegments([
        { start: "26-05-2026", end: "" },
        { start: "21-05-2026", end: "22-05-2026" },
        { start: "21-05-2026", end: "" },
      ])
    )
  ),
  [
    { start: "21-05-2026", end: "22-05-2026" },
    { start: "26-05-2026", end: "" },
  ]
);

assert.deepEqual(
  JSON.parse(
    JSON.stringify(
      context.addDateRangeToSegments(
        [{ start: "21-05-2026", end: "22-05-2026" }],
        "2026-05-22",
        "2026-05-24"
      )
    )
  ),
  [
    { start: "21-05-2026", end: "22-05-2026" },
    { start: "23-05-2026", end: "24-05-2026" },
  ]
);

assert.deepEqual(
  JSON.parse(
    JSON.stringify(
      context.removeDateFromSegments(
        [{ start: "21-05-2026", end: "24-05-2026" }],
        "2026-05-22"
      )
    )
  ),
  [
    { start: "21-05-2026", end: "" },
    { start: "23-05-2026", end: "24-05-2026" },
  ]
);

assert.deepEqual(
  JSON.parse(
    JSON.stringify(
      context.removeDateFromSegments(
        [{ start: "21-05-2026", end: "24-05-2026" }],
        "2026-05-21"
      )
    )
  ),
  [{ start: "22-05-2026", end: "24-05-2026" }]
);

assert.equal(
  context.isDateSelectedInSegments(
    [{ start: "21-05-2026", end: "23-05-2026" }],
    "2026-05-22"
  ),
  true
);

assert.deepEqual(
  JSON.parse(
    JSON.stringify(
      context.removeDateFromSegments([{ start: "21-05-2026", end: "" }], "2026-05-21")
    )
  ),
  []
);

console.log("date multi tests passed");
