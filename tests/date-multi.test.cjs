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
  "normalizeDateSegments",
  "isDateSelectedInSegments",
  "findNearestSingleDateSegmentIndex",
  "pairDateWithSingleSegment",
  "addDateRangeToSegments",
  "toggleDateInSegments",
  "removeDateFromSegments",
  "getDatePickerPreviewRange",
  "formatDatePickerRangeLabel",
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
  "21-05-2026 - 24-05-2026"
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
  [{ start: "21-05-2026", end: "24-05-2026" }]
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
        [{ start: "08-06-2026", end: "12-06-2026" }],
        "2026-06-08"
      )
    )
  ),
  [{ start: "09-06-2026", end: "12-06-2026" }]
);

assert.deepEqual(
  JSON.parse(
    JSON.stringify(
      context.pairDateWithSingleSegment(
        [{ start: "12-06-2026", end: "" }],
        "2026-06-07"
      )
    )
  ),
  [{ start: "07-06-2026", end: "12-06-2026" }]
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

const forwardRange = context.addDateRangeToSegments([], "2026-06-03", "2026-06-07");
assert.deepEqual(JSON.parse(JSON.stringify(forwardRange)), [
  { start: "03-06-2026", end: "07-06-2026" },
]);

assert.deepEqual(
  JSON.parse(JSON.stringify(context.addDateRangeToSegments([], "2026-06-07", "2026-06-03"))),
  JSON.parse(JSON.stringify(forwardRange))
);

assert.deepEqual(
  JSON.parse(
    JSON.stringify(context.addDateRangeToSegments(forwardRange, "2026-06-10", "2026-06-10"))
  ),
  [
    { start: "03-06-2026", end: "07-06-2026" },
    { start: "10-06-2026", end: "" },
  ]
);

assert.deepEqual(
  JSON.parse(
    JSON.stringify(
      context.addDateRangeToSegments(
        [{ start: "01-06-2026", end: "" }],
        "2026-06-03",
        "2026-06-05"
      )
    )
  ),
  [
    { start: "01-06-2026", end: "" },
    { start: "03-06-2026", end: "05-06-2026" },
  ]
);

assert.deepEqual(
  JSON.parse(JSON.stringify(context.toggleDateInSegments(forwardRange, "2026-06-05"))),
  [
    { start: "03-06-2026", end: "04-06-2026" },
    { start: "06-06-2026", end: "07-06-2026" },
  ]
);

assert.deepEqual(
  JSON.parse(
    JSON.stringify(
      context.normalizeDateSegments([
        { start: "03-06-2026", end: "07-06-2026" },
        { start: "05-06-2026", end: "10-06-2026" },
      ])
    )
  ),
  [{ start: "03-06-2026", end: "10-06-2026" }]
);

assert.deepEqual(
  JSON.parse(JSON.stringify(context.getDatePickerPreviewRange("2026-06-10", "2026-06-05"))),
  { start: "2026-06-05", end: "2026-06-10" }
);
assert.equal(
  context.formatDatePickerRangeLabel("2026-06-10", "2026-06-05"),
  "05-06-2026 - 10-06-2026"
);

assert.match(html, /document\.addEventListener\("pointerdown", handleDatePickerPointerDown\)/);
assert.match(html, /document\.addEventListener\("pointercancel", handleDatePickerPointerCancel\)/);
assert.match(html, /activeDatePicker\?\.picker === picker && activeDatePicker\.dragStart/);
assert.match(html, /aria-pressed="\$\{isSelected \? "true" : "false"\}"/);
assert.match(html, /Klik satu tanggal · tahan dan geser untuk rentang/);
assert.match(html, /Lepas untuk menambahkan/);

console.log("date multi tests passed");
