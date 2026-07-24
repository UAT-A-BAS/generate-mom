const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");

function extractFunctionSource(source, name) {
  const start = source.indexOf(`function ${name}`);
  assert.notEqual(start, -1, `${name} should exist`);

  const bodyStart = source.indexOf("{", start);
  assert.notEqual(bodyStart, -1, `${name} should have a body`);

  let depth = 0;
  for (let index = bodyStart; index < source.length; index += 1) {
    const char = source[index];
    if (char === "{") {
      depth += 1;
    } else if (char === "}") {
      depth -= 1;
      if (depth === 0) {
        return source.slice(start, index + 1);
      }
    }
  }

  throw new Error(`${name} body should close`);
}

const context = {};
Object.assign(context, {
  Blob,
  DataView,
  Date,
  TextDecoder,
  TextEncoder,
  Uint8Array,
});
vm.createContext(context);
vm.runInContext(extractFunctionSource(html, "normalizeTimeDisplayValue"), context);

assert.equal(context.normalizeTimeDisplayValue("08.30"), "08:30");
assert.equal(context.normalizeTimeDisplayValue("8.3"), "08:30");
assert.equal(context.normalizeTimeDisplayValue("09:30"), "09:30");
assert.equal(context.normalizeTimeDisplayValue("9:30"), "09:30");
assert.equal(context.normalizeTimeDisplayValue("08.30 - 9.3"), "08:30 - 09:30");
assert.equal(context.normalizeTimeDisplayValue("17:30-18:30"), "17:30-18:30");

const xlsxFunctions = [
  "escapeXml",
  "getXlsxColumnName",
  "getXlsxCellRef",
  "getXlsxCellColumnIndex",
  "addXlsxCell",
  "estimateXlsxRowHeight",
  "buildXlsxWorksheetXml",
  "buildXlsxWorkbookXml",
  "buildXlsxWorkbookRelsXml",
  "buildXlsxContentTypesXml",
  "buildXlsxStylesXml",
  "buildZipCrcTable",
  "getZipCrc32",
  "getZipDosTime",
  "getZipDosDate",
  "writeZipUint16",
  "writeZipUint32",
  "concatZipParts",
  "buildStoredZip",
  "buildXlsxWorkbookBlob",
];
const workbookSheet = {
  name: "Project MOM",
  rows: [
    [{ ref: "A1", value: "MOM Export", styleId: 1 }],
    [
      { ref: "A2", value: "08:30", styleId: 12 },
      { ref: "B2", value: "Ready & reviewed", styleId: 3 },
    ],
  ],
  rowHeights: [20, 18],
  merges: ["A1:B1"],
  columnWidths: [18, 32],
  maxColumnCount: 2,
};
context.getXlsxSheetFromResult = () => workbookSheet;
vm.runInContext(
  `${xlsxFunctions.map((name) => extractFunctionSource(html, name)).join("\n")}
const xlsxCrcTable = buildZipCrcTable();`,
  context,
);

function readStoredZip(blobBytes) {
  const entries = new Map();
  const decoder = new TextDecoder();
  const view = new DataView(blobBytes.buffer, blobBytes.byteOffset, blobBytes.byteLength);
  let offset = 0;
  while (offset + 4 <= blobBytes.length && view.getUint32(offset, true) === 0x04034b50) {
    const size = view.getUint32(offset + 18, true);
    const nameLength = view.getUint16(offset + 26, true);
    const extraLength = view.getUint16(offset + 28, true);
    const nameStart = offset + 30;
    const dataStart = nameStart + nameLength + extraLength;
    const name = decoder.decode(blobBytes.subarray(nameStart, nameStart + nameLength));
    entries.set(name, blobBytes.subarray(dataStart, dataStart + size));
    offset = dataStart + size;
  }
  return entries;
}

async function testXlsxWorkbook() {
  const workbookBlob = vm.runInContext("buildXlsxWorkbookBlob()", context);
  assert.equal(
    workbookBlob.type,
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  );
  const entries = readStoredZip(new Uint8Array(await workbookBlob.arrayBuffer()));
  assert.deepEqual(
    [...entries.keys()].sort(),
    [
      "[Content_Types].xml",
      "_rels/.rels",
      "docProps/app.xml",
      "docProps/core.xml",
      "xl/_rels/workbook.xml.rels",
      "xl/styles.xml",
      "xl/workbook.xml",
      "xl/worksheets/sheet1.xml",
    ].sort(),
  );
  const decoder = new TextDecoder();
  const workbookXml = decoder.decode(entries.get("xl/workbook.xml"));
  const worksheetXml = decoder.decode(entries.get("xl/worksheets/sheet1.xml"));
  assert.match(workbookXml, /sheet name="Project MOM"/);
  assert.match(worksheetXml, /<mergeCell ref="A1:B1"\/>/);
  assert.match(worksheetXml, /Ready &amp; reviewed/);
  assert.match(worksheetXml, /<c r="A2"[^>]*s="12"/);
}

assert.match(html, /Export XLSX/);
assert.match(html, /mom-meeting-export-\$\{new Date\(\)\.toISOString\(\)\.slice\(0, 10\)\}\.xlsx/);
assert.match(html, /application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet/);
assert.match(html, /function buildXlsxWorkbookBlob/);
assert.match(html, /function exportXlsxResult\(\)/);
assert.doesNotMatch(html, /exportXlsResult/);
assert.match(html, /exportXlsxBtn/);
assert.match(html, /previewDrawerExportXlsxBtn/);
assert.doesNotMatch(html, /exportXlsBtn/);
assert.doesNotMatch(html, /previewDrawerExportXlsBtn/);
assert.match(html, /function getXlsxSheetFromResult\(\)/);
assert.doesNotMatch(html, /function getXlsxSheetsFromResult\(\)/);
assert.match(html, /function getXlsxSheetName\(\)/);
assert.match(html, /getPrimaryProjectName\(\)/);
assert.match(html, /showGridLines="0"/);
assert.match(html, /getXlsxColumnWidthsForTable/);
assert.match(html, /rgb="FFF7EBC2"/);
assert.match(html, /rgb="FF9BD255"/);
assert.match(html, /rgb="FFD7E3F7"/);
assert.match(html, /function addXlsxCell\(/);
assert.match(html, /columnOffset === 0 && rowOffset === 0/);
assert.match(html, /getXlsxAlignmentStyleId/);
assert.match(html, /function getXlsxFlaggedStyleId/);
assert.match(html, /value === "-"/);
assert.match(html, /getXlsxFlaggedStyleId\(baseStyleId\)/);
assert.match(html, /horizontal="center" vertical="center"/);
assert.match(html, /vertical="top" wrapText="1"/);
assert.match(html, /cellXfs count="18"/);
assert.match(html, /content: buildXlsxWorksheetXml\(sheet\)/);
assert.doesNotMatch(html, /sheets\.map\(\(sheet, index\) => \(\{\s*path: `xl\/worksheets\/sheet\$\{index \+ 1\}\.xml`/);

testXlsxWorkbook()
  .then(() => console.log("mom-export tests passed"))
  .catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
