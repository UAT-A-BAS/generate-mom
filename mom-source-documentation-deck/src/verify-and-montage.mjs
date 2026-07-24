import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Canvas, loadImage } from "skia-canvas";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const scratch = path.join(root, "scratch");
const output = path.join(root, "output");
const pptxPath = path.join(output, "MOM_Source_Documentation.pptx");
const manifestPath = path.join(scratch, "manifest.json");
const reportPath = path.join(scratch, "quality-report.json");
const relative = (filePath) => path.relative(root, filePath).replaceAll("\\", "/");
const issues = [];

function readJson(filePath) {
  try {
    return JSON.parse(fs.readFileSync(filePath, "utf8"));
  } catch (error) {
    issues.push(`${relative(filePath)}: ${error.message}`);
    return null;
  }
}

function listZipEntries(filePath) {
  try {
    const buffer = fs.readFileSync(filePath);
    let endOffset = -1;
    for (let index = buffer.length - 22; index >= Math.max(0, buffer.length - 65557); index -= 1) {
      if (buffer.readUInt32LE(index) === 0x06054b50) {
        endOffset = index;
        break;
      }
    }
    if (endOffset < 0) throw new Error("ZIP end record was not found");

    const entryCount = buffer.readUInt16LE(endOffset + 10);
    let offset = buffer.readUInt32LE(endOffset + 16);
    const entries = [];
    for (let index = 0; index < entryCount; index += 1) {
      if (buffer.readUInt32LE(offset) !== 0x02014b50) {
        throw new Error(`invalid central-directory entry at offset ${offset}`);
      }
      const nameLength = buffer.readUInt16LE(offset + 28);
      const extraLength = buffer.readUInt16LE(offset + 30);
      const commentLength = buffer.readUInt16LE(offset + 32);
      entries.push(buffer.subarray(offset + 46, offset + 46 + nameLength).toString("utf8"));
      offset += 46 + nameLength + extraLength + commentLength;
    }
    return entries;
  } catch (error) {
    issues.push(`${relative(filePath)}: ${error.message}`);
    return [];
  }
}

const manifest = readJson(manifestPath);
const layoutFiles = fs.readdirSync(scratch).filter((f) => f.endsWith(".layout.json")).sort();
const pngFiles = fs.readdirSync(scratch).filter((f) => /^slide-\d+\.png$/.test(f)).sort();
const expectedLayouts = manifest?.artifacts?.map((entry) => entry.layout).sort() || [];
const expectedPngs = manifest?.artifacts?.map((entry) => entry.png).sort() || [];

if (!manifest || !Number.isInteger(manifest.slideCount) || manifest.slideCount < 1) {
  issues.push(`${relative(manifestPath)}: missing a positive slideCount`);
}
if (manifest && manifest.artifacts?.length !== manifest.slideCount) {
  issues.push(`${relative(manifestPath)}: artifact count does not match slideCount`);
}
if (JSON.stringify(layoutFiles) !== JSON.stringify(expectedLayouts)) {
  issues.push(`layout artifact set mismatch; expected ${expectedLayouts.join(", ")}, found ${layoutFiles.join(", ")}`);
}
if (JSON.stringify(pngFiles) !== JSON.stringify(expectedPngs)) {
  issues.push(`PNG artifact set mismatch; expected ${expectedPngs.join(", ")}, found ${pngFiles.join(", ")}`);
}

for (const file of layoutFiles) {
  const layout = readJson(path.join(scratch, file));
  if (!layout) continue;
  for (const element of layout.elements || []) {
    if (!element.textPreview) continue;
    const bbox = element.bbox || [];
    const textLayout = element.textLayout || {};
    const renderedHeight = textLayout.height || 0;
    const maxLineWidth = Math.max(0, ...(textLayout.lines || []).map((line) => line.width || 0));
    if (renderedHeight && bbox[3] && renderedHeight > bbox[3] + 2) {
      issues.push(`${file} ${element.name || element.id} height ${renderedHeight}>${bbox[3]}`);
    }
    if (maxLineWidth && bbox[2] && maxLineWidth > bbox[2] + 4) {
      issues.push(`${file} ${element.name || element.id} width ${maxLineWidth}>${bbox[2]}`);
    }
  }
}

const scale = 0.18;
const gap = 18;
const cols = 4;
const thumbW = Math.round(1920 * scale);
const thumbH = Math.round(1080 * scale);
const rows = Math.max(1, Math.ceil(pngFiles.length / cols));
const canvas = new Canvas(cols * thumbW + (cols + 1) * gap, rows * thumbH + (rows + 1) * gap);
const ctx = canvas.getContext("2d");
ctx.fillStyle = "#f1f5f9";
ctx.fillRect(0, 0, canvas.width, canvas.height);

for (let i = 0; i < pngFiles.length; i += 1) {
  try {
    const img = await loadImage(path.join(scratch, pngFiles[i]));
    const x = gap + (i % cols) * (thumbW + gap);
    const y = gap + Math.floor(i / cols) * (thumbH + gap);
    ctx.drawImage(img, x, y, thumbW, thumbH);
    ctx.fillStyle = "#111827";
    ctx.font = "18px Aptos";
    ctx.fillText(String(i + 1).padStart(2, "0"), x + 10, y + 24);
  } catch (error) {
    issues.push(`${relative(path.join(scratch, pngFiles[i]))}: ${error.message}`);
  }
}

const montagePath = path.join(scratch, "contact-sheet.png");
fs.writeFileSync(montagePath, await canvas.toBuffer("png"));

const pptxEntries = listZipEntries(pptxPath);
const pptxSlides = pptxEntries.filter((entry) => /^ppt\/slides\/slide\d+\.xml$/.test(entry));
for (const requiredEntry of ["[Content_Types].xml", "ppt/presentation.xml"]) {
  if (!pptxEntries.includes(requiredEntry)) issues.push(`${relative(pptxPath)}: missing ${requiredEntry}`);
}
if (manifest && pptxSlides.length !== manifest.slideCount) {
  issues.push(`${relative(pptxPath)}: expected ${manifest.slideCount} slides, found ${pptxSlides.length}`);
}

const report = {
  workspace: ".",
  pptx: relative(pptxPath),
  checks: {
    artifact_inventory: {
      expected_slides: manifest?.slideCount || 0,
      layout_slides: layoutFiles.length,
      png_slides: pngFiles.length,
    },
    pptx_package: {
      entries: pptxEntries.length,
      slides: pptxSlides.length,
    },
    text_fit: {
      issues: issues.filter((issue) => issue.includes(" height ") || issue.includes(" width ")),
    },
  },
  failures: issues,
  warnings: [],
  contactSheet: relative(montagePath),
};
fs.writeFileSync(reportPath, `${JSON.stringify(report, null, 2)}\n`, "utf8");
console.log(JSON.stringify(report, null, 2));

if (issues.length) process.exitCode = 1;
