import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import {
  Presentation,
  PresentationFile,
  row,
  column,
  grid,
  layers,
  panel,
  text,
  shape,
  rule,
  fill,
  hug,
  fixed,
  wrap,
  grow,
  fr,
  auto,
} from "@oai/artifact-tool";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const OUTPUT_DIR = path.join(ROOT, "output");
const SCRATCH = path.join(ROOT, "scratch");
const OUT = path.join(OUTPUT_DIR, "MOM_Source_Documentation.pptx");

const W = 1920;
const H = 1080;
const C = {
  bg: "#F7F9FC",
  paper: "#FFFFFF",
  ink: "#152033",
  muted: "#56657A",
  line: "#D8DEE8",
  navy: "#12263A",
  blue: "#2563EB",
  cyan: "#06A7B8",
  green: "#4D9F6F",
  amber: "#E9A23B",
  red: "#C44E52",
  lavender: "#7C6ADE",
  table1: "#F7EBC2",
  table2: "#9BD255",
  table3: "#D7E3F7",
  lampiran: "#E9EDF3",
  flag: "#F8FC46",
};

const FONT = "Aptos";
const MONO = "Cascadia Mono";

const deck = Presentation.create({ slideSize: { width: W, height: H } });

function t(value, options = {}) {
  return text(value, {
    width: fill,
    height: hug,
    style: {
      typeface: FONT,
      fontSize: 28,
      color: C.ink,
      lineSpacing: 1.05,
      ...options.style,
    },
    ...options,
  });
}

function label(value, color = C.blue) {
  return panel(
    {
      width: hug,
      height: fixed(38),
      padding: { x: 16, y: 7 },
      fill: `${color}18`,
      stroke: `${color}55`,
      borderRadius: 8,
    },
    t(value, {
      width: hug,
      style: { fontSize: 15, bold: true, color },
    }),
  );
}

function bulletList(items, opts = {}) {
  return column(
    { width: fill, height: hug, gap: opts.gap ?? 12 },
    items.map((item) =>
      row(
        { width: fill, height: hug, gap: 12, alignItems: "start" },
        [
          shape({
            geometry: "ellipse",
            width: fixed(9),
            height: fixed(9),
            fill: opts.dotColor ?? C.blue,
            stroke: "none",
            margin: { top: 13 },
          }),
          t(item, {
            width: fill,
            style: { fontSize: opts.fontSize ?? 24, color: opts.color ?? C.ink },
          }),
        ],
      ),
    ),
  );
}

function box(title, body, color = C.blue, opts = {}) {
  return panel(
    {
      name: opts.name,
      width: opts.width ?? fill,
      height: opts.height ?? hug,
      padding: { x: 24, y: 20 },
      fill: opts.fill ?? C.paper,
      stroke: opts.stroke ?? C.line,
      borderRadius: 8,
    },
    column(
      { width: fill, height: hug, gap: 9 },
      [
        row(
          { width: fill, height: hug, gap: 10, alignItems: "center" },
          [
            shape({ geometry: "rect", width: fixed(8), height: fixed(28), fill: color, stroke: "none" }),
            t(title, { style: { fontSize: opts.titleSize ?? 24, bold: true, color: C.ink } }),
          ],
        ),
        typeof body === "string"
          ? t(body, { style: { fontSize: opts.bodySize ?? 21, color: C.muted } })
          : body,
      ],
    ),
  );
}

function arrow(color = C.muted) {
  return t("->", {
    width: fixed(58),
    style: { fontSize: 30, bold: true, color, alignment: "center" },
  });
}

function footer(slide, index, source = "Source: index.html and ExportMOMToDraft.bas") {
  slide.compose(
    row(
      { width: fill, height: fixed(38), padding: { x: 74, y: 0 }, alignItems: "center" },
      [
        t(source, {
          width: grow(1),
          style: { fontSize: 13, color: "#7A8699" },
        }),
        t(String(index).padStart(2, "0"), {
          width: fixed(52),
          style: { fontSize: 14, bold: true, color: "#7A8699", alignment: "right" },
        }),
      ],
    ),
    { frame: { left: 0, top: 1024, width: W, height: 44 }, baseUnit: 8 },
  );
}

function background(slide) {
  slide.compose(
    layers(
      { width: fill, height: fill },
      [
        shape({ name: "bg", geometry: "rect", width: fill, height: fill, fill: C.bg, stroke: "none" }),
        shape({ name: "top-band", geometry: "rect", width: fill, height: fixed(12), fill: C.navy, stroke: "none" }),
      ],
    ),
    { frame: { left: 0, top: 0, width: W, height: H }, baseUnit: 8 },
  );
}

function standardSlide(title, subtitle, bodyNode, notes) {
  const slide = deck.slides.add();
  background(slide);
  slide.compose(
    grid(
      {
        name: "slide-root",
        width: fill,
        height: fill,
        rows: [auto, fr(1)],
        columns: [fr(1)],
        padding: { x: 74, y: 64 },
        rowGap: 34,
      },
      [
        column(
          { name: "title-stack", width: fill, height: hug, gap: 12 },
          [
            t(title, {
              name: "slide-title",
              style: { fontSize: 48, bold: true, color: C.ink },
            }),
            subtitle
              ? t(subtitle, {
                  name: "slide-subtitle",
                  width: wrap(1300),
                  style: { fontSize: 21, color: C.muted },
                })
              : rule({ width: fixed(160), stroke: C.blue, weight: 4 }),
          ],
        ),
        bodyNode,
      ],
    ),
    { frame: { left: 0, top: 0, width: W, height: 1020 }, baseUnit: 8 },
  );
  footer(slide, deck.slides.items.length);
  if (notes) slide.speakerNotes.setText(notes);
  return slide;
}

function tableGrid(columns, rows, colWidths, opts = {}) {
  const cells = [];
  columns.forEach((h, i) => {
    cells.push(
      panel(
        {
          width: fill,
          height: fixed(opts.headerH ?? 54),
          padding: { x: 14, y: 12 },
          fill: opts.headerFill ?? C.navy,
          stroke: opts.headerStroke ?? C.navy,
          borderRadius: 0,
        },
        t(h, { style: { fontSize: opts.headerSize ?? 17, bold: true, color: "#FFFFFF" } }),
      ),
    );
  });
  rows.forEach((r, rowIndex) => {
    r.forEach((cell, i) => {
      cells.push(
        panel(
          {
            width: fill,
            height: fixed(opts.rowH ?? 68),
            padding: { x: 14, y: 11 },
            fill: rowIndex % 2 ? "#FAFBFD" : C.paper,
            stroke: C.line,
            borderRadius: 0,
          },
          t(cell, {
            style: {
              fontSize: opts.bodySize ?? 17,
              color: i === 0 ? C.ink : C.muted,
              bold: i === 0,
            },
          }),
        ),
      );
    });
  });
  return grid(
    {
      width: fill,
      height: hug,
      columns: colWidths,
      rows: [fixed(opts.headerH ?? 54), ...rows.map(() => fixed(opts.rowH ?? 68))],
      columnGap: 0,
      rowGap: 0,
    },
    cells,
  );
}

function cover() {
  const slide = deck.slides.add();
  slide.compose(
    layers(
      { width: fill, height: fill },
      [
        shape({ geometry: "rect", width: fill, height: fill, fill: "#F4F7FB", stroke: "none" }),
        shape({ geometry: "rect", width: fixed(610), height: fill, fill: C.navy, stroke: "none" }),
        shape({ geometry: "rect", width: fixed(28), height: fill, fill: C.blue, stroke: "none", margin: { left: 610 } }),
      ],
    ),
    { frame: { left: 0, top: 0, width: W, height: H }, baseUnit: 8 },
  );
  slide.compose(
    grid(
      {
        width: fill,
        height: fill,
        columns: [fixed(570), fr(1)],
        rows: [fr(1)],
        columnGap: 92,
        padding: { x: 74, y: 86 },
      },
      [
        column(
          { width: fill, height: fill, gap: 30, justifyContent: "space-between" },
          [
            column(
              { width: fill, height: hug, gap: 28 },
              [
                label("SOURCE DOCUMENTATION", C.cyan),
                t("MOM Generator\nHTML + Outlook VBA", {
                  name: "cover-title",
                  style: { fontSize: 62, bold: true, color: "#FFFFFF", lineSpacing: 0.94 },
                }),
                t("Editable technical presentation\nfrom index.html and ExportMOMToDraft.bas", {
                  style: { fontSize: 24, color: "#D7E0ED", lineSpacing: 1.12 },
                }),
              ],
            ),
            t("Prepared for office review", { style: { fontSize: 18, color: "#A9B8CA" } }),
          ],
        ),
        column(
          { width: fill, height: fill, gap: 34, justifyContent: "center" },
          [
            t("One workflow, two artifacts", {
              style: { fontSize: 38, bold: true, color: C.ink },
            }),
            row(
              { width: fill, height: hug, gap: 20, alignItems: "center" },
              [
                box("HTML app", "Capture MOM data, validate, preview, export Outlook-safe HTML.", C.blue, {
                  height: fixed(168),
                }),
                arrow(C.blue),
                box("VBA macro", "Read exported HTML, create Outlook draft, repair table rendering.", C.green, {
                  height: fixed(168),
                }),
                arrow(C.green),
                box("Outlook draft", "Editable email body with tables ready for review and send.", C.amber, {
                  height: fixed(168),
                }),
              ],
            ),
            panel(
              {
                width: fill,
                height: fixed(170),
                padding: { x: 28, y: 24 },
                fill: C.paper,
                stroke: C.line,
                borderRadius: 8,
              },
              column(
                { width: fill, height: hug, gap: 10 },
                [
                  t("Cover thesis", { style: { fontSize: 19, bold: true, color: C.blue } }),
                  t(
                    "The browser file is the MOM authoring and export surface. The Outlook VBA file is the delivery bridge that turns the exported HTML into a saved, displayed Outlook draft.",
                    { style: { fontSize: 25, color: C.ink } },
                  ),
                ],
              ),
            ),
          ],
        ),
      ],
    ),
    { frame: { left: 0, top: 0, width: W, height: H }, baseUnit: 8 },
  );
  slide.speakerNotes.setText(
    "Open with the system boundary: this deck documents the delivered source files, not a new generator. The key integration point is the exported Outlook HTML file consumed by the VBA macro.",
  );
}

cover();

standardSlide(
  "Executive Summary",
  "The sources implement a browser-based MOM authoring tool plus an Outlook automation macro.",
  grid(
    { width: fill, height: fill, columns: [fr(1), fr(1)], rows: [auto, auto], columnGap: 28, rowGap: 24 },
    [
      box("HTML file role", "Single-page MOM Generator: form sections, validation, preview rendering, JSON draft save/load, XLSX export, Outlook HTML export.", C.blue),
      box("VBA file role", "Outlook macro: asks user for generated HTML and project name, creates a draft email, then corrects table header formatting inside Outlook WordEditor.", C.green),
      box("Integration", "User exports an Outlook-specific HTML document from the browser, then runs the macro and selects that file.", C.amber),
      box("Output", "A saved and displayed Outlook draft with subject, greeting, MOM content tables, closing copy, and Outlook-safe formatting.", C.lavender),
    ],
  ),
  "Main message: the files split responsibility cleanly. Browser handles data composition and export; Outlook VBA handles local mail-client injection and final Word table fixes.",
);

standardSlide(
  "Purpose Of The HTML File",
  "index.html is the front-end application and export engine for MOM preparation data.",
  row(
    { width: fill, height: fill, gap: 30, alignItems: "stretch" },
    [
      column(
        { width: grow(1), height: fill, gap: 22 },
        [
          box("Capture", "Collects certification, pre-implementation checklist, implementation strategy, optional verification appendix, and meeting notes.", C.blue),
          box("Control", "Validates required fields, normalizes dates/times, flags fields, merges repeated table values, and keeps dynamic rows in state.", C.green),
          box("Export", "Generates preview HTML, standalone HTML, Outlook-safe HTML, JSON draft files, and XLSX files.", C.amber),
        ],
      ),
      panel(
        { width: fixed(560), height: fill, padding: { x: 30, y: 28 }, fill: C.navy, stroke: C.navy, borderRadius: 8 },
        column(
          { width: fill, height: fill, gap: 20 },
          [
            t("Primary user actions", { style: { fontSize: 24, bold: true, color: "#FFFFFF" } }),
            bulletList(
              [
                "Load Draft Data",
                "Smart Import Raw Data",
                "Preview Table",
                "Export XLSX",
                "Export to Outlook",
                "Save Draft Data",
                "Clear All",
              ],
              { fontSize: 22, color: "#DCE6F3", dotColor: C.cyan, gap: 14 },
            ),
          ],
        ),
      ),
    ],
  ),
);

standardSlide(
  "Purpose Of The VBA File",
  "ExportMOMToDraft.bas is a targeted Outlook automation bridge.",
  grid(
    { width: fill, height: fill, columns: [fr(1.1), fr(0.9)], rows: [fr(1)], columnGap: 34 },
    [
      column(
        { width: fill, height: fill, gap: 22 },
        [
          box("Select input", "Opens a native Windows file picker in Downloads\\ExportMOM and accepts .html/.htm files.", C.blue),
          box("Create draft", "Reads the HTML as UTF-8, creates an Outlook MailItem, sets HTMLBody, subject, saves, and displays it.", C.green),
          box("Repair rendering", "Uses the displayed draft's WordEditor to identify key tables and adjust header rows and widths.", C.amber),
        ],
      ),
      tableGrid(
        ["Dependency", "Use"],
        [
          ["Outlook VBA", "Runs macro in Outlook session"],
          ["comdlg32.dll", "Native file picker"],
          ["ADODB.Stream", "UTF-8 file read"],
          ["WordEditor", "Table formatting after display"],
          ["VBScript.RegExp", "Prepatch table2 header HTML"],
        ],
        [fr(0.75), fr(1.25)],
        { headerFill: C.navy, bodySize: 18, rowH: 76 },
      ),
    ],
  ),
  "Stress that the macro is not a generator. It consumes the browser-generated export and adapts it to Outlook's HTML/Word rendering model.",
);

standardSlide(
  "How The Files Work Together",
  "The HTML file produces Outlook-ready content; the VBA file injects and finalizes it inside Outlook.",
  column(
    { width: fill, height: fill, gap: 30, justifyContent: "center" },
    [
      row(
        { width: fill, height: hug, gap: 18, alignItems: "center" },
        [
          box("1. Browser form", "User enters MOM data and clicks Preview Table.", C.blue, { height: fixed(142) }),
          arrow(C.blue),
          box("2. Export HTML", "exportOutlookResult() downloads mom-outlook-export-*.html.", C.cyan, { height: fixed(142) }),
          arrow(C.cyan),
          box("3. VBA macro", "PickHtmlFile() selects the exported file.", C.green, { height: fixed(142) }),
          arrow(C.green),
          box("4. Outlook draft", "HTMLBody is set, then headers/widths are fixed.", C.amber, { height: fixed(142) }),
        ],
      ),
      panel(
        { width: fill, height: fixed(150), padding: { x: 30, y: 24 }, fill: "#EFF6FF", stroke: "#BBD4FF", borderRadius: 8 },
        t(
          "Integration contract: the browser export must contain a complete HTML email document with predictable table classes: table1, table2, table3, and optional lampiran-table. The VBA macro depends on that structure to recognize and format tables.",
          { style: { fontSize: 27, color: C.ink } },
        ),
      ),
    ],
  ),
  "Use this slide to explain the handoff: HTML creates file artifact; VBA uses local Outlook object model. No server or API is involved.",
);

standardSlide(
  "End-To-End Workflow",
  "User path from data entry to email draft.",
  grid(
    { width: fill, height: fill, columns: [fr(1), fixed(72), fr(1), fixed(72), fr(1)], rows: [auto, auto], columnGap: 4, rowGap: 34 },
    [
      box("Author", "Fill Sertifikasi, Checklist, Strategi Implementasi, optional Lampiran, and Notes Meeting.", C.blue, { height: fixed(190) }),
      arrow(C.blue),
      box("Preview", "validateForm() checks completeness; buildResultMarkup() renders MOM tables into resultArea.", C.cyan, { height: fixed(190) }),
      arrow(C.cyan),
      box("Export", "buildOutlookEmailDocument() wraps Outlook-safe table markup in greeting and closing email copy.", C.green, { height: fixed(190) }),
      box("Select", "Outlook macro opens file picker, user chooses the generated Outlook HTML file.", C.amber, { height: fixed(190) }),
      arrow(C.amber),
      box("Draft", "Macro creates MailItem, sets subject and HTMLBody, saves, displays, then applies Word table fixes.", C.lavender, { height: fixed(190) }),
      arrow(C.lavender),
      box("Review", "User reviews/edit draft in Outlook before sending to meeting stakeholders.", C.red, { height: fixed(190) }),
    ],
  ),
  "Point out the manual checkpoints: preview before export, file selection in macro, final review in Outlook.",
);

standardSlide(
  "Key HTML Structure",
  "The page is organized around input panels, overlays, preview area, and one large script block.",
  tableGrid(
    ["Section", "Important IDs / Classes", "Purpose"],
    [
      ["Hero", "loadDraftBtn, draftFileInput", "Load saved JSON draft data into the app."],
      ["Sertifikasi", "table1Projects, smartImportBtn", "Dynamic project, release, BPRO, changes, blueprint, UAT data."],
      ["Checklist", "table2Rows", "Fixed readiness checklist with status, PIC, target, notes."],
      ["Strategi", "table3Rows, addTable3Row", "Implementation schedule rows with date, time, PIC, status."],
      ["Lampiran", "lampiranEnabled, lampiranRows", "Optional verification scenario appendix."],
      ["Preview/export", "resultArea, exportOutlookBtn", "Generated output and export commands."],
      ["Overlays", "previewDrawer, memoPanel, smartImportModal", "Auxiliary preview, notes, and raw import workflows."],
    ],
    [fr(0.7), fr(1.15), fr(1.35)],
    { headerFill: C.navy, bodySize: 16, rowH: 70 },
  ),
);

standardSlide(
  "HTML Logic Map",
  "Key JavaScript modules are implemented as grouped functions inside index.html.",
  grid(
    { width: fill, height: fill, columns: [fr(1), fr(1), fr(1)], rows: [auto, auto], columnGap: 22, rowGap: 22 },
    [
      box("State factories", "createEmptyTable1Project, createEmptyRelationPackage, createEmptyTable3Row, createEmptyLampiranDateGroup", C.blue),
      box("Format helpers", "escapeHtml, formatDateForDisplay, formatDateIndonesian, normalizeTimeDisplayValue, formatMultilineWithLinks", C.cyan),
      box("Dynamic rendering", "renderTable1Projects, renderTable2Rows, renderTable3Rows, renderLampiranSection", C.green),
      box("Validation", "validateForm plus missing-field helpers for active project, release, feature, checklist, strategy, appendix rows", C.amber),
      box("Output rendering", "renderTable1Result, renderChecklistResult, renderTable3Result, renderLampiranResult, buildResultMarkup", C.lavender),
      box("Exports", "buildExportMarkup, buildOutlookEmailDocument, exportXlsxResult, exportOutlookResult, saveDraftData, loadDraftFile", C.red),
    ],
  ),
  "The HTML file is function-heavy but internally coherent: input state -> validation -> result markup -> export variants.",
);

standardSlide(
  "Key VBA Functions",
  "The macro has one public entry point and focused private helpers.",
  tableGrid(
    ["Function / Sub", "Responsibility", "Notable behavior"],
    [
      ["ExportMOMToDraft", "Main orchestration", "Creates folder, prompts file/project, reads HTML, creates Outlook draft."],
      ["PickHtmlFile", "Input file selection", "Uses GetOpenFileNameW with HTML filter and default Downloads\\ExportMOM path."],
      ["ReadTextFileUtf8", "File read", "Uses ADODB.Stream with Charset = utf-8."],
      ["FixTable2HeaderForOutlook", "Pre-injection HTML fix", "Regex replaces table2 thead first row with Outlook-friendly header HTML."],
      ["FixDisplayedTableHeaders", "Post-display Word fix", "Loops WordEditor tables, detects known table types, applies row/width fixes."],
      ["IsCertificationTable / IsChecklistTable / IsStrategyTable", "Table recognition", "Detects tables by header keywords after cleaning Word cell text."],
      ["FixWordTableHeaderRow", "Header formatting", "Exact height, no page break, bold, 12 pt, centered, vertical middle."],
      ["SetWordTableColumnWidth", "Column sizing", "Sets target checklist Target width and strategy Date width in points."],
    ],
    [fr(0.9), fr(1.05), fr(1.35)],
    { headerFill: C.navy, bodySize: 15.5, rowH: 66 },
  ),
  "This slide is the implementation inventory for reviewers. Note that table recognition is content-based, not metadata-based.",
);

standardSlide(
  "Data Flow Diagram",
  "Data moves from user-entered browser state to a local HTML file, then into an Outlook MailItem.",
  column(
    { width: fill, height: fill, gap: 30, justifyContent: "center" },
    [
      row(
        { width: fill, height: hug, gap: 18, alignItems: "center" },
        [
          box("User inputs", "Project, release, BPRO, changes, checklist, strategy, lampiran, memo", C.blue, { height: fixed(128) }),
          arrow(C.blue),
          box("JS state", "table1ProjectsState, table3State, lampiranState, checklistRows", C.cyan, { height: fixed(128) }),
          arrow(C.cyan),
          box("Result markup", "buildResultMarkup() creates .result-document with tables", C.green, { height: fixed(128) }),
        ],
      ),
      row(
        { width: fill, height: hug, gap: 18, alignItems: "center" },
        [
          box("Export transform", "buildExportMarkup('outlook') applies Outlook-safe layout", C.amber, { height: fixed(128) }),
          arrow(C.amber),
          box("HTML file", "mom-outlook-export-<project>-<date>.html", C.lavender, { height: fixed(128) }),
          arrow(C.lavender),
          box("Outlook draft", "MailItem.HTMLBody + WordEditor table adjustments", C.red, { height: fixed(128) }),
        ],
      ),
      panel(
        { width: fill, height: fixed(112), padding: { x: 26, y: 22 }, fill: C.paper, stroke: C.line, borderRadius: 8 },
        t("Side output: XLSX export builds SpreadsheetML parts and a stored ZIP directly in the browser. Draft data saves as JSON version mom-generator-draft-v1.", {
          style: { fontSize: 24, color: C.muted },
        }),
      ),
    ],
  ),
);

standardSlide(
  "Technical Architecture",
  "All execution is local: browser, downloaded files, Outlook VBA, and Outlook Word rendering.",
  grid(
    { width: fill, height: fill, columns: [fr(1), fr(1), fr(1)], rows: [fr(1), fr(1)], columnGap: 24, rowGap: 24 },
    [
      box("Browser UI layer", "HTML, CSS, DOM events, dynamic tables, modal overlays, date picker, preview drawer.", C.blue, { height: fill }),
      box("Browser logic layer", "State factories, parsers, validators, renderers, draft persistence, XLSX ZIP builder.", C.cyan, { height: fill }),
      box("Browser export layer", "Inline styles, Excel/Outlook transforms, standalone document wrappers, download blobs.", C.green, { height: fill }),
      box("Local file boundary", "Generated .html, .xlsx, and .json files are user-managed downloads.", C.amber, { height: fill }),
      box("Outlook automation", "Application.CreateItem, MailItem.HTMLBody, Save, Display.", C.lavender, { height: fill }),
      box("WordEditor patch layer", "Displayed email tables are detected and corrected through Outlook's embedded Word editor.", C.red, { height: fill }),
    ],
  ),
);

standardSlide(
  "User Flow",
  "The intended user experience is guided but still includes manual review gates.",
  row(
    { width: fill, height: fill, gap: 28 },
    [
      column(
        { width: grow(1), height: fill, gap: 18 },
        [
          box("1. Start or reload", "Open index.html, optionally Load Draft Data JSON.", C.blue),
          box("2. Enter or import", "Use manual fields or Smart Import Raw Data for BPRO, Changes, and Release mapping.", C.cyan),
          box("3. Validate preview", "Click Preview Table. App focuses missing fields and shows feedback.", C.green),
          box("4. Export", "Use Export to Outlook to download the Outlook HTML document.", C.amber),
          box("5. Draft email", "Run Outlook macro, select HTML, enter project name, then review displayed draft.", C.lavender),
        ],
      ),
      panel(
        { width: fixed(520), height: fill, padding: { x: 30, y: 28 }, fill: "#FFF8EB", stroke: "#F1D19B", borderRadius: 8 },
        column(
          { width: fill, height: hug, gap: 18 },
          [
            t("User safeguards", { style: { fontSize: 27, bold: true, color: C.ink } }),
            bulletList(
              [
                "Required-field validation before preview/export",
                "Save/load draft JSON",
                "Preview drawer for review",
                "Manual Outlook draft review before send",
                "Flags for fields needing attention",
              ],
              { fontSize: 22, dotColor: C.amber },
            ),
          ],
        ),
      ),
    ],
  ),
);

standardSlide(
  "Benefits Of This Approach",
  "The split keeps business users in familiar tools while avoiding server-side infrastructure.",
  grid(
    { width: fill, height: fill, columns: [fr(1), fr(1), fr(1)], rows: [fr(1), fr(1)], columnGap: 24, rowGap: 24 },
    [
      box("Low infrastructure", "Runs as local browser file plus Outlook macro. No backend service required.", C.blue),
      box("Consistent MOM format", "Fixed tables and generated email wrapper reduce manual formatting drift.", C.green),
      box("Outlook-ready output", "Dedicated Outlook export strips unsafe layout patterns and fixes table widths.", C.amber),
      box("Reusable drafts", "JSON draft save/load supports interrupted work and repeat meetings.", C.cyan),
      box("Operational flexibility", "Smart Import, dynamic rows, optional appendix, and field flags fit real meeting prep.", C.lavender),
      box("Editable final email", "Output remains an Outlook draft that users can inspect and amend before sending.", C.red),
    ],
  ),
);

standardSlide(
  "Risks, Limitations, Dependencies",
  "Most risks sit at client-rendering boundaries and manual handoff steps.",
  tableGrid(
    ["Area", "Risk / limitation", "Why it matters"],
    [
      ["Outlook rendering", "HTML email support differs from browser rendering.", "Macro patches headers and widths because Outlook uses Word rendering."],
      ["Manual handoff", "User must export file, run macro, select file, enter project name.", "Missed step blocks draft creation."],
      ["Macro environment", "Requires Outlook VBA, trusted macro settings, Windows APIs.", "Security policies may restrict execution."],
      ["Content detection", "VBA identifies tables by header keywords.", "Header wording changes can break table recognition."],
      ["Large input", "Smart import supports many rows but browser performance can degrade.", "Source warning appears for more than 1000 rows."],
      ["Single-file scale", "HTML contains UI, styles, parsing, XLSX ZIP logic, export logic.", "Maintenance and testing become harder as features grow."],
      ["Encoding/text", "Email copy includes localized Indonesian text and symbols.", "Encoding must remain UTF-8 end to end."],
    ],
    [fr(0.7), fr(1.15), fr(1.35)],
    { headerFill: C.navy, bodySize: 16, rowH: 70 },
  ),
  "Use this as the sober risk slide: no server risk, but local client compatibility and macro trust are central.",
);

standardSlide(
  "Recommended Improvements",
  "Focus improvements on reliability, maintainability, and handoff clarity without changing the core workflow.",
  grid(
    { width: fill, height: fill, columns: [fr(1), fr(1)], rows: [auto, auto, auto], columnGap: 28, rowGap: 20 },
    [
      box("Add export metadata", "Embed hidden markers or data attributes in generated tables so VBA can identify table types without relying only on header text.", C.blue),
      box("Centralize table schemas", "Define shared table column names, widths, and class names in one HTML config object to reduce drift.", C.green),
      box("Add smoke tests", "Test validateForm, draft JSON roundtrip, Outlook export markup, and XLSX builder for representative inputs.", C.amber),
      box("Improve macro guidance", "Use clearer prompts and default export folder handling for users who download elsewhere.", C.cyan),
      box("Version export contract", "Add visible or hidden export version to Outlook HTML so the macro can warn on incompatible files.", C.lavender),
      box("Separate concerns later", "If maintenance grows, split CSS and JS modules while preserving same single-page user experience.", C.red),
    ],
  ),
  "Recommendations are intentionally scoped. They document better contracts and tests without turning this into a new generator.",
);

standardSlide(
  "Final Conclusion",
  "The current design is practical: browser for authoring, VBA for Outlook delivery.",
  row(
    { width: fill, height: fill, gap: 34, alignItems: "stretch" },
    [
      panel(
        { width: grow(1), height: fill, padding: { x: 42, y: 40 }, fill: C.navy, stroke: C.navy, borderRadius: 8 },
        column(
          { width: fill, height: fill, gap: 24, justifyContent: "center" },
          [
            t("Best reading", { style: { fontSize: 22, bold: true, color: C.cyan } }),
            t(
              "These files form a local documentation-to-email pipeline for MOM preparation. The HTML file owns data quality and output shape. The VBA file owns Outlook draft creation and rendering fixes.",
              { style: { fontSize: 38, bold: true, color: "#FFFFFF", lineSpacing: 1.02 } },
            ),
          ],
        ),
      ),
      column(
        { width: fixed(620), height: fill, gap: 24 },
        [
          box("What works", "Clear split of responsibility, editable Outlook output, repeatable tables, local operation.", C.green),
          box("What needs care", "Macro trust settings, Outlook HTML quirks, manual file handoff, table schema drift.", C.amber),
          box("Next best step", "Add explicit export metadata and focused regression tests around HTML export and VBA table recognition.", C.blue),
        ],
      ),
    ],
  ),
  "Close by reinforcing that the implementation is useful as-is, with most improvement value coming from stronger contracts and tests.",
);

async function saveBlob(blob, filePath) {
  const buffer = Buffer.from(await blob.arrayBuffer());
  await fs.writeFile(filePath, buffer);
}

async function pathExists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function publishBuild(tempOutputDir, tempScratchDir) {
  const suffix = `${process.pid}-${Date.now()}`;
  const outputBackup = `${OUTPUT_DIR}.backup-${suffix}`;
  const scratchBackup = `${SCRATCH}.backup-${suffix}`;
  const outputExisted = await pathExists(OUTPUT_DIR);
  const scratchExisted = await pathExists(SCRATCH);
  let outputPublished = false;
  let scratchPublished = false;

  try {
    if (outputExisted) await fs.rename(OUTPUT_DIR, outputBackup);
    if (scratchExisted) await fs.rename(SCRATCH, scratchBackup);
    await fs.rename(tempOutputDir, OUTPUT_DIR);
    outputPublished = true;
    await fs.rename(tempScratchDir, SCRATCH);
    scratchPublished = true;
  } catch (error) {
    if (scratchPublished) await fs.rm(SCRATCH, { recursive: true, force: true });
    if (outputPublished) await fs.rm(OUTPUT_DIR, { recursive: true, force: true });
    if (scratchExisted && (await pathExists(scratchBackup))) await fs.rename(scratchBackup, SCRATCH);
    if (outputExisted && (await pathExists(outputBackup))) await fs.rename(outputBackup, OUTPUT_DIR);
    throw error;
  } finally {
    await fs.rm(outputBackup, { recursive: true, force: true });
    await fs.rm(scratchBackup, { recursive: true, force: true });
  }
}

async function main() {
  const buildRoot = await fs.mkdtemp(path.join(ROOT, ".deck-build-"));
  const tempOutputDir = path.join(buildRoot, "output");
  const tempScratchDir = path.join(buildRoot, "scratch");

  try {
    await fs.mkdir(tempOutputDir, { recursive: true });
    await fs.mkdir(tempScratchDir, { recursive: true });
    const pptxBlob = await PresentationFile.exportPptx(deck);
    await pptxBlob.save(path.join(tempOutputDir, path.basename(OUT)));

    const slides = deck.slides.items;
    const artifacts = [];
    for (let i = 0; i < slides.length; i += 1) {
      const slideNumber = String(i + 1).padStart(2, "0");
      const pngName = `slide-${slideNumber}.png`;
      const layoutName = `slide-${slideNumber}.layout.json`;
      const slide = slides[i];
      const png = await deck.export({ slide, format: "png" });
      await saveBlob(png, path.join(tempScratchDir, pngName));
      const layout = await deck.export({ slide, format: "layout" });
      await saveBlob(layout, path.join(tempScratchDir, layoutName));
      artifacts.push({ slide: i + 1, png: pngName, layout: layoutName });
    }

    await fs.writeFile(
      path.join(tempScratchDir, "manifest.json"),
      `${JSON.stringify({ schemaVersion: 1, slideCount: slides.length, artifacts }, null, 2)}\n`,
      "utf8",
    );
    await publishBuild(tempOutputDir, tempScratchDir);
    console.log(
      JSON.stringify(
        {
          output: path.relative(ROOT, OUT).replaceAll("\\", "/"),
          scratch: path.relative(ROOT, SCRATCH).replaceAll("\\", "/"),
          slides: slides.length,
        },
        null,
        2,
      ),
    );
  } finally {
    await fs.rm(buildRoot, { recursive: true, force: true });
  }
}

await main();
