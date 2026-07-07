#!/usr/bin/env node
/*
 * Facilities LOF generator (CRR 20 / CRR 21).
 * Usage: node generate_lof.js <findings.json> [output_dir]
 *
 * Consumes the findings JSON described in ../CLAUDE.md and emits
 * [SchoolName]_Facilities_LOF.docx. Enforces, regardless of input:
 *   - the fixed 18-row order (missing rows emitted as None. / None.)
 *   - Program Access hard suppression (None. / None.)
 *   - warning when a corrective cite is not 2010 ADA
 *   - rejection of em-dash characters anywhere in the JSON
 * Table format and boilerplate come verbatim from
 * ../../references/boilerplate.md (do not edit strings here without
 * updating that file).
 */

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  UnderlineType, HeadingLevel,
} = require("docx");

const FIXED_ROW_ORDER = [
  "Accessible Parking",
  "Accessible Routes/Walkways from accessible parking",
  "Stairways and Steps",
  "Ramps",
  "Curb Ramps",
  "Entrances, Doors, and Gates",
  "Rooms, Offices, and Administration",
  "Elevators/Lifts",
  "Accessible Drinking Fountains",
  "Cafeteria",
  "Library",
  "CTE Classrooms",
  "Labs/Shops",
  "Gymnasium/Auditorium/Weight Room/Other",
  "Stadium/Field",
  "Dressing, Fitting, and Locker Rooms",
  "Restrooms",
  "Public Telephones",
];

// ---------------------------------------------------------------- helpers
const t = (text, o = {}) => new TextRun({ text, font: "Arial", size: 20, ...o });
const b = (text, o = {}) => t(text, { bold: true, ...o });
const u = (text) => t(text, { underline: { type: UnderlineType.SINGLE } });
const p = (children, o = {}) => {
  if (typeof children === "string") children = [t(children)];
  return new Paragraph({ children, spacing: { after: 80 }, ...o });
};

const cb = { style: BorderStyle.SINGLE, size: 4, color: "999999" };
const borders = { top: cb, bottom: cb, left: cb, right: cb };
const CW = [2808, 1368, 2592, 2592]; // ~30 / 14 / 28 / 28 percent of 9360 DXA

const mkCell = (paras, w) => new TableCell({
  borders, margins: { top: 80, bottom: 80, left: 120, right: 120 },
  width: { size: w, type: WidthType.DXA }, verticalAlign: VerticalAlign.TOP,
  children: paras,
});
const hCell = (text, w) => new TableCell({
  borders, margins: { top: 80, bottom: 80, left: 120, right: 120 },
  width: { size: w, type: WidthType.DXA },
  shading: { fill: "1A2744", type: ShadingType.CLEAR },
  verticalAlign: VerticalAlign.CENTER,
  children: [p([b(text, { color: "FFFFFF" })], { alignment: AlignmentType.CENTER })],
});

// Findings cell: each item is text plus a bold cite line.
function findingParas(items) {
  if (!items || items.length === 0) return [p([t("None.")])];
  const out = [];
  for (const it of items) {
    out.push(p([t(it.text)]));
    if (it.cite) out.push(p([b("Cite: " + it.cite)]));
  }
  return out;
}

function isProgramAccessOnly(row) {
  const stds = (row.standards || []).map((s) => String(s).toLowerCase());
  return stds.length > 0 && stds.every((s) => s.includes("program access"));
}

// ---------------------------------------------------------------- validate
function fail(msg) { console.error("ERROR: " + msg); process.exit(1); }

const [, , jsonPath, outDirArg] = process.argv;
if (!jsonPath) fail("usage: node generate_lof.js <findings.json> [output_dir]");

const raw = fs.readFileSync(jsonPath, "utf8");
if (raw.includes("—")) fail("em-dash character found in findings JSON; house style forbids em-dashes. Fix the JSON and rerun.");

const data = JSON.parse(raw);
const crrNum = String(data.crr || 20);
if (!["20", "21"].includes(crrNum)) fail("crr must be 20 or 21");
const school = data.school || "School";

const byArea = new Map();
for (const row of data.rows || []) {
  if (!FIXED_ROW_ORDER.includes(row.area)) {
    console.warn(`WARN: unknown area "${row.area}" ignored (not one of the fixed 18 rows). Fold its content into the closest fixed row.`);
    continue;
  }
  if (byArea.has(row.area)) fail(`duplicate row for area "${row.area}"`);
  byArea.set(row.area, row);
}

const warnings = [];
const stats = { findings: 0, programAccess: 0, populated: byArea.size };

// ---------------------------------------------------------------- table
const tableRows = [
  new TableRow({ tableHeader: true, children: [
    hCell("Area Reviewed with Construction and Alteration Date(s)", CW[0]),
    hCell("Applicable Building Standard", CW[1]),
    hCell("Accessibility Violation", CW[2]),
    hCell("Corrective Actions", CW[3]),
  ] }),
  ...FIXED_ROW_ORDER.map((area) => {
    const row = byArea.get(area) || {};
    const col1 = [p([b(area + ":")])];
    for (const loc of row.locations || []) {
      if (loc.name) col1.push(p([u(loc.name)]));
      if (loc.constructed) col1.push(p([t("Constructed: " + loc.constructed)]));
      if (loc.modified) col1.push(p([t("Modified: " + loc.modified)]));
    }
    const col2 = (row.standards || []).length
      ? row.standards.map((s) => p([t(s)]))
      : [p([t("")])];

    let violations = row.violations || [];
    let correctives = row.correctives || [];

    if (isProgramAccessOnly(row)) {
      // Hard suppression: Program Access never generates findings.
      if (violations.length || correctives.length) {
        warnings.push(`Program Access suppression applied to "${area}": ${violations.length} violation(s) and ${correctives.length} corrective(s) in the JSON were replaced with None.`);
      }
      violations = [];
      correctives = [];
      stats.programAccess++;
    } else {
      for (const c of correctives) {
        if (c.cite && !/2010\s*ADA/i.test(c.cite)) {
          warnings.push(`Corrective cite for "${area}" is "${c.cite}"; corrective actions must ALWAYS cite 2010 ADA. Review before delivering.`);
        }
      }
      stats.findings += violations.length;
    }

    return new TableRow({ children: [
      mkCell(col1, CW[0]),
      mkCell(col2, CW[1]),
      mkCell(findingParas(violations), CW[2]),
      mkCell(findingParas(correctives), CW[3]),
    ] });
  }),
];

// -------------------------------------------------------------- boilerplate
// Verbatim instrument language; see ../../references/boilerplate.md.
const boilerplate = [
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [b("CRR " + crrNum + ": Accessible Facilities", { size: 28 })], spacing: { after: 120 } }),
  p([b("Applicable Requirements: "), t("28 CFR Sections 35, 35.151, 35.151(a); 34 CFR Sections 104.22, 104.22(a), 104.23, 104.23(a), 104.23(c) and 36 Appendix D.")]),
  p([u("Program Access/Readily Accessible – Existing Facility under 504")]),
  p([t("For existing recipient facilities under 504 that were built or altered beginning June 3, 1977, or earlier, a recipient shall operate its program or activity so that when each part is viewed in its entirety, it is “readily accessible” to disabled persons. A recipient is not required to make each of its existing facilities or every part of a facility accessible to and useable by persons with disabilities.")], { spacing: { after: 120 } }),
  p([u("American National Standards Institute (ANSI) – New Construction under 504")]),
  p([t("Each facility or part of a facility constructed by, on behalf of, or for the use of a recipient under 504 that were built or altered between June 4, 1977, and January 17, 1991, inclusive, shall be designed and constructed in such manner that the facility or part of the facility is readily accessible to and usable by persons with disabilities. Conformance with the “American National Standard Specifications for Making Buildings and Facilities Accessible to, and Usable by, the Physically Disabled” published by the American National Standards Institute, Inc. (ANSI) A117.1–1961 (R1971) Later versions of ANSI A117.1 do not apply.")], { spacing: { after: 120 } }),
  p([u("Uniform Federal Accessibility Standards (UFAS) – New Construction under 504")]),
  p([t("Each facility or part of a facility constructed by, on behalf of, or for the use of a recipient or public entity under 504 that were built or altered between January 18, 1991, and January 26, 1992, inclusive, shall be designed and constructed in such manner that the facility or part of the facility is readily accessible to and usable by persons with disabilities. Conformance with the Uniform Federal Accessibility Standards (UFAS) (Appendix A to 41 CFR Section 101-19.6). Departures from particular technical and scoping requirements permitted where substantially equivalent or greater access to and usability of the building is provided.")], { spacing: { after: 120 } }),
  p([u("1991 Americans with Disabilities Act (ADA) – New Construction under 504")]),
  p([t("Each facility or part of a facility constructed by, on behalf of, or for the use of a recipient or public entity is designed and constructed in such manner that the facility or part of the facility is readily accessible to and usable by persons with disabilities. UFAS or the 1991 Americans with Disabilities Act (ADA) Standard apply to facilities constructed or altered on or after January 27, 1992 and before September 15, 2010.")], { spacing: { after: 120 } }),
  p([u("2010 Americans with Disabilities Act (ADA) – New Construction under 504")]),
  p([t("Each facility or part of a facility constructed by, on behalf of, or for the use of a recipient or public entity is designed and constructed in such a manner that the facility or part of the facility is readily accessible to and usable by persons with disabilities. The 2010 ADA Standards apply to facilities constructed on or after March 15, 2012.")], { spacing: { after: 200 } }),
  p([b("Summary of Analysis, Findings, and Required Corrective Actions (CRR " + crrNum + "):")]),
  p([t("Facilities listed below were assessed under program accessibility, ANSI, UFAS, 1991 ADA, and 2010 ADA standards. The Office of Civil Rights requires all corrective actions to be made in accordance with 2010 ADA standards; therefore, certain areas will not require corrective action as the noted deficiency is within the 2010 ADA standards.")], { spacing: { after: 200 } }),
];

// ---------------------------------------------------------------- document
const doc = new Document({
  creator: "Murjani McTier",
  lastModifiedBy: "Murjani McTier",
  styles: { default: { document: { run: { font: "Arial", size: 20 } } } },
  sections: [{
    properties: { page: { margin: { top: 1080, bottom: 1080, left: 1440, right: 1440 } } },
    children: [
      ...boilerplate,
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: CW, rows: tableRows }),
    ],
  }],
});

const outDir = outDirArg || process.cwd();
fs.mkdirSync(outDir, { recursive: true });
const safeName = school.replace(/[^A-Za-z0-9 ]/g, "").trim().replace(/\s+/g, "_") || "School";
const outPath = path.join(outDir, `${safeName}_Facilities_LOF.docx`);

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(outPath, buf);
  // Validation: the buffer must be a readable zip containing document.xml.
  const AdmZip = tryRequire("adm-zip");
  if (AdmZip) {
    const zip = new AdmZip(outPath);
    const entry = zip.getEntry("word/document.xml");
    if (!entry) fail("generated docx is missing word/document.xml");
    const xml = zip.readAsText(entry);
    if (!xml.includes("CRR " + crrNum)) fail("generated docx failed content validation");
  } else if (buf.slice(0, 2).toString() !== "PK") {
    fail("generated file is not a valid zip container");
  }
  console.log("LOF written: " + outPath + " (" + buf.length + " bytes)");
  console.log(`Summary: ${FIXED_ROW_ORDER.length} rows emitted, ${stats.populated} populated from JSON, ${stats.findings} finding(s) issued, ${stats.programAccess} Program Access row(s).`);
  for (const w of warnings) console.warn("WARN: " + w);
  if (!warnings.length) console.log("No rule warnings.");
}).catch((e) => fail(e.message));

function tryRequire(name) { try { return require(name); } catch { return null; } }
