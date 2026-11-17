import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

function readArguments(args) {
  const [template, planFile, output, ...options] = args;
  if (!template || !planFile || !output) {
    throw new Error("Usage: write_xlsx.mjs <template.xlsx> <plan.json> <output.xlsx> [--before before.png] [--after after.png]");
  }
  const result = { template, planFile, output };
  for (let index = 0; index < options.length; index += 2) {
    const key = options[index];
    const value = options[index + 1];
    if (!["--before", "--after"].includes(key) || !value) {
      throw new Error(`Unknown or incomplete option ${key}`);
    }
    result[key.slice(2)] = value;
  }
  return result;
}

function columnIndex(letters) {
  let index = 0;
  for (const letter of letters) index = index * 26 + letter.charCodeAt(0) - 64;
  return index;
}

function previewRange(writes) {
  const cells = writes.map(({ cell }) => {
    const match = /^([A-Z]+)([1-9]\d*)$/.exec(cell);
    if (!match) throw new Error(`Invalid cell address ${cell}`);
    return { column: match[1], index: columnIndex(match[1]), row: Number(match[2]) };
  });
  const left = cells.reduce((a, b) => a.index < b.index ? a : b).column;
  const right = cells.reduce((a, b) => a.index > b.index ? a : b).column;
  const firstRow = Math.max(1, Math.min(...cells.map(cell => cell.row)) - 1);
  const lastRow = Math.max(...cells.map(cell => cell.row));
  return `${left}${firstRow}:${right}${lastRow}`;
}

async function renderPreview(workbook, sheetName, range, destination) {
  if (!destination) return;
  const preview = await workbook.render({ sheetName, range, scale: 1.5, format: "png" });
  await fs.mkdir(path.dirname(path.resolve(destination)), { recursive: true });
  await fs.writeFile(destination, new Uint8Array(await preview.arrayBuffer()));
}

async function main() {
  const args = readArguments(process.argv.slice(2));
  if (path.resolve(args.template) === path.resolve(args.output)) {
    throw new Error("Output must be a new workbook path.");
  }
  try {
    await fs.access(args.output);
    throw new Error("Output already exists. Choose a new output filename.");
  } catch (error) {
    if (error.code !== "ENOENT") throw error;
  }
  const plan = JSON.parse(await fs.readFile(args.planFile, "utf8"));
  if (plan.pending?.length || !plan.sheet || !Array.isArray(plan.writes) || !plan.writes.length) {
    throw new Error("Plan must contain a selected sheet and confirmed cell writes.");
  }
  if (plan.template && path.resolve(plan.template) !== path.resolve(args.template)) {
    throw new Error("Plan template does not match the supplied workbook.");
  }
  if (new Set(plan.writes.map(write => write.cell)).size !== plan.writes.length) {
    throw new Error("Plan contains duplicate cell writes.");
  }
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(args.template));
  const sheet = workbook.worksheets.getItem(plan.sheet);
  const range = plan.preview_range || previewRange(plan.writes);
  await renderPreview(workbook, plan.sheet, range, args.before);

  for (const { cell, value } of plan.writes) {
    const literal = typeof value === "string" && value.startsWith("=") ? `'${value}` : value;
    sheet.getRange(cell).values = [[literal]];
  }
  workbook.recalculate();
  for (const { cell, value } of plan.writes) {
    const actual = sheet.getRange(cell).values[0]?.[0];
    if (actual !== value) throw new Error(`Written value mismatch at ${cell}`);
  }
  const errors = await workbook.inspect({
    kind: "match",
    sheetId: plan.sheet,
    range,
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
    options: { useRegex: true, maxResults: 30 },
    maxChars: 2000
  });
  await renderPreview(workbook, plan.sheet, range, args.after);
  await fs.mkdir(path.dirname(path.resolve(args.output)), { recursive: true });
  const temporary = `${args.output}.${process.pid}.tmp.xlsx`;
  try {
    await (await SpreadsheetFile.exportXlsx(workbook)).save(temporary);
    await fs.copyFile(temporary, args.output, fs.constants.COPYFILE_EXCL);
  } finally {
    await fs.rm(temporary, { force: true });
  }
  process.stdout.write(JSON.stringify({
    output: path.resolve(args.output),
    sheet: plan.sheet,
    written_cells: plan.writes.length,
    preview_range: range,
    formula_error_scan: errors.ndjson
  }, null, 2) + "\n");
}

main().catch(error => {
  process.stderr.write(`${error.message}\n`);
  process.exitCode = 1;
});
