import type * as ExcelJSType from 'exceljs';
import { computeLine, computeTotals } from '../timesheet/calc';
import { TEMPLATES } from '../timesheet/definitions';
import { Line, TemplateDef, TimesheetData } from '../timesheet/types';
import { normalizeDate, parseTime, timeToDayFraction, toExcelDate } from '../timesheet/values';

/**
 * The prebuilt browser bundle, loaded on demand: keeps ExcelJS out of the main
 * bundle and avoids webpack Node polyfill errors. It also runs under Jest.
 */
export async function loadExcelJs(): Promise<typeof ExcelJSType> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const mod: any = await import(/* webpackChunkName: 'exceljs' */ 'exceljs/dist/exceljs.min.js');
  return mod.default || mod;
}

/** Cell references (A22, $E$23) not followed by '(' or '!'; sheet-qualified refs are not in these formulas. */
const CELL_REF = /(?<![A-Za-z_])(\$?[A-Z]{1,3})(\$?)(\d+)(?![\d(!])/g;

/** Rewrites each row number in a formula or range with map. Absolute rows ($5) map too. */
export const mapRows = (formula: string, map: (row: number) => number): string =>
  formula.replace(CELL_REF, (_m, c: string, abs: string, r: string) => `${c}${abs}${map(Number(r))}`);

type Validations = Record<string, unknown>;
const validationsOf = (ws: ExcelJSType.Worksheet): Validations =>
  (ws as unknown as { dataValidations: { model: Validations } }).dataValidations.model;

/**
 * Inserts extra rows above the last template row, so the totals' SUM ranges stretch over them.
 * ExcelJS moves values and styles but not formulas, validations, formats or hidden flags; those are fixed here.
 * Returns the definition with the larger row count and the moved totals cells.
 */
export function growRows(ws: ExcelJSType.Worksheet, def: TemplateDef, extra: number): TemplateDef {
  if (extra <= 0) return def;
  const at = def.firstRow + def.rowCount - 1;
  const src = at - 1;
  const shift = (r: number): number => (r >= at ? r + extra : r);

  // Plain formulas first: a shared formula would be translated again after the move.
  ws.eachRow((row) => row.eachCell((cell) => {
    if (cell.formula) cell.value = { formula: cell.formula } as ExcelJSType.CellFormulaValue;
  }));
  const hidden: boolean[] = [];
  for (let r = at; r <= ws.rowCount; r++) hidden[r] = ws.getRow(r).hidden;

  ws.duplicateRow(src, extra, true);

  for (let r = ws.rowCount; r >= at + extra; r--) {
    const row = ws.getRow(r);
    row.hidden = !!hidden[r - extra];
    row.eachCell((cell) => {
      if (cell.formula) cell.value = { formula: mapRows(cell.formula, shift) } as ExcelJSType.CellFormulaValue;
    });
  }
  for (let r = at; r < at + extra; r++) {
    const row = ws.getRow(r);
    row.hidden = false;
    row.eachCell((cell) => {
      if (cell.formula) cell.value = { formula: mapRows(cell.formula, (x) => (x === src ? r : x)) } as ExcelJSType.CellFormulaValue;
    });
  }

  const validations = validationsOf(ws);
  const moved: Validations = {};
  Object.keys(validations).forEach((k) => {
    moved[mapRows(k, shift)] = validations[k];
    const m = /^([A-Z]+)(\d+)$/.exec(k);
    if (m && Number(m[2]) === src) for (let r = at; r < at + extra; r++) moved[`${m[1]}${r}`] = validations[k];
    delete validations[k];
  });
  Object.assign(validations, moved);

  // Missing from the typings.
  const formats = (ws as unknown as { conditionalFormattings: { ref: string; rules: { formulae?: string[] }[] }[] }).conditionalFormattings;
  formats.forEach((cf) => {
    cf.ref = mapRows(cf.ref, shift);
    cf.rules.forEach((rule) => {
      if (rule.formulae) rule.formulae = rule.formulae.map((x) => mapRows(String(x), shift));
    });
  });

  return {
    ...def,
    rowCount: def.rowCount + extra,
    totals: def.totals.map((t) => ({ ...t, cell: mapRows(t.cell, shift) }))
  };
}

/** Writes the column headings from the definitions into the template header row, keeping cell styles. */
export function applyLayout(ws: ExcelJSType.Worksheet, def: TemplateDef): void {
  def.columns.forEach((c) => {
    const cell = ws.getCell(`${c.col}${def.headerRow}`);
    cell.value = c.header;
    if (c.header.indexOf('\n') >= 0) cell.alignment = { ...cell.alignment, wrapText: true };
  });
  if (def.type === 'Medic') {
    // DAYS is an input in F, so the hours formulas go; start/end columns D:E are hidden (F4:G4 'From:' stays visible).
    ['D', 'E'].forEach((c) => {
      ws.getCell(`${c}${def.headerRow}`).value = null;
      ws.getColumn(c).hidden = true;
    });
    for (let r = def.firstRow; r < def.firstRow + def.rowCount; r++) {
      ws.getCell(`F${r}`).value = null;
    }
    // The template's last row has no K formula, and its per diem total sums J:K.
    const last = def.firstRow + def.rowCount - 1;
    ws.getCell(`K${last}`).value = { formula: `SUM(I${last}:J${last})` } as ExcelJSType.CellFormulaValue;
    ws.getCell(def.totals.find((t) => t.key === 'perDiem')!.cell).value = { formula: `SUM(J${def.firstRow}:J${last})` } as ExcelJSType.CellFormulaValue;
  }
  const validations = validationsOf(ws);
  // Yes/No columns (Admin STAT) go in a hidden helper column of the template: show it and style it like the rate column.
  def.columns.filter((c) => c.kind === 'yesno').forEach((c) => {
    const like = def.columns.find((x) => x.kind === 'number')!.col;
    const column = ws.getColumn(c.col);
    column.hidden = false;
    column.width = c.width / 7;
    ws.getCell(`${c.col}${def.headerRow}`).style = ws.getCell(`${like}${def.headerRow}`).style;
    for (let r = def.firstRow; r < def.firstRow + def.rowCount; r++) {
      const cell = ws.getCell(`${c.col}${r}`);
      cell.value = null;
      cell.style = { ...ws.getCell(`${like}${r}`).style, numFmt: 'General', alignment: { horizontal: 'center', vertical: 'middle' } };
      validations[cell.address] = { type: 'list', allowBlank: true, formulae: ['"Yes,No"'] };
    }
  });
  // PSL cells hold the level text; a hidden table maps it back to the amount, and the line total looks it up.
  def.columns.forEach((c) => {
    const t = c.levelTable;
    if (!t) return;
    (c.options || []).forEach((amount, i) => {
      ws.getCell(`${t.labelCol}${t.firstRow + i}`).value = c.optionLabels![i];
      ws.getCell(`${t.amountCol}${t.firstRow + i}`).value = amount;
    });
    const range = `$${t.labelCol}$${t.firstRow}:$${t.amountCol}$${t.firstRow + (c.options || []).length - 1}`;
    const lookup = (r: number): string => `VLOOKUP(${c.col}${r},${range},2,FALSE)`;
    for (let r = def.firstRow; r < def.firstRow + def.rowCount; r++) {
      const amount = `${def.type === 'Medic' ? 'F' : 'D'}${r}*${lookup(r)}`;
      const formula = `IF(${c.col}${r}="",0,${amount})`; // blank PSL gives 0, a bad label still errors
      ws.getCell(`${def.type === 'Medic' ? 'H' : 'F'}${r}`).value = { formula } as ExcelJSType.CellFormulaValue;
      delete validations[`${c.col}${r}`]; // the template list holds amounts and would reject the label
    }
  });
}

const STAT_ORANGE = 'FFFFC000';

/** Fills a copy of the stored template. Only cell values change; styles and formats stay. */
export async function fillTemplate(template: ArrayBuffer, data: TimesheetData): Promise<ArrayBuffer> {
  const ExcelJS = await loadExcelJs();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(template);
  const ws = wb.getWorksheet(TEMPLATES[data.type].sheetName)!;
  // Rows up to the last one with anything in it; blank rows added at the end don't grow the file.
  const used = data.lines.reduce((n, l, i) => (Object.keys(l).some((k) => l[k] !== null && l[k] !== '') ? i + 1 : n), 0);
  const def = growRows(ws, TEMPLATES[data.type], used - TEMPLATES[data.type].rowCount);
  applyLayout(ws, def);

  ws.getCell(def.nameCell).value = data.name;
  if (data.from) ws.getCell(def.fromCell).value = toExcelDate(data.from);
  if (data.to) ws.getCell(def.toCell).value = toExcelDate(data.to);
  // The template styles these bold italic 14pt with a long date, which shows ######## in the column width.
  [def.fromCell, def.toCell].forEach((a) => {
    const cell = ws.getCell(a);
    cell.font = { ...cell.font, italic: false };
    cell.numFmt = 'dd/mm/yy;@';
  });
  // PSL columns show the level label, far longer than the template's amount column.
  def.columns.forEach((c) => {
    if (c.levelTable && c.optionLabels) ws.getColumn(c.col).width = Math.max(...c.optionLabels.map((l) => l.length), c.header.length) + 2;
  });

  data.lines.slice(0, def.rowCount).forEach((line, i) => {
    def.columns.forEach((c) => {
      const raw = line[c.key];
      if (c.kind === 'computed' || raw === null || raw === undefined || String(raw).trim() === '') return;
      const cell = ws.getCell(`${c.col}${def.firstRow + i}`);
      if (c.kind === 'date') {
        const iso = normalizeDate(String(raw));
        if (iso) cell.value = toExcelDate(iso);
      } else if (c.kind === 'time') {
        const t = parseTime(String(raw));
        if (t) cell.value = timeToDayFraction(t);
      } else if (c.kind === 'upper') {
        cell.value = String(raw).toUpperCase();
      } else if (c.kind === 'yesno') {
        cell.value = String(raw);
        // Payroll marks stat holiday rows orange.
        if (raw === 'Yes') {
          const row = ws.getRow(def.firstRow + i);
          for (let x = 1; x <= ws.getColumn(c.col).number; x++) {
            // A new style object: ExcelJS shares one between cells, so setting .fill would colour other rows too.
            const cell = row.getCell(x);
            cell.style = { ...cell.style, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: STAT_ORANGE } } };
          }
        }
      } else {
        const n = Number(raw);
        const at = c.levelTable ? (c.options || []).indexOf(n) : -1;
        if (at >= 0) cell.value = c.optionLabels![at];
        else if (!isNaN(n)) cell.value = n;
      }
    });
  });

  // Cached results, so previewers that don't recalculate (Outlook, phones) show real totals.
  const setResult = (address: string, result: Line[string] | undefined): void => {
    const cell = ws.getCell(address);
    if (typeof result === 'number' && cell.formula) cell.value = { ...(cell.value as ExcelJSType.CellFormulaValue), result };
  };
  for (let i = 0; i < def.rowCount; i++) {
    const computed = computeLine(data.type, data.lines[i] || {});
    def.columns.filter((c) => c.kind === 'computed').forEach((c) => setResult(`${c.col}${def.firstRow + i}`, computed[c.key]));
  }
  const totals = computeTotals(data.type, data.lines);
  def.totals.forEach(({ key, cell }) => setResult(cell, totals[key]));

  // ExcelJS drops external link parts but keeps names that use them ('[1]Instructors'!…), and Excel then refuses the file.
  const names = wb.definedNames as unknown as { model: { ranges: string[] }[] };
  names.model = names.model.filter((n) => !n.ranges.some((r) => r.indexOf('[') >= 0));

  wb.calcProperties.fullCalcOnLoad = true;
  return hideUnusedRows((await wb.xlsx.writeBuffer()) as ArrayBuffer);
}

/**
 * The templates hide every row below the form with sheetFormatPr zeroHeight="1".
 * ExcelJS drops that attribute on write, so put it back in each sheet.
 */
export async function hideUnusedRows(xlsx: ArrayBuffer): Promise<ArrayBuffer> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const mod: any = await import(/* webpackChunkName: 'jszip' */ 'jszip/dist/jszip.min.js');
  const JSZip = mod.default || mod;
  const zip = await JSZip.loadAsync(new Uint8Array(xlsx));
  for (const name of Object.keys(zip.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))) {
    const xml: string = await zip.file(name).async('string');
    zip.file(name, xml.replace(/<sheetFormatPr(?![^>]*zeroHeight)/, '<sheetFormatPr zeroHeight="1"'));
  }
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}
