import type * as ExcelJSType from 'exceljs';
import { computeLine, computeTotals } from '../timesheet/calc';
import { TEMPLATES } from '../timesheet/definitions';
import { Line, TemplateDef, TimesheetData } from '../timesheet/types';
import { normalizeDate, parseTime, timeToDayFraction, toExcelDate } from '../timesheet/values';

/**
 * The prebuilt browser bundle, loaded on demand: keeps ExcelJS out of the main
 * bundle and avoids webpack Node polyfill errors. It also runs under Jest.
 */
async function loadExcelJs(): Promise<typeof ExcelJSType> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const mod: any = await import(/* webpackChunkName: 'exceljs' */ 'exceljs/dist/exceljs.min.js');
  return mod.default || mod;
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
    ws.getCell('K20').value = { formula: 'SUM(I20:J20)' } as ExcelJSType.CellFormulaValue;
    ws.getCell('J22').value = { formula: 'SUM(J8:J20)' } as ExcelJSType.CellFormulaValue;
  }
  // PSL cells hold the level text; a hidden table maps it back to the amount, and the line total looks it up.
  const validations = (ws as unknown as { dataValidations: { model: Record<string, unknown> } }).dataValidations.model;
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

/** Fills a copy of the stored template. Only cell values change; styles and formats stay. */
export async function fillTemplate(template: ArrayBuffer, data: TimesheetData): Promise<ArrayBuffer> {
  const ExcelJS = await loadExcelJs();
  const def = TEMPLATES[data.type];
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(template);
  const ws = wb.getWorksheet(def.sheetName)!;
  applyLayout(ws, def);

  ws.getCell(def.nameCell).value = data.name;
  if (data.from) ws.getCell(def.fromCell).value = toExcelDate(data.from);
  if (data.to) ws.getCell(def.toCell).value = toExcelDate(data.to);

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
  return (await wb.xlsx.writeBuffer()) as ArrayBuffer;
}
