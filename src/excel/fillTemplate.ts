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

/** Medic formula corrections from the spec (Global Constraints); written as plain formulas over the shared ones. */
function fixMedicFormulas(ws: ExcelJSType.Worksheet, firstRow: number, lastRow: number): void {
  for (let r = firstRow; r <= lastRow; r++) {
    ws.getCell(`H${r}`).value = { formula: `G${r}` } as ExcelJSType.CellFormulaValue;
    ws.getCell(`F${r}`).value = { formula: `MOD(E${r}-D${r},1)*24` } as ExcelJSType.CellFormulaValue;
  }
  ws.getCell('K20').value = { formula: 'SUM(I20:J20)' } as ExcelJSType.CellFormulaValue;
  ws.getCell('J22').value = { formula: 'SUM(J8:J20)' } as ExcelJSType.CellFormulaValue;
}

/** Writes the column headings from the definitions into the template header row, keeping cell styles. */
export function applyLayout(ws: ExcelJSType.Worksheet, def: TemplateDef): void {
  def.columns.forEach((c) => {
    const cell = ws.getCell(`${c.col}${def.headerRow}`);
    cell.value = c.header;
    if (c.header.indexOf('\n') >= 0) cell.alignment = { ...cell.alignment, wrapText: true };
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
        if (!isNaN(n)) cell.value = n;
      }
    });
  });

  if (data.type === 'Medic') fixMedicFormulas(ws, def.firstRow, def.firstRow + def.rowCount - 1);

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

  wb.calcProperties.fullCalcOnLoad = true;
  return (await wb.xlsx.writeBuffer()) as ArrayBuffer;
}
