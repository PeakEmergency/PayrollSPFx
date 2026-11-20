import * as React from 'react';
import jspreadsheet from 'jspreadsheet-ce';
import 'jsuites/dist/jsuites.css';
import 'jspreadsheet-ce/dist/jspreadsheet.css';
import { computeLine } from '../../../timesheet/calc';
import { ColumnDef, Line, TemplateDef } from '../../../timesheet/types';
import { normalizeDate, parseTime } from '../../../timesheet/values';
import { fillRate, PayRates, RATE_KEY, rateOptions } from '../../../timesheet/rates';

export interface ITimesheetGridProps {
  def: TemplateDef;
  lines: Line[];
  /** Fills each row's rate as the user types; undefined while loading or if the person has none. */
  rates?: PayRates;
  onChange(lines: Line[]): void;
}

// Option lists that are counts or hours, not money.
const plainNumberKeys: string[] = ['days', 'adminHours', 'deductions'];

const money = (n: number): string => `$${n.toFixed(2)}`;

// 24-hour times in 15-minute steps; the browser's own time input can't be forced to 24-hour.
// ponytail: 15-minute granularity; shorten the step if shifts need exact minutes.
const times: string[] = Array.from({ length: 96 }, (_, i) => `${String(Math.floor(i / 4)).padStart(2, '0')}:${String((i % 4) * 15).padStart(2, '0')}`);

const formatComputed = (key: string, v: Line[string]): string =>
  typeof v !== 'number' ? '' : key === 'hours' ? v.toFixed(2) : money(v);

/** allowed: the only option values to offer (the person's PSL levels); undefined offers them all. */
function toColumn(c: ColumnDef, allowed?: number[]): jspreadsheet.Column {
  // Excel shows a column ~5px wider than width × 7; 90px fits the bold two-line headers.
  const base = { title: c.header, width: Math.max(c.width + 5, 90), align: 'center' as const };
  switch (c.kind) {
    case 'date':
      // position (missing from the typings) makes jSuites use position: fixed, so the grid's overflow can't clip it.
      return { ...base, type: 'calendar',
        options: { format: 'DD/MM/YY', position: true } as { format: string } };
    case 'dropdown':
      return {
        ...base,
        type: 'dropdown',
        // Labelled rates show the classification; the cell value stays the dollar amount.
        width: c.optionLabels ? 230 : base.width,
        source: (c.options || [])
          .map((o, i) => ({
            id: o,
            name: c.optionLabels ? c.optionLabels[i] : plainNumberKeys.indexOf(c.key) >= 0 ? String(o) : money(o)
          }))
          .filter((o) => !allowed || allowed.indexOf(o.id) >= 0)
      };
    case 'time':
      return { ...base, type: 'dropdown', autocomplete: true, source: times };
    case 'computed':
      return { ...base, type: 'text', readOnly: true };
    case 'yesno':
      return { ...base, width: Math.max(c.width + 5, 70), type: 'dropdown', source: ['Yes', 'No'] };
    default:
      return { ...base, type: 'text', align: c.kind === 'upper' ? 'left' : 'center' };
  }
}

/**
 * Columns on screen. The template's sub-total columns are ~1 char wide, so they are
 * left off screen like in Excel; they are still computed and written to the file.
 */
const STAT_ORANGE = '#ffc000';

const shownColumns = (def: TemplateDef): ColumnDef[] => def.columns.filter((c) => !(c.kind === 'computed' && c.width < 14));

const groupOf = (def: TemplateDef, c: ColumnDef): NonNullable<TemplateDef['groups']>[number] | undefined =>
  (def.groups || []).find((g) => c.col >= g.from && c.col <= g.to);

/** One header row above the column titles: each group spans its columns, other columns get a blank cell. */
function nestedHeaders(def: TemplateDef, cols: ColumnDef[]): jspreadsheet.NestedHeaderCell[][] {
  const cells: jspreadsheet.NestedHeaderCell[] = [];
  let last: unknown;
  cols.forEach((c) => {
    const g = groupOf(def, c);
    if (g && g === last) cells[cells.length - 1].colspan!++;
    else cells.push({ title: g ? g.title : '', colspan: 1 });
    last = g;
  });
  return [cells];
}

/** Grid row (cells in column order) → Line. Blank cells become null; times stay as typed until valid. */
function toLine(cols: ColumnDef[], row: jspreadsheet.CellValue[]): Line {
  const line: Line = {};
  cols.forEach((c, x) => {
    if (c.kind === 'computed') return;
    const raw = row[x] === undefined || row[x] === null ? '' : String(row[x]).trim();
    if (!raw) line[c.key] = null;
    else if (c.kind === 'date') line[c.key] = normalizeDate(raw);
    else if (c.kind === 'number' || c.kind === 'dropdown') line[c.key] = isNaN(Number(raw)) ? raw : Number(raw);
    else line[c.key] = raw;
  });
  return line;
}

const toRow = (cols: ColumnDef[], line: Line | undefined): jspreadsheet.CellValue[] =>
  cols.map((c) => {
    const v = line ? line[c.key] : null;
    return v === null || v === undefined ? '' : c.kind === 'computed' ? formatComputed(c.key, v) : v;
  });

/** Vanilla Jspreadsheet CE in a ref; fixed rows, one column per template column. */
export default function TimesheetGrid({ def, lines, rates, onChange }: ITimesheetGridProps): React.ReactElement {
  const host = React.useRef<HTMLDivElement>(null);
  const onChangeRef = React.useRef(onChange);
  onChangeRef.current = onChange;
  const ratesRef = React.useRef(rates);
  ratesRef.current = rates;

  React.useEffect(() => {
    const el = host.current!;
    const cols = shownColumns(def);
    let updating = false;

    const sync = (ws: jspreadsheet.WorksheetInstance): void => {
      const rows = ws.getData() as jspreadsheet.CellValue[][];
      const next = rows.map((r) => toLine(cols, r));
      updating = true;
      next.forEach((line, y) => {
        const computed = computeLine(def.type, line);
        cols.forEach((c, x) => {
          if (c.kind === 'computed') ws.setValueFromCoords(x, y, formatComputed(c.key, computed[c.key]), true);
        });
      });
      updating = false;
      // Stat holiday rows are orange, as payroll marks them in Excel.
      const yesno = cols.filter((c) => c.kind === 'yesno');
      next.forEach((line, y) => {
        const on = yesno.some((c) => line[c.key] === 'Yes');
        el.querySelectorAll<HTMLElement>(`tbody td[data-y="${y}"]`).forEach((td) => (td.style.backgroundColor = on ? STAT_ORANGE : ''));
      });
      onChangeRef.current(next);
    };

    const [ws] = jspreadsheet(el as HTMLDivElement, {
      tabs: false,
      toolbar: false,
      worksheets: [
        {
          data: Array.from({ length: Math.max(def.rowCount, lines.length) }, (_, i) => toRow(cols, lines[i])),
          columns: cols.map((c) => toColumn(c, c.key === RATE_KEY[def.type] ? rateOptions(def.type, rates) : undefined)),
          nestedHeaders: def.groups ? nestedHeaders(def, cols) : undefined,
          minDimensions: [cols.length, def.rowCount],
          defaultRowHeight: 36, // matches the Expense Report rows
          // Enter on the last row adds a row (Tab at the end of the last row too, below); the file grows to match.
          allowInsertRow: true,
          allowManualInsertRow: true,
          allowDeleteRow: false,
          allowInsertColumn: false,
          allowDeleteColumn: false,
          allowRenameColumn: false,
          columnSorting: false,
          allowComments: false
        }
      ],
      onchange: (instance, _cell, colIndex, rowIndex, newValue) => {
        if (updating) return;
        const c = cols[Number(colIndex)];
        const raw = String(newValue ?? '');
        let fixed = raw;
        if (c.kind === 'time') fixed = parseTime(raw) || raw;
        if (c.kind === 'upper') fixed = raw.toUpperCase();
        if (fixed !== raw) {
          updating = true;
          instance.setValueFromCoords(Number(colIndex), Number(rowIndex), fixed);
          updating = false;
        }
        const rateX = cols.findIndex((col) => col.key === RATE_KEY[def.type]);
        if (ratesRef.current && rateX >= 0) {
          const y = Number(rowIndex);
          const line = toLine(cols, (instance.getData() as jspreadsheet.CellValue[][])[y]);
          const filled = fillRate(def.type, line, ratesRef.current, c.key);
          if (filled[RATE_KEY[def.type]] !== line[RATE_KEY[def.type]]) {
            updating = true;
            instance.setValueFromCoords(rateX, y, filled[RATE_KEY[def.type]] as number);
            updating = false;
          }
        }
        sync(instance);
      },
      oninsertrow: (instance) => sync(instance)
    });
    ws.hideIndex(); // the template has no row numbers
    sync(ws);

    // Tab at the end of a row goes to the start of the next one, adding a row after the last.
    const onKeyDown = (e: KeyboardEvent): void => {
      const sel = ws.selectedCell;
      if (e.key !== 'Tab' || e.shiftKey || !sel || jspreadsheet.current !== ws) return;
      const [x, y] = [Number(sel[0]), Number(sel[1])];
      if (x !== cols.length - 1) return;
      e.preventDefault();
      e.stopImmediatePropagation();
      if (ws.edition) ws.closeEditor(ws.edition[0], true);
      if (y === ws.options.data!.length - 1) ws.insertRow();
      ws.updateSelectionFromCoords(0, y + 1, 0, y + 1);
    };
    document.addEventListener('keydown', onKeyDown, true);

    // Template fills: group colour on the group heading and its column titles, grey elsewhere.
    const headRows = el.querySelectorAll('thead tr');
    const titles = headRows[headRows.length - 1];
    cols.forEach((c, x) => {
      const td = titles.querySelector<HTMLElement>(`td[data-x="${x}"]`);
      if (td) td.style.background = groupOf(def, c)?.fill || '#f1f1f1';
    });
    if (def.groups && headRows.length > 1) {
      // First cell is the corner above the row numbers.
      const groupCells = Array.from(headRows[0].querySelectorAll<HTMLElement>('td')).slice(1);
      groupCells.forEach((td) => {
        const g = (def.groups || []).find((x) => x.title === td.textContent?.trim());
        td.style.background = g ? g.fill : '#f1f1f1';
      });
    }

    return () => {
      document.removeEventListener('keydown', onKeyDown, true);
      jspreadsheet.destroy(el as jspreadsheet.JspreadsheetInstanceElement, true);
    };
    // Lines only seed the grid; the parent remounts it (key) to reset. A new person's rates rebuild the PSL options.
  }, [def, rates]);

  return <div ref={host} />;
}
