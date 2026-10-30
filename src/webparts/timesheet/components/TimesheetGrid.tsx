import * as React from 'react';
import jspreadsheet from 'jspreadsheet-ce';
import 'jsuites/dist/jsuites.css';
import 'jspreadsheet-ce/dist/jspreadsheet.css';
import { computeLine } from '../../../timesheet/calc';
import { ColumnDef, Line, TemplateDef } from '../../../timesheet/types';
import { normalizeDate, parseTime } from '../../../timesheet/values';

export interface ITimesheetGridProps {
  def: TemplateDef;
  lines: Line[];
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

function toColumn(c: ColumnDef): jspreadsheet.Column {
  // Excel shows a column ~5px wider than width × 7; 90px fits the bold two-line headers.
  const base = { title: c.header, width: Math.max(c.width + 5, 90), align: 'center' as const };
  switch (c.kind) {
    case 'date':
      return { ...base, type: 'calendar', options: { format: 'DD/MM/YY' } };
    case 'dropdown':
      return {
        ...base,
        type: 'dropdown',
        // Labelled rates show the classification; the cell value stays the dollar amount.
        width: c.optionLabels ? 230 : base.width,
        source: (c.options || []).map((o, i) => ({
          id: o,
          name: c.optionLabels ? c.optionLabels[i] : plainNumberKeys.indexOf(c.key) >= 0 ? String(o) : money(o)
        }))
      };
    case 'time':
      return { ...base, type: 'dropdown', autocomplete: true, source: times };
    case 'computed':
      return { ...base, type: 'text', readOnly: true };
    default:
      return { ...base, type: 'text', align: c.kind === 'upper' ? 'left' : 'center' };
  }
}

/**
 * Columns on screen. The template's sub-total columns are ~1 char wide, so they are
 * left off screen like in Excel; they are still computed and written to the file.
 */
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
export default function TimesheetGrid({ def, lines, onChange }: ITimesheetGridProps): React.ReactElement {
  const host = React.useRef<HTMLDivElement>(null);
  const onChangeRef = React.useRef(onChange);
  onChangeRef.current = onChange;

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
      onChangeRef.current(next);
    };

    const [ws] = jspreadsheet(el as HTMLDivElement, {
      tabs: false,
      toolbar: false,
      worksheets: [
        {
          data: Array.from({ length: def.rowCount }, (_, i) => toRow(cols, lines[i])),
          columns: cols.map(toColumn),
          nestedHeaders: def.groups ? nestedHeaders(def, cols) : undefined,
          minDimensions: [cols.length, def.rowCount],
          allowInsertRow: false,
          allowManualInsertRow: false,
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
        sync(instance);
      }
    });
    ws.hideIndex(); // the template has no row numbers
    sync(ws);

    // Template fills: group colour on the group heading and its column titles, grey elsewhere.
    const headRows = el.querySelectorAll('thead tr');
    const titles = headRows[headRows.length - 1];
    cols.forEach((c, x) => {
      const td = titles.querySelector<HTMLElement>(`td[data-x="${x}"]`);
      if (td) td.style.background = groupOf(def, c)?.fill || '#f2f2f2';
    });
    if (def.groups && headRows.length > 1) {
      // First cell is the corner above the row numbers.
      const groupCells = Array.from(headRows[0].querySelectorAll<HTMLElement>('td')).slice(1);
      groupCells.forEach((td) => {
        const g = (def.groups || []).find((x) => x.title === td.textContent?.trim());
        td.style.background = g ? g.fill : '#f2f2f2';
      });
    }

    return () => jspreadsheet.destroy(el as jspreadsheet.JspreadsheetInstanceElement, true);
    // Lines only seed the grid; the parent remounts it (key) to reset.
  }, [def]);

  return <div ref={host} />;
}
