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

const formatComputed = (key: string, v: Line[string]): string =>
  typeof v !== 'number' ? '' : key === 'hours' ? v.toFixed(2) : money(v);

function toColumn(c: ColumnDef): jspreadsheet.Column {
  // Template's computed sub-total columns are ~1 char wide; keep them readable on screen.
  const base = { title: c.header, width: Math.max(c.width, 70), align: 'center' as const };
  switch (c.kind) {
    case 'date':
      return { ...base, type: 'calendar', options: { format: 'DD/MM/YY' } };
    case 'dropdown':
      return {
        ...base,
        type: 'dropdown',
        source: (c.options || []).map((o) => ({ id: o, name: plainNumberKeys.indexOf(c.key) >= 0 ? String(o) : money(o) }))
      };
    case 'computed':
      return { ...base, type: 'text', readOnly: true };
    default:
      return { ...base, type: 'text', align: c.kind === 'upper' ? 'left' : 'center' };
  }
}

/** Grid row (cells in column order) → Line. Blank cells become null; times stay as typed until valid. */
function toLine(def: TemplateDef, row: jspreadsheet.CellValue[]): Line {
  const line: Line = {};
  def.columns.forEach((c, x) => {
    if (c.kind === 'computed') return;
    const raw = row[x] === undefined || row[x] === null ? '' : String(row[x]).trim();
    if (!raw) line[c.key] = null;
    else if (c.kind === 'date') line[c.key] = normalizeDate(raw);
    else if (c.kind === 'number' || c.kind === 'dropdown') line[c.key] = isNaN(Number(raw)) ? raw : Number(raw);
    else line[c.key] = raw;
  });
  return line;
}

const toRow = (def: TemplateDef, line: Line | undefined): jspreadsheet.CellValue[] =>
  def.columns.map((c) => {
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
    let updating = false;

    const sync = (ws: jspreadsheet.WorksheetInstance): void => {
      const rows = ws.getData() as jspreadsheet.CellValue[][];
      const next = rows.map((r) => toLine(def, r));
      updating = true;
      next.forEach((line, y) => {
        const computed = computeLine(def.type, line);
        def.columns.forEach((c, x) => {
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
          data: Array.from({ length: def.rowCount }, (_, i) => toRow(def, lines[i])),
          columns: def.columns.map(toColumn),
          minDimensions: [def.columns.length, def.rowCount],
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
        const c = def.columns[Number(colIndex)];
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
    sync(ws);

    return () => jspreadsheet.destroy(el as jspreadsheet.JspreadsheetInstanceElement, true);
    // Lines only seed the grid; the parent remounts it (key) to reset.
  }, [def]);

  return <div ref={host} />;
}
