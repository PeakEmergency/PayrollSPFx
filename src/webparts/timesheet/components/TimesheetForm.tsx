import * as React from 'react';
import styles from './TimesheetForm.module.scss';
import TimesheetGrid from './TimesheetGrid';
import logo from '../assets/peak-logo.jpg';
import { computeTotals } from '../../../timesheet/calc';
import { TEMPLATES } from '../../../timesheet/definitions';
import { Line, TimesheetData, TimesheetType } from '../../../timesheet/types';
import { validate } from '../../../timesheet/validate';
import { SubmitResult } from '../../../services/submit';

export interface ITimesheetFormProps {
  emailTo: string;
  getName(): Promise<string>;
  submit(data: TimesheetData): Promise<SubmitResult>;
}

const types: TimesheetType[] = ['Admin', 'Instructor', 'Medic'];

// Totals that are counts or hours, not money.
const plainTotals: string[] = ['days', 'adminHours', 'deductions', 'hours'];

const hasEntries = (lines: Line[]): boolean =>
  lines.some((l) => Object.keys(l).some((k) => l[k] !== null && l[k] !== undefined && l[k] !== ''));

export default function TimesheetForm({ emailTo, getName, submit }: ITimesheetFormProps): React.ReactElement {
  const [type, setType] = React.useState<TimesheetType | undefined>();
  const [pendingType, setPendingType] = React.useState<TimesheetType | undefined>();
  const [name, setName] = React.useState('');
  const [from, setFrom] = React.useState('');
  const [to, setTo] = React.useState('');
  const [lines, setLines] = React.useState<Line[]>([]);
  const [gridKey, setGridKey] = React.useState(0);
  const [errors, setErrors] = React.useState<string[]>([]);
  const [sent, setSent] = React.useState('');
  const [submitting, setSubmitting] = React.useState(false);

  React.useEffect(() => {
    getName().then(setName, (err: Error) => setErrors([`Couldn't load your name: ${err.message}`]));
  }, []);

  const switchTo = (t: TimesheetType): void => {
    setType(t);
    setPendingType(undefined);
    setLines([]);
    setErrors([]);
    setSent('');
    setGridKey((k) => k + 1);
  };

  const pick = (t: TimesheetType): void => {
    if (t === type) return;
    if (hasEntries(lines)) setPendingType(t);
    else switchTo(t);
  };

  const onSubmit = async (): Promise<void> => {
    if (!type) return;
    const data: TimesheetData = { type, name, from: from || null, to: to || null, lines };
    setSent('');
    const problems = validate(data);
    if (problems.length) {
      setErrors(problems);
      return;
    }
    setErrors([]);
    setSubmitting(true);
    try {
      const result = await submit(data);
      if (result.ok) {
        setSent(`Timesheet sent to ${emailTo}.`);
        setLines([]);
        setFrom('');
        setTo('');
        setGridKey((k) => k + 1);
      } else {
        setErrors(result.errors);
      }
    } catch (err) {
      setErrors([err instanceof Error ? err.message : String(err)]);
    } finally {
      setSubmitting(false);
    }
  };

  const def = type ? TEMPLATES[type] : undefined;
  const totals = def ? computeTotals(def.type, lines) : {};

  return (
    <div className={styles.form}>
      <div className={styles.picker} role="group" aria-label="Timesheet type">
        {types.map((t) => (
          <button key={t} type="button" className={t === type ? styles.active : ''} aria-pressed={t === type} onClick={() => pick(t)}>
            {t}
          </button>
        ))}
      </div>

      {pendingType && (
        <div className={styles.confirm} role="alertdialog" aria-live="assertive">
          Switching to {pendingType} clears the rows you entered.
          <button type="button" onClick={() => switchTo(pendingType)}>Switch</button>
          <button type="button" onClick={() => setPendingType(undefined)}>Cancel</button>
        </div>
      )}

      {def && (
        <>
          <div className={styles.header}>
            <div className={styles.titleRow}>
              <div className={styles.title}>{def.title}</div>
              <img className={styles.logo} src={logo} alt="Peak Emergency" />
            </div>
            <div className={styles.headerRow}>
              <div className={styles.name} aria-label="Name">{name}</div>
              <div className={styles.period}>
                <div className={styles.periodLabel}>PAY PERIOD</div>
                <label>
                  From: <input type="date" className={from ? '' : styles.empty} value={from} onChange={(e) => setFrom(e.target.value)} />
                </label>
                <label>
                  To: <input type="date" className={to ? '' : styles.empty} value={to} onChange={(e) => setTo(e.target.value)} />
                </label>
              </div>
            </div>
          </div>

          <div className={styles.grid}>
            <TimesheetGrid key={`${def.type}-${gridKey}`} def={def} lines={lines} onChange={setLines} />
          </div>

          <div className={styles.totals}>
            <span className={styles.totalsLabel}>TOTALS</span>
            {def.totals.map(({ key }) => (
              <span key={key}>
                {def.columns.find((c) => c.key === key)!.header}:{' '}
                <strong>{plainTotals.indexOf(key) >= 0 ? totals[key].toFixed(2) : `$${totals[key].toFixed(2)}`}</strong>
              </span>
            ))}
          </div>

          <button type="button" className={styles.submit} disabled={submitting || !from || !to} onClick={onSubmit}>
            {submitting ? 'Sending…' : 'Submit'}
          </button>
        </>
      )}

      <div aria-live="polite">
        {errors.length > 0 && (
          <ul className={styles.errors}>
            {errors.map((e) => (
              <li key={e}>{e}</li>
            ))}
          </ul>
        )}
        {sent && <div className={styles.sent}>{sent}</div>}
      </div>
    </div>
  );
}
