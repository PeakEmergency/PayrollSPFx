import * as React from 'react';
import { DefaultButton, MessageBar, MessageBarType, PrimaryButton } from '@fluentui/react';
import styles from './TimesheetForm.module.scss';
import TimesheetGrid from './TimesheetGrid';
import logo from '../assets/peak-logo.jpg';
import { computeTotals } from '../../../timesheet/calc';
import { TEMPLATES } from '../../../timesheet/definitions';
import { Line, TimesheetData, TimesheetType } from '../../../timesheet/types';
import { validate } from '../../../timesheet/validate';
import { SubmitResult } from '../../../services/submit';
import { PayPeriod } from '../../../timesheet/payPeriods';
import PayPeriodPicker from './PayPeriodPicker';
import { hasRate, PayRates } from '../../../timesheet/rates';

export interface ITimesheetFormProps {
  emailTo: string;
  getName(): Promise<string>;
  getPayPeriods(): Promise<PayPeriod[]>;
  /** The signed-in person's Pay Rates row; undefined when they have none. */
  getPayRates(): Promise<PayRates | undefined>;
  submit(data: TimesheetData): Promise<SubmitResult>;
  /** Shown in place of the timesheet when the Expense Report button is picked. */
  expenseReport: React.ReactNode;
}

const types: TimesheetType[] = ['Admin', 'Instructor', 'Medic'];

// Totals that are counts or hours, not money.
const plainTotals: string[] = ['days', 'adminHours', 'deductions', 'hours'];

const hasEntries = (lines: Line[]): boolean =>
  lines.some((l) => Object.keys(l).some((k) => l[k] !== null && l[k] !== undefined && l[k] !== ''));

export default function TimesheetForm({ emailTo, getName, getPayPeriods, getPayRates, submit, expenseReport }: ITimesheetFormProps): React.ReactElement {
  const [type, setType] = React.useState<TimesheetType | undefined>();
  // Both forms stay mounted once opened (just hidden), so switching never loses entries.
  const [showExpense, setShowExpense] = React.useState(false);
  const [expenseOpened, setExpenseOpened] = React.useState(false);
  const [pendingType, setPendingType] = React.useState<TimesheetType | undefined>();
  const [name, setName] = React.useState('');
  const [from, setFrom] = React.useState('');
  const [to, setTo] = React.useState('');
  const [periods, setPeriods] = React.useState<PayPeriod[] | undefined>();
  const [periodsError, setPeriodsError] = React.useState('');
  const [rates, setRates] = React.useState<PayRates | undefined>();
  const [ratesLoaded, setRatesLoaded] = React.useState(false);
  const [ratesError, setRatesError] = React.useState('');
  const [lines, setLines] = React.useState<Line[]>([]);
  const [gridKey, setGridKey] = React.useState(0);
  const [errors, setErrors] = React.useState<string[]>([]);
  const [sent, setSent] = React.useState('');
  const [submitting, setSubmitting] = React.useState(false);

  React.useEffect(() => {
    getName().then(setName, (err: Error) => setErrors([`Couldn't load your name: ${err.message}`]));
  }, []);

  React.useEffect(() => {
    getPayPeriods().then(setPeriods, (err: Error) => setPeriodsError(`Pay periods couldn't be loaded: ${err.message}`));
  }, []);

  React.useEffect(() => {
    getPayRates().then(
      (r) => {
        setRates(r);
        setRatesLoaded(true);
      },
      (err: Error) => setRatesError(`Your pay rates couldn't be loaded: ${err.message}`)
    );
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
    setShowExpense(false);
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
        setSent(`Timesheet submitted. It will be emailed to ${emailTo} within a few minutes.`);
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
  const rateProblem = !def ? '' : ratesError || (ratesLoaded && !hasRate(def.type, rates) ? 'No pay rate is set up for you on this timesheet. Contact accounts.' : '');

  return (
    <>
    <div className={styles.form}>
      <div className={styles.picker} role="group" aria-label="Form type">
        <DefaultButton
          text="Expense Report"
          toggle
          checked={showExpense}
          onClick={() => {
            setShowExpense(true);
            setExpenseOpened(true);
            setPendingType(undefined);
          }}
        />
        {types.map((t) => (
          <DefaultButton key={t} text={t} toggle checked={t === type && !showExpense} onClick={() => pick(t)} />
        ))}
      </div>

      <div hidden={showExpense}>
      {pendingType && (
        <div className={styles.confirm} role="alertdialog" aria-live="assertive">
          Switching to {pendingType} clears the rows you entered.
          <PrimaryButton text="Switch" onClick={() => switchTo(pendingType)} />
          <DefaultButton text="Cancel" onClick={() => setPendingType(undefined)} />
        </div>
      )}

      {def && (
        <>
          <div className={styles.header}>
            <div className={styles.titleRow}>
              <h1 className={styles.title}>{def.title}</h1>
              <img className={styles.logo} src={logo} alt="Peak Emergency" />
            </div>
            <div className={styles.headerRow}>
              <div className={styles.name} aria-label="Name">{name}</div>
              <div className={styles.period}>
                <h2 className={styles.periodLabel}>PAY PERIOD</h2>
                <PayPeriodPicker
                  id="pay-period"
                  periods={periods}
                  error={periodsError}
                  from={from}
                  to={to}
                  onChange={(p) => {
                    setFrom(p.from);
                    setTo(p.to);
                  }}
                />
              </div>
            </div>
          </div>

          <div className={styles.grid}>
            <TimesheetGrid key={`${def.type}-${gridKey}`} def={def} lines={lines} rates={rates} onChange={setLines} />
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

          {rateProblem && <MessageBar messageBarType={MessageBarType.error}>{rateProblem}</MessageBar>}
          <PrimaryButton text={submitting ? 'Sending…' : 'Submit'} disabled={submitting || !from || !to || !ratesLoaded || !!rateProblem} onClick={onSubmit} />
        </>
      )}

      <div aria-live="polite">
        {errors.length > 0 && (
          <div className={styles.errorSummary} role="alert">
            <strong>Can&apos;t submit yet</strong>
            <ul>
              {errors.map((e) => (
                <li key={e}>{e}</li>
              ))}
            </ul>
          </div>
        )}
        {sent && <MessageBar messageBarType={MessageBarType.success}>{sent}</MessageBar>}
      </div>
      </div>
    </div>
    {expenseOpened && <div hidden={!showExpense}>{expenseReport}</div>}
    </>
  );
}
