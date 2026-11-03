import * as React from 'react';
import { DefaultButton, Dropdown, IDropdownOption, MessageBar, MessageBarType, PrimaryButton, Spinner, SpinnerSize, ThemeProvider, createTheme } from '@fluentui/react';
import { IReportService } from '../services/IReportService';
import { ExpenseLine, ExpenseReport as Report, Instructor, ReceiptInfo, ReportSummary } from '../domain/models';
import { createReport, validateReport } from '../domain/report';
import ExpenseTable from './ExpenseTable';
import PayPeriodDateField from './PayPeriodDateField';
import Receipts from './Receipts';
import peakLogo from './assets/peak-logo.jpg';
import styles from './ExpenseReport.module.scss';

const confirmationStyles = styles as unknown as Record<string, string>;

export interface IExpenseReportProps {
  service: IReportService;
  defaultRate: string;
  instanceId: string;
  isDarkTheme?: boolean;
  isPreview?: boolean;
}

type PendingAction = { type: 'new' } | { type: 'open'; id: number };
const errorText = (error: unknown): string => error instanceof Error ? error.message : String(error);
const sameInputs = (left: Report, right: Report): boolean => {
  const headerFields: (keyof Report)[] = ['schemaVersion', 'clientId', 'id', 'instructorId', 'instructorName', 'fromDate', 'toDate'];
  const lineFields: (keyof ExpenseLine)[] = ['id', 'date', 'startCity', 'endCity', 'kilometres', 'rate', 'description', 'gst', 'cost'];
  return headerFields.every(field => left[field] === right[field]) && left.lines.length === right.lines.length &&
    left.lines.every((line, index) => lineFields.every(field => line[field] === right.lines[index][field]));
};

export default function ExpenseReport(props: IExpenseReportProps): React.ReactElement {
  const [report, setReport] = React.useState<Report>(() => createReport(props.defaultRate));
  const [baseline, setBaseline] = React.useState(() => JSON.stringify(report));
  const [instructors, setInstructors] = React.useState<Instructor[]>([]);
  const [reports, setReports] = React.useState<ReportSummary[]>([]);
  const [receipts, setReceipts] = React.useState<ReceiptInfo[]>([]);
  const [pendingFiles, setPendingFiles] = React.useState<File[]>([]);
  const [saveUncertain, setSaveUncertain] = React.useState(false);
  const [loading, setLoading] = React.useState(true);
  const [busy, setBusy] = React.useState<'saving' | 'opening' | undefined>();
  const busyRef = React.useRef(false);
  const [loadError, setLoadError] = React.useState('');
  const [listError, setListError] = React.useState('');
  const [operationError, setOperationError] = React.useState('');
  const [status, setStatus] = React.useState('');
  const [showValidation, setShowValidation] = React.useState(false);
  const [selectedReportId, setSelectedReportId] = React.useState<number | undefined>();
  const [pendingAction, setPendingAction] = React.useState<PendingAction>();
  const confirmRef = React.useRef<HTMLElement>(null);
  const restoreFocusRef = React.useRef<HTMLElement | null>(null);
  const [reload, setReload] = React.useState(0);
  const summary = React.useRef<HTMLDivElement>(null);
  const issues = showValidation ? validateReport(report, instructors) : [];
  const dirty = baseline !== JSON.stringify(report) || pendingFiles.length > 0 || saveUncertain;
  const locked = !!busy;
  const setOperationBusy = (phase?: 'saving' | 'opening'): void => {
    // A synchronous ref also blocks duplicate events before React commits the disabled controls.
    busyRef.current = !!phase;
    setBusy(phase);
  };
  const darkTheme = React.useMemo(() => createTheme({ isInverted: true, palette: {
    themePrimary: '#8bbcf0', neutralPrimary: '#ffffff', neutralLighterAlt: '#292929', neutralLighter: '#333333',
    neutralLight: '#454545', neutralQuaternaryAlt: '#555555', neutralQuaternary: '#666666', neutralTertiary: '#b5b5b5',
    neutralSecondary: '#dddddd', neutralDark: '#ffffff', white: '#202020', black: '#ffffff'
  } }), []);

  React.useEffect(() => {
    let active = true;
    setLoading(true); setLoadError(''); setListError('');
    const roster = props.service.getInstructors().then(items => {
      if (active) { setInstructors(items); if (!items.length) setLoadError('No instructors are configured. Ask your site administrator to add the instructor roster.'); }
    }).catch(error => { if (active) { setInstructors([]); setLoadError(errorText(error)); } });
    const list = props.service.listReports().then(items => { if (active) setReports(items); })
      .catch(error => { if (active) setListError(errorText(error)); });
    Promise.all([roster, list]).then(() => { if (active) setLoading(false); }).catch(() => undefined);
    return () => { active = false; };
  }, [props.service, reload]);

  React.useEffect(() => {
    if (!dirty) return;
    const prevent = (event: BeforeUnloadEvent): void => { event.preventDefault(); event.returnValue = ''; };
    window.addEventListener('beforeunload', prevent);
    return () => window.removeEventListener('beforeunload', prevent);
  }, [dirty]);

  React.useEffect(() => { if (showValidation) summary.current?.focus(); }, [showValidation]);
  React.useEffect(() => {
    if (pendingAction) {
      restoreFocusRef.current = document.activeElement instanceof HTMLElement ? document.activeElement : null;
      confirmRef.current?.querySelector<HTMLElement>('button')?.focus();
    } else if (!busyRef.current) {
      restoreFocusRef.current?.focus();
      restoreFocusRef.current = null;
    }
  }, [pendingAction, busy]);

  const setHeader = (field: 'fromDate' | 'toDate', value: string): void => {
    if (busyRef.current) return;
    setReport(current => ({ ...current, [field]: value })); setStatus('');
  };
  const setLine = (index: number, field: keyof ExpenseLine, value: string): void => {
    if (busyRef.current) return;
    setReport(current => ({ ...current, lines: current.lines.map((line, i) => i === index ? { ...line, [field]: value } : line) }));
    setStatus('');
  };
  const refreshList = async (): Promise<void> => {
    try { setReports(await props.service.listReports()); setListError(''); }
    catch (error) { setListError(`Saved reports could not be refreshed: ${errorText(error)}`); }
  };
  const save = async (): Promise<void> => {
    if (busyRef.current || loading || loadError) return;
    setShowValidation(true); setStatus(''); setOperationError('');
    if (validateReport(report, instructors).length) { summary.current?.focus(); return; }
    setOperationBusy('saving');
    try {
      const saved = report.id !== undefined && pendingFiles.length > 0 && JSON.stringify(report) === baseline
        ? report : await props.service.saveReport(report);
      setReport(saved); setBaseline(JSON.stringify(saved)); setSelectedReportId(saved.id);
      const failed: File[] = [];
      const failures: string[] = [];
      const uploaded: ReceiptInfo[] = [];
      for (const file of pendingFiles) {
        try {
          if (saved.id === undefined) throw new Error('The saved report did not return a report ID.');
          uploaded.push(await props.service.uploadReceipt(saved.id, file));
        } catch (error) { failed.push(file); failures.push(`${file.name}: ${errorText(error)}`); }
      }
      setReceipts(current => [...current.filter(receipt => !uploaded.some(item => item.name === receipt.name)), ...uploaded]);
      setPendingFiles(failed);
      let refreshError = '';
      if (pendingFiles.length && saved.id !== undefined) {
        try {
          const latest = await props.service.loadReport(saved.id);
          if (!sameInputs(saved, latest.report)) refreshError = 'The report changed elsewhere while receipts were uploading. Reopen the report before making further changes.';
          else {
            const versioned = { ...saved, etag: latest.report.etag };
            setReport(versioned); setBaseline(JSON.stringify(versioned));
            setSaveUncertain(false);
          }
        } catch (error) { refreshError = `The report version could not be refreshed after receipt upload. Reopen before making further changes. ${errorText(error)}`; }
        if (refreshError) setSaveUncertain(true);
      }
      if (failed.length) setOperationError(`Report saved, but ${failed.length} receipt upload${failed.length === 1 ? '' : 's'} failed. Retry Save to upload the remaining files. ${failures.join(' ')} ${refreshError}`);
      else if (refreshError) setOperationError(`Report saved and receipts uploaded. ${refreshError}`);
      else setStatus('Report saved successfully.');
      await refreshList();
    } catch (error) { setOperationError(`Report could not be saved: ${errorText(error)}`); }
    finally { setOperationBusy(); }
  };
  const performAction = async (action: PendingAction): Promise<void> => {
    if (busyRef.current) return;
    setPendingAction(undefined);
    if (action.type === 'new') {
      const blank = createReport(props.defaultRate);
      setReport(blank); setBaseline(JSON.stringify(blank)); setReceipts([]); setPendingFiles([]); setSaveUncertain(false);
      setSelectedReportId(undefined); setShowValidation(false); setOperationError(''); setStatus('');
      return;
    }
    setOperationBusy('opening'); setOperationError(''); setStatus('');
    try {
      const loaded = await props.service.loadReport(action.id);
      setReport(loaded.report); setBaseline(JSON.stringify(loaded.report)); setReceipts(loaded.receipts); setPendingFiles([]);
      setSaveUncertain(false);
      setShowValidation(false); setStatus('Report opened.');
    } catch (error) { setOperationError(`Report could not be opened: ${errorText(error)}`); }
    finally { setOperationBusy(); }
  };
  const requestAction = (action: PendingAction): void => {
    if (busyRef.current) return;
    if (dirty) setPendingAction(action); else performAction(action).catch(error => setOperationError(errorText(error)));
  };
  const options: IDropdownOption[] = instructors.filter(person => person.active !== false || person.id === report.instructorId)
    .map(person => ({ key: person.id, text: person.id === report.instructorId ? report.instructorName : person.name }));
  if (report.instructorId && !options.some(option => option.key === report.instructorId)) {
    options.push({ key: report.instructorId, text: report.instructorName, disabled: true });
  }
  const issue = (field: string): string | undefined => issues.find(item => item.field === field)?.message;

  return <ThemeProvider theme={props.isDarkTheme ? darkTheme : undefined} className={`${styles.expenseReport} ${props.isDarkTheme ? styles.dark : ''}`}>
    <section aria-labelledby={`${props.instanceId}-title`} className={styles.reportSurface}>
      {props.isPreview && <MessageBar className={styles.noPrint}>Local preview — reports and receipts are stored in memory for this session.</MessageBar>}
      <div className={styles.brandRow}>
        <h1 id={`${props.instanceId}-title`}>EXPENSE REPORT</h1>
        <img className={styles.logo} src={peakLogo} alt="Peak paramedic services" />
      </div>
      <div className={styles.reportHeader}>
        <Dropdown id={`${props.instanceId}-instructor`} label="Instructor" className={styles.instructor}
          placeholder="PICK FROM DROPDOWN" selectedKey={report.instructorId || null} options={options}
          disabled={loading || locked || !instructors.length} required errorMessage={issue('instructorId')}
          onChange={(_, option) => {
            if (busyRef.current || !option) return;
            const selected = instructors.find(person => person.id === option.key && person.active !== false);
            if (selected) { setReport(current => ({ ...current, instructorId: selected.id, instructorName: selected.name })); setStatus(''); }
          }} />
        <div className={styles.printInstructor}><span>Instructor</span><strong>{report.instructorName || '—'}</strong></div>
        <section className={styles.payPeriod} aria-labelledby={`${props.instanceId}-pay-period`}>
          <h2 id={`${props.instanceId}-pay-period`}>PAY PERIOD</h2>
          <div className={styles.payFields}>
            <PayPeriodDateField id={`${props.instanceId}-from`} label="From" value={report.fromDate} otherValue={report.toDate}
              disabled={locked} errorMessage={issue('fromDate')} onChange={value => setHeader('fromDate', value)} />
            <PayPeriodDateField id={`${props.instanceId}-to`} label="To" value={report.toDate} otherValue={report.fromDate}
              disabled={locked} errorMessage={issue('toDate')} onChange={value => setHeader('toDate', value)} />
          </div>
        </section>
      </div>
      <div className={styles.noPrint}>
        {loading && <Spinner size={SpinnerSize.small} label="Loading instructors and saved reports…" />}
        {(loadError || listError) && <MessageBar messageBarType={MessageBarType.error} isMultiline>
          {loadError && <p>{loadError}</p>}{listError && <p>{listError}</p>}
          <DefaultButton text="Retry loading" disabled={locked || loading} onClick={() => setReload(value => value + 1)} />
        </MessageBar>}
        {!!issues.length && <div ref={summary} className={styles.errorSummary} role="alert" aria-label="Required fields missing" tabIndex={-1}>
          <strong>Required fields missing or invalid</strong>
          <p>This expense report can&apos;t be saved yet. Please fill in or correct:</p>
          <ul>{issues.map((item, index) => <li key={`${item.field}-${index}`}>{item.message}</li>)}</ul>
        </div>}
        {operationError && <MessageBar messageBarType={MessageBarType.error} isMultiline>{operationError}</MessageBar>}
        {status && <MessageBar messageBarType={MessageBarType.success} role="status">{status}</MessageBar>}
      </div>
      <ExpenseTable report={report} instanceId={props.instanceId} disabled={locked} issues={issues} onChange={setLine} />
      <Receipts instanceId={props.instanceId} receipts={receipts} pending={pendingFiles} disabled={locked}
        onAdd={files => { if (!busyRef.current) { setPendingFiles(current => [...current, ...files]); setStatus(''); } }}
        onRemove={index => { if (!busyRef.current) setPendingFiles(current => current.filter((_, i) => i !== index)); }} />
      <div className={`${styles.actionArea} ${styles.noPrint}`}>
        <div className={styles.actions}>
          <PrimaryButton text="Save" disabled={locked || loading || !!loadError} onClick={() => { save().catch(error => setOperationError(errorText(error))); }} />
          <DefaultButton text="New report" disabled={locked} onClick={() => requestAction({ type: 'new' })} />
          <DefaultButton text="Print" disabled={locked} onClick={() => window.print()} />
          {busy && <Spinner size={SpinnerSize.small} label={busy === 'saving' ? 'Saving report and receipts…' : 'Opening report…'} />}
        </div>
        <div className={styles.reopen}>
          <Dropdown id={`${props.instanceId}-saved-reports`} label="Saved reports" placeholder="Choose a report"
            selectedKey={selectedReportId ?? null} disabled={locked || loading || !reports.length}
            options={reports.map(item => ({ key: item.id, text: `${item.instructorName} · ${item.fromDate} – ${item.toDate} (#${item.id})` }))}
            onChange={(_, option) => setSelectedReportId(option ? Number(option.key) : undefined)} />
          <DefaultButton text="Reopen" disabled={locked || selectedReportId === undefined || loading} onClick={() => selectedReportId !== undefined && requestAction({ type: 'open', id: selectedReportId })} />
        </div>
      </div>
      {pendingAction && <div className={confirmationStyles.confirmOverlay}>
        <section ref={confirmRef} className={confirmationStyles.confirmDialog} role="dialog" aria-modal="true" aria-labelledby={`${props.instanceId}-discard-title`}
          onKeyDown={event => {
            if (event.key === 'Escape') { event.preventDefault(); setPendingAction(undefined); }
            if (event.key === 'Tab') {
              const buttons = confirmRef.current?.querySelectorAll<HTMLElement>('button');
              if (!buttons?.length) return;
              const first = buttons[0]; const last = buttons[buttons.length - 1];
              if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus(); }
              else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus(); }
            }
          }}>
          <h2 id={`${props.instanceId}-discard-title`}>Discard unsaved changes?</h2>
          <p>Your report has unsaved changes or pending receipts. Discard them to continue.</p>
          <div className={confirmationStyles.confirmActions}>
            <PrimaryButton text="Keep editing" onClick={() => setPendingAction(undefined)} />
            <DefaultButton text="Discard changes" onClick={() => { if (pendingAction) performAction(pendingAction).catch(error => setOperationError(errorText(error))); }} />
          </div>
        </section>
      </div>}
    </section>
  </ThemeProvider>;
}
