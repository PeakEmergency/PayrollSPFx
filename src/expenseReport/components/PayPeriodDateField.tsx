import * as React from 'react';
import { Calendar, Callout, DayOfWeek, DefaultButton, DirectionalHint, IconButton, TextField } from '@fluentui/react';
import { calendarStartDate, fromCalendarDate, parseDateInput } from '../domain/dates';
import styles from './ExpenseReport.module.scss';

export interface IPayPeriodDateFieldProps {
  id: string;
  label: 'From' | 'To';
  value: string;
  otherValue: string;
  disabled: boolean;
  errorMessage?: string;
  onChange: (value: string) => void;
}

const calendarStrings = {
  months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],
  goToToday: 'Today', prevMonthAriaLabel: 'Previous month', nextMonthAriaLabel: 'Next month',
  prevYearAriaLabel: 'Previous year', nextYearAriaLabel: 'Next year', closeButtonAriaLabel: 'Close calendar',
  monthPickerHeaderAriaLabel: '{0}, select to change month', yearPickerHeaderAriaLabel: '{0}, select to change year'
};

export default function PayPeriodDateField(props: IPayPeriodDateFieldProps): React.ReactElement {
  const [open, setOpen] = React.useState(false);
  const [rangeError, setRangeError] = React.useState('');
  const anchor = React.useRef<HTMLDivElement>(null);
  const button = React.useRef<HTMLButtonElement>(null);
  const close = (): void => { setOpen(false); button.current?.focus(); };
  const choose = (date: Date): void => {
    const raw = fromCalendarDate(date);
    const iso = parseDateInput(raw);
    const other = parseDateInput(props.otherValue);
    if (iso && other && ((props.label === 'To' && iso < other) || (props.label === 'From' && iso > other))) {
      setRangeError(props.label === 'To'
        ? `The TO date can't be before the FROM date (${props.otherValue}).`
        : `The FROM date can't be after the TO date (${props.otherValue}).`);
      return;
    }
    setRangeError(''); props.onChange(raw); close();
  };
  return <div className={styles.payDate} ref={anchor}>
    <div className={styles.dateEntry}>
      <TextField id={props.id} label={`${props.label}:`} value={props.value} placeholder="DD/MM/YY"
        className={!props.value ? styles.emptyDate : undefined} disabled={props.disabled} required
        errorMessage={props.errorMessage} autoComplete="off"
        onChange={(_, value) => { setRangeError(''); props.onChange(value || ''); }}
        onDoubleClick={() => !props.disabled && setOpen(true)}
        onKeyDown={event => { if (event.altKey && event.key === 'ArrowDown' && !props.disabled) { event.preventDefault(); setOpen(true); } }} />
      <IconButton elementRef={button} className={styles.calendarButton} iconProps={{ iconName: 'Calendar' }}
        ariaLabel={`Open ${props.label} calendar`} title={`Open ${props.label} calendar`}
        aria-expanded={open} aria-controls={open ? `${props.id}-calendar` : undefined} disabled={props.disabled}
        onClick={() => { setRangeError(''); setOpen(!open); }} />
    </div>
    <div className={styles.printDate}><strong>{props.label}:</strong> {props.value || '—'}</div>
    {open && !props.disabled && <Callout target={anchor} role="dialog" ariaLabel={`${props.label} calendar`}
      id={`${props.id}-calendar`} className={styles.calendarPopup} directionalHint={DirectionalHint.bottomLeftEdge}
      isBeakVisible={false} setInitialFocus calloutMaxWidth={320} onDismiss={() => setOpen(false)}
      onKeyDown={event => { if (event.key === 'Escape') { event.preventDefault(); close(); } }}>
      <Calendar value={calendarStartDate(props.value, props.label === 'To' ? props.otherValue : undefined)}
        firstDayOfWeek={DayOfWeek.Sunday} isMonthPickerVisible={false} showGoToToday={false}
        strings={calendarStrings} onSelectDate={choose} navigationIcons={{ leftNavigation: 'ChevronLeft', rightNavigation: 'ChevronRight' }}
        calendarDayProps={{ styles: {
          dayCell: { selectors: { ':first-child:not([aria-selected=true]), :last-child:not([aria-selected=true])': { background: '#f3f3f3' } } },
          daySelected: { background: '#17365d', color: '#fff' },
          dayIsToday: { background: '#dfe6f5', color: '#17365d' }
        } }} />
      <div className={styles.calendarActions}>
        <DefaultButton text="Today" onClick={() => choose(new Date())} />
        <DefaultButton text="Clear" onClick={() => { props.onChange(''); setRangeError(''); close(); }} />
        <DefaultButton text="Close" ariaLabel="Close calendar" onClick={close} />
      </div>
      {rangeError && <p className={styles.calendarError} role="alert">{rangeError}</p>}
    </Callout>}
  </div>;
}
