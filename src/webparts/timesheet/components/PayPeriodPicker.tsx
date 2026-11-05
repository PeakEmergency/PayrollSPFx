import * as React from 'react';
import { Dropdown, type IDropdownOption } from '@fluentui/react';
import { PayPeriod, displayDate, findPeriod, periodLabel, selectablePeriods, todayIso } from '../../../timesheet/payPeriods';

export interface IPayPeriodPickerProps {
  id: string;
  /** undefined while loading */
  periods: PayPeriod[] | undefined;
  error?: string;
  /** ISO dates or '' */
  from: string;
  to: string;
  disabled?: boolean;
  onChange(p: PayPeriod): void;
}

const key = (p: PayPeriod): string => `${p.year}-${p.number}`;

export default function PayPeriodPicker({ id, periods, error, from, to, disabled, onChange }: IPayPeriodPickerProps): React.ReactElement {
  const list = React.useMemo(() => selectablePeriods(periods || [], todayIso()), [periods]);
  const selected = findPeriod(list, from, to);
  const options: IDropdownOption[] = list.map((p) => ({ key: key(p), text: periodLabel(p) }));
  const message = error || (periods && !periods.length ? 'No pay periods are set up yet. Contact accounts.' : '');

  return (
    // Client: no "Pay period" label (the PAY PERIOD heading says it), From and To on one line beside the dropdown.
    <div style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'flex-start', gap: '6px 16px' }}>
      <Dropdown
        id={id}
        ariaLabel="Pay period"
        styles={{ root: { minWidth: 200 } }}
        dropdownWidth="auto"
        placeholder="Pick a pay period"
        options={options}
        selectedKey={selected ? key(selected) : null}
        disabled={disabled || !periods}
        onChange={(_e, o) => {
          const p = list.filter((x) => key(x) === o?.key)[0];
          if (p) onChange(p);
        }}
        errorMessage={message || undefined}
      />
      <div style={{ lineHeight: '32px', whiteSpace: 'nowrap' }}>
        From: <strong>{displayDate(from)}</strong>&nbsp;&nbsp; To: <strong>{displayDate(to)}</strong>
      </div>
    </div>
  );
}
