import * as React from 'react';
import { TextField } from '@fluentui/react';
import { ExpenseLine, ExpenseReport, ValidationIssue } from '../domain/models';
import { calculateReport, formatCurrency, formatQuantity } from '../domain/calculations';
import DateField from './DateField';
import styles from './ExpenseReport.module.scss';

export interface IExpenseTableProps {
  report: ExpenseReport;
  instanceId: string;
  disabled: boolean;
  issues: ValidationIssue[];
  onChange: (index: number, field: keyof ExpenseLine, value: string) => void;
}

export default function ExpenseTable({ report, instanceId, disabled, issues, onChange }: IExpenseTableProps): React.ReactElement {
  const totals = calculateReport(report);
  const header = (field: string): string => `${instanceId}-head-${field}`;
  const dateInput = (index: number): React.ReactElement => {
    const error = issues.find(issue => issue.field === `lines.${index}.date`)?.message;
    return <DateField id={`${instanceId}-row-${index}-date`} label="" ariaLabel={`Row ${index + 1} date`}
      value={report.lines[index].date} disabled={disabled} errorMessage={error} onChange={value => onChange(index, 'date', value)} />;
  };
  const input = (index: number, field: keyof ExpenseLine, name: string, numeric = false): React.ReactElement => {
    const error = issues.find(issue => issue.field === `lines.${index}.${field}`)?.message;
    return <><TextField id={`${instanceId}-row-${index}-${field}`} ariaLabel={`Row ${index + 1} ${name}`}
      value={report.lines[index][field]} disabled={disabled} borderless errorMessage={error}
      aria-invalid={!!error} autoComplete="off"
      className={numeric ? styles.numericInput : styles.cellInput} inputMode={numeric ? 'decimal' : undefined}
      onChange={(_, value) => onChange(index, field, value || '')} />
      <span className={styles.printValue}>{report.lines[index][field] || '\u00a0'}</span></>;
  };
  return <div className={styles.tableScroll} role="region" aria-label="Expense details, 13 rows" tabIndex={0}>
    <table className={styles.expenseTable} aria-label="Expense report details">
      <colgroup><col className={styles.dateColumn} /><col className={styles.cityColumn} /><col className={styles.cityColumn} />
        <col className={styles.kmColumn} /><col className={styles.moneyColumn} /><col className={styles.descriptionColumn} />
        <col className={styles.moneyColumn} /><col className={styles.moneyColumn} /><col className={styles.amountColumn} /></colgroup>
      <thead>
        <tr><th rowSpan={2} scope="col" id={header('date')} className={styles.dateHeading}>DATE</th>
          <th colSpan={4} scope="colgroup" id={header('mileage')} className={styles.mileageGroup}>PERSONAL VEHICLE MILEAGE FEES</th>
          <th colSpan={3} scope="colgroup" id={header('other')} className={styles.otherGroup}>OTHER FEES</th>
          <th rowSpan={2} scope="col" id={header('amount')} className={styles.amountHeading}>AMOUNT</th></tr>
        <tr><th scope="col" id={header('start')} className={styles.mileageStart}>START<small>CITY &amp; POSTAL CODE</small></th>
          <th scope="col" id={header('end')} className={styles.mileageHeading}>END<small>CITY &amp; POSTAL CODE</small></th>
          <th scope="col" id={header('km')} className={styles.mileageHeading}>TOTAL<br />km</th>
          <th scope="col" id={header('mileageTotal')} className={styles.mileageHeading}>TOTAL</th>
          <th scope="col" id={header('description')} className={styles.otherStart}>DESCRIPTION</th>
          <th scope="col" id={header('gst')} className={styles.otherHeading}>GST</th>
          <th scope="col" id={header('cost')} className={styles.otherHeading}>COST<br />(excl. GST)</th></tr>
      </thead>
      <tbody>{report.lines.map((line, index) => <tr key={line.id}>
        <td headers={header('date')} className={styles.dateCell}>{dateInput(index)}</td>
        <td headers={`${header('mileage')} ${header('start')}`} className={styles.sectionStart}>{input(index, 'startCity', 'start city and postal code')}</td>
        <td headers={`${header('mileage')} ${header('end')}`}>{input(index, 'endCity', 'end city and postal code')}</td>
        <td headers={`${header('mileage')} ${header('km')}`}>{input(index, 'kilometres', 'total km', true)}</td>
        <td headers={`${header('mileage')} ${header('mileageTotal')}`} aria-label={`Row ${index + 1} mileage total`} className={styles.calculated}>{formatCurrency(totals.lines[index].mileage)}</td>
        <td headers={`${header('other')} ${header('description')}`} className={styles.sectionStart}>{input(index, 'description', 'description')}</td>
        <td headers={`${header('other')} ${header('gst')}`}>{input(index, 'gst', 'GST', true)}</td>
        <td headers={`${header('other')} ${header('cost')}`}>{input(index, 'cost', 'cost excluding GST', true)}</td>
        <td headers={header('amount')} aria-label={`Row ${index + 1} amount`} className={`${styles.calculated} ${styles.amountCell}`}>{formatCurrency(totals.lines[index].amount)}</td>
      </tr>)}</tbody>
      <tfoot><tr><th colSpan={3} scope="row">TOTALS</th>
        <td aria-label="Report total km" className={`${styles.calculated} ${styles.plainTotal}`}>{formatQuantity(totals.kilometres)}</td>
        <td aria-label="Report mileage total" className={styles.calculated}>{formatCurrency(totals.mileage)}</td>
        <td className={styles.sectionStart} />
        <td aria-label="Report GST total" className={styles.calculated}>{formatCurrency(totals.gst)}</td>
        <td aria-label="Report cost total" className={`${styles.calculated} ${styles.plainTotal}`}>{formatCurrency(totals.cost)}</td>
        <td aria-label="Report amount" className={`${styles.calculated} ${styles.amountCell}`}>{formatCurrency(totals.amount)}</td>
      </tr></tfoot>
    </table>
  </div>;
}
