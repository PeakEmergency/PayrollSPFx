import * as React from 'react';
import { DefaultButton, Label, Link } from '@fluentui/react';
import { ReceiptInfo } from '../domain/models';
import styles from './ExpenseReport.module.scss';

export interface IReceiptsProps {
  instanceId: string;
  receipts: ReceiptInfo[];
  pending: File[];
  disabled: boolean;
  onAdd: (files: File[]) => void;
  onRemove: (index: number) => void;
}

export default function Receipts(props: IReceiptsProps): React.ReactElement {
  const id = `${props.instanceId}-receipts`;
  return <section className={styles.receipts} aria-labelledby={`${id}-title`}>
    <h2 id={`${id}-title`}>Receipts <span>(optional)</span></h2>
    <div className={styles.noPrint}>
      <Label htmlFor={id}>Add receipts</Label>
      <input id={id} type="file" multiple disabled={props.disabled} onChange={event => {
        props.onAdd(Array.from(event.currentTarget.files || [])); event.currentTarget.value = '';
      }} />
    </div>
    {!!props.receipts.length && <ul className={styles.receiptList}>{props.receipts.map(receipt => <li key={receipt.name}>
      <Link href={receipt.url} target="_blank" rel="noopener noreferrer">{receipt.name}</Link> <span>Saved</span>
    </li>)}</ul>}
    {!!props.pending.length && <ul className={`${styles.receiptList} ${styles.noPrint}`}>{props.pending.map((file, index) => <li key={`${file.name}-${file.lastModified}-${index}`}>
      <span>{file.name}</span> <span>Pending upload</span>
      <DefaultButton text="Remove" ariaLabel={`Remove ${file.name}`} disabled={props.disabled} onClick={() => props.onRemove(index)} />
    </li>)}</ul>}
  </section>;
}
