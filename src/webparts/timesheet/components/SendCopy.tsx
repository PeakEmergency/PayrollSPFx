import * as React from 'react';
import { ChoiceGroup, TextField } from '@fluentui/react';
import { copyProblem } from '../../../services/copyTo';

export interface ISendCopyProps {
  id: string;
  /** undefined: no copy. A string (maybe empty): send a copy to this address. */
  value: string | undefined;
  disabled?: boolean;
  /** Show the address error (after a submit attempt). */
  showError?: boolean;
  onChange(value: string | undefined): void;
}

/** "Send copy to" on both forms. The copy goes out with the submitted email only, never on Save. */
export default function SendCopy({ id, value, disabled, showError, onChange }: ISendCopyProps): React.ReactElement {
  return (
    <div style={{ margin: '10px 0' }}>
      <ChoiceGroup
        id={id}
        label="Send a copy to me"
        selectedKey={value === undefined ? 'no' : 'yes'}
        options={[{ key: 'no', text: 'No' }, { key: 'yes', text: 'Yes' }]}
        styles={{ flexContainer: { display: 'flex', gap: 16 } }}
        disabled={disabled}
        onChange={(_e, o) => onChange(o?.key === 'yes' ? '' : undefined)}
      />
      {value !== undefined && (
        <TextField
          id={`${id}-email`}
          label="Send copy to"
          type="email"
          placeholder="name@example.com"
          value={value}
          disabled={disabled}
          styles={{ root: { maxWidth: 360 } }}
          errorMessage={showError ? copyProblem(value) || undefined : undefined}
          onChange={(_e, v) => onChange(v || '')}
        />
      )}
    </div>
  );
}
