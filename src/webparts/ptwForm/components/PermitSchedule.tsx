import * as React from 'react';
import {
  Label, Checkbox, TextField, IDatePickerStyles, defaultDatePickerStrings,
  DatePicker, IColumn, DetailsList, SelectionMode,
  DetailsListLayoutMode,
  MessageBar,
  ComboBox,
  IComboBoxOption,
  IPersonaProps
} from '@fluentui/react';
import { IPermitScheduleProps } from '../../../Interfaces/PtwForm/IPermitSchedule';

const datePickerBlackStyles: Partial<IDatePickerStyles> = {
  root: { width: '100%', selectors: { '> *': { marginBottom: 15 } } },
  readOnlyTextField: {
    selectors: {
      '&.is-disabled .ms-TextField-field': { color: '#000 !important', fontWeight: 500, '-webkit-text-fill-color': '#000 !important' },
      '.field': { color: '#000 !important', fontWeight: 500, },
    }
  },
  textField: {
    selectors: {
      '&.is-disabled .ms-TextField-field': { color: '#000 !important', fontWeight: 500, '-webkit-text-fill-color': '#000 !important' },
      '.field': { color: '#000 !important', fontWeight: 500, },
    },
    field: { color: '#000 !important', fontWeight: 500, },
    root: { color: '#000  !important' },
    suffix: { color: '#000' },
    description: { color: '#000  !important' },
    fieldGroup: {
      // keep disabled background clean
      selectors: { '&.is-disabled': { background: 'transparent' } }
    }
  },
  icon: { color: '#000  !important' }
};

// Styling Components
// const comboBoxBlackStyles: Partial<IComboBoxStyles> = {
//   root: {
//     selectors: {
//       '.ms-ComboBox-Input': { color: '#000', fontWeight: 500, },
//       '&.is-disabled .ms-ComboBox-Input': { color: '#000', fontWeight: 500, },
//       '.ms-ComboBox-Input::placeholder': { color: '#000', fontWeight: 500, },
//     }
//   },
//   input: { color: '#000' } // supported in v8; safe no-op if ignored
// };

const PermitSchedule: React.FC<IPermitScheduleProps> = ({ workCategories,
  permitRows,
  selectedPermitTypeList,
  onPermitTypeChange,
  onPermitRowUpdate,
  styles,
  permitsValidityDays,
  permitStatus,
  isPermitIssuer,
  piApproverList,
  isIssued
}) => {

  const piStatusOptions: IComboBoxOption[] = React.useMemo(
    () => ['Approved', 'Rejected', 'Closed'].map(s => ({ key: s, text: s })),
    []
  );

  // Define DetailsList columns
  const columns: IColumn[] = React.useMemo(() => [
    {
      key: 'col-type', name: 'Type', minWidth: 165, maxWidth: 175,
      onRender: (row) => {
        const isNumericId = (id: string) => /^[0-9]+$/.test(String(id || ''));
        const hasPermitID = isNumericId(row.id);
        const isClosed = String(row.statusRecord ? row.statusRecord : '').toLowerCase() === 'closed';
        const disabled = (hasPermitID || isClosed);

        return (
          <Checkbox label={row.type === 'new' ? 'New Permit' : 'Permit Renewal'}
            checked={row.isChecked}
            onChange={(e, checked) => onPermitRowUpdate(row.id, 'type', row.id === "permit-row-0" ? 'new' : 'renewal', checked)}
            disabled={disabled}
          />
        );
      }

    },
    {
      key: 'col-date', name: 'Date', minWidth: 160, maxWidth: 170,
      onRender: (row) => (
        <DatePicker value={row.date ? new Date(row.date) : undefined} style={{ maxWidth: '100%' }} strings={defaultDatePickerStrings}
          onSelectDate={(date) => onPermitRowUpdate(row.id, 'date', date ? date.toISOString() : '', row.isChecked)}
          disabled={!row.isChecked}
          styles={datePickerBlackStyles}
        />
      )
    },
    {
      key: 'col-start', name: 'Starting Time', minWidth: 130, maxWidth: 140,
      onRender: (row) => (
        <TextField type="time"
          value={row.startTime || ''}
          style={{ width: '100%' }}
          max={row.endTime || undefined}
          step={60}
          onChange={(_, newValue) => onPermitRowUpdate(row.id, 'startTime', newValue || '', row.isChecked)}
          readOnly={(row.statusRecord?.toLowerCase() === 'closed')}
          disabled={!(row.isChecked || (row.statusRecord?.toLowerCase() === 'closed'))}
        />
      )
    },
    {
      key: 'col-end', name: 'Expiry Time', minWidth: 130, maxWidth: 140,
      onRender: (row) => (
        <TextField type="time"
          value={row.endTime || ''}
          style={{ width: '100%' }}
          min={row.startTime || undefined}
          step={60}
          onChange={(_, newValue) => onPermitRowUpdate(row.id, 'endTime', newValue || '', row.isChecked)}
          readOnly={(row.statusRecord?.toLowerCase() === 'closed')}
          disabled={!(row.isChecked || (row.statusRecord?.toLowerCase() === 'closed'))}
        />
      )
    },
    {
      key: 'col-piApprover', name: 'Permit Approver', minWidth: 140, maxWidth: 160,
      onRender: (row) => {
        const effectiveList = piApproverList ?? [];
        const isClosed = String(row.statusRecord ? row.statusRecord : '').toLowerCase() === 'closed';
        if (isClosed) {
          return <TextField value={row.piApprover ? row.piApprover?.text : ''} readOnly />
        }
        const selectedKey = row.piApprover && row.piApprover?.id !== undefined ? String(row.piApprover.id) : undefined;
        const options = effectiveList.map((m: IPersonaProps) => ({
          key: String(m.id),
          text: m.title || m.text || ''
        }));
        return (
          <ComboBox
            placeholder={row.isChecked ? "Select Approver" : ""}
            options={options}
            selectedKey={selectedKey}
            onChange={(_, opt) => onPermitRowUpdate(row.id, 'piApproverList', (opt?.key as string) || '', row.isChecked)}
            useComboBoxAsMenuWidth
            disabled={!row.isChecked}
          />
        );
      }
    },
    {
      key: 'col-status', name: 'Permit Status', minWidth: 140, maxWidth: 160,
      onRender: (row) => {
        const isClosed = String(row.statusRecord ? row.statusRecord : '').toLowerCase() === 'closed';
        if (isClosed) {
          return <TextField value={row.piStatus} readOnly />;
        }
        else {
          return (
            <ComboBox
              placeholder={row.isChecked ? "Select status" : ""}
              options={piStatusOptions.filter(opt => opt.key !== 'Closed')}
              selectedKey={row.piStatus || undefined}
              onChange={(_, opt) => onPermitRowUpdate(row.id, 'piStatus', (opt?.key as string) || '', row.isChecked)}
              useComboBoxAsMenuWidth
              disabled={!isPermitIssuer}
            />
          );
        }
      }
    },

  ], [onPermitRowUpdate, isPermitIssuer, piApproverList]);

  return (
    <div id="permitTypeScheduleSection" className={styles?.formBody} style={{ marginTop: '20px' }}>
      {/* Permit Type Selection */}
      <div className="row pb-3">
        <div>
          <Label className={styles?.ptwLabel}>Select Type of Permit / Work Category</Label>
        </div>
        <div className="form-group col-md-12">
          <div className="row p-2">
            {workCategories
              ?.sort((a, b) => {
                const aOrder = a.orderRecord !== undefined && a.orderRecord !== null ? Number(a.orderRecord) : Number.POSITIVE_INFINITY;
                const bOrder = b.orderRecord !== undefined && b.orderRecord !== null ? Number(b.orderRecord) : Number.POSITIVE_INFINITY;
                return aOrder - bOrder;
              })
              ?.map(category => {
                const checked = selectedPermitTypeList.some(p => p.id === category.id);
                return (
                  <div key={category.id} className="col-xl-2 col-lg-3 col-3 col-md-3 col-sm-6 col-12" style={{ marginBottom: '10px' }}>
                    <Checkbox
                      disabled={isIssued}
                      label={category.title}
                      checked={checked}
                      onChange={(e, checked) => onPermitTypeChange(checked, category)}
                    />
                  </div>
                );
              }
              )}
          </div>
        </div>
      </div>

      {/* Permit validity info */}
      {permitsValidityDays > 0 && (
        <div className="col-md-12" style={{ marginBottom: 5 }}>
          <MessageBar>
            {`Valid Permits: ${permitsValidityDays} (1 New${permitsValidityDays > 1 ? ` + ${permitsValidityDays - 1} renewal${permitsValidityDays - 1 > 1 ? 's' : ''}` : ''}) based on selected work categories.`}
          </MessageBar>
        </div>
      )}

      {/* Permit Schedule Table */}
      {workCategories && permitRows.length > 0 && (
        <div className="row pb-3">
          <div className="form-group col-md-12">
            <div className={styles?.permitTable} style={{ border: '1px solid #ccc', borderRadius: '4px', overflow: 'hidden', padding: '8px' }}>
              <DetailsList
                items={permitRows.sort((a, b) => { return a.orderRecord! - b.orderRecord! })}
                columns={columns}
                selectionMode={SelectionMode.none}
                layoutMode={DetailsListLayoutMode.fixedColumns}
                compact={true}
                getKey={(item) => item.id}
              />
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default PermitSchedule;