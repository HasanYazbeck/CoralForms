import * as React from 'react';
import { Label, Checkbox, TextField, IDatePickerStyles, defaultDatePickerStrings, DatePicker, IColumn, DetailsList, SelectionMode, DetailsListLayoutMode } from '@fluentui/react';
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
const PermitSchedule: React.FC<IPermitScheduleProps> = ({ workCategories,
  permitRows,
  selectedPermitTypeList,
  onPermitTypeChange,
  onPermitRowUpdate,
  styles
}) => {

  // Define DetailsList columns
  const columns: IColumn[] = React.useMemo(() => [
    {
      key: 'col-type', name: 'Type', minWidth: 165, maxWidth: 175,
      onRender: (row) => (
        <Checkbox label={row.type === 'new' ? 'New Permit' : 'Permit Renewal'} checked={row.isChecked}
          onChange={(e, checked) => onPermitRowUpdate(row.id, 'type', row.id === "permit-row-0" ? 'new' : 'renewal', checked )}
        />
      )
    },
    {
      key: 'col-date', name: 'Date', minWidth: 160, maxWidth: 170,
      onRender: (row) => (
        <DatePicker value={row.date ? new Date(row.date) : undefined} style={{ maxWidth: '100%' }} strings={defaultDatePickerStrings}
          onSelectDate={(date) => onPermitRowUpdate(row.id, 'date', date ? date.toISOString() : '',  row.isChecked)}
          disabled={!row.isChecked}
          styles={datePickerBlackStyles}
        />
      )
    },
    {
      key: 'col-start', name: 'Starting Time', minWidth: 130, maxWidth: 140,
      onRender: (row) => (
        <TextField type="time" value={row.startTime} style={{ width: '100%' }}
          onChange={(_, newValue) => onPermitRowUpdate(row.id, 'startTime', newValue || '',  row.isChecked)}
          disabled={!row.isChecked}
        />
      )
    },
    {
      key: 'col-end', name: 'Expiry Time', minWidth: 130, maxWidth: 140,
      onRender: (row) => (
        <TextField type="time" value={row.endTime} style={{ width: '100%' }}
          onChange={(_, newValue) => onPermitRowUpdate(row.id, 'endTime', newValue || '', row.isChecked)}
          disabled={!row.isChecked}
        />
      )
    }
  ], [onPermitRowUpdate]);

  return (
    <div id="permitTypeScheduleSection" className={styles?.formBody} style={{ marginTop: '20px' }}>
      {/* Permit Type Selection */}
      <div className="row pb-3">
        <div>
          <Label className={styles?.ptwLabel}>Select Type of Permit / Work Category</Label>
        </div>
        <div className="form-group col-md-12">
          <div className="row">
            {workCategories
              ?.sort((a, b) => {
                const aOrder = a.orderRecord !== undefined && a.orderRecord !== null ? Number(a.orderRecord) : Number.POSITIVE_INFINITY;
                const bOrder = b.orderRecord !== undefined && b.orderRecord !== null ? Number(b.orderRecord) : Number.POSITIVE_INFINITY;
                return aOrder - bOrder;
              })
              ?.map(category => {
                // const checked = selectedPermitTypeList.some(p => p.id === category.id);
                return (
                  <div key={category.id} className="col-xl-2 col-lg-3 col-3 col-md-3 col-sm-6 col-12" style={{ marginBottom: '10px' }}>
                    <Checkbox
                      label={category.title}
                      checked={category.isChecked}
                      onChange={(e, checked) => onPermitTypeChange(checked, category)}
                    />
                  </div>
                );
              }
              )}
          </div>
        </div>
      </div>

      {/* Permit Schedule Table */}
      {workCategories && permitRows.length > 0 && (
        <div className="row pb-3">
          {/* <div>
            <Label className={styles?.ptwLabel}>Permit Schedule - {selectedPermitType.title} {selectedPermitType.renewalValidity && selectedPermitType.renewalValidity > 0 &&
              ` (Renewable ${selectedPermitType.renewalValidity} times)`}
            </Label>
          </div> */}
          <div className="form-group col-md-12">
            <div className={styles?.permitTable} style={{ border: '1px solid #ccc', borderRadius: '4px', overflow: 'hidden', padding: '8px' }}>
              <DetailsList
                items={permitRows}
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