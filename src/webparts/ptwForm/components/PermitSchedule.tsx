import * as React from 'react';
import { Label, Checkbox, TextField } from '@fluentui/react';
import { IPermitScheduleProps } from '../../../Interfaces/PtwForm/IPermitSchedule';

const PermitSchedule: React.FC<IPermitScheduleProps> = ({
  workCategories,
  selectedPermitType,
  permitRows,
  onPermitTypeChange,
  onPermitRowUpdate,
  styles
}) => {

  return (
    <div id="permitTypeScheduleSection" className={styles?.formBody} style={{ marginTop: '20px' }}>
      {/* Permit Type Selection */}
      <div className="row pb-3">
        <div>
          <Label className={styles?.ptwLabel}>Select Type of Permit / Work Category</Label>
        </div>
        <div className="form-group col-md-12">
          <div className={styles?.checkboxContainer}>
            {workCategories
              ?.sort((a, b) => {
                const aOrder = a.orderRecord !== undefined && a.orderRecord !== null ? Number(a.orderRecord) : Number.POSITIVE_INFINITY;
                const bOrder = b.orderRecord !== undefined && b.orderRecord !== null ? Number(b.orderRecord) : Number.POSITIVE_INFINITY;
                return aOrder - bOrder;
              })
              ?.map(category => (
                <div key={category.id} className="col-xl-2 col-lg-3 col-md-4 col-sm-6 col-12" style={{ marginBottom: '10px' }}>
                  <Checkbox 
                    label={category.title}
                    checked={selectedPermitType?.id === category.id}
                    onChange={(_, checked) => {
                      if (checked) {
                        onPermitTypeChange(category);
                      } else {
                        onPermitTypeChange(undefined);
                      }
                    }}
                  />
                </div>
              ))}
          </div>
        </div>
      </div>

      {/* Permit Schedule Table */}
      {selectedPermitType && permitRows.length > 0 && (
        <div className="row pb-3">
          <div>
            <Label className={styles?.ptwLabel}>
              Permit Schedule - {selectedPermitType.title} 
              {selectedPermitType.renewalValidity && selectedPermitType.renewalValidity > 0 && 
                ` (Renewable ${selectedPermitType.renewalValidity} times)`
              }
            </Label>
          </div>
          <div className="form-group col-md-12">
            <div className={styles?.permitTable} style={{ 
              border: '1px solid #ccc', 
              borderRadius: '4px',
              overflow: 'hidden'
            }}>
              {/* Table Header */}
              <div style={{ 
                display: 'grid', 
                gridTemplateColumns: '1fr 2fr 2fr 2fr',
                backgroundColor: '#f5f5f5',
                borderBottom: '1px solid #ccc',
                fontWeight: 'bold',
                padding: '10px'
              }}>
                <div style={{ padding: '5px', borderRight: '1px solid #ccc' }}>Type</div>
                <div style={{ padding: '5px', borderRight: '1px solid #ccc' }}>Date</div>
                <div style={{ padding: '5px', borderRight: '1px solid #ccc' }}>Starting Time</div>
                <div style={{ padding: '5px' }}>Expiry Time</div>
              </div>

              {/* Table Rows */}
              {permitRows.map((row, index) => (
                <div key={row.id} style={{ 
                  display: 'grid', 
                  gridTemplateColumns: '1fr 2fr 2fr 2fr',
                  borderBottom: index < permitRows.length - 1 ? '1px solid #ccc' : 'none',
                  backgroundColor: index % 2 === 0 ? '#fff' : '#f9f9f9'
                }}>
                  {/* Type Column */}
                  <div style={{ 
                    padding: '10px', 
                    borderRight: '1px solid #ccc',
                    display: 'flex',
                    alignItems: 'center'
                  }}>
                    <Checkbox 
                      label={row.type === 'new' ? 'New Permit' : 'Permit Renewal'}
                      checked={true}
                      disabled={true}
                    />
                  </div>

                  {/* Date Column */}
                  <div style={{ 
                    padding: '10px', 
                    borderRight: '1px solid #ccc' 
                  }}>
                    <TextField
                      type="date"
                      value={row.date}
                      onChange={(e, newValue) => onPermitRowUpdate(row.id, 'date', newValue || '')}
                      style={{
                        width: '100%',
                        padding: '5px',
                        border: '1px solid #ccc',
                        borderRadius: '4px'
                      }}
                    />
                  </div>

                  {/* Starting Time Column */}
                  <div style={{ 
                    padding: '10px', 
                    borderRight: '1px solid #ccc',
                    display: 'flex',
                    gap: '5px',
                    alignItems: 'center'
                  }}>
                    <TextField
                      type="time"
                      value={row.startTime}
                      onChange={(e, newValue) => onPermitRowUpdate(row.id, 'startTime', newValue || '')}
                      style={{
                        flex: 1,
                        padding: '5px',
                        border: '1px solid #ccc',
                        borderRadius: '4px'
                      }}
                    />
                    {/* <div style={{ display: 'flex', gap: '5px' }}>
                      <label style={{ display: 'flex', alignItems: 'center', gap: '2px' }}>
                        <input
                          type="radio"
                          name={`${row.id}-start-ampm`}
                          checked={row.startAmPm === 'AM'}
                          onChange={() => onPermitRowUpdate(row.id, 'startAmPm', 'AM')}
                        />
                        AM
                      </label>
                      <label style={{ display: 'flex', alignItems: 'center', gap: '2px' }}>
                        <input
                          type="radio"
                          name={`${row.id}-start-ampm`}
                          checked={row.startAmPm === 'PM'}
                          onChange={() => onPermitRowUpdate(row.id, 'startAmPm', 'PM')}
                        />
                        PM
                      </label>
                    </div> */}
                  </div>

                  {/* Expiry Time Column */}
                  <div style={{ 
                    padding: '10px',
                    display: 'flex',
                    gap: '5px',
                    alignItems: 'center'
                  }}>
                    <TextField
                      type="time"
                      value={row.endTime}
                      onChange={(e, newValue) => onPermitRowUpdate(row.id, 'endTime', newValue || '')}
                      style={{
                        flex: 1,
                        padding: '5px',
                        border: '1px solid #ccc',
                        borderRadius: '4px'
                      }}
                    />
                    {/* <div style={{ display: 'flex', gap: '5px' }}>
                      <label style={{ display: 'flex', alignItems: 'center', gap: '2px' }}>
                        <input
                          type="radio"
                          name={`${row.id}-end-ampm`}
                          checked={row.endAmPm === 'AM'}
                          onChange={() => onPermitRowUpdate(row.id, 'endAmPm', 'AM')}
                        />
                        AM
                      </label>
                      <label style={{ display: 'flex', alignItems: 'center', gap: '2px' }}>
                        <input
                          type="radio"
                          name={`${row.id}-end-ampm`}
                          checked={row.endAmPm === 'PM'}
                          onChange={() => onPermitRowUpdate(row.id, 'endAmPm', 'PM')}
                        />
                        PM
                      </label>
                    </div> */}
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default PermitSchedule;