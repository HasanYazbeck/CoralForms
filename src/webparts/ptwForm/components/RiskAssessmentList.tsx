import * as React from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  TextField,
  ComboBox,
  IComboBoxOption,
  Dropdown,
  IDropdownOption,
  IconButton,
  Stack,
  Label,
  ChoiceGroup,
  IChoiceGroupOption,
  Checkbox
} from '@fluentui/react';
import { ILookupItem } from '../../../Interfaces/PtwForm/IPTWForm';

export interface IRiskTaskRow {
  id: string;
  task: string;
  initialRisk?: string;       // from _ptwFormStructure.initialRisk[]
  safeguardIds: number[];     // multi-select ILookupItem[]
  residualRisk?: string;      // from _ptwFormStructure.residualRisk[]
}

export interface IRiskAssessmentListProps {
  initialRiskOptions: string[];
  residualRiskOptions: string[];
  safeguards: ILookupItem[];                // ILookupItem[] source (e.g., precuationsItems)
  defaultRows?: IRiskTaskRow[];
  overallRiskOptions?: string[];            // from _ptwFormStructure.overallRiskAssessment
  onChange?: (state: {
    rows: IRiskTaskRow[];
    overallRisk?: string;
    l2Required: boolean;
    l2Ref?: string;
  }) => void;
}

const toComboOptions = (values: string[]): IComboBoxOption[] =>
  (values || []).map(v => ({ key: v, text: v }));

const toDropdownOptions = (items: ILookupItem[]): IDropdownOption[] =>
  (items || []).map(i => ({ key: i.id, text: i.title }));

const newRow = (): IRiskTaskRow => ({
  id: `riskrow-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
  task: '',
  initialRisk: undefined,
  safeguardIds: [],
  residualRisk: undefined
});

const RiskAssessmentList: React.FC<IRiskAssessmentListProps> = ({
  initialRiskOptions,
  residualRiskOptions,
  safeguards,
  defaultRows,
  overallRiskOptions,
  onChange
}) => {
  const [rows, setRows] = React.useState<IRiskTaskRow[]>(defaultRows?.length ? defaultRows : [newRow()]);
  const [overallRisk, setOverallRisk] = React.useState<string | undefined>(undefined);
  const [l2Required, setL2Required] = React.useState<boolean>(false);
  const [l2Ref, setL2Ref] = React.useState<string>('');

  // Notify parent
  React.useEffect(() => {
    onChange?.({ rows, overallRisk, l2Required, l2Ref: l2Required ? l2Ref : undefined });
  }, [rows, overallRisk, l2Required, l2Ref, onChange]);

  const handleTaskChange = (id: string, value: string | undefined) => {
    setRows(prev => prev.map(r => (r.id === id ? { ...r, task: value || '' } : r)));
  };

  const handleInitialRiskChange = (id: string, option?: IComboBoxOption) => {
    setRows(prev => prev.map(r => (r.id === id ? { ...r, initialRisk: option?.key as string | undefined } : r)));
  };

  const handleSafeguardsChange = (id: string, options?: IDropdownOption[]) => {
    const selectedIds = (options || []).filter(o => !!o.selected).map(o => Number(o.key));
    setRows(prev => prev.map(r => (r.id === id ? { ...r, safeguardIds: selectedIds } : r)));
  };

  const handleResidualRiskChange = (id: string, option?: IComboBoxOption) => {
    setRows(prev => prev.map(r => (r.id === id ? { ...r, residualRisk: option?.key as string | undefined } : r)));
  };

  const addRow = () => setRows(prev => [...prev, newRow()]);
  const deleteRow = (id: string) => setRows(prev => prev.filter(r => r.id !== id));

  const columns: IColumn[] = React.useMemo(() => [
    {
      key: 'col-task',
      name: 'Task',
      minWidth: 220,
      onRender: (row: IRiskTaskRow) => (
        <TextField
          value={row.task}
          onChange={(_, v) => handleTaskChange(row.id, v)}
          placeholder="Enter task"
        />
      )
    },
    {
      key: 'col-ir',
      name: 'Initial Risk (IR)',
      minWidth: 180,
      onRender: (row: IRiskTaskRow) => (
        <ComboBox
          placeholder="Select"
          options={toComboOptions(initialRiskOptions || [])}
          selectedKey={row.initialRisk}
          onChange={(_, option) => handleInitialRiskChange(row.id, option)}
          useComboBoxAsMenuWidth
        />
      )
    },
    {
      key: 'col-safe',
      name: 'Safeguards',
      minWidth: 260,
      onRender: (row: IRiskTaskRow) => (
        <Dropdown
          placeholder="Select"
          options={toDropdownOptions(safeguards || [])}
          selectedKeys={row.safeguardIds}
          onChange={() => handleSafeguardsChange(row.id, [])}
        //   onChange={(_, __, ___, selectedOptions) => handleSafeguardsChange(row.id, selectedOptions)}
          multiSelect
        />
      )
    },
    {
      key: 'col-rr',
      name: 'Residual Risk (RR)',
      minWidth: 180,
      onRender: (row: IRiskTaskRow) => (
        <ComboBox
          placeholder="Select"
          options={toComboOptions(residualRiskOptions || [])}
          selectedKey={row.residualRisk}
          onChange={(_, option) => handleResidualRiskChange(row.id, option)}
          useComboBoxAsMenuWidth
        />
      )
    },
    {
      key: 'col-actions',
      name: 'Actions',
      minWidth: 80,
      maxWidth: 100,
      onRender: (row: IRiskTaskRow) => (
        <IconButton
          iconProps={{ iconName: 'Delete' }}
          ariaLabel="Delete row"
          title="Delete row"
          onClick={() => deleteRow(row.id)}
        />
      )
    }
  ], [initialRiskOptions, residualRiskOptions, safeguards]);

  const overallOptions: IChoiceGroupOption[] = (overallRiskOptions || []).map(o => ({
    key: o,
    text: o
  }));

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Label style={{ margin: 0 }}>Job Description / Tasks</Label>
        <IconButton
          iconProps={{ iconName: 'Add' }}
          text="Add row"
          onClick={addRow}
          styles={{ label: { fontWeight: 600 } as any }}
        />
      </Stack>

      <DetailsList
        items={rows}
        columns={columns}
        compact
        selectionMode={SelectionMode.none}
        layoutMode={DetailsListLayoutMode.fixedColumns}
        getKey={(item: IRiskTaskRow) => item.id}
      />

      {/* Overall Risk Assessment */}
      {overallOptions.length > 0 && (
        <Stack>
          <Label>Overall Risk Assessment</Label>
          <ChoiceGroup
            selectedKey={overallRisk}
            options={overallOptions}
            onChange={(_, option) => setOverallRisk(option?.key)}
          />
          <Label styles={{ root: { fontStyle: 'italic', fontSize: 12, color: '#6b6b6b' } as any }}>
            If the Overall Risk Assessment is ranked as High (as per COR-HSE-01-MTX-001), HSE & terminal management approval is required.
          </Label>
        </Stack>
      )}

      {/* Detailed (L2) Risk Assessment */}
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
        <Checkbox
          label="Detailed (L2) Risk Assessment required"
          checked={l2Required}
          onChange={(_, chk) => setL2Required(!!chk)}
        />
        {l2Required && (
          <TextField
            label="Risk Assessment Ref. Nbr."
            value={l2Ref}
            onChange={(_, v) => setL2Ref(v || '')}
            styles={{ root: { maxWidth: 360 } }}
          />
        )}
      </Stack>
    </Stack>
  );
};

export default RiskAssessmentList;