import * as React from 'react';
import {
    DetailsList,
    DetailsListLayoutMode,
    IColumn,
    SelectionMode,
    TextField,
    ComboBox,
    IComboBoxOption,
    IconButton,
    Stack,
    Label,
    ChoiceGroup,
    IChoiceGroupOption,
    Checkbox,
    DefaultButton
} from '@fluentui/react';
import { ILookupItem } from '../../../Interfaces/PtwForm/IPTWForm';

export interface IRiskTaskRow {
    id: string;
    task: string;
    initialRisk?: string;       // from _ptwFormStructure.initialRisk[]
    safeguardIds: number[];     // multi-select ILookupItem[]
    residualRisk?: string;      // from _ptwFormStructure.residualRisk[]
    safeguardsNote?: string;
    disabledFields: boolean;    // custom text entered in the safeguards combobox
    orderRecord: number;    // for sorting
}

export interface IRiskAssessmentListProps {
    initialRiskOptions: string[];
    residualRiskOptions: string[];
    safeguards: ILookupItem[];
    defaultRows?: IRiskTaskRow[];
    overallRiskOptions?: string[];
    disableRiskControls?: boolean;
    onChange?: (state: {
        rows: IRiskTaskRow[];
        overallRisk?: string;
        l2Required: boolean;
        l2Ref?: string;
    }) => void;
}

const toComboOptions = (values: string[]): IComboBoxOption[] =>
    (values || []).map(v => ({ key: v, text: v }));

// (Dropdown implementation removed)

const newRow = (): IRiskTaskRow => ({
    id: `riskrow-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
    task: '',
    initialRisk: undefined,
    safeguardIds: [],
    residualRisk: undefined,
    disabledFields: true,
    orderRecord: 0
});

const RiskAssessmentList: React.FC<IRiskAssessmentListProps> = ({
    initialRiskOptions,
    residualRiskOptions,
    safeguards,
    defaultRows,
    overallRiskOptions,
    onChange,
    disableRiskControls = false
}) => {
    const [rows, setRows] = React.useState<IRiskTaskRow[]>(defaultRows?.length ? defaultRows : [newRow()]);
    const [overallRisk, setOverallRisk] = React.useState<string | undefined>(undefined);
    const [l2Required, setL2Required] = React.useState<boolean>(false);
    const [l2Ref, setL2Ref] = React.useState<string>('');
    const [safeFilterByRow, setSafeFilterByRow] = React.useState<Record<string, string>>({});
    const allSafeguardsById = React.useRef<Map<number, ILookupItem>>(new Map());

    React.useEffect(() => {
        (safeguards || []).forEach(s => {
            if (s?.id !== undefined && !allSafeguardsById.current.has(Number(s.id))) {
                allSafeguardsById.current.set(Number(s.id), s);
            }
        });
    }, [safeguards]);

    // Notify parent
    React.useEffect(() => {
        onChange?.({ rows, overallRisk, l2Required, l2Ref: l2Required ? l2Ref : undefined });
    }, [rows, overallRisk, l2Required, l2Ref, onChange]);

    const handleTaskChange = (id: string, value: string | undefined) => {
        setRows(prev => prev.map(r => (r.id === id ? { ...r, task: value || '', disabledFields: value === "" } : r)));
    };

    const handleInitialRiskChange = (id: string, option?: IComboBoxOption) => {
        if (disableRiskControls) return;
        setRows(prev => prev.map(r => (r.id === id ? { ...r, initialRisk: option?.key as string | undefined } : r)));
    };

    // Build multi-select ComboBox options with selected state based on row.safeguardIds
    const buildSafeguardComboOptions = React.useCallback((row: IRiskTaskRow): IComboBoxOption[] => {
        const filterText = (safeFilterByRow[row.id] || '').trim().toLowerCase();
        const list = (safeguards || []).filter(i => !filterText || (i.title || '').toLowerCase().includes(filterText));
        const base = list.map(i => ({ key: i.id, text: i.title } as IComboBoxOption));
        return base.map(opt => ({ ...opt, selected: row.safeguardIds?.includes(Number(opt.key)) }));
    }, [safeguards, safeFilterByRow]);

    // Toggle selection for a single option in multi-select ComboBox
    const handleSafeguardComboChange = React.useCallback((row: IRiskTaskRow, option?: IComboBoxOption, _index?: number, _value?: string) => {
        // When typing in the ComboBox, option is undefined and _value has the current input text
        if (!option) {
            if (typeof _value === 'string') {
                setSafeFilterByRow(prev => ({ ...prev, [row.id]: _value }));
            }
            return;
        }
        const idNum = Number(option.key);
        setRows(prev => prev.map(r => {
            if (r.id !== row.id) return r;
            const current = new Set(r.safeguardIds || []);
            if (option.selected) {
                current.add(idNum);
            } else {
                current.delete(idNum);
            }
            return { ...r, safeguardIds: Array.from(current) };
        }));
    }, [setRows]);

    // Remove safeguard from chips
    const removeSafeguard = React.useCallback((rowId: string, id: number) => {
        setRows(prev => prev.map(r => (
            r.id === rowId
                ? { ...r, safeguardIds: (r.safeguardIds || []).filter(x => Number(x) !== Number(id)) }
                : r
        )));
    }, [setRows]);

    const handleResidualRiskChange = (id: string, option?: IComboBoxOption) => {
        if (disableRiskControls) return;
        setRows(prev => prev.map(r => (r.id === id ? { ...r, residualRisk: option?.key as string | undefined } : r)));
    };

    const getRiskColors = React.useCallback((key: string) => {
        const k = (key || '').toLowerCase();
        if (k.includes('low')) return { bg: '#22B14C', fg: '#ffffff' };      // green
        if (k.includes('medium') || k.includes('med')) return { bg: '#FFF200', fg: '#323130' }; // yellow
        if (k.includes('high')) return { bg: '#ED1C24', fg: '#ffffff' };     // red
        return { bg: '#rgb(241 241 241)', fg: '#323130' }; // fallback
    }, []);

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
            minWidth: 120,
            onRender: (row: IRiskTaskRow) => (
                <ComboBox
                    placeholder="Select"
                    options={toComboOptions(initialRiskOptions || [])}
                    selectedKey={row.initialRisk}
                    onChange={(_, option) => handleInitialRiskChange(row.id, option)}
                    useComboBoxAsMenuWidth
                    disabled={disableRiskControls || row.disabledFields}
                />
            )
        },
        {
            key: 'col-safe',
            name: 'Safeguards',
            minWidth: 400,
            resizable: true,
            onRender: (row: IRiskTaskRow) => {
                const selectedItems = (row.safeguardIds || [])
                    .sort((a, b) => Number(a) - Number(b))
                    .map(id => allSafeguardsById.current.get(Number(id)))
                    .filter(Boolean) as ILookupItem[];
                return (
                    <div>
                        <ComboBox
                            key={`saf-${row.id}-${(row.safeguardIds || []).slice().sort((a, b) => Number(a) - Number(b)).join('_')}`}
                            placeholder="Select safeguards"
                            multiSelect
                            options={buildSafeguardComboOptions(row)}
                            onChange={(_, option, index, value) => handleSafeguardComboChange(row, option, index, value)}
                            // Control input text to implement search filtering
                            text={safeFilterByRow[row.id] || ''}
                            allowFreeform
                            useComboBoxAsMenuWidth
                            disabled={row.disabledFields}
                        />

                        <div style={{ border: '1px solid #e1e1e1', borderRadius: 4, padding: 6, marginTop: 6, width: '100%' }}>
                            {selectedItems.length === 0 ? (
                                <span style={{ color: '#605e5c', fontStyle: 'italic' }}>No safeguards selected</span>
                            ) : (
                                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                                    {selectedItems.map(s => (
                                        <span key={s.id}
                                            style={{
                                                background: '#f3f2f1', border: '1px solid #c8c6c4', lineHeight: 1.4,
                                                whiteSpace: 'break-spaces', borderRadius: 2, padding: '2px 6px',
                                                display: 'inline-flex', alignItems: 'center', gap: 6
                                            }}>
                                            <IconButton
                                                iconProps={{ iconName: 'Cancel' }}
                                                ariaLabel={`Remove ${s.title}`}
                                                title={`Remove ${s.title}`}
                                                onClick={() => removeSafeguard(row.id, Number(s.id))}
                                                styles={{ root: { height: 20, width: 20, minWidth: 20 }, icon: { fontSize: 12 } }}
                                            />
                                            <span style={{ color: '#323130' }}>{s.title}</span>
                                        </span>
                                    ))}
                                </div>
                            )}
                        </div>
                    </div>
                );
            }
        },
        {
            key: 'col-rr',
            name: 'Residual Risk (RR)',
            minWidth: 120,
            onRender: (row: IRiskTaskRow) => (
                <ComboBox
                    placeholder="Select"
                    options={toComboOptions(residualRiskOptions || [])}
                    selectedKey={row.residualRisk}
                    onChange={(_, option) => handleResidualRiskChange(row.id, option)}
                    useComboBoxAsMenuWidth
                    disabled={disableRiskControls || row.disabledFields}
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
    ], [initialRiskOptions, residualRiskOptions, safeguards, safeFilterByRow, disableRiskControls]);

    const overallOptions: IChoiceGroupOption[] = (overallRiskOptions || []).map(o => {
        const { bg, fg } = getRiskColors(o);
        return {
            key: o,
            text: o,
            onRenderField: (props, defaultRender) => {
                // Wrap the default radio+label inside a colored tile
                return (
                    <div
                        style={{
                            backgroundColor: bg,
                            color: fg,
                            padding: '8px 10px',
                            borderRadius: 4,
                            border: '1px solid transparent',
                            minWidth: 100,
                            display: 'inline-flex',
                            alignItems: 'center',
                            justifyContent: 'center'
                        }}
                    >
                        {defaultRender ? defaultRender(props) : null}
                    </div>
                );
            }
        };
    });

    return (
        <Stack tokens={{ childrenGap: 12 }}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                <Label style={{ margin: 0 }}>Job Description / Tasks</Label>
                <DefaultButton
                    iconProps={{ iconName: 'Add' }}
                    text="Add Task"
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
                <div className='row'>
                    <div className='form-group' style={{
                        display: 'flex',
                        flexWrap: 'wrap', alignItems: 'center',
                        justifyContent: 'flex-end',
                        border: '1px solid #BFBFBF',
                        backgroundColor: '#f1f1f1'
                    }}>
                        <Label style={{ marginRight: "10px", paddingTop: "10px" }}>Overall Risk Assessment</Label>
                        <ChoiceGroup
                            disabled={disableRiskControls || rows.some(r => r.disabledFields)}
                            selectedKey={overallRisk}
                            options={overallOptions}
                            onChange={(_, option) => setOverallRisk(option?.key)}
                            style={{ color: '#232020', fontWeight: "700" }}
                            styles={{
                                flexContainer: {
                                    display: 'flex',
                                    flexDirection: 'row',
                                    flexWrap: 'nowrap', // keep on one line
                                    columnGap: "12px"
                                },
                                root: {
                                    selectors: {
                                        '.ms-ChoiceFieldGroup-flexContainer': {
                                            display: 'flex !important',
                                            flexDirection: 'row !important',
                                            flexWrap: 'nowrap !important',
                                            columnGap: "12px"
                                        }
                                    }
                                }
                            }}
                        />
                        <Label styles={{ root: { fontStyle: 'italic', fontSize: 12, color: '#6b6b6b' } as any }}>
                            If the Overall Risk Assessment is ranked as High (as per COR-HSE-03-MTX-001), HSE & terminal management approval is required.
                        </Label>
                    </div>
                </div>
            )
            }

            {/* Detailed (L2) Risk Assessment */}
            <div className="row pt-2">
                <div className="form-group" style={{ display: "flex", alignItems: "center" }}>
                    <div className='col-md-12' style={{
                        display: "flex",
                        justifyContent: "space-between",
                        alignItems: "center"
                    }}>
                        <div className='col-md-5'>
                            <Checkbox
                                label="Detailed (L2) Risk Assessment required"
                                checked={l2Required}
                                onChange={(_, chk) => setL2Required(!!chk)}
                                disabled={disableRiskControls}
                            />
                        </div>

                        <div className='col-md-7' style={{ display: 'flex', flexWrap: "nowrap", gap: "10px" }}>
                            <Label style={{ fontStyle: 'italic' }}>Risk Assessment Ref. Nbr.</Label>
                            <TextField
                                value={l2Ref}
                                disabled={disableRiskControls || !l2Required}
                                onChange={(_, v) => setL2Ref(v || '')}
                                styles={{ root: { width: '68%' } }}
                            />
                        </div>

                    </div>
                </div>
            </div>
        </Stack >
    );
};

export default RiskAssessmentList;