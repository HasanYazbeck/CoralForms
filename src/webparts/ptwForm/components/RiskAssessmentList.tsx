import * as React from 'react';
import { ILookupItem } from '../../../Interfaces/PtwForm/IPTWForm';
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
    DefaultButton,
    IComboBoxStyles,
    TooltipHost, DirectionalHint
} from '@fluentui/react';

export interface IRiskTaskRow {
    id: string;
    task: string;
    initialRisk?: string;
    safeguardIds: number[];
    residualRisk?: string;
    disabledFields: boolean;
    orderRecord: number;
    customSafeguards: string[];
}

export interface IRiskAssessmentListProps {
    initialRiskOptions: string[];
    residualRiskOptions: string[];
    safeguards: ILookupItem[];
    defaultRows?: IRiskTaskRow[];
    overallRiskOptions?: string[];
    disableRiskControls?: boolean;
    selectedOverallRisk?: string;
    l2Required: boolean;
    l2Ref: string;
    onOverallRiskChange: (overallRisk: string | undefined) => void;
    onDetailedRiskChange: (required: boolean) => void;
    onDetailedRiskRefChange: (ref: string) => void;
    onChange?: (state: {
        rows: IRiskTaskRow[];
    }) => void;
}

const toComboOptions = (values: string[]): IComboBoxOption[] =>
    (values || []).map(v => ({ key: v, text: v }));

const comboBoxBlackStyles: Partial<IComboBoxStyles> = {
    root: {
        selectors: {
            '.ms-ComboBox-Input': { color: '#000', fontWeight: 500, },
            '&.is-disabled .ms-ComboBox-Input': { color: '#000', fontWeight: 500, },
            '.ms-ComboBox-Input::placeholder': { color: '#000', fontWeight: 500, },
        }
    },
    input: { color: '#000' } // supported in v8; safe no-op if ignored
};

const newRow = (): IRiskTaskRow => ({
    id: `riskrow-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
    task: '',
    initialRisk: undefined,
    safeguardIds: [],
    residualRisk: undefined,
    disabledFields: true,
    orderRecord: 0,
    customSafeguards: []
});

const RiskAssessmentList: React.FC<IRiskAssessmentListProps> = ({
    initialRiskOptions,
    residualRiskOptions,
    safeguards,
    defaultRows,
    overallRiskOptions,
    selectedOverallRisk,
    l2Required,
    l2Ref,
    onOverallRiskChange,
    onDetailedRiskChange,
    onDetailedRiskRefChange,
    onChange,
    disableRiskControls = false
}) => {
    const [rows, setRows] = React.useState<IRiskTaskRow[]>(defaultRows?.length ? defaultRows : [newRow()]);
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
    React.useEffect(() => { onChange?.({ rows }); }, [rows, onChange]);

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

    // Helper: commit one freeform token into row (match existing or add custom)
    const commitFreeformToken = React.useCallback((rowId: string, rawText: string) => {
        const raw = (rawText || '').trim();
        if (!raw) return;

        const match = (safeguards || []).find(s => (s.title || '').toLowerCase() === raw.toLowerCase());
        if (match?.id != null) {
            const idNum = Number(match.id);
            setRows(prev => prev.map(r =>
                r.id === rowId
                    ? { ...r, safeguardIds: Array.from(new Set([...(r.safeguardIds || []), idNum])) }
                    : r
            ));
        } else {
            setRows(prev => prev.map(r =>
                r.id === rowId
                    ? { ...r, customSafeguards: Array.from(new Set([...(r.customSafeguards || []), raw])) }
                    : r
            ));
        }
    }, [safeguards, setRows]);

    // Commit any pending text from the filter box
    const commitPendingFreeform = React.useCallback((row: IRiskTaskRow) => {
        const raw = (safeFilterByRow[row.id] || '').trim();
        if (!raw) return;
        commitFreeformToken(row.id, raw);
        setSafeFilterByRow(prev => ({ ...prev, [row.id]: '' }));
    }, [safeFilterByRow, commitFreeformToken]);

    // Toggle selection for a single option in multi-select ComboBox + handle freeform typing
    const handleSafeguardComboChange = React.useCallback((row: IRiskTaskRow, option?: IComboBoxOption, _index?: number, _value?: string) => {
        // Freeform typing path
        if (!option) {
            if (typeof _value === 'string') {
                // Support comma as a token delimiter: commit tokens before the last comma
                if (_value.includes(',')) {
                    const parts = _value.split(',');
                    const tokens = parts.slice(0, -1).map(t => t.trim()).filter(Boolean);
                    tokens.forEach(tok => commitFreeformToken(row.id, tok));
                    const remainder = parts[parts.length - 1]; // keep last incomplete token in the input
                    setSafeFilterByRow(prev => ({ ...prev, [row.id]: remainder }));
                } else {
                    // Just update the live filter text
                    setSafeFilterByRow(prev => ({ ...prev, [row.id]: _value }));
                }
            }
            return;
        }

        // Option toggle path
        const idNum = Number(option.key);
        setRows(prev => prev.map(r => {
            if (r.id !== row.id) return r;
            const current = new Set(r.safeguardIds || []);
            if (option.selected) current.add(idNum);
            else current.delete(idNum);
            return { ...r, safeguardIds: Array.from(current) };
        }));
    }, [setRows, commitFreeformToken, setSafeFilterByRow]);

    // NEW: remove a custom safeguard chip
    const removeCustomSafeguard = React.useCallback((rowId: string, label: string) => {
        setRows(prev => prev.map(r => r.id === rowId ? { ...r, customSafeguards: (r.customSafeguards || []).filter(x => x !== label) } : r));
    }, []);

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

    // Render colored options (Low/Medium/High)
    const renderRiskOption = React.useCallback((option?: IComboBoxOption) => {
        const label = String(option?.text ?? option?.key ?? '');
        const { bg, fg } = getRiskColors(label);
        // The option root has horizontal padding (8px). Use negative margins to fill the row.
        return (
            <div style={{
                backgroundColor: bg,
                color: fg,
                padding: '6px 8px',
                margin: '0 -8px',
                display: 'block',
                fontWeight: 600
            }}>
                {label}
            </div>
        );
    }, [getRiskColors]);

    // Fill the whole ComboBox (input + chevron)
    const getRiskComboStyles = React.useCallback((selectedLabel?: string): Partial<IComboBoxStyles> => {
        if (!selectedLabel) return {};
        const { bg, fg } = getRiskColors(selectedLabel);
        return {
            root: {
                backgroundColor: bg,
                borderColor: bg,
                borderRadius: 2,
                selectors: {
                    // input area
                    '.ms-ComboBox': { backgroundColor: bg },
                    '.ms-ComboBox input': { backgroundColor: bg, color: fg, fontWeight: 600 },
                    // caret (chevron) button
                    '.ms-ComboBox-CaretDown-button': {
                        backgroundColor: bg,
                        borderLeft: '1px solid rgba(0,0,0,0.1)'
                    },
                    '.ms-ComboBox-CaretDown-button .ms-Button-menuIcon': { color: fg },
                    '&:hover': { backgroundColor: bg, borderColor: bg },
                    '&:focus-within': { backgroundColor: bg, borderColor: bg }
                }
            },
            input: { backgroundColor: bg, color: fg, fontWeight: 600 }
        };
    }, [getRiskColors]);

    const addRow = () => setRows(prev => [...prev, newRow()]);
    const deleteRow = (id: string) => setRows(prev => prev.filter(r => r.id !== id));

    const columns: IColumn[] = React.useMemo(() => [
        {
            key: 'col-task',
            name: 'Task',
            minWidth: 220,
            onRender: (row: IRiskTaskRow) => (<TextField value={row.task} onChange={(_, v) => handleTaskChange(row.id, v)} placeholder="Enter task" />)
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
                    onRenderOption={renderRiskOption}
                    styles={getRiskComboStyles(row.initialRisk) && comboBoxBlackStyles}
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

                // count non-empty custom safeguards
                const customList = (row.customSafeguards || []).filter(t => (t || '').trim().length > 0);
                const hasAny = selectedItems.length > 0 || customList.length > 0;

                return (
                    <div>
                        <ComboBox
                            key={`saf-${row.id}-${(row.safeguardIds || []).slice().sort((a, b) => Number(a) - Number(b)).join('_')}`}
                            placeholder="Select Safeguards"
                            multiSelect
                            options={buildSafeguardComboOptions(row)}
                            onChange={(_, option, index, value) => handleSafeguardComboChange(row, option, index, value)}
                            // Control input text to implement search filtering
                            text={safeFilterByRow[row.id] || ''}
                            allowFreeform
                            useComboBoxAsMenuWidth
                            disabled={row.disabledFields}
                            // Commit pending freeform text when the control loses focus or menu closes
                            onBlur={() => commitPendingFreeform(row)}
                            onMenuDismissed={() => commitPendingFreeform(row)}
                        />

                        <div style={{ border: '1px solid #e1e1e1', borderRadius: 4, padding: 6, marginTop: 6, width: '100%' }}>
                            {!hasAny ? (
                                // <span style={{ color: '#605e5c', fontStyle: 'italic' }}>
                                //     No safeguards selected (type to search or add your own )
                                // </span>
                                <TooltipHost
                                    content="To add your own safeguard, type it in the box, then press comma then Enter to add it."
                                    directionalHint={DirectionalHint.topLeftEdge}
                                >
                                    <span style={{ color: '#605e5c', fontStyle: 'italic', cursor: 'help' }}>
                                        No safeguards selected
                                    </span>
                                </TooltipHost>
                            ) : (
                                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                                    {selectedItems.map(s => (
                                        <span key={`lkp-${s.id}`}
                                            style={{
                                                background: '#f3f2f1', border: '1px solid #c8c6c4', lineHeight: 1.4,
                                                whiteSpace: 'break-spaces', borderRadius: 2, padding: '2px 6px',
                                                display: 'inline-flex', alignItems: 'center', gap: 6
                                            }}>
                                            <IconButton iconProps={{ iconName: 'Cancel' }} ariaLabel={`Remove ${s.title}`} title={`Remove ${s.title}`}
                                                onClick={() => removeSafeguard(row.id, Number(s.id))} styles={{ root: { height: 20, width: 20, minWidth: 20 }, icon: { fontSize: 12 } }} />
                                            <span style={{ color: '#323130' }}>{s.title}</span>
                                        </span>
                                    ))}

                                    {(customList || []).map(txt => (
                                        <span key={`custom-${txt}`} style={{ whiteSpace: 'break-spaces', background: '#e5f1ff', border: '1px solid #99c7ff', borderRadius: 2, padding: '2px 6px', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                                            <IconButton iconProps={{ iconName: 'Cancel' }} ariaLabel={`Remove ${txt}`} title={`Remove ${txt}`}
                                                onClick={() => removeCustomSafeguard(row.id, txt)} styles={{ root: { height: 20, width: 20, minWidth: 20 }, icon: { fontSize: 12 } }} />
                                            <span style={{ color: '#1a3b5d', fontStyle: 'italic' }}>{txt}</span>
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
                    onRenderOption={renderRiskOption}
                    styles={getRiskComboStyles(row.residualRisk) && comboBoxBlackStyles}
                    
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
    ], [initialRiskOptions, residualRiskOptions, safeguards, safeFilterByRow, disableRiskControls, renderRiskOption, getRiskComboStyles]);

    const overallOptions: IChoiceGroupOption[] = (overallRiskOptions || []).map(o => {
        const normalized = o.trim();
        const { bg, fg } = getRiskColors(o);
        return {
            key: normalized,
            text: normalized,
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
                            selectedKey={selectedOverallRisk}
                            options={overallOptions}
                            onChange={(_, option) => onOverallRiskChange(option?.key)}
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
                                onChange={(_, chk) => onDetailedRiskChange(!!chk)}
                                disabled={disableRiskControls}
                            />
                        </div>

                        <div className='col-md-7' style={{ display: 'flex', flexWrap: "nowrap", gap: "10px" }}>
                            <Label style={{ fontStyle: 'italic' }}>Risk Assessment Ref. Nbr.</Label>
                            <TextField
                                value={l2Ref}
                                disabled={disableRiskControls || !l2Required}
                                onChange={(_, v) => onDetailedRiskRefChange(v || '')}
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