import * as React from 'react';
import { Stack, CommandBar, ICommandBarItemProps } from '@fluentui/react';
import PpeForm from './PpeForm';
import SubmittedPpeFormsList from './SubmittedPpeFormsList';
import type { IPpeFormWebPartProps } from './IPpeFormProps';

type Mode = 'list' | 'add' | 'edit';

const PPEFormListGuid = '7afa2286-c552-4ff6-952e-1c09f32734cd';

const PpeFormHost: React.FC<IPpeFormWebPartProps> = (props) => {
    const [mode, setMode] = React.useState<Mode>('list');
    // const [formId, setFormId] = React.useState<number | undefined>(undefined);

    React.useEffect(() => {
        const url = new URL(window.location.href);
        const m = (url.searchParams.get('mode') || '').toLowerCase();
        // const id = url.searchParams.get('formId') || undefined;
        if (m === 'add') setMode('add');
        else if (m === 'edit') { setMode('edit'); /* setFormId(id ? Number(id) : undefined); */ }
        else setMode('list');
    }, []);

    const topBarItems: ICommandBarItemProps[] = React.useMemo(() => {
        if (mode === 'list') return [];
        return [{ key: 'back', text: 'Back to list', iconProps: { iconName: 'Back' }, onClick: () => setMode('list') }
        ];
    }, [mode]);

    return (
        <Stack tokens={{ childrenGap: 8 }}>
            {mode !== 'list' && <CommandBar items={topBarItems} />}
            {mode === 'list' && (
                <SubmittedPpeFormsList context={props.context} listGuid={PPEFormListGuid}
                    onAddNew={() => setMode('add')} onEdit={(_id) => { setMode('edit'); }} />
            )}
            {mode !== 'list' && (
                <PpeForm
                    context={props.context}
                    ThemeColor={props.ThemeColor}
                    IsDarkTheme={props.IsDarkTheme}
                    HasTeamsContext={props.HasTeamsContext}
                    onClose={() => setMode('list')}
                    onSubmitted={() => setMode('list')}
                />
            )}
        </Stack>
    );
};

export default PpeFormHost;
