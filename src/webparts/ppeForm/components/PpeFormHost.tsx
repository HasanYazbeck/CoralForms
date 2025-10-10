import * as React from 'react';
import { Stack, CommandBar, ICommandBarItemProps } from '@fluentui/react';
import PpeForm from './PpeForm';
import SubmittedPpeFormsList from './SubmittedPpeFormsList';
import type { IPpeFormWebPartProps } from './IPpeFormProps';
import { SPCrudOperations } from '../../../Classes/SPCrudOperations';

type Mode = 'list' | 'add' | 'edit';

const PpeFormHost: React.FC<IPpeFormWebPartProps> = (props) => {
    const [mode, setMode] = React.useState<Mode>('list');
    const [formId, setFormId] = React.useState<number | undefined>(undefined);
    const [PPEFormListGuid, setPPEFormListGuid] = React.useState<string>('');

    React.useEffect(() => {
        if (PPEFormListGuid) return;
        let cancelled = false;
        (async () => {
            try {
                const sp = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PPE_Form', '');
                const guid = await sp._getSharePointListGUID();
                if (!cancelled) {
                    setPPEFormListGuid(guid || '');
                }
            } catch {
                if (!cancelled) setPPEFormListGuid('');
            }
        })();

        return () => { cancelled = true; };
    }, [PPEFormListGuid, props.context]);

    // Initialize mode based on (priority) props.formId > URL formId > explicit mode param
    React.useEffect(() => {
        const href = (window.top?.location?.href) || window.location.href;
        const url = new URL(href);
        const urlId = url.searchParams.get('formId');
        const m = (url.searchParams.get('mode') || '').toLowerCase();

        const propId = props.formId && props.formId > 0 ? props.formId : undefined;
        const queryId = urlId && Number(urlId) > 0 ? Number(urlId) : undefined;

        if (propId) {
            setFormId(propId);
            setMode('edit');
            return;
        }

        if (queryId) {
            setFormId(queryId);
            setMode('edit');
            return;
        }

        if (m === 'add') setMode('add');
        else if (m === 'edit') { setMode('edit'); setFormId(queryId); }
        else setMode('list');
    }, [props.formId]);

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
                    onAddNew={() => { setFormId(undefined); setMode('add'); }}
                    onEdit={(id) => { setFormId(id); setMode('edit'); }} />
            )}
            {mode !== 'list' && (
                <PpeForm
                    context={props.context}
                    ThemeColor={props.ThemeColor}
                    IsDarkTheme={props.IsDarkTheme}
                    HasTeamsContext={props.HasTeamsContext}
                    formId={formId}
                    onClose={() => setMode('list')}
                    onSubmitted={() => setMode('list')}
                />
            )}
        </Stack>
    );
};

export default PpeFormHost;
