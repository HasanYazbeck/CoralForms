import * as React from 'react';
import { Stack, CommandBar, ICommandBarItemProps } from '@fluentui/react';
// import PpeForm from './PpeForm';
// import SubmittedPpeFormsList from './SubmittedPpeFormsList';
import { SPCrudOperations } from '../../../Classes/SPCrudOperations';
// import PtwForm from './PtwForm';
import SubmittedPTWFormsList from './SubmittedPTWFormsList';
import { IPTWFormProps } from './IPTWFormProps';
import PTWForm from './PtwForm';

type Mode = 'list' | 'add' | 'edit';

const PTWFormHost: React.FC<IPTWFormProps> = (props) => {
    const [mode, setMode] = React.useState<Mode>('list');
    const [formId, setFormId] = React.useState<number | undefined>(undefined);
    const [PTWFormListGuid, setPTWFormListGuid] = React.useState<string>('');

    React.useEffect(() => {
        if (PTWFormListGuid) return;
        let cancelled = false;
        (async () => {
            try {
                const sp = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form', '');
                const guid = await sp._getSharePointListGUID();
                if (!cancelled) {
                    setPTWFormListGuid(guid || '');
                }
            } catch {
                if (!cancelled) setPTWFormListGuid('');
            }
        })();

        return () => { cancelled = true; };
    }, [PTWFormListGuid, props.context]);

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
                <SubmittedPTWFormsList context={props.context} listGuid={PTWFormListGuid}
                    onAddNew={() => { setFormId(undefined); setMode('add'); }}
                    onEdit={(id) => { setFormId(id); setMode('edit'); }} />
            )}
            {mode !== 'list'
                && (
                    <PTWForm
                        context={props.context}
                        ThemeColor={props.ThemeColor}
                        IsDarkTheme={props.IsDarkTheme}
                        formId={formId}
                        onClose={() => setMode('list')}
                        onSubmitted={() => setMode('list')}
                        useTargetAudience={props.useTargetAudience}
                        targetAudience={props.targetAudience}
                    />
                )
            }
        </Stack>
    );
};

export default PTWFormHost;
