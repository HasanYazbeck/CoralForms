import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import LocalStorageComponent from '../../../Classes/LocalStorageComponent';
import {
  DetailsList, DetailsListLayoutMode, IColumn, Selection, SelectionMode, MarqueeSelection, CommandBar, ICommandBarItemProps, Stack, Text, Spinner, DefaultButton,
  IPersonaProps,
  PersonaSize,
  Persona
} from '@fluentui/react';


export type Row = {
  id: number;
  coralReferenceNumber?: string;
  company?: string;
  created?: Date;
  permitOriginator?: IPersonaProps;
  assetId?: string;
  projectTitle?: string;
  assetCategory?: string;
  assetDetails?: string;
  FormStatusRecord: string;
};

interface SubmittedPTWFormsListProps {
  context: WebPartContext;
  /** PTWForm list GUID. If not supplied, the component will try to operate but delete will be disabled. */
  listGuid?: string;
  title?: string;
  /** Called when user clicks Add New. If not provided, navigates with ?mode=add */
  onAddNew?: () => void;
  /** Called when user clicks Edit and exactly one row is selected. If not provided, navigates with ?formId=ID&mode=edit */
  onEdit?: (formId: number, row: Row) => void;
  /** Called after successful delete with the deleted IDs */
  onDelete?: (deletedIds: number[]) => void;
  onView?: (formId: number, row: Row) => void;
}

const SubmittedPTWFormsList: React.FC<SubmittedPTWFormsListProps> = ({
  context,
  listGuid,
  title = 'PERMIT TO WORK (PTW) SUBMISSIONS',
  onAddNew,
  onEdit,
  onDelete,
  onView
}) => {
  const [loading, setLoading] = React.useState(true);
  const [loadingMore, setLoadingMore] = React.useState(false);
  const [items, setItems] = React.useState<Row[]>([]);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [view, setView] = React.useState<'submitted' | 'closed' | 'rejected' | 'saved'>('saved');
  const [selectionVersion, setSelectionVersion] = React.useState(0);
  const [nextLink, setNextLink] = React.useState<string | undefined>(undefined);
  const [hasMore, setHasMore] = React.useState<boolean>(false);
  const [isPOEligible, setIsPOEligible] = React.useState<boolean>(false);
  const currentUserEmail = (context?.pageContext?.user?.email || '').toLowerCase();
  const webUrl = context.pageContext.web.absoluteUrl;
  const PAGE_SIZE = 50;
  const [formStatusRecord, setFormStatusRecord] = LocalStorageComponent('FormStatusRecord', { value: '' });

  const selectionRef = React.useRef(
    new Selection({
      selectionMode: SelectionMode.multiple,
      onSelectionChanged: () => setSelectionVersion(v => v + 1)
    })
  );

  const selectedRows = React.useMemo(() => (selectionRef.current.getSelection() as Row[]) || [], [selectionVersion]);

  // NEW: resolve eligibility from SP group membership
  React.useEffect(() => {
    let disposed = false;
    const esc = (s: string) => s.replace(/'/g, "''");
    const groupName = 'PermitOriginatorGroup';

    async function checkEligibility() {
      try {
        const url = `${webUrl}/_api/web/sitegroups/getbyname('${esc(groupName)}')/users?$select=EMail,LoginName`;
        const res = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (!res.ok) { if (!disposed) setIsPOEligible(false); return; }
        const json: any = await res.json();
        const users: any[] = json?.value || json?.d?.results || [];
        const found = users.some(u => String(u.EMail || u.Email || u.LoginName || '').toLowerCase() === currentUserEmail);
        if (!disposed) setIsPOEligible(found);
      } catch {
        if (!disposed) setIsPOEligible(false);
      }
    }

    if (currentUserEmail) checkEligibility();
    return () => { disposed = true; };
  }, [context.spHttpClient, webUrl, currentUserEmail]);

  // Build the shared select/expand pieces (without $top so we can append it)
  const baseSelect = React.useMemo(() => (
    `?$select=Id,CoralReferenceNumber,AssetID,ProjectTitle,Created,FormStatusRecord,` +
    `PermitOriginator/Title,PermitOriginator/EMail,` +
    `AssetCategory/Id,AssetCategory/Title,` +
    `AssetDetails/Id,AssetDetails/Title,` +
    `CompanyRecord/Id,CompanyRecord/Title` +
    `&$expand=PermitOriginator,AssetCategory,AssetDetails,CompanyRecord`
  ), []);

  const mapRows = React.useCallback((data: any[]): Row[] => {
    return (data || []).map((obj: any): Row => {
      return {
        id: Number(obj.Id),
        coralReferenceNumber: obj.CoralReferenceNumber ?? undefined,
        created: obj.Created ? new Date(obj.Created) : undefined,
        permitOriginator: obj.PermitOriginator ? { text: obj.PermitOriginator.Title, secondaryText: obj.PermitOriginator.EMail } : undefined,
        assetId: obj.AssetID ?? undefined,
        projectTitle: obj.ProjectTitle ?? undefined,
        assetCategory: obj.AssetCategory ? obj.AssetCategory.Title : undefined,
        assetDetails: obj.AssetDetails ? obj.AssetDetails.Title : undefined,
        company: obj.CompanyRecord ? obj.CompanyRecord.Title : undefined,
        FormStatusRecord: obj.FormStatusRecord ?? undefined
      };
    });
  }, []);

  const columns = React.useMemo<IColumn[]>(
    () => [
      { key: 'colCoralReferenceNumber', name: 'Ref #', fieldName: 'coralReferenceNumber', minWidth: 160, maxWidth: 180, isResizable: true },
      { key: 'colAssetId', name: 'Asset Id', fieldName: 'assetId', minWidth: 110, isResizable: true },

      { key: 'colProjectTitle', name: 'Project Title', fieldName: 'projectTitle', minWidth: 160, isResizable: true },
      { key: 'colAssetCategory', name: 'Asset Category', fieldName: 'assetCategory', minWidth: 200, isResizable: true },
      { key: 'colAssetDetails', name: 'Asset Details', fieldName: 'assetDetails', minWidth: 200, isResizable: true },
      { key: 'colCompany', name: 'Company', fieldName: 'company', minWidth: 200, isResizable: true },
      {
        key: 'colPermitOriginator',
        name: 'Permit Originator',
        fieldName: 'permitOriginator',
        minWidth: 150,
        isResizable: true,
        onRender: (row: Row) => {
          if (!row.permitOriginator) return '';
          return (
            <Persona
              text={row.permitOriginator.text}
              secondaryText={row.permitOriginator.secondaryText}
              size={PersonaSize.size24}
              hidePersonaDetails={false}
            />
          );
        }
      },
      {
        key: 'colCreated', name: 'Date Submitted', fieldName: 'created', minWidth: 140,
        onRender: (row: Row) => (row.created ? row.created.toLocaleDateString() : '')
      }
    ],
    []
  );

  const loadItems = React.useCallback(async (scope: 'submitted' | 'closed' | 'rejected' | 'saved' = view, reset: boolean = false) => {

    if (!listGuid) {
      setItems([]);
      setHasMore(false);
      setNextLink(undefined);
      return;
    }

    const filterActive = `&$filter=FormStatusRecord eq 'Submitted'`;
    const filterClosed = `&$filter=WorkflowStatus eq 'Closed By System'`;
    const filterRejected = `&$filter=WorkflowStatus eq 'Rejected'`;
    const filterSaved = `&$filter=FormStatusRecord eq 'Saved' and PermitOriginator/EMail eq '${context.pageContext.user.email}'`;
    const orderBy = `&$orderby=Created desc`;

    const headers = { Accept: 'application/json;odata=nometadata', 'odata-version': '' } as any;
    const webUrl = context.pageContext.web.absoluteUrl;
    // Decide which URL to call
    let url: string;
    if (!reset && nextLink) {
      // Continue from SharePoint’s paging link (absolute URL)
      url = nextLink;
    } else {
      const filter = scope === 'closed' ? filterClosed : scope === 'rejected' ? filterRejected : scope === 'saved' ? filterSaved : filterActive;
      url = `${webUrl}/_api/web/lists(guid'${listGuid}')/items${baseSelect}${filter}${orderBy}&$top=${PAGE_SIZE}`;
    }

    // Toggle the right spinner
    reset ? setLoading(true) : setLoadingMore(true);
    setError(undefined);

    try {
      const res = await context.spHttpClient.get(url, SPHttpClient.configurations.v1, { headers });
      if (!res.ok) {
        const t = await res.text();
        throw new Error(t || res.statusText);
      }
      const json: any = await res.json();

      // Read rows and the continuation link (SharePoint uses different keys depending on odata mode)
      const rows = mapRows(json.value || json.d?.results || []);
      const next = json['odata.nextLink'] || json['@odata.nextLink'] || json.d?.__next || undefined;

      setItems(prev => (reset ? rows : prev.concat(rows)));
      setNextLink(next);
      setHasMore(!!next);

      // Clear selection when view resets
      if (reset) {
        selectionRef.current.setAllSelected(false);
        setSelectionVersion(v => v + 1);
      }
    } catch (e: any) {
      setError(`Failed: ${e?.message || e} - Please reload page again.`);
      if (reset) {
        setItems([]);
        setHasMore(false);
        setNextLink(undefined);
      }
    } finally {
      reset ? setLoading(false) : setLoadingMore(false);
    }
  }, [context, listGuid, view, baseSelect, nextLink, mapRows]);

  const deleteSelected = React.useCallback(async () => {
    const rows = selectedRows;
    if (!rows.length || !listGuid) return;
    const ids = rows.map(r => r.id);
    const confirmMsg = `Delete ${ids.length} item(s)? This cannot be undone.`;
    if (!window.confirm(confirmMsg)) return;

    try {
      const base = `${context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${listGuid}')/items`;
      await Promise.all(
        ids.map(async id => {
          const url = `${base}(${id})`;

          const res = await context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
            headers: {
              'IF-MATCH': '*',
              'X-HTTP-Method': 'DELETE',
              // Headers to satisfy SharePoint DELETE over POST
              'Accept': 'application/json;odata=nometadata',
              'Content-Type': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          });
          // 204 No Content is normal; allow 404 as "already deleted"
          if (!res.ok && res.status !== 404) {
            const t = await res.text();
            throw new Error(`Delete failed for ID ${id}: ${t || res.statusText}`);
          }
        })
      );
      onDelete?.(ids);
      // await loadItems();
      setNextLink(undefined);
      setHasMore(false);
      loadItems(view, true);
    } catch (e: any) {
      setError(`Delete error: ${e?.message || e}`);
    }
  }, [selectedRows, listGuid, context, loadItems, onDelete]);

  const switchState = React.useCallback((next: 'submitted' | 'rejected' | 'closed' | 'saved') => {
    setView(next);
    setNextLink(undefined);
    setHasMore(false);
    // Clear selection and load first page for the new scope
    selectionRef.current.setAllSelected(false);
    setSelectionVersion(v => v + 1);
    loadItems(next, true);
  }, [loadItems]);

  React.useEffect(() => {
    if (!listGuid) return;
    // initial load
    loadItems(view, true);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [listGuid]);

  const navigateWithParams = (params: Record<string, string | number | undefined>) => {
    const url = new URL(window.location.href);
    Object.entries(params).forEach(([k, v]) => {
      if (v === undefined || v === null) return;
      url.searchParams.set(k, String(v));
    });
    window.location.href = url.toString();
    console.log(formStatusRecord.value);
  };

  const viewLabel = React.useMemo(() => (
    view === 'submitted' ? 'Submitted' :
      view === 'rejected' ? 'Rejected' :
        view === 'closed' ? 'Closed' : 'Saved'
  ), [view]);

  const cmdItems = React.useMemo<ICommandBarItemProps[]>(() => {
    const editDisabled = selectedRows.length !== 1;
    const delDisabled = selectedRows.length === 0 || !listGuid;
    return [
      {
        key: 'add',
        text: 'Add New',
        disabled: !isPOEligible,
        iconProps: { iconName: 'Add' },
        onClick: () => (onAddNew ? onAddNew() : navigateWithParams({ mode: 'add' }))
      },
      {
        key: 'edit',
        text: 'Edit',
        iconProps: { iconName: 'Edit' },
        disabled: editDisabled,
        onClick: () => {
          const row = selectedRows[0];
          if (row) {
            setFormStatusRecord({ "value": row.FormStatusRecord });
            if (onEdit) onEdit(row.id, row);
            else {
              navigateWithParams({ mode: 'edit', formId: row.id });
            }
          }
        }
      },
      {
        key: 'delete',
        text: 'Delete',
        iconProps: { iconName: 'Delete' },
        disabled: delDisabled,
        onClick: deleteSelected
      },
      {
        key: 'refresh',
        text: 'Refresh',
        iconProps: { iconName: 'Refresh' },
        onClick: () => {
          setNextLink(undefined);
          setHasMore(false);
          loadItems(view, true);
        }
      },
      {
        key: 'view',
        text: `${viewLabel}`,
        iconProps: { iconName: 'View' },
        subMenuProps: {
          items: [
            {
              key: 'saved',
              text: 'Saved',
              iconProps: { iconName: 'Save' },
              onClick: () => switchState('saved')
            },
            {
              key: 'submitted',
              text: 'Submitted',
              iconProps: { iconName: 'ActivateOrders' },
              onClick: () => switchState('submitted')
            },
            {
              key: 'rejected',
              text: 'Rejected',
              iconProps: { iconName: 'StatusErrorFull' },
              onClick: () => switchState('rejected')
            },
            {
              key: 'closed',
              text: 'Closed',
              iconProps: { iconName: 'Cancel' },
              onClick: () => switchState('closed')
            },
          ],
        },
      },
    ];
  }, [selectedRows, listGuid, onAddNew, onEdit, deleteSelected, loadItems, view]);

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      <Text variant="xLarge">{title}</Text>
      <CommandBar items={cmdItems} />
      {loading && <Spinner label="Loading..." />}
      {error && <Text styles={{ root: { color: 'red' } }}>{error}</Text>}
      <MarqueeSelection selection={selectionRef.current}>
        <DetailsList items={items}
          columns={columns}
          selection={selectionRef.current}
          selectionMode={SelectionMode.multiple}
          layoutMode={DetailsListLayoutMode.justified}
          setKey={`ptwForms-${view}`}
          compact
          isHeaderVisible
          styles={{ root: { minHeight: '350px', height: 'auto' } }}
        />
      </MarqueeSelection>

      {/* Paged footer */}
      <Stack horizontal horizontalAlign="center" tokens={{ childrenGap: 8 }}>
        {hasMore && !loading && (
          <DefaultButton
            text={loadingMore ? 'Loading…' : 'Load more'}
            disabled={loadingMore}
            onClick={() => loadItems(view, false)}
          />
        )}
        {loadingMore && <Spinner size={0} />} {/* small spinner */}
        {!hasMore && !loading && items.length >= PAGE_SIZE && (
          <Text styles={{ root: { color: '#605e5c' } }}>No More Results</Text>
        )}
      </Stack>
    </Stack>
  );

};

export default SubmittedPTWFormsList;