import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  Selection,
  SelectionMode,
  MarqueeSelection,
  CommandBar,
  ICommandBarItemProps,
  Stack,
  Text,
  Spinner,
  Pivot,
  PivotItem,
  DefaultButton
} from '@fluentui/react';

export type Row = {
  id: number;
  employeeName?: string;
  coralEmployeeID: number;
  reason?: string;
  replacementReason?: string;
  jobTitle?: string;
  department?: string;
  company?: string;
  requester?: string;
  requesterEmail?: string;
  submitter?: string;
  submitterEmail?: string;
  created?: Date;
  workflowStatus?: string;
  rejectionReason?: string;
  coralReferenceNumber?: string;
};


interface SubmittedPpeFormsListProps {
  context: WebPartContext;
  /** PPEForm list GUID. If not supplied, the component will try to operate but delete will be disabled. */
  listGuid?: string;
  title?: string;
  /** Called when user clicks Add New. If not provided, navigates with ?mode=add */
  onAddNew?: () => void;
  /** Called when user clicks Edit and exactly one row is selected. If not provided, navigates with ?formId=ID&mode=edit */
  onEdit?: (formId: number, row: Row) => void;
  /** Called after successful delete with the deleted IDs */
  onDelete?: (deletedIds: number[]) => void;
}

const SubmittedPpeFormsList: React.FC<SubmittedPpeFormsListProps> = ({ context, listGuid, title = 'PERSONAL PROTECTIVE EQUIPMENT (PPE) SUBMISSIONS', onAddNew, onEdit, onDelete }) => {
  const [loading, setLoading] = React.useState(true);
  const [loadingMore, setLoadingMore] = React.useState(false);
  const [items, setItems] = React.useState<Row[]>([]);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [view, setView] = React.useState<'active' | 'closed'>('active');
  const [selectionVersion, setSelectionVersion] = React.useState(0);
  const [nextLink, setNextLink] = React.useState<string | undefined>(undefined);
  const [hasMore, setHasMore] = React.useState<boolean>(false);
  const PAGE_SIZE = 50;

  const selectionRef = React.useRef(
    new Selection({
      selectionMode: SelectionMode.multiple,
      onSelectionChanged: () => setSelectionVersion(v => v + 1)
    })
  );

  const selectedRows = React.useMemo(() => (selectionRef.current.getSelection() as Row[]) || [], [selectionVersion]);

  // Build the shared select/expand pieces (without $top so we can append it)
  const baseSelect = React.useMemo(() => (
    `?$select=Id,ReasonForRequest,ReasonRecord,Created,WorkflowStatus,RejectionReason,EmployeeRecord/FullName,EmployeeRecord/CoralEmployeeID,` +
    `JobTitleRecord/Title,DepartmentRecord/Title,CompanyRecord/Title,CoralReferenceNumber,` +
    `RequesterName/Title,RequesterName/EMail,SubmitterName/Title,SubmitterName/EMail` +
    `&$expand=EmployeeRecord,JobTitleRecord,DepartmentRecord,CompanyRecord,RequesterName,SubmitterName`
  ), []);

  const mapRows = React.useCallback((data: any[]): Row[] => {
    return (data || []).map((obj: any): Row => {
      const created = obj.Created ? new Date(obj.Created) : undefined;
      return {
        id: Number(obj.Id),
        coralEmployeeID: obj.EmployeeRecord?.CoralEmployeeID ?? undefined,
        employeeName: obj.EmployeeRecord?.FullName ?? undefined,
        reason: obj.ReasonForRequest ?? undefined,
        replacementReason: obj.ReplacementReason ?? undefined,
        jobTitle: obj.JobTitleRecord?.Title ?? undefined,
        department: obj.DepartmentRecord?.Title ?? undefined,
        company: obj.CompanyRecord?.Title ?? undefined,
        requester: obj.RequesterName?.Title ?? undefined,
        requesterEmail: obj.RequesterName?.EMail ?? undefined,
        submitter: obj.SubmitterName?.Title ?? undefined,
        submitterEmail: obj.SubmitterName?.EMail ?? undefined,
        workflowStatus: obj.WorkflowStatus ?? undefined,
        rejectionReason: obj.RejectionReason ?? undefined,
        created,
        coralReferenceNumber: obj.CoralReferenceNumber ?? undefined
      };
    });
  }, []);

  const columns = React.useMemo<IColumn[]>(
    () => [
      { key: 'colCoralReferenceNumber', name: 'Ref #', fieldName: 'coralReferenceNumber', minWidth: 100, maxWidth: 150, isResizable: true },
      { key: 'colEmpId', name: 'Emp #', fieldName: 'coralEmployeeID', minWidth: 70, maxWidth: 90 },
      { key: 'colEmployee', name: 'Employee', fieldName: 'employeeName', minWidth: 150, isResizable: true },
      { key: 'colReason', name: 'Reason', fieldName: 'reason', minWidth: 110 },
      { key: 'colJobTitle', name: 'Job Title', fieldName: 'jobTitle', minWidth: 120, isResizable: true },
      { key: 'colDept', name: 'Department', fieldName: 'department', minWidth: 160, isResizable: true },
      { key: 'colCompany', name: 'Company', fieldName: 'company', minWidth: 120 },
      { key: 'colRequester', name: 'Requester', fieldName: 'requester', minWidth: 140 },
      { key: 'colSubmitter', name: 'Submitter', fieldName: 'submitter', minWidth: 140 },
      { key: 'colworkflowStatus', name: 'Status', fieldName: 'workflowStatus', minWidth: 200, isResizable: true },
      { key: 'colRejectionReason', name: 'Rejection Reason', fieldName: 'rejectionReason', minWidth: 200, isResizable: true },
      {
        key: 'colCreated', name: 'Date Submitted', fieldName: 'created', minWidth: 140,
        onRender: (row: Row) => (row.created ? row.created.toLocaleDateString() : '')
      }
    ],
    []
  );

  const loadItems = React.useCallback(async (scope: 'active' | 'closed' = view, reset: boolean = false) => {

    if (!listGuid) {
      setItems([]);
      setHasMore(false);
      setNextLink(undefined);
      return;
    }

    // Null-safe, startswith filter to keep "Closed By System" and any "Closed ..." statuses separate
    const filterActive = `&$filter=WorkflowStatus ne 'Closed By System'`;
    const filterClosed = `&$filter=WorkflowStatus eq 'Closed By System'`;
    const orderBy = `&$orderby=Created desc`;

    const headers = {
      Accept: 'application/json;odata=nometadata',
      'odata-version': ''
    } as any;
    const webUrl = context.pageContext.web.absoluteUrl;
    // Decide which URL to call
    let url: string;
    if (!reset && nextLink) {
      // Continue from SharePoint’s paging link (absolute URL)
      url = nextLink;
    } else {
      // First page
      const filter = scope === 'closed' ? filterClosed : filterActive;
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
      const next =
        json['odata.nextLink'] ||
        json['@odata.nextLink'] ||
        json.d?.__next ||
        undefined;

      setItems(prev => (reset ? rows : prev.concat(rows)));
      setNextLink(next);
      setHasMore(!!next);

      // Clear selection when view resets
      if (reset) {
        selectionRef.current.setAllSelected(false);
        setSelectionVersion(v => v + 1);
      }
    } catch (e: any) {
      setError(`Failed to load forms: ${e?.message || e}`);
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
      await loadItems();
    } catch (e: any) {
      setError(`Delete error: ${e?.message || e}`);
    }
  }, [selectedRows, listGuid, context, loadItems, onDelete]);

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
  };

  const cmdItems = React.useMemo<ICommandBarItemProps[]>(() => {
    const editDisabled = selectedRows.length !== 1;
    const delDisabled = selectedRows.length === 0 || !listGuid;
    return [
      {
        key: 'add',
        text: 'Add New',
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
            if (onEdit) onEdit(row.id, row);
            else navigateWithParams({ mode: 'edit', formId: row.id });
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
      }
    ];
  }, [selectedRows, listGuid, onAddNew, onEdit, deleteSelected, loadItems, view]);

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      <Text variant="xLarge">{title}</Text>
      <CommandBar items={cmdItems} />
      <Pivot
        selectedKey={view}
        onLinkClick={(item) => {
          const key = (item?.props.itemKey as 'active' | 'closed') ?? 'active';
          setView(key);
          setNextLink(undefined);
          setHasMore(false);
          loadItems(key, true); // reset paging
        }}
      >
        <PivotItem headerText="Active" itemKey="active" />
        <PivotItem headerText="Closed" itemKey="closed" />
      </Pivot>
      {loading && <Spinner label="Loading..." />}
      {error && <Text styles={{ root: { color: 'red' } }}>{error}</Text>}
      <MarqueeSelection selection={selectionRef.current}>
        <DetailsList
          items={items}
          columns={columns}
          selection={selectionRef.current}
          selectionMode={SelectionMode.multiple}
          layoutMode={DetailsListLayoutMode.justified}
          setKey={`ppeForms-${view}`}
          compact
          isHeaderVisible
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

export default SubmittedPpeFormsList;