import * as React from 'react';
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
  Spinner
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPCrudOperations } from '../../../Classes/SPCrudOperations';

type Row = {
  id: number;
  employeeName?: string;
  employeeID?: number;
  reason?: string;
  replacementReason?: string;
  jobTitle?: string;
  department?: string;
  division?: string;
  company?: string;
  requester?: string;
  requesterEmail?: string;
  submitter?: string;
  submitterEmail?: string;
  created?: Date;
  workflowStatus?: string;
  rejectionReason?: string;
};

export interface SubmittedPpeFormsListProps {
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
  const [items, setItems] = React.useState<Row[]>([]);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const selectionRef = React.useRef(
    new Selection({
      selectionMode: SelectionMode.multiple,
      onSelectionChanged: () => setSelectionVersion(v => v + 1)
    })
  );
  const [selectionVersion, setSelectionVersion] = React.useState(0);

  const selectedRows = React.useMemo(() => (selectionRef.current.getSelection() as Row[]) || [], [selectionVersion]);

  const columns = React.useMemo<IColumn[]>(
    () => [
      // { key: 'colId', name: 'ID', fieldName: 'id', minWidth: 50, maxWidth: 70 },
      { key: 'colEmpId', name: 'Emp #', fieldName: 'employeeID', minWidth: 70, maxWidth: 90 },
      { key: 'colEmployee', name: 'Employee', fieldName: 'employeeName', minWidth: 150, isResizable: true },
      { key: 'colReason', name: 'Reason', fieldName: 'reason', minWidth: 110 },
      { key: 'colJobTitle', name: 'Job Title', fieldName: 'jobTitle', minWidth: 120, isResizable: true },
      { key: 'colDept', name: 'Department', fieldName: 'department', minWidth: 160, isResizable: true },
      { key: 'colDivision', name: 'Division', fieldName: 'division', minWidth: 120, isResizable: true },
      { key: 'colCompany', name: 'Company', fieldName: 'company', minWidth: 120 },
      { key: 'colRequester', name: 'Requester', fieldName: 'requester', minWidth: 140 },
      { key: 'colSubmitter', name: 'Submitter', fieldName: 'submitter', minWidth: 140 },
      // { key: 'colworkflowStatus', name: 'Workflow Status', fieldName: 'workflowStatus', minWidth: 200, isResizable: true },
      { key: 'colRejectionReason', name: 'Rejection Reason', fieldName: 'rejectionReason', minWidth: 200, isResizable: true },
      {
        key: 'colCreated', name: 'Date Submitted', fieldName: 'created', minWidth: 140,
        onRender: (row: Row) => (row.created ? row.created.toLocaleDateString() : '')
      }
    ],
    []
  );

  const loadItems = React.useCallback(async () => {
    setLoading(true);
    setError(undefined);
    try {
      const guid = listGuid;
      if (!guid) {
        setItems([]);
        setError('PPEForm list GUID not provided.');
        setLoading(false);
        return;
      }

      const select = `?$select=Id,EmployeeID,ReasonForRequest,ReplacementReason,Created,WorkflowStatus,RejectionReason,EmployeeRecord/FullName,` +
        `JobTitleRecord/Title,DepartmentRecord/Title,DivisionRecord/Title,CompanyRecord/Title,` +
        `RequesterName/Title,RequesterName/EMail,SubmitterName/Title,SubmitterName/EMail` +
        `&$expand=EmployeeRecord,JobTitleRecord,DepartmentRecord,DivisionRecord,CompanyRecord,RequesterName,SubmitterName` +
        `&$orderby=Created desc`;

      const crud = new SPCrudOperations(context.spHttpClient, context.pageContext.web.absoluteUrl, guid, select);
      const data: any[] = await crud._getItemsWithQuery();
      const filteredItems = data.filter( item => !item.WorkflowStatus?.toLowerCase().includes('closed'));
      const mapped: Row[] = (filteredItems || []).map((obj: any): Row => {
        const created = obj.Created ? new Date(obj.Created) : undefined;
        return {
          id: Number(obj.Id),
          employeeID: obj.EmployeeID ?? undefined,
          employeeName: obj.EmployeeRecord?.FullName ?? undefined,
          reason: obj.ReasonForRequest ?? undefined,
          replacementReason: obj.ReplacementReason ?? undefined,
          jobTitle: obj.JobTitleRecord?.Title ?? undefined,
          department: obj.DepartmentRecord?.Title ?? undefined,
          division: obj.DivisionRecord?.Title ?? undefined,
          company: obj.CompanyRecord?.Title ?? undefined,
          requester: obj.RequesterName?.Title ?? undefined,
          requesterEmail: obj.RequesterName?.EMail ?? undefined,
          submitter: obj.SubmitterName?.Title ?? undefined,
          submitterEmail: obj.SubmitterName?.EMail ?? undefined,
          workflowStatus: obj.WorkflowStatus ?? undefined,
          rejectionReason: obj.RejectionReason ?? undefined,
          created,
        };
      });

    setItems(mapped);
    } catch (e: any) {
      setError(`Failed to load forms: ${e?.message || e}`);
    } finally {
      setLoading(false);
    }
  }, [context, listGuid]);

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
    loadItems();
  }, [loadItems]);

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
        onClick: () => loadItems()
      }
    ];
  }, [selectedRows, listGuid, onAddNew, onEdit, deleteSelected, loadItems]);

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      <Text variant="xLarge">{title}</Text>
      <CommandBar items={cmdItems} />
      {loading && <Spinner label="Loading..." />}
      {error && <Text styles={{ root: { color: 'red' } }}>{error}</Text>}
      <MarqueeSelection selection={selectionRef.current}>
        <DetailsList items={items} columns={columns} selection={selectionRef.current} selectionMode={SelectionMode.multiple}
          layoutMode={DetailsListLayoutMode.justified} setKey="ppeForms" compact isHeaderVisible />
      </MarqueeSelection>
    </Stack>
  );
};

export default SubmittedPpeFormsList;
