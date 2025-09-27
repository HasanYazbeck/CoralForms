import * as React from "react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { IGraphResponse, IGraphUserResponse, ILKPItemInstructionsForUse } from "../../../Interfaces/ICommon";


// Components
import { DefaultPalette, DetailsListLayoutMode } from "@fluentui/react";
import type { IPpeFormWebPartProps } from "./IPpeFormProps";
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';
import { TextField } from '@fluentui/react/lib/TextField';
import { ComboBox, IComboBoxOption } from '@fluentui/react/lib/ComboBox';
import { Stack, IStackStyles } from '@fluentui/react/lib/Stack';
import { DetailsList, IColumn, Selection, SelectionMode } from '@fluentui/react';
import { DatePicker, mergeStyleSets, defaultDatePickerStrings } from '@fluentui/react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Label } from '@fluentui/react/lib/Label';
import { Checkbox } from '@fluentui/react';
import { Separator } from '@fluentui/react/lib/Separator';
import { MessageBar } from '@fluentui/react/lib/MessageBar';
import { CommandBar, ICommandBarItemProps } from '@fluentui/react/lib/CommandBar';

// Styles
import "bootstrap/dist/css/bootstrap.min.css";
import styles from "./PpeForm.module.scss";
import { SPCrudOperations } from "../../../Classes/SPCrudOperations";
import { SPHelpers } from "../../../Classes/SPHelpers";
import { ICoralFormsList } from "../../../Interfaces/ICoralFormsList";
import { IUser } from "../../../Interfaces/IUser";
import { IPPEItem } from "../../../Interfaces/IPPEItem";
import { IPPEItemDetails } from "../../../Interfaces/IPPEItemDetails";
import { IEmployeeProps, IEmployeesPPEItemsCriteria } from "../../../Interfaces/IEmployeeProps";
import { IFormsApprovalWorkflow } from "../../../Interfaces/IFormsApprovalWorkflow";

const stackStyles: IStackStyles = {
  root: {
    background: DefaultPalette.themeTertiary,
    display: "inline",
  },
};

const datePickerStyles = mergeStyleSets({
  root: { selectors: { '> *': { marginBottom: 15 } } },
  control: { maxWidth: 300, marginBottom: 15 },
});

export default function PPEFormDynamic(props: IPpeFormWebPartProps) {
  // Helpers and refs
  const formName = "PERSONAL PROTECTIVE EQUIPMENT";
  const spHelpers = useMemo(() => new SPHelpers(), []);
  const spCrudRef = useRef<SPCrudOperations | undefined>(undefined);
  const selectionRef = useRef(new Selection());

  // Local state (converted from class state)
  const [jobTitle, setJobTitle] = useState("");

  const [department, setDepartment] = useState("");
  const [division, setDivision] = useState("");
  const [company, setCompany] = useState("");
  const [_employee, setEmployee] = useState<IPersonaProps[]>([]);
  const [_employeeId, setEmployeeId] = useState<number | undefined>(undefined);
  const [submitter, setSubmitter] = useState<IPersonaProps[]>([]);
  const [requester, setRequester] = useState<IPersonaProps[]>([]);
  const [isReplacementChecked, setIsReplacementChecked] = useState(false);

  // New hook state
  const [users, setUsers] = useState<IUser[]>([]);
  // Employees list items fetched via _getEmployees (used for picker search)
  const [employees, setEmployees] = useState<IEmployeeProps[]>([]);
  const [, setEmployeePPEItemsCriteria] = useState<IEmployeesPPEItemsCriteria>({ Id: '' });
  const [ppeItems, setPpeItems] = useState<IPPEItem[]>([]);
  const [ppeItemDetails, setPpeItemDetails] = useState<IPPEItemDetails[]>([]);
  const [itemInstructionsForUse, setItemInstructionsForUse] = useState<ILKPItemInstructionsForUse[]>([]);
  const [, setFormsApprovalWorkflow] = useState<IFormsApprovalWorkflow[]>([]);
  const [, setCoralFormsList] = useState<ICoralFormsList>({ Id: "" });
  const [loading, setLoading] = useState<boolean>(true);

  // Rows for the items table
  const [ppeItemsRows, setPpeItemsRows] = useState<Array<any>>([]);
  // Approvals sign-off rows (Department, HR, HSE, Warehouse)
  const [approvalsRows, setApprovalsRows] = useState<Array<any>>([
    { SignOff: 'Department Approval', Name: '', Status: '', Reason: '', Date: undefined, __index: 0 },
    { SignOff: 'HR Approval', Name: '', Status: '', Reason: '', Date: undefined, __index: 1 },
    { SignOff: 'HSE Approval', Name: '', Status: '', Reason: '', Date: undefined, __index: 2 },
    { SignOff: 'Warehouse Approval', Name: '', Status: '', Reason: '', Date: undefined, __index: 3 }
  ]);

  // Helper: ensure we return a string[] or undefined from strings or arrays
  const normalizeToStringArray = useCallback((val: any): string[] | undefined => {
    if (val === undefined || val === null) return undefined;
    if (Array.isArray(val)) return val.map((v) => (v !== undefined && v !== null ? String(v) : '')).filter(Boolean);
    if (typeof val === 'string') return val.split(',').map(s => s.trim()).filter(Boolean);
    // If it's an object with 'results' (SharePoint REST sometimes), handle that
    if (val && typeof val === 'object' && Array.isArray((val as any).results)) return (val as any).results.map((v: any) => String(v)).filter(Boolean);
    return undefined;
  }, []);

  // ---------------------------
  // Data-loading functions (ported)
  // ---------------------------
  const _getUsers = useCallback(async (): Promise<IUser[]> => {
    let fetched: IUser[] = [];
    let endpoint: string | null = "/users?$select=id,displayName,mail,department,jobTitle,mobilePhone,officeLocation&$expand=manager($select=id,displayName)";
    try {
      do {
        const client: MSGraphClientV3 = await (props.context as any).msGraphClientFactory.getClient("3");
        const response: IGraphResponse = await client.api(endpoint).get();
        if (response?.value && Array.isArray(response.value)) {
          const seenIds = new Set<string>();
          const mappedUsers = response.value
            .filter((u: IGraphUserResponse) => u.mail)
            .filter((user) => user.mail && !user.mail?.toLowerCase().includes("healthmailbox") && !user.mail?.toLowerCase().includes("softflow-intl.com") && !user.mail?.toLowerCase().includes("sync"))
            .filter(user => {
              if (seenIds.has(user.id)) return false;
              seenIds.add(user.id);
              return true;
            })
            .map((user: IGraphUserResponse) => ({
              id: user.id,
              displayName: user.displayName,
              email: user.mail,
              jobTitle: user.jobTitle,
              department: user.department,
              officeLocation: user.officeLocation,
              mobilePhone: user.mobilePhone,
              profileImageUrl: undefined,
              isSelected: false,
              manager: user.manager ? { id: user.manager.id, displayName: user.manager.displayName } : undefined,
            } as IUser));

          fetched.push(...mappedUsers);
          endpoint = response["@odata.nextLink"] || null;
        } else {
          break;
        }
      } while (endpoint);

      if (fetched.length > 0) {
        setUsers(fetched);
      }
      return fetched;
    } catch (error) {
      console.error("Error fetching users:", error);
      setUsers([]);
      return [];
    }
  }, [props.context]);

  const _getEmployees = useCallback(async (usersArg?: IUser[], employeeFullName?: string): Promise<IEmployeeProps[]> => {
    try {
      const query: string = `?$select=Id,EmployeeID,FullName,Division/Id,Division/Title,Company/Id,Company/Title,EmploymentStatus,JobTitle/Id,JobTitle/Title,Department/Id,Department/Title,Manager/Id,Manager/FullName,Created&$expand=Division,Company,JobTitle,Department,Manager&$filter=substringof('${employeeFullName}', FullName)&$orderby=Order asc`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, '6331f331-caa7-4732-a205-7abcd1f7d53f', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: IEmployeeProps[] = [];
      const usersToUse = usersArg && usersArg.length ? usersArg : users;
      data.forEach((obj: any) => {
        if (obj) {
          const createdBy = usersToUse && usersToUse.length ? usersToUse.filter(u => u.id.toString() === obj.AuthorId?.toString())[0] : undefined;
          let created: Date | undefined;
          if (obj.Created !== undefined) created = new Date(spHelpers.adjustDateForGMTOffset(obj.Created));
          const temp: IEmployeeProps = {
            Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            employeeID: obj.EmployeeID !== undefined && obj.EmployeeID !== null ? obj.EmployeeID : 0,
            fullName: obj.FullName !== undefined && obj.FullName !== null ? obj.FullName : undefined,
            jobTitle: obj.JobTitle !== undefined && obj.JobTitle !== null ? { id: obj.JobTitle.Id, title: obj.JobTitle.Title } : undefined,
            company: obj.Company !== undefined && obj.Company !== null ? { id: obj.Company.Id, title: obj.Company.Title } : undefined,
            department: obj.Department !== undefined && obj.Department !== null ? { id: obj.Department.Id, title: obj.Department.Title } : undefined,
            manager: obj.Manager !== undefined && obj.Manager !== null ? { Id: obj.Manager.Id, fullName: obj.Manager.FullName } as IEmployeeProps : undefined,
            employmentStatus: obj.EmploymentStatus !== undefined && obj.EmploymentStatus !== null ? obj.EmploymentStatus : undefined,
            division: obj.Division !== undefined && obj.Division !== null ? { id: obj.Division.Id, title: obj.Division.Title } : undefined,
            Created: created !== undefined ? created : undefined,
            CreatedBy: createdBy !== undefined ? createdBy : undefined,
          };

          result.push(temp);
        }
      });
      setEmployees(result);
      return result;
    } catch (error) {
      console.error('An error has occurred while retrieving items!', error);
      setEmployees([]);
      return [];
    }
  }, [props.context, spHelpers]);

  const _getEmployeesPPEItemsCriteria = useCallback(async (usersArg?: IUser[], employeeID?: string) => {
    try {
      const query: string = `?$select=Id,Employee/EmployeeID,Employee/FullName,Created,SafetyHelmet,ReflectiveVest,SafetyShoes,` +
        `Employee/ID,Employee/FullName,RainSuit/Id,RainSuit/DisplayText,UniformCoveralls/Id,UniformCoveralls/DisplayText,UniformTop/Id,UniformTop/DisplayText,` +
        `UniformPants/Id,UniformPants/DisplayText,WinterJacket/Id,WinterJacket/DisplayText` +
        `&$expand=Employee,RainSuit,UniformCoveralls,UniformTop,UniformPants,WinterJacket` +
        `&$filter=Employee/EmployeeID eq ${employeeID}&$orderby=Order asc`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, '2f3c099b-5355-4796-b40a-6f2c728b849a', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      let result: IEmployeesPPEItemsCriteria;

      if (data && data.length > 0) {
        const obj = data[0]; // Get the first object
        result = {
          Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
          employeeID: obj.Employee !== undefined && obj.Employee !== null ? obj.Employee.EmployeeID : undefined,
          fullName: obj.Employee !== undefined && obj.Employee !== null ? obj.Employee.FullName : undefined,
          reflectiveVest: obj.ReflectiveVest !== undefined && obj.ReflectiveVest !== null ? obj.ReflectiveVest : undefined,
          safetyHelmet: obj.SafetyHelmet !== undefined && obj.SafetyHelmet !== null ? obj.SafetyHelmet : undefined,
          safetyShoes: obj.SafetyShoes !== undefined && obj.SafetyShoes !== null ? obj.SafetyShoes : undefined,
          rainSuit: obj.RainSuit !== undefined && obj.RainSuit !== null ? obj.RainSuit.DisplayText : undefined,
          uniformCoveralls: obj.UniformCoveralls !== undefined && obj.UniformCoveralls !== null ? obj.UniformCoveralls.DisplayText : undefined,
          uniformTop: obj.UniformTop !== undefined && obj.UniformTop !== null ? obj.UniformTop.DisplayText : undefined,
          uniformPants: obj.UniformPants !== undefined && obj.UniformPants !== null ? obj.UniformPants.DisplayText : undefined,
          winterJacket: obj.WinterJacket !== undefined && obj.WinterJacket !== null ? obj.WinterJacket.DisplayText : undefined,
          Created: undefined, CreatedBy: undefined,
        };
        setEmployeePPEItemsCriteria(result);

      }

    } catch (error) {
      console.error('An error has occurred while retrieving items!', error);
      setEmployeePPEItemsCriteria({ Id: '' });

    }
  }, [props.context, spHelpers]);

  const _getCoralFormsList = useCallback(async (usersArg?: IUser[]): Promise<ICoralFormsList | undefined> => {
    try {
      const searchEscaped = formName.replace(/'/g, "''");
      const query: string = `?$select=Id,Title,hasInstructionForUse,hasWorkflow,Created&$filter=substringof('${searchEscaped}', Title)`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, '22a9fee1-dfe4-4ad0-8ce4-89d014c63049', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const usersToUse = usersArg && usersArg.length ? usersArg : users;
      const ppeform = data.find((obj: any) => obj !== null);
      let result: ICoralFormsList = { Id: "" };

      if (ppeform) {
        const createdBy = usersToUse?.find(u => u.id.toString() === ppeform.AuthorId?.toString());
        const created = ppeform.Created ? new Date(spHelpers.adjustDateForGMTOffset(ppeform.Created)) : undefined;
        result = {
          Id: ppeform.Id ?? undefined,
          Title: ppeform.Title ?? undefined,
          CreatedBy: createdBy,
          Created: created,
          hasInstructionForUse: ppeform.hasInstructionForUse ?? undefined,
          hasWorkflow: ppeform.hasWorkflow ?? undefined,
        };
      }
      setCoralFormsList(result);
      return result;
    } catch (error) {
      // console.error('An error has occurred!', error);
      setCoralFormsList({ Id: '' });
      return undefined;
    }
  }, [props.context, spHelpers]);

  const _getPPEItems = useCallback(async (usersArg?: IUser[]) => {
    try {
      const query: string = `?$select=Id,Title,Brands,Created`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, '00f4b4fc-896d-40bb-9a03-3889e651d244', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: IPPEItem[] = [];
      const usersToUse = usersArg && usersArg.length ? usersArg : users;
      data.forEach((obj: any) => {
        if (obj) {
          const createdBy = usersToUse && usersToUse.length ? usersToUse.filter(u => u.id.toString() === obj.AuthorId?.toString())[0] : undefined;
          let created: Date | undefined;
          if (obj.Created !== undefined) created = new Date(spHelpers.adjustDateForGMTOffset(obj.Created));
          const temp: IPPEItem = {
            Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            CreatedBy: createdBy !== undefined ? createdBy : undefined,
            Created: created !== undefined ? created : undefined,
            Title: obj.Title !== undefined && obj.Title !== null ? obj.Title : undefined,
            Brands: normalizeToStringArray(obj.Brands),
            Order: obj.Order !== undefined && obj.Order !== null ? obj.Order : undefined,
            PPEItemsDetails: []
          };
          result.push(temp);
        }
      });
      setPpeItems(result);
    } catch (error) {
      console.error('An error has occurred while retrieving items!', error);
      setPpeItems([]);
    }
  }, [props.context, spHelpers]);

  const _getPPEItemsDetails = useCallback(async (usersArg?: IUser[]) => {
    try {
      const query: string = `?$select=Id,Title,PPEItem,Sizes,Types,Created,PPEItem/Id,PPEItem/Title&$expand=PPEItem`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, '3435bbde-cb56-43cf-aacf-e975c65b68c3', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: IPPEItemDetails[] = [];
      const usersToUse = usersArg && usersArg.length ? usersArg : users;
      data.forEach((obj: any) => {
        if (obj) {
          const createdBy = usersToUse && usersToUse.length ? usersToUse.filter(u => u.id.toString() === obj.AuthorId?.toString())[0] : undefined;
          let created: Date | undefined;
          if (obj.Created !== undefined) created = new Date(spHelpers.adjustDateForGMTOffset(obj.Created));
          const temp: IPPEItemDetails = {
            Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            Title: obj.Title !== undefined && obj.Title !== null ? obj.Title : undefined,
            CreatedBy: createdBy !== undefined ? createdBy : undefined,
            Created: created !== undefined ? created : undefined,
            Sizes: normalizeToStringArray(obj.Sizes),
            Types: normalizeToStringArray(obj.Types),
            PPEItem: obj.PPEItem !== undefined ? {
              Id: obj.PPEItem.Id !== undefined && obj.PPEItem.Id !== null ? obj.PPEItem.Id : undefined,
              Title: obj.PPEItem.Title !== undefined && obj.PPEItem.Title !== null ? obj.PPEItem.Title : undefined,
              Brands: normalizeToStringArray(obj.PPEItem.Brands),
              Order: obj.PPEItem.Order !== undefined && obj.PPEItem.Order !== null ? obj.PPEItem.Order : undefined,
            } : undefined,
          };
          result.push(temp);
        }
      });
      setPpeItemDetails(result);
    } catch (error) {
      console.error('An error has occurred while retrieving items!', error);
      setPpeItemDetails([]);
    }
  }, [props.context, spHelpers]);

  const _getLKPItemInstructionsForUse = useCallback(async (usersArg?: IUser[], formName?: string) => {
    try {
      const query: string = `?$select=Id,FormName,Order,Description,Created&$filter=substringof('${formName}', FormName)&$orderby=Order asc`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, '2edbaa23-948a-4560-a553-acbe7bc60e7b', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: ILKPItemInstructionsForUse[] = [];
      const usersToUse = usersArg && usersArg.length ? usersArg : users;
      data.forEach((obj: any) => {
        if (obj) {
          const createdBy = usersToUse && usersToUse.length ? usersToUse.filter(u => u.id.toString() === obj.AuthorId?.toString())[0] : undefined;
          let created: Date | undefined;
          if (obj.Created !== undefined) created = new Date(spHelpers.adjustDateForGMTOffset(obj.Created));
          const temp: ILKPItemInstructionsForUse = {
            Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            FormName: obj.FormName !== undefined && obj.FormName !== null ? obj.FormName : undefined,
            Order: obj.Order !== undefined && obj.Order !== null ? obj.Order : undefined,
            Description: obj.Description !== undefined && obj.Description !== null ? obj.Description : undefined,
            Created: created !== undefined ? created : undefined,
            CreatedBy: createdBy !== undefined ? createdBy : undefined,
          };
          result.push(temp);
        }
      });
      // sort by Order (ascending). If Order is missing, place those items at the end.
      result.sort((a, b) => {
        const aOrder = (a && a.Order !== undefined && a.Order !== null) ? Number(a.Order) : Number.POSITIVE_INFINITY;
        const bOrder = (b && b.Order !== undefined && b.Order !== null) ? Number(b.Order) : Number.POSITIVE_INFINITY;
        return aOrder - bOrder;
      });
      setItemInstructionsForUse(result);
    } catch (error) {
      console.error('An error has occurred while retrieving items!', error);
      setItemInstructionsForUse([]);
    }
  }, [props.context, spHelpers]);

  const _getFormsApprovalWorkflow = useCallback(async (usersArg?: IUser[], formName?: string) => {
    try {
      const query: string = `?$select=Id,FormName/Id,FormName/Title,ManagerName/Id,RecordOrder,Created,SignOffName&$expand=FormName,ManagerName` +
        `&$filter=substringof('${formName}', FormName/Title)&$orderby=RecordOrder asc`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'd084f344-63cb-4426-ae51-d7f875f3f99a', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: IFormsApprovalWorkflow[] = [];
      const usersToUse = usersArg && usersArg.length ? usersArg : users;

      data.forEach((obj: any) => {
        if (obj) {
          const createdBy = usersToUse && usersToUse.length ? usersToUse.filter(u => u.id.toString() === obj.AuthorId?.toString())[0] : undefined;
          let created: Date | undefined;
          const deptManagerPersonas: IPersonaProps | undefined = usersToUse
            .filter(u => u.email?.toString() === obj.DepartmentManager?.EMail?.toString())
            .map(u => ({
              text: u.displayName || '',
              secondaryText: u.email || '',
              id: u.id
            }) as IPersonaProps)[0];

          if (obj.Created !== undefined) created = new Date(spHelpers.adjustDateForGMTOffset(obj.Created));
          const temp: IFormsApprovalWorkflow = {
            Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            FormName: obj.FormName !== undefined && obj.FormName !== null ? obj.FormName : undefined,
            Order: obj.Order !== undefined && obj.Order !== null ? obj.Order : undefined,
            SignOffName: obj.SignOffName !== undefined && obj.SignOffName !== null ? obj.SignOffName : undefined,
            EmployeeId: obj.ManagerName !== undefined && obj.ManagerName !== null ? obj.ManagerName.Id : undefined,
            DepartmentManager: deptManagerPersonas,
            Status: undefined,
            Reason: undefined,
            Date: undefined,
            Created: created !== undefined ? created : undefined,
            CreatedBy: createdBy !== undefined ? createdBy : undefined,
          };
          result.push(temp);
        }
      });
      // sort by Order (ascending). If Order is missing, place those items at the end.
      result.sort((a, b) => {
        const aOrder = (a && a.Order !== undefined && a.Order !== null) ? Number(a.Order) : Number.POSITIVE_INFINITY;
        const bOrder = (b && b.Order !== undefined && b.Order !== null) ? Number(b.Order) : Number.POSITIVE_INFINITY;
        return aOrder - bOrder;
      });
      setFormsApprovalWorkflow(result);
    } catch (error) {
      console.error('An error has occurred while retrieving items!', error);
      setFormsApprovalWorkflow([]);
    }
  }, [props.context, spHelpers]);
  // ---------------------------
  // useEffect: load data on mount
  // ---------------------------
  useEffect(() => {
    let cancelled = false;

    const load = async () => {
      setLoading(true);
      const fetchedUsers = await _getUsers();
      const coralListResult = await _getCoralFormsList(fetchedUsers);
      // await _getEmployeesPPEItemsCriteria(fetchedUsers, _employeeId ? String(_employeeId) : '');
      await _getPPEItems(fetchedUsers);
      await _getPPEItemsDetails(fetchedUsers);

      // Use the returned result from _getCoralFormsList instead of the (possibly stale) coralFormsList state
      if (coralListResult && coralListResult.hasInstructionForUse) {
        if (coralListResult.hasInstructionForUse) await _getLKPItemInstructionsForUse(fetchedUsers, formName);
        if (coralListResult.hasWorkflow) await _getFormsApprovalWorkflow(fetchedUsers, formName);
      }

      if (!cancelled) {
        try {
          const currentUserEmail = props.context.pageContext.user.email;
          const current = fetchedUsers.find(u => u.email === currentUserEmail);
          if (current) setSubmitter([{ text: current.displayName || '', secondaryText: current.email || '', id: current.id }]);
        } catch (e) {
          // ignore if context not available
        }
        setLoading(false);
      }
    };

    load();

    return () => { cancelled = true; };
  }, [_getEmployees, _getUsers, _getPPEItems, _getPPEItemsDetails, _getCoralFormsList, _getLKPItemInstructionsForUse, _getFormsApprovalWorkflow, props.context]);

  // ---------------------------
  // Row helpers
  // ---------------------------
  const createEmptyRow = useCallback(() => ({ Item: '', Brands: '', Required: false, Details: [] as string[], Qty: '', Size: '', SizesSelected: [] as string[], Selected: false }), []);

  const addRow = useCallback(() => {
    setPpeItemsRows(prev => {
      const base = prev && prev.length > 0 ? [...prev] : [createEmptyRow()];
      base.push(createEmptyRow());
      return base;
    });
  }, [createEmptyRow]);

  // Single size selection handler
  const handleSizeChange = useCallback((rowIndex: number, sizeVal: string | undefined) => {
    setPpeItemsRows(prev => {
      const rows = prev && prev.length ? [...prev] : [createEmptyRow()];
      while (rows.length <= rowIndex) rows.push(createEmptyRow());
      const val = sizeVal ? String(sizeVal).trim() : '';
      // @ts-ignore
      rows[rowIndex].SizesSelected = val ? [val] : [];
      // @ts-ignore
      rows[rowIndex].Size = val;
      return rows;
    });
  }, [createEmptyRow]);

  // Removed toggleSizeByType (type-based sizing deprecated)

  const deleteSelectedRows = useCallback(() => {
    setPpeItemsRows(prev => {
      const rows = prev && prev.length > 0 ? [...prev] : [];
      const selectedIndices = selectionRef.current ? selectionRef.current.getSelectedIndices() : [];
      if (selectedIndices && selectedIndices.length > 0) {
        const filtered = rows.filter((r, idx) => selectedIndices.indexOf(idx) === -1);
        selectionRef.current.setAllSelected(false);
        return filtered;
      } else {
        return rows.filter(r => !r.Selected);
      }
    });
  }, []);

  const onRowChange = useCallback((index: number, field: string, value: any) => {
    setPpeItemsRows(prev => {
      const rows = prev && prev.length > 0 ? [...prev] : [createEmptyRow()];
      while (rows.length <= index) rows.push(createEmptyRow());
      // @ts-ignore
      rows[index][field] = value;
      return rows;
    });
  }, [createEmptyRow]);

  // Approvals handlers
  const onApprovalChange = useCallback((index: number, field: string, value: any) => {
    setApprovalsRows(prev => {
      const rows = prev && prev.length > 0 ? [...prev] : [];
      while (rows.length <= index) rows.push({ SignOff: '', Name: '', Signature: '', Date: undefined, __index: rows.length });
      // @ts-ignore
      rows[index][field] = value;
      return rows;
    });
  }, []);

  // Build map: itemTitleLower -> detail titles, and item+detail -> sizes
  const detailsByItem = useMemo(() => {
    const map: Record<string, string[]> = {};
    (ppeItemDetails || []).forEach(d => {
      const it = d && d.PPEItem && d.PPEItem.Title ? String(d.PPEItem.Title).trim() : '';
      const dt = d && d.Title ? String(d.Title).trim() : '';
      if (!it || !dt) return;
      const key = it.toLowerCase();
      if (!map[key]) map[key] = [];
      if (map[key].indexOf(dt) === -1) map[key].push(dt);
    });
    Object.keys(map).forEach(k => map[k].sort((a, b) => a.localeCompare(b)));
    return map;
  }, [ppeItemDetails]);

  const sizesByItemDetail = useMemo(() => {
    const map: Record<string, string[]> = {};
    (ppeItemDetails || []).forEach(d => {
      const it = d && d.PPEItem && d.PPEItem.Title ? String(d.PPEItem.Title).trim() : '';
      const dt = d && d.Title ? String(d.Title).trim() : '';
      if (!it || !dt) return;
      const key = `${it.toLowerCase()}||${dt.toLowerCase()}`;
      const arr = Array.isArray(d.Sizes) ? d.Sizes.map(s => String(s).trim()).filter(Boolean) : [];
      if (!map[key]) map[key] = [];
      arr.forEach(s => { if (map[key].indexOf(s) === -1) map[key].push(s); });
      map[key].sort((a, b) => a.localeCompare(b));
    });
    return map;
  }, [ppeItemDetails]);

  const handleDetailChange = useCallback((rowIndex: number, detail: string | undefined) => {
    setPpeItemsRows(prev => {
      const rows = prev && prev.length ? [...prev] : [createEmptyRow()];
      while (rows.length <= rowIndex) rows.push(createEmptyRow());
      // @ts-ignore
      rows[rowIndex].Details = detail ? [detail] : [];
      // reset sizes selection when detail changes
      // @ts-ignore
      rows[rowIndex].SizesSelected = [];
      // @ts-ignore
      rows[rowIndex].Size = '';
      return rows;
    });
  }, [createEmptyRow]);

  // ---------------------------
  // Handlers
  // ---------------------------

  // Map of Item Title -> Brands[] (deduped)
  const brandsMap = useMemo(() => {
    const map: Record<string, string[]> = {};
    // First, populate from the main PPE items list (ppeItems) which contains Brands
    (ppeItems || []).forEach((pi: any) => {
      const title = pi && pi.Title ? String(pi.Title).trim() : undefined;
      const brandsArr = normalizeToStringArray(pi && pi.Brands ? pi.Brands : undefined) || [];
      if (title) {
        if (!map[title]) map[title] = [];
        map[title] = Array.from(new Set(map[title].concat(brandsArr)));
      }
    });

    // Then merge in any brands from PPEItemDetails (to capture detail-level brands)
    (ppeItemDetails || []).forEach((p: any) => {
      const title = p && p.PPEItem && p.PPEItem.Title ? String(p.PPEItem.Title).trim() : (p && p.Title ? String(p.Title).trim() : undefined);
      const brandsArr = normalizeToStringArray(p && p.PPEItem && p.PPEItem.Brands ? p.PPEItem.Brands : p && p.Brands ? p.Brands : undefined) || [];
      if (title) {
        if (!map[title]) map[title] = [];
        map[title] = Array.from(new Set(map[title].concat(brandsArr)));
      }
    });

    return map;
  }, [ppeItemDetails, ppeItems, normalizeToStringArray]);

  // Removed sizesMap / sizesByTypeMap; sizes now derived from selected detail (single select)

  // When the Item for a row is changed, also pre-fill the Brands field with the first matching brand (if any)
  const handleItemChange = useCallback((index: number, newItem: any) => {
    const newItemStr = newItem !== undefined && newItem !== null ? String(newItem).trim() : '';
    setPpeItemsRows(prev => {
      const rows = prev && prev.length > 0 ? [...prev] : [createEmptyRow()];
      while (rows.length <= index) rows.push(createEmptyRow());
      rows[index].Item = newItemStr;
      rows[index].Brands = '';
      // Reset details & sizes when item changes
      // @ts-ignore
      rows[index].Details = [];
      // @ts-ignore
      rows[index].SizesSelected = [];
      // @ts-ignore
      rows[index].Size = '';
      const brandOptions = brandsMap[newItemStr] || [];
      if (brandOptions.length) {
        setTimeout(() => {
          setPpeItemsRows(current => {
            const copy = current && current.length > 0 ? [...current] : [createEmptyRow()];
            while (copy.length <= index) copy.push(createEmptyRow());
            copy[index].Brands = brandOptions[0];
            return copy;
          });
        }, 0);
      }
      return rows;
    });
  }, [createEmptyRow, brandsMap]);

  const handleEmployeeChange = useCallback(async (items: IPersonaProps[]) => {
    if (items && items.length > 0) {
      const selected = items[0];
      // First try to find in employees list by FullName (fullName -> persona.text)
      const emp = employees.find(e => (e.fullName || '').toLowerCase() === (selected.text || '').toLowerCase());
      // Fallback to users (Graph) if not found
      const user = users.find(u => u.displayName?.toLowerCase() === (selected.text || '').toLowerCase() || u.id === selected.id);
      setEmployee([selected]);
      setEmployeeId(emp?.employeeID);
      setJobTitle(emp?.jobTitle?.title || user?.jobTitle || '');
      setDepartment(emp?.department?.title || user?.department || '');
      // Ensure division stored as a simple string
      setDivision(emp?.division?.title || '');
      setCompany(emp?.company?.title || user?.company || '');
      // Auto-set requester ONLY if Employee list record has a manager; otherwise leave empty
      if (emp?.manager?.fullName) {
        setRequester([{ text: emp.manager.fullName, id: emp.manager.Id ? String(emp.manager.Id) : emp.manager.fullName }]);
      } else {
        setRequester([]);
      }
      try {
        // Fetch PPE items criteria for this employee ID
        await _getEmployeesPPEItemsCriteria(users, emp?.employeeID ? String(emp.employeeID) : '');
      } catch (e) {
        console.warn('Failed to load PPE items criteria for employee', e);
      }
    } else {
      setEmployee([]);
      setEmployeeId(undefined);
      setJobTitle('');
      setDepartment('');
      setDivision('');
      setCompany('');
      setRequester([]);
      try {
        // Optionally clear criteria when no employee selected
        await _getEmployeesPPEItemsCriteria(users, '');
      } catch (e) {
        console.warn('Failed to clear PPE items criteria', e);
      }
    }
  }, [employees, users, _getEmployeesPPEItemsCriteria]);

  // Employee picker dynamic resolver using Employee list instead of raw users
  const employeeOnFilterChanged = useCallback((filterText: string, currentPersonas: IPersonaProps[], limitResults?: number): Promise<IPersonaProps[]> | IPersonaProps[] => {
    if (!filterText || filterText.trim().length === 0) return [];
    // Always return a promise so the picker waits for async results
    return _getEmployees(undefined, filterText).then(fetched => {
      const list = fetched.length ? fetched : employees; // fallback to existing state
      const matches = list
        .filter(e => (e.fullName || '').toLowerCase().includes(filterText.toLowerCase()))
        .map(e => ({ text: e.fullName || '', secondaryText: e.jobTitle?.title, id: (e.Id ? String(e.Id) : e.fullName) }) as IPersonaProps);
      const deduped: IPersonaProps[] = [];
      const seen = new Set<string>();
      matches.forEach(m => { const key = (m.text || '').toLowerCase(); if (!seen.has(key)) { seen.add(key); deduped.push(m); } });
      return limitResults ? deduped.slice(0, limitResults) : deduped;
    });
  }, [_getEmployees, employees]);

  // Requester resolver (merge employees and Graph users for broader search)
  const requesterOnFilterChanged = useCallback((filterText: string, currentPersonas: IPersonaProps[], limitResults?: number): IPersonaProps[] | Promise<IPersonaProps[]> => {
    if (!filterText || filterText.trim().length === 0) return [];
    const lower = filterText.toLowerCase();
    const employeeMatches = employees
      .filter(e => (e.fullName || '').toLowerCase().includes(lower))
      .map(e => ({ text: e.fullName || '', secondaryText: e.jobTitle?.title, id: (e.Id ? String(e.Id) : e.fullName) }) as IPersonaProps);
    const userMatches = users
      .filter(u => (u.displayName || '').toLowerCase().includes(lower))
      .map(u => ({ text: u.displayName || '', secondaryText: u.jobTitle, id: u.id }) as IPersonaProps);
    const combined = employeeMatches.concat(userMatches);
    const deduped: IPersonaProps[] = [];
    const seen = new Set<string>();
    combined.forEach(p => { const key = (p.text || '').toLowerCase(); if (!seen.has(key)) { seen.add(key); deduped.push(p); } });
    return limitResults ? deduped.slice(0, limitResults) : deduped;
  }, [employees, users]);

  const handleRequesterChange = useCallback((items: IPersonaProps[] | undefined) => {
    if (items && items.length) setRequester([items[0]]); else setRequester([]);
  }, []);

  const handleNewRequestChange = useCallback((ev: React.FormEvent<HTMLElement>, checked?: boolean) => {
    if (checked) setIsReplacementChecked(false);
  }, []);

  const handleReplacementChange = useCallback((ev: React.FormEvent<HTMLElement>, checked?: boolean) => {
    setIsReplacementChecked(!!checked);
  }, []);

  // ---------------------------
  // Render
  // ---------------------------
  if (loading) {
    return (
      <div className={styles.loadingContainer}>
        <Spinner label={"Preparing PPE form — fresh items coming right up!"} size={SpinnerSize.large} />
      </div>
    );
  }

  const delayResults = false;
  const logoUrl = `${props.context.pageContext.web.absoluteUrl}/SiteAssets/coral-logo.png`;
  const peopleList: IPersonaProps[] = users.map(user => ({ text: user.displayName || '', secondaryText: user.email || '', id: user.id }));

  const filterPromise = (personasToReturn: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> => {
    if (delayResults) return convertResultsToPromise(personasToReturn);
    return personasToReturn;
  };

  const onFilterChanged = (filterText: string, currentPersonas: IPersonaProps[], limitResults?: number): IPersonaProps[] | Promise<IPersonaProps[]> => {
    if (filterText) {
      let filteredPersonas: IPersonaProps[] = filterPersonasByText(filterText);
      filteredPersonas = removeDuplicates(filteredPersonas, currentPersonas);
      filteredPersonas = limitResults ? filteredPersonas.slice(0, limitResults) : filteredPersonas;
      return filterPromise(filteredPersonas);
    }
    return [];
  };

  const filterPersonasByText = (filterText: string): IPersonaProps[] => peopleList.filter(item => doesTextContain(item.text as string, filterText));
  function doesTextContain(text: string, filterText: string): boolean { if (!text || !filterText) return false; return text.toLowerCase().includes(filterText.toLowerCase()); }
  function removeDuplicates(personas: IPersonaProps[], possibleDupes: IPersonaProps[]) { return personas.filter(persona => !listContainsPersona(persona, possibleDupes)); }
  function listContainsPersona(persona: IPersonaProps, personas: IPersonaProps[]) { if (!personas || !personas.length) return false; return personas.filter(item => item.text === persona.text).length > 0; }
  function convertResultsToPromise(results: IPersonaProps[]): Promise<IPersonaProps[]> { return new Promise<IPersonaProps[]>((resolve) => setTimeout(() => resolve(results), 2000)); }
  function onInputChange(input: string): string { const outlookRegEx = /<.*>/g; const emailAddress = outlookRegEx.exec(input); if (emailAddress && emailAddress[0]) return emailAddress[0].substring(1, emailAddress[0].length - 1); return input; }

  return (
    <div className={styles.ppeFormBackground}>
      <form>
        <div className={styles.formHeader}>
          <img src={logoUrl} alt="Logo" className={styles.formLogo} />
          <span className={styles.formTitle}>PERSONAL PROTECTIVE EQUIPMENT (PPE) REQUISITION FORM</span>
        </div>

        <Stack horizontal styles={stackStyles}>
          <div className="row">
            <div className="form-group col-md-6"><TextField label="Employee ID" value={_employeeId?.toString()} disabled={true} /></div>
          </div>

          <div className="row">
            <div className="form-group col-md-6">
              <NormalPeoplePicker
                label={"Employee Name"}
                itemLimit={1}
                // Use employee list based resolver
                onResolveSuggestions={employeeOnFilterChanged}
                className={'ms-PeoplePicker'}
                key={'employee'}
                removeButtonAriaLabel={'Remove'}
                inputProps={{ onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'), onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'), 'aria-label': 'Employee Picker' }}
                onInputChange={onInputChange}
                resolveDelay={50}
                disabled={false}
                onChange={handleEmployeeChange}
              />
            </div>

            <div className="form-group col-md-6">
              <DatePicker disabled value={new Date(Date.now())} label="Date Requested" className={datePickerStyles.control} strings={defaultDatePickerStrings} />
            </div>
          </div>

          <div className="row">
            <div className="form-group col-md-6">
              <TextField label="Job Title" value={jobTitle} disabled={true} />
            </div>
            <div className="form-group col-md-6">
              <TextField label="Department" value={department} disabled={true} />
            </div>
          </div>

          <div className="row">
            <div className="form-group col-md-6"><TextField label="Division" value={division} disabled={true} /></div>
            <div className="form-group col-md-6"><TextField label="Company" value={company} disabled={true} /></div>
          </div>

          <div className="row">
            <div className="form-group col-md-6">
              <NormalPeoplePicker
                label={"Requester Name"}
                itemLimit={1}
                onResolveSuggestions={requesterOnFilterChanged}
                className={'ms-PeoplePicker'}
                key={'requester'}
                removeButtonAriaLabel={'Remove'}
                inputProps={{ onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'), onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'), 'aria-label': 'Requester Picker' }}
                onInputChange={onInputChange}
                resolveDelay={150}
                disabled={false}
                onChange={handleRequesterChange}
                selectedItems={requester}
              />
            </div>

            <div className="form-group col-md-6">
              <NormalPeoplePicker label={"Submitter Name"} itemLimit={1} onResolveSuggestions={onFilterChanged} className={'ms-PeoplePicker'} key={'normal'} removeButtonAriaLabel={'Remove'} inputProps={{ onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'), onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'), 'aria-label': 'People Picker' }} onInputChange={onInputChange} resolveDelay={300} disabled={true} selectedItems={submitter} />
            </div>
          </div>

          <div className={`row  ${styles.mt10}`}>
            <div className="form-group col-md-12 d-flex justify-content-between" >
              <Label htmlFor={""}>Reason for Request</Label>

              <Checkbox label="New Request" className="align-items-center" checked={!isReplacementChecked} onChange={handleNewRequestChange} />

              <Checkbox label="Replacement" className="align-items-center" checked={isReplacementChecked} onChange={handleReplacementChange} />

              <TextField placeholder="Reason" disabled={!isReplacementChecked} />
            </div>
          </div>
        </Stack>

        <Separator />

        <div className="mb-2 text-center">
          <small className="text-muted" style={{ fontStyle: 'italic', fontSize: '1.0rem' }}>Please complete the table below in the blank spaces; grey spaces are for administrative use only.</small>
        </div>

        <Separator />
        {/* PPE Items Grid */}
        <Stack horizontal styles={stackStyles}>
          <div className="row">
            <div className="form-group col-md-12">
              {(() => {
                const commandBarItems: ICommandBarItemProps[] = [
                  { key: 'addItem', text: 'Add Item', iconProps: { iconName: 'Add' }, onClick: addRow },
                  { key: 'deleteSelected', text: 'Delete', iconProps: { iconName: 'Delete' }, onClick: deleteSelectedRows }
                ];
                return <CommandBar items={commandBarItems} styles={{ root: { marginBottom: 8 } }} />;
              })()}

              {(() => {
                const defaultRows = [createEmptyRow()];
                const rows = ppeItemsRows && ppeItemsRows.length > 0 ? ppeItemsRows : defaultRows;
                const items = rows.map((r, idx) => ({ ...r, __index: idx } as any));
                const itemDetailsFromDetails: string[] = (ppeItemDetails || []).map((p: any) => (p && p.PPEItem && p.PPEItem.Title) ? String(p.PPEItem.Title).trim() : (p && p.Title ? String(p.Title).trim() : undefined)).filter(Boolean) as string[];
                const itemDetailsFromItems: string[] = (ppeItems || []).map((pi: any) => pi && pi.Title ? String(pi.Title).trim() : undefined).filter(Boolean) as string[];
                const allTitles = itemDetailsFromDetails.concat(itemDetailsFromItems).filter(Boolean) as string[];
                const distinctTitles = Array.from(new Set(allTitles)).sort((a, b) => a.localeCompare(b));
                const itemOptions: IComboBoxOption[] = distinctTitles.map(t => ({ key: t, text: t }));

                const columns: IColumn[] = [
                  {
                    key: 'columnItem', name: 'Item', fieldName: 'Item', minWidth: 150, isResizable: true, onRender: (item: any) => (
                      <div className={styles.comboCell}>
                        <ComboBox allowFreeform autoComplete="on" selectedKey={item.Item || undefined} options={itemOptions} onChange={(ev, option, index, value) => { const newVal = option ? option.key : value; handleItemChange(item.__index, newVal || ''); }} />
                      </div>
                    )
                  },
                  {
                    key: 'columnRequired', name: 'Required', fieldName: 'Required', minWidth: 80, maxWidth: 90, isResizable: false, onRender: (item: any) => (
                      <div className={styles.comboCell} style={{ textAlign: 'center' }}>
                        <Checkbox ariaLabel="Required" checked={!!item.Required} onChange={(_ev, checked) => onRowChange(item.__index, 'Required', !!checked)} />
                      </div>
                    )
                  },

                  {
                    key: 'columnBrand', name: 'Brand', fieldName: 'Brands', minWidth: 120, isResizable: true, onRender: (item: any) => {
                      const options = (brandsMap && brandsMap[item.Item]) ? brandsMap[item.Item].map((b: string) => ({ key: b, text: b })) : [];
                      return (
                        <div className={styles.comboCell}>
                          <div style={{ width: '100%' }}>
                            <ComboBox allowFreeform autoComplete="on" selectedKey={item.Brands || undefined} options={options} onChange={(ev, option, index, value) => { const newVal = option ? option.key : value; onRowChange(item.__index, 'Brands', newVal || ''); }} />
                          </div>
                        </div>
                      );
                    }
                  },
                  {
                    key: 'columnDetails', name: 'Specific Details', fieldName: 'Details', minWidth: 220, isResizable: true, onRender: (item: any) => {
                      const itemTitleRaw = item && item.Item ? String(item.Item).trim() : '';
                      if (itemTitleRaw === 'Others') {
                        const purposeVal = item && item.Purpose ? item.Purpose : '';
                        return (<div className={styles.comboCell}><TextField value={purposeVal} onChange={(ev, val) => onRowChange(item.__index, 'Purpose', val || '')} /></div>);
                      }
                      if (!itemTitleRaw) return <div className={`${styles.tableSecondaryBg} ${styles.detailsCell}`}><small className="text-muted">Select Item first</small></div>;
                      const options = (detailsByItem[itemTitleRaw.toLowerCase()] || []).map(t => ({ key: t, text: t }));
                      if (!options.length) return <div className={`${styles.tableSecondaryBg} ${styles.detailsCell}`}><small className="text-muted">No details</small></div>;
                      const selected = Array.isArray(item.Details) && item.Details.length ? item.Details[0] : undefined;
                      return (
                        <div className={styles.comboCell}>
                          <ComboBox
                            placeholder="Select detail"
                            allowFreeform
                            autoComplete="on"
                            selectedKey={selected}
                            options={options}
                            onChange={(ev, option, _i, value) => {
                              const val = option ? String(option.key) : (value || '');
                              handleDetailChange(item.__index, val);
                            }} />
                        </div>
                      );
                    }
                  },
                  { key: 'columnQty', name: 'Qty', fieldName: 'Qty', minWidth: 70, maxWidth: 90, isResizable: false, onRender: (item: any) => <div className={`${styles.tableSecondaryBg} text-center align-middle`}><TextField value={item.Qty} onChange={(ev, val) => onRowChange(item.__index, 'Qty', val || '')} underlined={true} /></div> },
                  {
                    key: 'columnSize', name: 'Size', fieldName: 'Size', minWidth: 140, maxWidth: 200, isResizable: true, onRender: (item: any) => {
                      const itemTitle = item && item.Item ? String(item.Item).trim() : '';
                      if (itemTitle === 'Others') {
                        const sizeVal = item && item.Size ? item.Size : '';
                        return <div className={styles.comboCell}><TextField value={sizeVal} onChange={(ev, val) => onRowChange(item.__index, 'Size', val || '')} /></div>;
                      }
                      const selectedDetail: string | undefined = Array.isArray(item.Details) && item.Details.length ? String(item.Details[0]).trim() : undefined;
                      if (!selectedDetail) return <div className={`${styles.tableSecondaryBg} ${styles.detailsCell}`}><small className="text-muted">Select detail first</small></div>;
                      const key = `${itemTitle.toLowerCase()}||${selectedDetail.toLowerCase()}`;
                      const sizeOptions = sizesByItemDetail[key] || [];
                      if (!sizeOptions.length) return <div className={styles.tableSecondaryBg}><small className="text-muted">No sizes</small></div>;
                      const options = sizeOptions.map(s => ({ key: s, text: s }));
                      const selected = Array.isArray(item.SizesSelected) && item.SizesSelected.length ? item.SizesSelected[0] : undefined;
                      return (
                        <div className={styles.comboCell}>
                          <ComboBox
                            placeholder="Select size"
                            allowFreeform
                            autoComplete="on"
                            selectedKey={selected}
                            options={options}
                            onChange={(ev, option, _i, value) => {
                              const val = option ? String(option.key) : (value || '');
                              handleSizeChange(item.__index, val);
                            }} />
                        </div>
                      );
                    }
                  },
                ];

                return (
                  <>
                    <DetailsList items={items} columns={columns} selection={selectionRef.current} selectionMode={SelectionMode.single} setKey="ppeItemsList" layoutMode={DetailsListLayoutMode.fixedColumns} isHeaderVisible={true} className={styles.detailsListHeaderCenter} />
                  </>
                );
              })()}
            </div>
          </div>
        </Stack>

        <Separator />
        {/* Instructions For Use */}
        <Stack horizontal styles={stackStyles}>
          {itemInstructionsForUse && itemInstructionsForUse.length > 0 && (
            <div style={{ marginTop: 12 }}>
              <Label>Instructions for Use:</Label>
              {itemInstructionsForUse.map((instr: ILKPItemInstructionsForUse, idx: number) => (
                <MessageBar key={instr.Id ?? instr.Order} isMultiline styles={{ root: { marginBottom: 6 } }}>
                  <strong>{`${idx + 1}. `}</strong>
                  {instr.Description}
                </MessageBar>
              ))}
            </div>
          )}
        </Stack>

        <Separator />
        {/* Approvals sign-off table */}
        <Stack horizontal styles={stackStyles} className="mt-3 mb-3">
          <div style={{ marginTop: 18 }}>
            <Label>Approvals / Sign-off</Label>
            <DetailsList
              items={approvalsRows}
              columns={[
                { key: 'colSignOff', name: 'Sign off', fieldName: 'SignOff', minWidth: 160, isResizable: true },
                {
                  key: 'colName', name: 'Name', fieldName: 'Name', minWidth: 260, isResizable: true, onRender: (item: any) => (
                    <div style={{ minWidth: 220 }}>
                      <NormalPeoplePicker
                        itemLimit={1}
                        onResolveSuggestions={onFilterChanged}
                        onChange={(items: IPersonaProps[] | undefined) => {
                          const sel = items && items.length ? items[0] : undefined;
                          // store display name and id (if available) in the Name field
                          onApprovalChange(item.__index, 'Name', sel ? (sel.text || '') : '');
                          // also store an id field for potential persistence
                          onApprovalChange(item.__index, 'NameId', sel ? sel.id : undefined);
                        }}
                        selectedItems={item.Name ? [{ text: item.Name, id: item.NameId }] : []}
                        resolveDelay={300}
                        inputProps={{ 'aria-label': 'Approvee' }}
                      />
                    </div>
                  )
                },
                { key: 'colStatus', name: 'Status', fieldName: 'Status', minWidth: 220, isResizable: true, onRender: (item: any) => (<TextField value={item.Signature || ''} onChange={(ev, val) => onApprovalChange(item.__index, 'Signature', val || '')} />) },
                { key: 'colReason', name: 'Reason', fieldName: 'Reason', minWidth: 220, isResizable: true, onRender: (item: any) => (<TextField value={item.Reason || ''} onChange={(ev, val) => onApprovalChange(item.__index, 'Reason', val || '')} />) },
                { key: 'colDate', name: 'Date', fieldName: 'Date', minWidth: 140, isResizable: true, onRender: (item: any) => (<DatePicker value={item.Date ? new Date(item.Date) : undefined} onSelectDate={(date) => onApprovalChange(item.__index, 'Date', date)} strings={defaultDatePickerStrings} />) }
              ]}
              selectionMode={SelectionMode.none}
              setKey="approvalsList"
              layoutMode={DetailsListLayoutMode.fixedColumns}
            />
          </div>
        </Stack>

      </form>
    </div>
  );
}
