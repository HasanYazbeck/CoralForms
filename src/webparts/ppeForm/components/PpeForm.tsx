import * as React from "react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { IGraphResponse, IGraphUserResponse, ILKPItemInstructionsForUse } from "../../../Interfaces/ICommon";

// Components
import { ComboBox, DefaultPalette, DetailsListLayoutMode } from "@fluentui/react";
import type { IPpeFormWebPartProps } from "./IPpeFormProps";
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';
import { TextField } from '@fluentui/react/lib/TextField';
import { Stack, IStackStyles } from '@fluentui/react/lib/Stack';
import { DetailsList, SelectionMode } from '@fluentui/react';
import { DatePicker, mergeStyleSets, defaultDatePickerStrings } from '@fluentui/react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Label } from '@fluentui/react/lib/Label';
import { Checkbox } from '@fluentui/react';
import { Separator } from '@fluentui/react/lib/Separator';
import { MessageBar } from '@fluentui/react/lib/MessageBar';
import { PrimaryButton, DefaultButton } from '@fluentui/react';

// Styles
import "bootstrap/dist/css/bootstrap.min.css";
import styles from "./PpeForm.module.scss";

// Classes
import { SPCrudOperations } from "../../../Classes/SPCrudOperations";
import { SPHelpers } from "../../../Classes/SPHelpers";
import { ICoralFormsList } from "../../../Interfaces/ICoralFormsList";
import { IUser } from "../../../Interfaces/IUser";
import { IPPEItemDetails } from "../../../Interfaces/IPPEItemDetails";
import { IEmployeeProps, IEmployeesPPEItemsCriteria } from "../../../Interfaces/IEmployeeProps";
import { IFormsApprovalWorkflow } from "../../../Interfaces/IFormsApprovalWorkflow";
import { IPPEItem } from "../../../Interfaces/IPPEItem";

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

export default function PpeForm(props: IPpeFormWebPartProps) {
  // Helpers and refs
  const formName = "PERSONAL PROTECTIVE EQUIPMENT";
  const spHelpers = useMemo(() => new SPHelpers(), []);
  const spCrudRef = useRef<SPCrudOperations | undefined>(undefined);
  const [jobTitle, setJobTitle] = useState("");
  const [department, setDepartment] = useState("");
  const [division, setDivision] = useState("");
  const [company, setCompany] = useState("");
  const [_employee, setEmployee] = useState<IPersonaProps[]>([]);
  const [_employeeId, setEmployeeId] = useState<number | undefined>(undefined);
  const [submitter, setSubmitter] = useState<IPersonaProps[]>([]);
  const [requester, setRequester] = useState<IPersonaProps[]>([]);
  const [isReplacementChecked, setIsReplacementChecked] = useState(false);
  const containerRef = React.useRef<HTMLDivElement>(null);
  const [users, setUsers] = useState<IUser[]>([]);
  const [employees, setEmployees] = useState<IEmployeeProps[]>([]);
  const [employeePPEItemsCriteria, setEmployeePPEItemsCriteria] = useState<IEmployeesPPEItemsCriteria>({ Id: '' });
  const [ppeItems, setPpeItems] = useState<IPPEItem[]>([]);
  const [ppeItemDetails, setPpeItemDetails] = useState<IPPEItemDetails[]>([]);
  const [itemInstructionsForUse, setItemInstructionsForUse] = useState<ILKPItemInstructionsForUse[]>([]);
  const [, setFormsApprovalWorkflow] = useState<IFormsApprovalWorkflow[]>([]);
  const [, setCoralFormsList] = useState<ICoralFormsList>({ Id: "" });
  const [loading, setLoading] = useState<boolean>(true);
  const [reason, setReason] = useState<string>('');           // capture Reason text
  const [isSaving, setIsSaving] = useState<boolean>(false);    // Save button state
  const [isSubmitting, setIsSubmitting] = useState<boolean>(false); // Submit button state
  const [bannerText, setBannerText] = useState<string>();

  // Aggregated rows (one per unique Item) replacing manual add/remove paradigm
  interface ItemRowState {
    item: string;
    order?: number;             // original order for sorting
    brands: string[];            // all available brands for item
    brandSelected?: string;      // chosen brand
    required: boolean | undefined;           // required flag per item
    qty?: string;                // overall quantity (if applies per item)
    details: string[];           // all available detail titles for this item
    selectedDetail?: string;   // checked details
    itemSizes: string[];         // available sizes at item-level
    itemSizeSelected?: string;   // chosen size for the item
    othersItemdetailsText?: Record<string, string>; // Added: holds free-text per detail for "Others"
  }
  const [itemRows, setItemRows] = useState<ItemRowState[]>([]);
  // Approvals sign-off rows (Department, HR, HSE, Warehouse)
  const [approvalsRows, setApprovalsRows] = useState<Array<any>>([
    { SignOff: 'Department Approval', Name: '', Status: '', Reason: '', Date: undefined, __index: 0 },
    { SignOff: 'HR Approval', Name: '', Status: '', Reason: '', Date: undefined, __index: 1 },
    { SignOff: 'HSE Approval', Name: '', Status: '', Reason: '', Date: undefined, __index: 2 },
    { SignOff: 'Warehouse Approval', Name: '', Status: '', Reason: '', Date: undefined, __index: 3 }
  ]);

  const formPayload = useCallback((status: 'Draft' | 'Submitted') => {
    return {
      formName,
      status,
      employeeId: _employeeId,
      employeeName: _employee?.[0]?.text,
      jobTitle,
      department,
      division,
      company,
      requestType: isReplacementChecked ? 'Replacement' : 'New',
      reason: isReplacementChecked ? reason : '',
      items: itemRows.map(r => ({
        item: r.item,
        required: !!r.required,
        brand: r.brandSelected,
        qty: r.qty ? Number(r.qty) : undefined,
        size: r.itemSizeSelected,
        selectedDetails: r.selectedDetail,
        othersText: r.item.toLowerCase() === 'others' ? r.othersItemdetailsText : undefined
      })),
      approvals: approvalsRows
    };
  }, [_employee, _employeeId, jobTitle, department, division, company, isReplacementChecked, reason, itemRows, approvalsRows, formName]);


  const validateBeforeSubmit = useCallback((): string | undefined => {
    // If “Others” is required, ensure a size is chosen (since you show size ComboBox when required)
    const othersMissingSize = itemRows.some(r =>
      r.item.toLowerCase() === 'others' && r.required && (!r.itemSizeSelected || !r.itemSizeSelected.trim())
    );
    if (othersMissingSize) return 'Please choose a size for "Others" since it is marked Required.';

    // Example: ensure at least one item is required or has any selection (tweak as you need)
    const anySelection = itemRows.some(r =>
      r.required || r.brandSelected || r.qty || r.itemSizeSelected || (r.selectedDetail)
    );
    if (!anySelection) return 'Please select at least one item or mark one as Required.';

    // Example: if Replacement, require a reason
    if (isReplacementChecked && !reason.trim()) return 'Please provide a reason for Replacement.';

    return undefined;
  }, [itemRows, isReplacementChecked, reason]);


  const handleSave = useCallback(async () => {
    try {
      setBannerText(undefined);
      setIsSaving(true);
      const payload = formPayload('Draft');
      // TODO: Wire to SharePoint persistence here.
      // console.log as a placeholder so you can see the shape:
      console.log('Save payload (Draft):', payload);
      setBannerText('Draft saved (demo). Hook this up to SharePoint to persist.');
    } catch (e) {
      console.error(e);
      setBannerText('Failed to save draft.');
    } finally {
      setIsSaving(false);
    }
  }, [formPayload]);

  const handleSubmit = useCallback(async () => {
    try {
      setBannerText(undefined);
      const validationError = validateBeforeSubmit();
      if (validationError) {
        setBannerText(validationError);
        return;
      }

      setIsSubmitting(true);
      const payload = formPayload('Submitted');
      // TODO: Wire to SharePoint persistence and/or workflow trigger here.
      console.log('Submit payload:', payload);
      setBannerText('Form submitted (demo). Hook this up to SharePoint to persist/trigger workflow.');
    } catch (e) {
      console.error(e);
      setBannerText('Failed to submit.');
    } finally {
      setIsSubmitting(false);
    }
  }, [formPayload, validateBeforeSubmit]);
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
          endpoint = (response as any)["@odata.nextLink"] || null;
        } else {
          endpoint = null;
        }
      } while (endpoint);
      if (fetched.length > 0) setUsers(fetched);
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
          rainSuit: obj.RainSuit !== undefined && obj.RainSuit !== null ? { id: obj.RainSuit.Id, label: obj.RainSuit.DisplayText } : undefined,
          uniformCoveralls: obj.UniformCoveralls !== undefined && obj.UniformCoveralls !== null ? { id: obj.UniformCoveralls.Id, label: obj.UniformCoveralls.DisplayText } : undefined,
          uniformTop: obj.UniformTop !== undefined && obj.UniformTop !== null ? { id: obj.UniformTop.Id, label: obj.UniformTop.DisplayText } : undefined,
          uniformPants: obj.UniformPants !== undefined && obj.UniformPants !== null ? { id: obj.UniformPants.Id, label: obj.UniformPants.DisplayText } : undefined,
          winterJacket: obj.WinterJacket !== undefined && obj.WinterJacket !== null ? { id: obj.WinterJacket.Id, label: obj.WinterJacket.DisplayText } : undefined,
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
      const query: string = `?$select=Id,Title,Brands,Order,Created&$orderby=Order asc`;
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
            Order: obj.Order !== undefined && obj.Order !== null ? obj.Order : undefined,
            Brands: spHelpers.NormalizeToStringArray(obj.Brands),
            PPEItemsDetails: []
          };
          result.push(temp);
        }
      });

      const items = result.sort((a: any, b: any) => a.Order - b.Order);
      // Map a sequence 1, 2, 3 instead of 100, 200, 300
      const normalizedItems = items.map((item: any, index: number) => ({
        ...item,
        Order: index + 1 // This will be 1,2,3
      }));


      // console.log("PPE Item:", result);
      setPpeItems(normalizedItems);
    } catch (error) {
      console.error('An error has occurred while retrieving items!', error);
      setPpeItems([]);
    }
  }, [props.context, spHelpers]);

  const _getPPEItemsDetails = useCallback(async (usersArg?: IUser[]) => {
    try {
      const query: string = `?$select=Id,Title,PPEItem,Sizes,Created,PPEItem/Id,PPEItem/Title&$expand=PPEItem`;
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
            Sizes: spHelpers.NormalizeToStringArray(obj.Sizes),
            PPEItem: obj.PPEItem !== undefined ? {
              Id: obj.PPEItem.Id !== undefined && obj.PPEItem.Id !== null ? obj.PPEItem.Id : undefined,
              Title: obj.PPEItem.Title !== undefined && obj.PPEItem.Title !== null ? obj.PPEItem.Title : undefined,
              Order: obj.PPEItem.Order !== undefined && obj.PPEItem.Order !== null ? obj.PPEItem.Order : undefined,
              Brands: spHelpers.NormalizeToStringArray(obj.PPEItem.Brands),
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
      const query: string = `?$select=Id,FormName/Id,FormName/Title,Order,Created,DepartmentName,Manager&$expand=FormName,ManagerName($select=Id,FullName)&$filter=substringof('${formName}', FormName)&$orderby=Order asc`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'd084f344-63cb-4426-ae51-d7f875f3f99a', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: IFormsApprovalWorkflow[] = [];
      const usersToUse = usersArg && usersArg.length ? usersArg : users;
      data.forEach((obj: any) => {
        if (obj) {
          const createdBy = usersToUse && usersToUse.length ? usersToUse.filter(u => u.id.toString() === obj.AuthorId?.toString())[0] : undefined;
          let created: Date | undefined;
          if (obj.Created !== undefined) created = new Date(spHelpers.adjustDateForGMTOffset(obj.Created));
          const temp: IFormsApprovalWorkflow = {
            Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            FormName: obj.FormName !== undefined && obj.FormName !== null ? obj.FormName : undefined,
            Order: obj.Order !== undefined && obj.Order !== null ? obj.Order : undefined,
            DepartmentName: obj.DepartmentName !== undefined && obj.DepartmentName !== null ? obj.DepartmentName : undefined,
            Manager: obj.Manager !== undefined && obj.Manager !== null ? obj.Manager : undefined,
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
      // Fetch base PPE Items first (brands) then details (sizes & detail rows)
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
  // Map of Item Title -> Brands[] (deduped) from base items
  const brandsMap = useMemo(() => {
    const map: Record<string, { order: number; brands: string[] }> = {};
    (ppeItems || []).forEach((pi: any) => {
      const title = pi?.Title ? String(pi.Title).trim() : undefined;
      const brandsArr = spHelpers.NormalizeToStringArray(pi?.Brands) || [];
      const order = pi?.Order ?? 0;

      if (title) {
        if (!map[title]) map[title] = { order, brands: [] };
        // Ensure unique brands
        map[title].brands = Array.from(new Set(map[title].brands.concat(brandsArr)));
        // Keep the order updated if needed
        map[title].order = order;
      }
    });
    const sortedBrands = Object.keys(map).sort((a, b) => map[a].order - map[b].order).map(key => ({ key, ...map[key] }));

    return sortedBrands;
  }, [ppeItems, spHelpers.NormalizeToStringArray]);

  // Map of Item Title -> Sizes[] (deduped) from details
  // const sizesMap = useMemo(() => {
  //   const map: Record<string, string[]> = {};
  //   (ppeItemDetails || []).forEach((p: any) => {
  //     const title = p && p.PPEItem && p.PPEItem.Title ? String(p.PPEItem.Title).trim() : (p && p.Title ? String(p.Title).trim() : undefined);
  //     const sizesArr = spHelpers.NormalizeToStringArray(p && p.Sizes ? p.Sizes : undefined) || [];
  //     if (title) {
  //       if (!map[title]) map[title] = [];
  //       map[title] = Array.from(new Set(map[title].concat(sizesArr)));
  //     }
  //   });
  //   return map;
  // }, [ppeItemDetails, spHelpers.NormalizeToStringArray]);

  const ppeItemMap = useMemo(() => {
    // Create a map from item title to details
    const map: { [title: string]: IPPEItemDetails[] } = {};

    (ppeItemDetails || []).forEach((detail: IPPEItemDetails) => {
      const title = detail?.PPEItem?.Title ? String(detail.PPEItem.Title).trim() : undefined;
      if (!title) return;

      if (!map[title]) {
        map[title] = [];
      }
      map[title].push(detail);
    });

    // Now fill each ppeItem with its details
    return (ppeItems || []).map(item => {
      const title = item.Title ? String(item.Title).trim() : "";
      return {
        ...item,
        Brands: brandsMap.find(b => b.key === title)?.brands || [],
        PPEItemsDetails: map[title] || []  // fill with matching details or empty array
      };
    });
  }, [ppeItems, ppeItemDetails, brandsMap]);

  useEffect(() => {
    if (!ppeItemMap || !ppeItemMap.length) return;

    const rows: ItemRowState[] = ppeItemMap.map(item => ({
      item: item.Title || "",
      order: item.Order ?? undefined,            // comes directly from ppeItems
      brands: item.Brands || [],                  // brands set in filledPpeItems
      brandSelected: undefined,                   // default or pre-selected brand if needed
      required: undefined,                             // or use a flag if available in data
      qty: undefined,                             // set later if needed
      details: (item.PPEItemsDetails || []).map(d => d.Title || ""),  // all detail titles
      selectedDetail: "",                        // empty initially
      itemSizes: (item.PPEItemsDetails || []).map(d => d.Sizes || []).reduce((acc, val) => acc.concat(val), []),
      itemSizeSelected: undefined,                // default if needed
      othersItemdetailsText: {},                  // empty initially
    }));

    setItemRows(rows);
  }, [ppeItemMap]);

  // Apply employee PPE criteria to pre-select details (assumption: label matches detail title)
  useEffect(() => {
    if (!employeePPEItemsCriteria || !employeePPEItemsCriteria.employeeID) return;
    const map: Record<string, string> = {};

    Object.entries(employeePPEItemsCriteria).forEach(([key, value]) => {
      const itemDetail = spHelpers.CamelString(key.split(/(?=[A-Z])/).join(" "));
      const itemValue = value || "";

      // const itemTitle = ppeItemMap.find(i => i.Title?.toLowerCase() === itemDetail.toLowerCase())?.Title;

      switch (itemDetail) {
        case "Reflective Vest":
          console.log(itemDetail + " - " + itemValue);
          break;

        case "Safety Helmet":
          console.log(itemDetail + " - " + itemValue);
          break;

        case "Safety Shoes":
          console.log(itemDetail + " - " + itemValue);
          break;

        case "Rain Suit":
          console.log(itemDetail + " - " + itemValue);
          break;

        case "Winter Jacket":
          console.log(itemDetail + " - " + itemValue);
          break;

        case "Uniform Coveralls":
          console.log(itemDetail + " - " + itemValue);
          break;

        case "Uniform Top":
          console.log(itemDetail + " - " + itemValue);
          break;

        case "Uniform Pants":
          console.log(itemDetail + " - " + itemValue);
          break;

        default:
          break;
      }


      if (itemDetail && spHelpers.CamelString(itemDetail) === ppeItemMap.find(i => i.Title?.toLowerCase() === itemDetail.toLowerCase())?.Title) {
        map[key] = itemValue;
      }
    });

    const labelFields: (string | undefined)[] = [
      employeePPEItemsCriteria.rainSuit?.label,
      employeePPEItemsCriteria.uniformCoveralls?.label,
      employeePPEItemsCriteria.uniformTop?.label,
      employeePPEItemsCriteria.uniformPants?.label,
      employeePPEItemsCriteria.winterJacket?.label
    ].filter(Boolean);

    if (!labelFields.length) return;

    setItemRows(prev => prev.map(r => {
      const matched = r.details.filter(d => labelFields.some(l => l && l.toLowerCase() === d.toLowerCase()));
      if (!matched.length) return r;
      return { ...r, selectedDetails: r.selectedDetail };
    }));

  }, [employeePPEItemsCriteria]);

  const toggleItemDetail = useCallback((rowIndex: number, detail: string, checked?: boolean) => {
    setItemRows(prev =>
      prev.map((r, idx) => {
        if (idx !== rowIndex) return r;
        if (!detail) return r;

        if (typeof checked === 'boolean') {
          return {
            ...r,
            selectedDetail: checked
              ? detail
              : (r.selectedDetail === detail ? undefined : r.selectedDetail),
          };
        }

        // Fallback (no 'checked' provided): toggle if the same size was clicked
        return {
          ...r,
          selectedDetail: r.selectedDetail === detail ? undefined : detail,
        };
      })
    );
  }, []);

  const toggleBrand = useCallback((rowIndex: number, brandVal?: string, checked?: boolean) => {
    setItemRows(prev =>
      prev.map((r, idx) => {
        if (idx !== rowIndex) return r;
        if (!brandVal) return r;

        if (typeof checked === 'boolean') {
          return {
            ...r,
            brandSelected: checked
              ? brandVal
              : (r.brandSelected === brandVal ? undefined : r.brandSelected),
          };
        }

        // Fallback (no 'checked' provided): toggle if the same size was clicked
        return {
          ...r,
          brandSelected: r.brandSelected === brandVal ? undefined : brandVal,
        };
      })
    );
  }, []);

  const toggleRequired = useCallback((rowIndex: number, checked?: boolean) => {
    setItemRows(prev => prev.map((r, i) => i === rowIndex ? { ...r, required: !!checked } : r));
  }, []);

  const toggleSize = useCallback((rowIndex: number, sizeVal?: string, checked?: boolean) => {
    setItemRows(prev =>
      prev.map((r, idx) => {
        if (idx !== rowIndex) return r;
        if (!sizeVal) return r;

        if (typeof checked === 'boolean') {
          return {
            ...r,
            itemSizeSelected: checked
              ? sizeVal
              : (r.itemSizeSelected === sizeVal ? undefined : r.itemSizeSelected),
          };
        }

        // Fallback (no 'checked' provided): toggle if the same size was clicked
        return {
          ...r,
          itemSizeSelected: r.itemSizeSelected === sizeVal ? undefined : sizeVal,
        };
      })
    );
  }, []);

  // const updateOtherDetailText = useCallback((rowIndex: number, detail?: string, checked?: boolean) => {
  //   setItemRows(prev =>
  //     prev.map((r, idx) => {
  //       if (idx !== rowIndex) return r;
  //       if (!detail) return r;

  //       if (typeof checked === 'boolean') {
  //         return {
  //           ...r,
  //           itemDetail: checked
  //             ? detail
  //             : (r.selectedDetail === detail ? undefined : r.selectedDetail),
  //         };
  //       }

  //       // Fallback (no 'checked' provided): toggle if the same size was clicked
  //       return {
  //         ...r,
  //         itemSizeSelected: r.itemSizeSelected === detail ? undefined : detail,
  //       };
  //     })
  //   );
  // }, []);

  const updateItemQty = useCallback((rowIndex: number, qty?: string) => {
    setItemRows(prev => prev.map((r, i) => i === rowIndex ? { ...r, qty: qty } : r));
  }, []);

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

  // Removed per-detail sizes map (sizes now at item level)

  // ---------------------------
  // Handlers
  // ---------------------------


  // Ensure brandSelected & itemSizeSelected always within available options
  // useEffect(() => {
  //   setItemRows(prev => prev.map(r => {
  //     const available = brandsMap.find(b => b.key.toLowerCase() === r.item.toLowerCase())?.brands;
  //     let brandSelected = r.brandSelected;
  //     if (!available) brandSelected = undefined; else if (!brandSelected || available.indexOf(brandSelected) === -1) brandSelected = available.length === 1 ? available[0] : undefined;
  //     const itemSizes = sizesMap[r.item] || r.itemSizes || [];
  //     let itemSizeSelected = r.itemSizeSelected;
  //     if (!itemSizes.length) itemSizeSelected = undefined; else if (!itemSizeSelected || itemSizes.indexOf(itemSizeSelected) === -1) itemSizeSelected = itemSizes.length === 1 ? itemSizes[0] : undefined;
  //     return { ...r, brandSelected, itemSizes, itemSizeSelected };
  //   }));
  // }, [brandsMap, sizesMap]);

  const handleEmployeeChange = useCallback(async (items?: IPersonaProps[], selectedOption?: string) => {

    if (items && items.length > 0) {
      const selected = items[0];
      // First try to find in employees list by FullName (fullName -> persona.text)
      const emp = employees.find(e => (e.fullName || '').toLowerCase() === (selected?.text || '').toLowerCase());
      // Fallback to users (Graph) if not found
      const user = users.find(u => u.displayName?.toLowerCase() === (selected?.text || '').toLowerCase() || u.id === selected?.id);
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
        await _getEmployeesPPEItemsCriteria(users, selected?.tertiaryText ? String(selected.tertiaryText) : '');

        if (employeePPEItemsCriteria && employeePPEItemsCriteria.employeeID !== selected?.tertiaryText) {
          setItemRows(prev => prev.map(r => ({
            ...r,
            brandSelected: undefined,
            itemSizeSelected: undefined,
            qty: undefined,
            required: undefined,
            selectedDetails: [],            // added: clear details too
            othersItemdetailsText: {}
          })));
        }

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
      setEmployeePPEItemsCriteria({ Id: '' });
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
        .map(e => ({ text: e.fullName || '', secondaryText: e.jobTitle?.title, id: (e.Id ? String(e.Id) : e.fullName), tertiaryText: (e.employeeID ? String(e.employeeID) : e.employeeID) }) as IPersonaProps);
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
      {bannerText && <MessageBar styles={{ root: { marginBottom: 8 } }}>{bannerText}</MessageBar>}
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
                onChange={(items) => {
                  const selectedText = items?.[0]?.text || '';
                  const empId = employees.find(e => (e.fullName || '').toLowerCase() === selectedText.toLowerCase())?.employeeID;
                  return handleEmployeeChange(items, empId ? String(empId) : undefined);
                }}
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

              <TextField placeholder="Reason" disabled={!isReplacementChecked} value={reason}
                onChange={(_e, v) => setReason(v || '')} />
            </div>
          </div>
        </Stack>

        <Separator />

        <div className="mb-2 text-center">
          <small className="text-muted" style={{ fontStyle: 'italic', fontSize: '1.0rem' }}>Please complete the table below in the blank spaces; grey spaces are for administrative use only.</small>
        </div>

        <Separator />
        {/* Aggregated PPE Items Grid with detail checkboxes */}
        <Stack horizontal styles={stackStyles}>
          <div className="row">
            <div className="form-group col-md-12">
              <DetailsList
                items={itemRows}
                setKey="ppeAggregatedItemsList"
                selectionMode={SelectionMode.none}
                layoutMode={DetailsListLayoutMode.fixedColumns}
                columns={[
                  {
                    key: 'colItem', name: 'Item', fieldName: 'item', minWidth: 60, isResizable: true,
                    onRender: (r: ItemRowState) => <span style={{
                      display: 'block', whiteSpace: 'normal',
                      wordWrap: 'break-word', overflowWrap: 'anywhere', lineHeight: 1.3
                    }}>{r.item}</span>
                  },
                  {
                    key: 'colRequired', name: 'Required', fieldName: 'required', minWidth: 50, maxWidth: 70,
                    onRender: (r: ItemRowState) =>
                      <Checkbox checked={r.required} ariaLabel="Required" id={r.item}
                        onChange={(_e, ch) => toggleRequired(itemRows.indexOf(r), ch)}
                        styles={{ root: { display: 'flex', justifyContent: 'center', alignItems: 'center', width: '100%' } }} />
                  },
                  {
                    key: 'colBrand', name: 'Brand', fieldName: 'brand', minWidth: 140, isResizable: false,
                    onRender: (r: ItemRowState) => {
                      return (
                        <>
                          {r.brands.length === 0 && <span>N/A</span>}
                          {
                            r.brands.map(brand => {
                              const brandChecked = r.brandSelected === brand;
                              return (
                                <div key={brand} style={{ display: 'flex', alignItems: 'center', marginBottom: 4 }}>
                                  <Checkbox label={brand} checked={brandChecked}
                                    onChange={(_e, ch) => toggleBrand(itemRows.indexOf(r), brand, !!ch)}
                                    styles={{
                                      root: { alignItems: 'flex-start' }, // top-align text if wrapped
                                      label: { whiteSpace: 'normal', wordWrap: 'break-word', overflowWrap: 'anywhere', lineHeight: '1.3' }
                                    }}
                                  />
                                </div>
                              );
                            })
                          }
                        </>
                      );
                    }
                  },
                  {
                    key: 'colDetails', name: 'Specific Detail', fieldName: 'itemDetails', minWidth: 230, isResizable: true, onRender: (r: ItemRowState) => (
                      <div>
                        {r.details.map(detail => {
                          // ...inside the onRender of colDetails...
                          {
                            const itemLabel = r.item.toLowerCase() === 'others';
                            if (itemLabel) {
                              return (
                                <div key={detail} style={{ display: 'flex', flexDirection: 'column', marginBottom: 8 }}>

                                  <TextField placeholder={detail} multiline autoAdjustHeight
                                    scrollContainerRef={containerRef} styles={{ root: { width: '100%' } }}
                                  // value={r.othersItemdetailsText?.[detail] || ''}
                                  // eslint-disable-next-line react/jsx-no-bind
                                  // onChange={(_e, ch) => updateOtherDetailText(itemRows.indexOf(r), detail, !!ch)}
                                  />
                                </div>
                              );
                            }
                            // Special case: Winter Jacket - no checkboxes, just show the label (detail)
                            if (r.item.toLowerCase() === 'winter jacket') return (<Label>{detail || ''}</Label>)

                            const checked = r.selectedDetail === detail;
                            return (
                              <div key={detail} style={{ display: 'flex', alignItems: 'center', marginBottom: 4 }}>
                                <Checkbox
                                  label={detail}
                                  checked={checked}
                                  onChange={(_e, ch) => toggleItemDetail(itemRows.indexOf(r), detail, !!ch)}
                                  styles={{
                                    root: { alignItems: 'flex-start' }, // top-align text if wrapped
                                    label: { whiteSpace: 'normal', wordWrap: 'break-word', overflowWrap: 'anywhere', lineHeight: '1.3' }
                                  }}
                                />
                              </div>
                            );
                          }
                        })
                        }
                      </div>
                    )
                  },
                  {
                    key: 'colQty', name: 'Qty', fieldName: 'qty', minWidth: 30, maxWidth: 40, onRender: (r: ItemRowState) => (
                      <TextField value={r.qty || ''} type='number'
                        onChange={(_e, v) => updateItemQty(itemRows.indexOf(r), v || '')} min={0} max={99}
                        styles={{
                          root: { display: 'flex', justifyContent: 'center', alignItems: 'center', width: '100%' },
                          field: {
                            '&::-webkit-outer-spin-button': { WebkitAppearance: 'none', margin: 0 },
                            '&::-webkit-inner-spin-button': { WebkitAppearance: 'none', margin: 0 },
                            // Remove arrows in Chrome, Edge, Safari
                            MozAppearance: 'textfield',
                            appearance: 'textfield',
                          }
                        }}
                      />
                    )
                  },
                  {
                    key: 'colSizes', name: 'Size', fieldName: 'size', minWidth: 140, isResizable: true,
                    onRender: (r: ItemRowState) => {
                      if (r.item.toLowerCase() === 'others') {
                        // Show Sizes only if Required is checked
                        if (!r.required) return <span />;

                        const sizes = Array.from(new Set((r.itemSizes || []).map(s => String(s).trim()).filter(Boolean))).sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
                        return (
                          <div key={r.item} style={{ display: 'flex', alignItems: 'center', marginBottom: 4 }}>
                            <ComboBox placeholder={sizes.length ? 'Size' : 'No sizes'}
                              selectedKey={r.itemSizeSelected || undefined}
                              options={sizes.map(s => ({ key: s, text: s }))}
                              styles={{ root: { width: 140 } }}
                              disabled={!sizes.length}
                            />
                          </div>
                        );
                      }
                      else {
                        return (<>
                          {(() => {
                            const sizes = Array.from(new Set((r.itemSizes || []).map(s => String(s).trim()).filter(Boolean)))
                              .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));

                            if (!sizes.length) return <span>N/A</span>;

                            const cols = sizes.length > 12 ? 2 : (sizes.length > 6 ? 2 : 1);
                            return (
                              <div style={{ display: 'grid', gridTemplateColumns: `repeat(${cols}, minmax(0, 1fr))`, gap: 4 }}>
                                {sizes.map(size => {
                                  const sizeChecked = r.itemSizeSelected === size;
                                  return (
                                    <div key={size} style={{ display: 'flex', alignItems: 'center' }}>
                                      <Checkbox label={size} checked={sizeChecked} onChange={(_e, _ch) => toggleSize(itemRows.indexOf(r), size)}
                                        styles={{
                                          root: { alignItems: 'flex-start' }, label: { whiteSpace: 'normal', wordWrap: 'break-word', overflowWrap: 'anywhere', lineHeight: '1.3' }
                                        }}
                                      />
                                    </div>
                                  );
                                })}
                              </div>
                            );
                          })()
                          }
                        </>
                        );
                      }
                    }
                  }
                ]}
                isHeaderVisible={true}
                className={styles.detailsListHeaderCenter}
              />
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
                { key: 'colSignOff', name: 'Sign off', fieldName: 'SignOff', minWidth: 120, isResizable: true },
                {
                  key: 'colName', name: 'Name', fieldName: 'Name', minWidth: 180, isResizable: true, onRender: (item: any) => (
                    <div style={{ minWidth: 130 }}>
                      <NormalPeoplePicker
                        itemLimit={1}
                        required={true}
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
                { key: 'colStatus', name: 'Status', fieldName: 'Status', minWidth: 120, isResizable: true, onRender: (item: any) => (<TextField value={item.Signature || ''} onChange={(ev, val) => onApprovalChange(item.__index, 'Signature', val || '')} />) },
                { key: 'colReason', name: 'Reason', fieldName: 'Reason', minWidth: 160, isResizable: true, onRender: (item: any) => (<TextField value={item.Reason || ''} onChange={(ev, val) => onApprovalChange(item.__index, 'Reason', val || '')} />) },
                { key: 'colDate', name: 'Date', fieldName: 'Date', minWidth: 120, isResizable: true, onRender: (item: any) => (<DatePicker value={item.Date ? new Date(item.Date) : undefined} onSelectDate={(date) => onApprovalChange(item.__index, 'Date', date)} strings={defaultDatePickerStrings} />) }
              ]}
              selectionMode={SelectionMode.none}
              setKey="approvalsList"
              layoutMode={DetailsListLayoutMode.fixedColumns}
            />
          </div>
        </Stack>
        <Separator />

        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 8 }}>
          <DefaultButton
            text={isSaving ? 'Saving…' : 'Save as Draft'}
            onClick={handleSave}
            disabled={isSaving || isSubmitting}
          />
          <PrimaryButton
            text={isSubmitting ? 'Submitting…' : 'Submit'}
            onClick={handleSubmit}
            disabled={isSubmitting}
          />
        </div>
      </form>
    </div>
  );
}
