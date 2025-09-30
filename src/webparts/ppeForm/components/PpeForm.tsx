import * as React from "react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { ISPHttpClientOptions, MSGraphClientV3, SPHttpClient } from "@microsoft/sp-http";
import { ICommon, IGraphResponse, IGraphUserResponse, ILKPItemInstructionsForUse, ISPListItem } from "../../../Interfaces/ICommon";

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
import {
  PrimaryButton,
  // DefaultButton 
} from '@fluentui/react';

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
import { DocumentMetaBanner } from "./DocumentMetaBanner";
// import { IPPEForm } from "../../../Interfaces/IPPEForm";
// import { IPPEForm } from "../../../Interfaces/IPPEForm";
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
  const containerRef = React.useRef<HTMLDivElement>(null);
  const bannerTopRef = useRef<HTMLDivElement>(null);
  const [_jobTitle, setJobTitleId] = useState<ICommon>({ id: '', title: '' });
  const [_department, setDepartmentId] = useState<ICommon>({ id: '', title: '' });
  const [_division, setDivisionId] = useState<ICommon>({ id: '', title: '' });
  const [_company, setCompanyId] = useState<ICommon>({ id: '', title: '' });
  const [_employee, setEmployee] = useState<IPersonaProps[]>([]);
  const [_employeeId, setEmployeeId] = useState<number | undefined>(undefined);
  const [_submitter, setSubmitter] = useState<IPersonaProps[]>([]);
  const [_requester, setRequester] = useState<IPersonaProps[]>([]);
  const [_isReplacementChecked, setIsReplacementChecked] = useState(false);
  const [_replacementReason, setReplacementReason] = useState<string>('');
  const [users, setUsers] = useState<IUser[]>([]);
  const [employees, setEmployees] = useState<IEmployeeProps[]>([]);
  const [employeePPEItemsCriteria, setEmployeePPEItemsCriteria] = useState<IEmployeesPPEItemsCriteria>({ Id: '' });
  const [ppeItems, setPpeItems] = useState<IPPEItem[]>([]);
  const [ppeItemDetails, setPpeItemDetails] = useState<IPPEItemDetails[]>([]);
  const [itemInstructionsForUse, setItemInstructionsForUse] = useState<ILKPItemInstructionsForUse[]>([]);
  const [formsApprovalWorkflow, setFormsApprovalWorkflow] = useState<IFormsApprovalWorkflow[]>([]);
  const [coralFormsList, setCoralFormsList] = useState<ICoralFormsList>({ Id: "" });
  const [loading, setLoading] = useState<boolean>(true);
  // const [, setIsSaving] = useState<boolean>(false);    // Save button state
  const [isSubmitting, setIsSubmitting] = useState<boolean>(false); // Submit button state
  const [bannerText, setBannerText] = useState<string>();
  const [bannerTick, setBannerTick] = useState(0);
  const [lKPWorkflowStatus, setlKPWorkflowStatus] = useState<ISPListItem[]>([]);

  // TODO: Replace these with your actual list GUIDs or titles
  const sharePointLists = {
    PPEForm: { value: '7afa2286-c552-4ff6-952e-1c09f32734cd' },
    PPEFormItems: { value: '43a83414-6b55-4856-aaea-b7447a37a024' },
    PPEFormApprovalWorkflow: { value: 'd0f9db49-8f4d-4685-9176-d639baec4b88' },
  };
  const webUrl = props.context.pageContext.web.absoluteUrl;

  // Resolve a SharePoint user by email or login and return its numeric user Id
  const ensureUserId = useCallback(async (loginOrEmail?: string): Promise<number | undefined> => {
    if (!loginOrEmail) return undefined;
    const url = `${webUrl}/_api/web/ensureuser`;
    const options: ISPHttpClientOptions = {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=verbose',
        'odata-version': '',
      },
      body: JSON.stringify({ 'logonName': 'i:0#.f|membership|' + loginOrEmail })
      // body: JSON.stringify({ logonName: loginOrEmail })
    };
    const res = await props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
    if (!res.ok) {
      const t = await res.text();
      throw new Error(`ensureUser failed for ${loginOrEmail}: ${t}`);
    }
    const u = await res.json();
    return u?.Id;
  }, [props.context.spHttpClient, webUrl]);

  // Try to get an email from a Persona; fallback: look up from loaded Graph users by display name
  const emailFromPersona = useCallback((p?: IPersonaProps): string | undefined => {
    if (!p) return undefined;
    const sec = String(p.secondaryText || '').trim();
    if (sec.includes('@')) return sec; // already an email
    // fallback by displayName
    const byName = users.find(u => (u.displayName || '').toLowerCase() === String(p.text || '').toLowerCase());
    return byName?.email;
  }, [users]);

  interface ItemRowState {
    itemId: number | undefined;  // unique key per row
    item: string;
    order?: number | undefined;             // original order for sorting
    brands: string[];            // all available brands for item
    brandSelected?: string;      // chosen brand
    required: boolean | undefined;           // required flag per item
    qty?: string;                // overall quantity (if applies per item)
    details: string[];           // all available detail titles for this item
    selectedDetail?: string;   // checked details
    itemSizes: string[];         // available sizes at item-level
    itemSizeSelected?: string;   // chosen size for the item
    otherPurpose?: string | undefined; // Added: holds free-text per detail for "Others"
    types?: string[];
    selectedType?: string;              // unique list of Types for this item (if any)
    typeSizesMap?: Record<string, string[]>;
    selectedSizesByType?: Record<string, string | undefined>; // NEW: one size per type
  }

  const [itemRows, setItemRows] = useState<ItemRowState[]>([]);

  const loggedInUserEmail = useMemo(
    () => (props.context.pageContext?.user?.email || '').toLowerCase(),
    [props.context]
  );

  const loggedInUser = useMemo(
    () => users.find(u => (u.email || '').toLowerCase() === loggedInUserEmail),
    [users, loggedInUserEmail]
  );

  const formPayload = useCallback((status: 'Draft' | 'Submitted') => {
    return {
      formName,
      status,
      employeeId: _employeeId,
      employeeName: _employee?.[0]?.text,
      _jobTitle,
      _department,
      _division,
      _company,
      requestType: _isReplacementChecked ? 'Replacement' : 'New Request',
      replacementReason: _isReplacementChecked ? _replacementReason : '',
      items: itemRows.map(r => {
        const hasTypes = r.types && r.types.length > 0;
        const sizeCsv = hasTypes ? r.types!.map(t => (r.selectedSizesByType?.[t] ?? '')).join(',') : (r.itemSizeSelected || '');
        const typeCsv = hasTypes ? r.types!.join(',') : (r.selectedType || '');
        return {
          itemId: r.itemId,
          item: r.item,
          required: !!r.required,
          brand: r.brandSelected,
          qty: r.qty ? Number(r.qty) : undefined,
          size: sizeCsv,
          selectedDetails: r.selectedDetail,
          selectedDetailId: r.selectedDetail ? Number(ppeItemDetails.find(d => d.Title === r.selectedDetail && d.PPEItem?.Id === r.itemId)?.Id) : undefined,
          type: typeCsv,
          othersText: r.item.toLowerCase() === 'others' ? r.otherPurpose : undefined
        };
      }),
      approvals: formsApprovalWorkflow
    };
  }, [_employee, _employeeId, _jobTitle, _department, _division, _company, _isReplacementChecked, _replacementReason, itemRows, formsApprovalWorkflow, formName]);

  const validateBeforeSubmit = useCallback((): string | undefined => {
    // If “Others” is required, ensure a size is chosen (since you show size ComboBox when required)
    const missing: string[] = [];
    if (!_employee?.[0]?.text?.trim()) missing.push('Employee Name');
    if (!_jobTitle?.title?.trim()) missing.push('Job Title');
    if (!_department.title?.trim()) missing.push('Department');
    if (!_company.title?.trim()) missing.push('Company');
    if (!_division.title?.trim()) missing.push('Division');
    if (_requester.length === 0) missing.push('Requester');

    if (missing.length) {
      return `Please fill in the required fields: ${missing.join(', ')}.`;
    }

    // Example: if Replacement, require a reason
    if (_isReplacementChecked && !_replacementReason.trim()) return 'Please provide a reason for Replacement.';

    // Ensure at least one item is required or has any selection
    const anyRequired = itemRows.some(r => r.required);
    if (!anyRequired) return 'Please select at least one item or mark one as Required.';

    // const itemRequiredWithoutQty = itemRows.find(r => r.required && r.qty == undefined);
    // if (itemRequiredWithoutQty) return `Please fill in the Qty field for the item "${itemRequiredWithoutQty.item}".`;

    if (anyRequired) {
      const othersMissingPurpose = itemRows.some(r => r.item.toLowerCase() === 'others' && r.required && (r.otherPurpose === undefined || !r.otherPurpose.trim()));
      if (othersMissingPurpose) return 'Please fill in the Purpose field for "Others" since it is marked Required.';

      const othersMissingSize = itemRows.some(r => r.item.toLowerCase() === 'others' && r.required && (!r.itemSizeSelected || !r.itemSizeSelected.trim()));
      if (othersMissingSize) return 'Please choose a size for "Others" since it is marked Required.';

      // Validate each required item individually and stop on first failure
      for (const r of itemRows.filter(r => r.required)) {
        if (!r.required) continue;

        // 1) Detail is required when the item is marked required
        if (!r.selectedDetail && r.item.toLowerCase() !== 'others') {
          if (r.item.toLowerCase() === 'winter jacket') continue;
          return `Please select a Specific Detail for the required item "${r.item}".`;
        }

        if (r.item.toLowerCase() === 'others' && (r.otherPurpose === undefined || r.otherPurpose.toString().trim() === '')) {
          return `Please fill in the Purpose field for the item "${r.item}".`;
        }

        if (r.qty === undefined || r.qty.toString().trim() === '') {
          return `Please enter a quantity for the required item "${r.item}".`;
        }

        // Validate quantity for all items, but only if a value is provided
        const qtyStr = (r.qty ?? '').toString().trim();

        if (!qtyStr) continue; // only validate if set

        const isWholeNumber = /^\d+$/.test(qtyStr);
        const n = Number(qtyStr);

        if (!isWholeNumber || !Number.isFinite(n) || n <= 0) {
          return `Please enter a valid quantity (whole number > 0) for the item "${r.item}".`;
        }

        // 2) If sizes exist, validate size selection
        const hasTypes = Array.isArray(r.types) && r.types.length > 0 && r.item.toLowerCase() !== 'others';
        const hasAnySizes = (Array.isArray(r.itemSizes) && r.itemSizes.length > 0 && r.item.toLowerCase() !== 'others') || hasTypes;

        if (hasAnySizes) {
          if (hasTypes) {
            // typed sizes: at least one type must have a selection
            const anyTypeHasSelection = Object.values(r.selectedSizesByType || {}).some(v => !!v && String(v).trim().length > 0);
            if (!anyTypeHasSelection) {
              return `Please choose a size for the required item "${r.item}".`;
            }

            const isCoverallsDetail = /coveralls/i.test(r.selectedDetail || '');
            if (isCoverallsDetail) {
              const coverallsKey = (r.types || []).find(t => /coveralls/i.test(t));
              const coverallsSel = coverallsKey ? r.selectedSizesByType?.[coverallsKey] : undefined;
              if (coverallsKey && (!coverallsSel || !String(coverallsSel).trim())) {
                return `Please choose a size for Coveralls for the required item "${r.item}".`;
              }
            } else {
              // Require both Top and Pants when not Coveralls
              const topKey = (r.types || []).find(t => /top/i.test(t));
              const pantsKey = (r.types || []).find(t => /pants/i.test(t));
              const topSel = topKey ? r.selectedSizesByType?.[topKey] : undefined;
              const pantsSel = pantsKey ? r.selectedSizesByType?.[pantsKey] : undefined;

              // If both types exist, both must be selected
              if (topKey && pantsKey) {
                if (!topSel || !String(topSel).trim() || !pantsSel || !String(pantsSel).trim()) {
                  return `Please choose both Top and Pants sizes for the required item "${r.item}".`;
                }
              }
            }

          } else {
            // non-typed sizes
            if (!r.itemSizeSelected || !String(r.itemSizeSelected).trim()) {
              return `Please choose a size for the required item "${r.item}".`;
            }
          }
        }
      }
    }

    const nonApprovedForm = formsApprovalWorkflow.filter(item => item.DepartmentManager?.id === loggedInUser?.id && item.Status === undefined);
    if (nonApprovedForm && nonApprovedForm.length >= 1) { return 'Please update your approval status before submitting the form.'; }
    const rejectedWorkflowStatusId = lKPWorkflowStatus.find(p => p.Title?.toLowerCase().includes("rejected"));
    const rejectedForm = formsApprovalWorkflow.filter(item => item.DepartmentManager?.id === loggedInUser?.id && item.Status === rejectedWorkflowStatusId?.Id?.toString());
    if (rejectedForm && rejectedForm.length > 0 && rejectedForm[0]?.Reason === undefined) { return 'Please provide a reason for rejection before submitting the form.' };

    return undefined;
  }, [_employee, _jobTitle, _department, _company, _division, _requester, itemRows, _isReplacementChecked, _replacementReason, formsApprovalWorkflow]);

  const canEditApprovalRow = useCallback((row: IFormsApprovalWorkflow): boolean => {
    const dm = row?.DepartmentManager as IPersonaProps | undefined;
    if (!dm) return false;

    // Prefer email match (we stored email in secondaryText when we found a Graph match)
    const dmEmail = (dm.secondaryText || '').toLowerCase();
    if (dmEmail && loggedInUserEmail && dmEmail === loggedInUserEmail) {
      return true;
    }

    // Fallback to Graph id match
    const dmId = dm.id ? String(dm.id).toLowerCase() : '';
    const currId = loggedInUser?.id ? String(loggedInUser.id).toLowerCase() : '';
    if (dmId && currId && dmId === currId) {
      return true;
    }

    // Last resort: display name match
    const dmName = (dm.text || '').toLowerCase();
    const currName = (loggedInUser?.displayName || '').toLowerCase();
    if (dmName && currName && dmName === currName) {
      return true;
    }
    return false;
  }, [loggedInUserEmail, loggedInUser]);

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
      // console.error("Error fetching users:", error);
      setUsers([]);
      return [];
    }
  }, [props.context]);

  const _getEmployees = useCallback(async (usersArg?: IUser[], employeeFullName?: string): Promise<IEmployeeProps[]> => {
    try {
      const query: string = `?$select=Id,EmployeeID,FullName,Division/Id,Division/Title,Company/Id,Company/Title,EmploymentStatus,JobTitle/Id,JobTitle/Title,` +
        `Department/Id,Department/Title,Manager/Id,Manager/FullName,Created,Author/EMail` +
        `&$expand=Author,Division,Company,JobTitle,Department,Manager,Author` +
        `&$filter=substringof('${employeeFullName}', FullName)&$orderby=Order asc`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, '6331f331-caa7-4732-a205-7abcd1f7d53f', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: IEmployeeProps[] = [];
      const usersToUse = usersArg && usersArg.length ? usersArg : users;
      data.forEach((obj: any) => {
        if (obj) {
          const createdBy = usersToUse && usersToUse.length ? usersToUse.filter(u => u.email?.toString() === obj.Author?.EMail?.toString())[0] : undefined;
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
      // console.error('An error has occurred while retrieving items!', error);
      setEmployees([]);
      return [];
    }
  }, [props.context, spHelpers]);

  const _getEmployeesPPEItemsCriteria = useCallback(async (usersArg?: IUser[], employeeID?: string) => {
    try {
      const query: string = `?$select=Id,Employee/EmployeeID,Employee/FullName,Created,SafetyHelmet,ReflectiveVest,SafetyShoes,` +
        `Employee/ID,Employee/FullName,RainSuit/Id,RainSuit/DisplayText,UniformCoveralls/Id,UniformCoveralls/DisplayText,UniformTop/Id,UniformTop/DisplayText,` +
        `UniformPants/Id,UniformPants/DisplayText,WinterJacket/Id,WinterJacket/DisplayText,Author/EMail` +
        `&$expand=Author,Employee,RainSuit,UniformCoveralls,UniformTop,UniformPants,WinterJacket` +
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
      // console.error('An error has occurred while retrieving items!', error);
      setEmployeePPEItemsCriteria({ Id: '' });
    }
  }, [props.context, spHelpers]);

  const _getCoralFormsList = useCallback(async (usersArg?: IUser[]): Promise<ICoralFormsList | undefined> => {
    try {
      const searchEscaped = formName.replace(/'/g, "''");
      const query: string = `?$select=Id,Title,hasInstructionForUse,hasWorkflow,Created,Author/EMail&$expand=Author&$filter=substringof('${searchEscaped}', Title)`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, '22a9fee1-dfe4-4ad0-8ce4-89d014c63049', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const usersToUse = usersArg && usersArg.length ? usersArg : users;
      const ppeform = data.find((obj: any) => obj !== null);
      let result: ICoralFormsList = { Id: "" };

      if (ppeform) {
        const createdBy = usersToUse?.find(u => u.email?.toString() === ppeform.Author?.EMail?.toString());
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
      const query: string = `?$select=Id,Title,Brands,RecordOrder,Created,Author/EMail&$expand=Author&$orderby=RecordOrder asc`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, '00f4b4fc-896d-40bb-9a03-3889e651d244', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: IPPEItem[] = [];
      const usersToUse = usersArg && usersArg.length ? usersArg : users;

      data.forEach((obj: any) => {
        if (obj) {
          const createdBy = usersToUse && usersToUse.length ? usersToUse.filter(u => u.email?.toString() === obj.Author?.EMail?.toString())[0] : undefined;
          let created: Date | undefined;
          if (obj.Created !== undefined) created = new Date(spHelpers.adjustDateForGMTOffset(obj.Created));
          const temp: IPPEItem = {
            Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            CreatedBy: createdBy !== undefined ? createdBy : undefined,
            Created: created !== undefined ? created : undefined,
            Title: obj.Title !== undefined && obj.Title !== null ? obj.Title : undefined,
            Order: obj.RecordOrder !== undefined && obj.RecordOrder !== null ? obj.RecordOrder : undefined,
            Brands: spHelpers.NormalizeToStringArray(obj.Brands),
            PPEItemsDetails: []
          };
          result.push(temp);
        }
      });
      setPpeItems(result);
    } catch (error) {
      // console.error('An error has occurred while retrieving items!', error);
      setPpeItems([]);
    }
  }, [props.context, spHelpers]);

  const _getPPEItemsDetails = useCallback(async (usersArg?: IUser[]) => {
    try {
      const query: string = `?$select=Id,Title,PPEItem,Sizes,Types,RecordOrder,Created,PPEItem/Id,PPEItem/Title,Author/EMail&$expand=Author,PPEItem`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, '3435bbde-cb56-43cf-aacf-e975c65b68c3', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: IPPEItemDetails[] = [];
      const usersToUse = usersArg && usersArg.length ? usersArg : users;
      data.forEach((obj: any) => {
        if (obj) {
          const createdBy = usersToUse && usersToUse.length ? usersToUse.filter(u => u.email?.toString() === obj.Author?.EMail?.toString())[0] : undefined;
          let created: Date | undefined;
          if (obj.Created !== undefined) created = new Date(spHelpers.adjustDateForGMTOffset(obj.Created));
          const temp: IPPEItemDetails = {
            Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            Title: obj.Title !== undefined && obj.Title !== null ? obj.Title : undefined,
            CreatedBy: createdBy !== undefined ? createdBy : undefined,
            Created: created !== undefined ? created : undefined,
            Sizes: spHelpers.NormalizeToStringArray(obj.Sizes),
            Types: spHelpers.NormalizeToStringArray(obj.Types),
            Order: obj.RecordOrder !== undefined && obj.RecordOrder !== null ? obj.RecordOrder : undefined,
            PPEItem: obj.PPEItem !== undefined ? {
              Id: obj.PPEItem.Id !== undefined && obj.PPEItem.Id !== null ? obj.PPEItem.Id : undefined,
              Title: obj.PPEItem.Title !== undefined && obj.PPEItem.Title !== null ? obj.PPEItem.Title : undefined,
              Order: obj.PPEItem.RecordOrder !== undefined && obj.PPEItem.RecordOrder !== null ? obj.PPEItem.RecordOrder : undefined,
              Brands: spHelpers.NormalizeToStringArray(obj.PPEItem.Brands),
            } : undefined,
          };
          result.push(temp);
        }
      });
      setPpeItemDetails(result);
    } catch (error) {
      setPpeItemDetails([]);
      // console.error('An error has occurred while retrieving items!', error);
    }
  }, [props.context, spHelpers]);

  const _getLKPItemInstructionsForUse = useCallback(async (usersArg?: IUser[], formName?: string) => {
    try {
      const query: string = `?$select=Id,FormName,RecordOrder,Description,Created,Author/EMail&$expand=Author&$filter=substringof('${formName}', FormName)&$orderby=RecordOrder asc`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, '2edbaa23-948a-4560-a553-acbe7bc60e7b', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: ILKPItemInstructionsForUse[] = [];
      const usersToUse = usersArg && usersArg.length ? usersArg : users;
      data.forEach((obj: any) => {
        if (obj) {
          const createdBy = usersToUse && usersToUse.length ? usersToUse.filter(u => u.email?.toString() === obj.Author?.EMail?.toString())[0] : undefined;
          let created: Date | undefined;
          if (obj.Created !== undefined) created = new Date(spHelpers.adjustDateForGMTOffset(obj.Created));
          const temp: ILKPItemInstructionsForUse = {
            Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            FormName: obj.FormName !== undefined && obj.FormName !== null ? obj.FormName : undefined,
            Order: obj.RecordOrder !== undefined && obj.RecordOrder !== null ? obj.Order : undefined,
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
      setItemInstructionsForUse([]);
      // console.error('An error has occurred while retrieving items!', error);
    }
  }, [props.context, spHelpers]);

  // const _getFormsApprovalWorkflow = useCallback(async (usersArg?: IUser[], formName?: string) => {
  //   try {
  //     const query: string = `?$select=Id,Author/EMail,FormName/Id,FormName/Title,IsFinalFormApprover,ManagerName/Id,RecordOrder,Created,SignOffName,DepartmentManager/Id,DepartmentManager/Title,DepartmentManager/EMail` +
  //       `&$expand=Author,FormName,ManagerName,DepartmentManager` +
  //       `&$filter=substringof('${formName}', FormName/Title)&$orderby=RecordOrder asc`;
  //     spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'd084f344-63cb-4426-ae51-d7f875f3f99a', query);
  //     const data = await spCrudRef.current._getItemsWithQuery();
  //     const result: IFormsApprovalWorkflow[] = [];
  //     const usersToUse = usersArg && usersArg.length ? usersArg : users;
  //     data.forEach((obj: any) => {
  //       if (obj) {
  //         const createdBy = usersToUse && usersToUse.length ? usersToUse.filter(u => u.email?.toString() === obj.Author?.EMail?.toString())[0] : undefined;
  //         let created: Date | undefined;
  //         const deptEmail = obj?.DepartmentManager?.EMail;
  //         const deptTitle = obj?.DepartmentManager?.Title;
  //         const match = (deptEmail && usersToUse.find(u => (u.email || '').toLowerCase() === String(deptEmail).toLowerCase()));

  //         const deptManagerPersona: IPersonaProps | undefined = match
  //           ? { text: match.displayName || deptTitle || '', secondaryText: match.email || match.jobTitle || '', id: match.id }
  //           : (deptTitle ? { text: deptTitle, secondaryText: deptEmail || '', id: String(obj.DepartmentManager?.Id ?? deptTitle) } as IPersonaProps : undefined);


  //         if (obj.Created !== undefined) created = new Date(spHelpers.adjustDateForGMTOffset(obj.Created));
  //         const temp: IFormsApprovalWorkflow = {
  //           Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
  //           FormName: obj.FormName !== undefined && obj.FormName !== null ? obj.FormName : undefined,
  //           Order: obj.RecordOrder !== undefined && obj.RecordOrder !== null ? obj.RecordOrder : undefined,
  //           SignOffName: obj.SignOffName !== undefined && obj.SignOffName !== null ? obj.SignOffName : undefined,
  //           EmployeeId: obj.ManagerName !== undefined && obj.ManagerName !== null ? obj.ManagerName.Id : undefined,
  //           DepartmentManager: deptManagerPersona,
  //           IsFinalFormApprover: obj.IsFinalFormApprover !== undefined && obj.IsFinalFormApprover !== null ? obj.IsFinalFormApprover : false,
  //           Status: undefined,
  //           Reason: undefined,
  //           Date: undefined,
  //           Created: created !== undefined ? created : undefined,
  //           CreatedBy: createdBy !== undefined ? createdBy : undefined,
  //         };
  //         result.push(temp);
  //       }
  //     });
  //     // sort by Order (ascending). If Order is missing, place those items at the end.
  //     result.sort((a, b) => {
  //       const aOrder = (a && a.Order !== undefined && a.Order !== null) ? Number(a.Order) : Number.POSITIVE_INFINITY;
  //       const bOrder = (b && b.Order !== undefined && b.Order !== null) ? Number(b.Order) : Number.POSITIVE_INFINITY;
  //       return aOrder - bOrder;
  //     });
  //     setFormsApprovalWorkflow(result);
  //   } catch (error) {
  //     setFormsApprovalWorkflow([]);
  //     // console.error('An error has occurred while retrieving items!', error);
  //   }
  // }, [props.context, spHelpers]);

  const _getLKPWorkflowStatus = useCallback(async (usersArg?: IUser[]): Promise<ISPListItem[]> => {
    try {
      const query: string = `?$select=Id,Title,RecordOrder,Created,Author/EMail&$expand=Author&$orderby=RecordOrder asc`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, '80eabd4a-6467-40d4-ae8f-fafcc77d334e', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: ISPListItem[] = [];
      const usersToUse = usersArg && usersArg.length ? usersArg : users;
      data.forEach((obj: any) => {
        if (obj) {
          const createdBy = usersToUse && usersToUse.length ? usersToUse.filter(u => u.email?.toString() === obj.Author?.EMail?.toString())[0] : undefined;
          let created: Date | undefined;
          if (obj.Created !== undefined) created = new Date(spHelpers.adjustDateForGMTOffset(obj.Created));
          const temp: ISPListItem = {
            Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            Title: obj.Title !== undefined && obj.Title !== null ? obj.Title : undefined,
            Order: obj.RecordOrder !== undefined && obj.RecordOrder !== null ? obj.RecordOrder : undefined,
            Created: created !== undefined ? created : undefined,
            CreatedBy: createdBy !== undefined ? createdBy : undefined,
          };

          result.push(temp);
        }
      });
      setlKPWorkflowStatus(result);
      return result;
    } catch (error) {
      setlKPWorkflowStatus([]);
      return [];
    }
  }, [props.context, spHelpers]);


  const _getPPEFormApprovalWorkflows = useCallback(async (usersArg?: IUser[], formId: number = 1) => {
    try {
      const PPEFormApprovalWorkflowGUID = sharePointLists.PPEFormApprovalWorkflow.value;
      const query: string = `?$select=Id,SignOffName,Approver/Id,Approver/EMail,Approver/Title,Author/EMail,PPEForm/Id,PPEForm/Title,IsFinalApprover,OrderRecord,Created,StatusRecord/Id,StatusRecord/Title,Reason` +
        `&$expand=Author,PPEForm,StatusRecord,Approver` +
        `&$filter=PPEForm/Id eq 61`+
        `&$orderby=OrderRecord asc`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, PPEFormApprovalWorkflowGUID, query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: IFormsApprovalWorkflow[] = [];
      const usersToUse = usersArg && usersArg.length ? usersArg : users;
      data.forEach((obj: any) => {
        if (obj) {
          const createdBy = usersToUse && usersToUse.length ? usersToUse.filter(u => u.email?.toString() === obj.Author?.EMail?.toString())[0] : undefined;
          let created: Date | undefined;
          const approverEmail = obj?.Approver?.EMail;
          const approverTitle = obj?.Approver?.Title;
          const match = (approverEmail && usersToUse.find(u => (u.email || '').toLowerCase() === String(approverEmail).toLowerCase()));

          const deptApproverPersona: IPersonaProps | undefined = match
            ? { text: match.displayName || approverTitle || '', secondaryText: match.email || match.jobTitle || '', id: match.id }
            : (approverTitle ? { text: approverTitle, secondaryText: approverEmail || '', id: String(obj.Approver?.Id ?? approverTitle) } as IPersonaProps : undefined);

          if (obj.Created !== undefined) created = new Date(spHelpers.adjustDateForGMTOffset(obj.Created));
          const temp: IFormsApprovalWorkflow = {
            Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            FormName: obj.FormName !== undefined && obj.FormName !== null ? { title: obj.FormName.Title, id: obj.FormName.Id } : undefined,
            Order: obj.RecordOrder !== undefined && obj.RecordOrder !== null ? obj.RecordOrder : undefined,
            SignOffName: obj.SignOffName !== undefined && obj.SignOffName !== null ? obj.SignOffName : undefined,
            EmployeeId: obj.ManagerName !== undefined && obj.ManagerName !== null ? obj.ManagerName.Id : undefined,
            DepartmentManager: deptApproverPersona,
            IsFinalFormApprover: obj.IsFinalFormApprover !== undefined && obj.IsFinalFormApprover !== null ? obj.IsFinalFormApprover : false,
            Status: obj.StatusRecord !== undefined && obj.StatusRecord !== null ? { id: obj.StatusRecord.Id?.toString(), title: obj.StatusRecord.Title } : undefined,
            Reason: undefined,
            Date: created !== undefined ? created : undefined,
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
      setFormsApprovalWorkflow([]);
      // console.error('An error has occurred while retrieving items!', error);
    }
  }, [props.context, spHelpers]);

  // const _getPPEFormPerEmployee = useCallback(async (usersArg?: IUser[], employeeListId?: number, employeeHRId?: number): Promise<ISPListItem[]> => {
  //   try {
  //     const ppeFormGUID = sharePointLists.PPEForm.value;
  //     const query: string = `?$select=Id,EmployeeID,EmployeeRecord/Id,EmployeeRecord/FullName,ReasonForRequest,ReplacementReason,JobTitleRecord/Id,JobTitleRecord/Title,CompanyRecord/Id,CompanyRecord/Title,DivisionRecord/Id,DivisionRecord/Title,DepartmentRecord/Id,DepartmentRecord/Title,RecordOrder,Created,Author/EMail` +
  //       `RequesterName/Id,RequesterName/EMail,RequesterName/Title,SubmitterName/Id,SubmitterName/EMail,SubmitterName/Title` +
  //       `&$expand=EmployeeRecord,Author,JobTitleRecord,CompanyRecord,DivisionRecord,DepartmentRecord,RequesterName,SubmitterName` +
  //       `&$orderby=RecordOrder asc`;
  //     spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, ppeFormGUID, query);
  //     const data = await spCrudRef.current._getItemsWithQuery();
  //     const result: IPPEForm[] = [];
  //     const usersToUse = usersArg && usersArg.length ? usersArg : users;

  //     const toIUser = (p?: { Id?: any; Title?: string; EMail?: string }): IUser => {
  //       if (!p) {
  //         return { id: '', displayName: '' }; // safe fallback to satisfy non-optional IUser
  //       }
  //       const email = (p.EMail || '').toLowerCase();
  //       const match = usersToUse.find(u => (u.email || '').toLowerCase() === email);
  //       if (match) return match;
  //       return {
  //         id: String(p.Id ?? p.EMail ?? p.Title ?? ''),
  //         displayName: p.Title || '',
  //         email: p.EMail
  //       };
  //     };

  //     data.forEach((obj: any) => {
  //       if (obj) {
  //         const createdBy = usersToUse && usersToUse.length ? usersToUse.filter(u => u.email?.toString() === obj.Author?.EMail?.toString())[0] : undefined;
  //         let created: Date | undefined;
  //         if (obj.Created !== undefined) created = new Date(spHelpers.adjustDateForGMTOffset(obj.Created));
  //         const temp: IPPEForm = {
  //           Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
  //           employeeID: obj.EmployeeID !== undefined && obj.EmployeeID !== null ? obj.EmployeeID : undefined,
  //           employeeRecord: obj.EmployeeRecord !== undefined && obj.EmployeeRecord !== null ? { Id: obj.EmployeeRecord.Id, primaryText: obj.EmployeeRecord.FullName } as IPersonaProps : undefined,
  //           reasonForRequest: obj.ReasonForRequest !== undefined && obj.ReasonForRequest !== null ? obj.ReasonForRequest : undefined,
  //           replacementReason: obj.ReplacementReason !== undefined && obj.ReplacementReason !== null ? obj.ReplacementReason : undefined,
  //           jobTitle: obj.JobTitleRecord !== undefined && obj.JobTitleRecord !== null ? { id: obj.JobTitleRecord.Id, title: obj.JobTitleRecord.Title } : undefined,
  //           company: obj.CompanyRecord !== undefined && obj.CompanyRecord !== null ? { id: obj.CompanyRecord.Id, title: obj.CompanyRecord.Title } : undefined,
  //           division: obj.DivisionRecord !== undefined && obj.DivisionRecord !== null ? { id: obj.DivisionRecord.Id, title: obj.DivisionRecord.Title } : undefined,
  //           department: obj.DepartmentRecord !== undefined && obj.DepartmentRecord !== null ? { id: obj.DepartmentRecord.Id, title: obj.DepartmentRecord.Title } : undefined,
  //           requesterName: toIUser(obj.RequesterName),
  //           submitterName: toIUser(obj.SubmitterName),
  //           dateRequested: created !== undefined ? created : undefined,
  //           Order: obj.RecordOrder !== undefined && obj.RecordOrder !== null ? obj.RecordOrder : undefined,
  //           ppeItems: [],
  //           Created: created !== undefined ? created : undefined,
  //           CreatedBy: createdBy !== undefined ? createdBy : undefined,
  //         };

  //         result.push(temp);
  //       }
  //     });
  //     // setlKPWorkflowStatus(result);
  //     return result;
  //   } catch (error) {
  //     // setlKPWorkflowStatus([]);
  //     return [];
  //   }
  // }, [props.context, spHelpers,users, sharePointLists.PPEForm]);


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
      await _getLKPWorkflowStatus(fetchedUsers);
      await _getPPEItems(fetchedUsers);
      await _getPPEItemsDetails(fetchedUsers);

      // Use the returned result from _getCoralFormsList instead of the (possibly stale) coralFormsList state
      if (coralListResult && coralListResult.hasInstructionForUse) {
        if (coralListResult.hasInstructionForUse) await _getLKPItemInstructionsForUse(fetchedUsers, formName);
        //  if (coralListResult.hasWorkflow) await _getFormsApprovalWorkflow(fetchedUsers, formName);
        if (coralListResult.hasWorkflow) await _getPPEFormApprovalWorkflows(fetchedUsers, Number(coralFormsList?.Id));
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
  }, [_getEmployees, _getUsers, _getLKPWorkflowStatus, _getPPEItems, _getPPEItemsDetails, _getCoralFormsList, _getLKPItemInstructionsForUse, _getPPEFormApprovalWorkflows, props.context]);

  useEffect(() => {
    if (!bannerText) return;

    // Determine current scrollTop (container or window)
    const currentScrollTop = (containerRef.current && typeof containerRef.current.scrollTop === 'number'
      ? containerRef.current.scrollTop
      : (window.scrollY || document.documentElement.scrollTop || 0));

    if (currentScrollTop >= 0) {
      // Wait a tick so the banner renders, then scroll to it
      requestAnimationFrame(() => {
        if (bannerTopRef.current) {
          bannerTopRef.current.scrollIntoView({ behavior: 'smooth', block: 'start' });
        } else if (containerRef.current) {
          containerRef.current.scrollTo({ top: 0, behavior: 'smooth' });
        } else {
          window.scrollTo({ top: 0, behavior: 'smooth' });
        }
      });
    }
  }, [bannerText, bannerTick]);
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

  const ppeItemMap = useMemo(() => {
    // Create a map from item title to details``
    const map: { [id: number]: { [title: string]: IPPEItemDetails[] } } = {};

    (ppeItemDetails || []).forEach((detail: IPPEItemDetails) => {
      const title = detail?.PPEItem?.Title ? String(detail.PPEItem.Title).trim() : undefined;
      const id = detail?.PPEItem?.Id ? Number(detail.PPEItem.Id) : undefined;
      if (!title || !id) return;

      if (!map[id]) {
        map[id] = {};
      }

      if (!map[id][title]) {
        map[id][title] = [];
      }

      map[id][title].push(detail);
    });

    // Now fill each ppeItem with its details
    return (ppeItems || []).map(item => {
      const title = item.Title ? String(item.Title).trim() : "";
      const id = Number(item.Id);
      return {
        ...item,
        Brands: brandsMap.find(b => b.key === title)?.brands || [],
        PPEItemsDetails: (map[id] && map[id][title]) ? map[id][title] : []  // fill with matching details or empty array
      };
    });
  }, [ppeItems, ppeItemDetails, brandsMap]);

  useEffect(() => {
    if (!ppeItemMap || !ppeItemMap.length) return;

    const rows: ItemRowState[] = ppeItemMap.map(item => {
      const details = (item.PPEItemsDetails || []);
      // normalize types and sizes
      const allSizes = Array.from(
        new Set(details.flatMap(d => (spHelpers.NormalizeToStringArray((d as any).Sizes) || []))
          .map(s => String(s).trim())
          .filter(Boolean)
        )
      );

      const typesArr = Array.from(
        new Set(details
          .flatMap(d => (spHelpers.NormalizeToStringArray((d as any).Types) || []))
          .map(t => String(t).trim())
          .filter(Boolean)
        )
      );

      const typeSizesMap: Record<string, string[]> = {};
      details.forEach(d => {
        const dTypes = spHelpers.NormalizeToStringArray((d as any).Types) || [];
        const dSizes = (spHelpers.NormalizeToStringArray((d as any).Sizes) || []).map(s => String(s).trim()).filter(Boolean);
        dTypes.forEach(t => {
          const key = String(t).trim();
          if (!key) return;
          const prev = typeSizesMap[key] || [];
          typeSizesMap[key] = Array.from(new Set(prev.concat(dSizes)));
        });
      });

      return {
        itemId: Number(item.Id) || undefined,
        item: item.Title || "",
        order: item.Order ?? undefined,
        brands: item.Brands || [],
        brandSelected: undefined,
        required: undefined,
        qty: undefined,
        details: details.map(d => d.Title || ""),
        selectedDetail: undefined,
        itemSizes: allSizes,              // union of sizes across details
        itemSizeSelected: undefined,
        // NEW: attach types and per-type sizes
        types: typesArr,
        typeSizesMap,
        selectedSizesByType: {},
        othersItemdetailsText: {},
      };
    });

    setItemRows(rows);
  }, [ppeItemMap, spHelpers]);

  // Apply employee PPE criteria to pre-select details (assumption: label matches detail title)
  // useEffect(() => {
  //   if (!employeePPEItemsCriteria || !employeePPEItemsCriteria.employeeID) return;
  //   const map: Record<string, string> = {};

  //   Object.entries(employeePPEItemsCriteria).forEach(([key, value]) => {
  //     const itemDetail = spHelpers.CamelString(key.split(/(?=[A-Z])/).join(" "));
  //     const itemValue = value || "";

  //     if (itemDetail && spHelpers.CamelString(itemDetail) === ppeItemMap.find(i => i.Title?.toLowerCase() === itemDetail.toLowerCase())?.Title) {
  //       map[key] = itemValue;
  //     }
  //     switch (itemDetail) {
  //       case "Reflective Vest":
  //         map[key] = itemValue;
  //         break;

  //       case "Safety Helmet":
  //         map[key] = itemValue;
  //         break;

  //       case "Safety Shoes":
  //         map[key] = itemValue;
  //         break;

  //       case "Rain Suit":
  //         map[key] = itemValue;
  //         break;

  //       case "Winter Jacket":
  //         map[key] = itemValue;
  //         break;

  //       case "Uniform Coveralls":
  //         map[key] = itemValue;
  //         break;

  //       case "Uniform Top":
  //         map[key] = itemValue;
  //         break;

  //       case "Uniform Pants":
  //         map[key] = itemValue;
  //         break;

  //       default:
  //         break;
  //     }
  //   });

  //   const labelFields: (string | undefined)[] = [
  //     employeePPEItemsCriteria.rainSuit,
  //     employeePPEItemsCriteria.uniformCoveralls,
  //     employeePPEItemsCriteria.uniformTop,
  //     employeePPEItemsCriteria.uniformPants,
  //     employeePPEItemsCriteria.winterJacket
  //   ].filter(Boolean);

  //   if (!labelFields.length) return;

  //   setItemRows(prev => prev.map(r => {
  //     const matched = r.details.filter(d => labelFields.some(l => l && l.toLowerCase() === d.toLowerCase()));
  //     if (!matched.length) return r;
  //     return { ...r, selectedDetails: r.selectedDetail };
  //   }));

  // }, [employeePPEItemsCriteria]);

  const toggleRequired = useCallback((rowIndex: number, checked?: boolean) => {
    setItemRows(prev => prev.map((r, i) => {
      if (i !== rowIndex) return r;
      if (checked) return { ...r, required: true };
      // when unchecking, clear selections
      return {
        ...r,
        required: false,
        brandSelected: undefined,
        selectedDetail: undefined,
        itemSizeSelected: undefined,
        otherPurpose: undefined,
        selectedSizesByType: {},
        qty: undefined,
      };
    }));
  }, []);

  const toggleItemDetail = useCallback((rowIndex: number, detail: string, checked?: boolean) => {
    setItemRows(prev =>
      prev.map((r, idx) => {
        if (idx !== rowIndex) return r;
        if (!detail) return r;
        if (!r.required) return r;

        // Compute next selected detail first
        let nextDetail: string | undefined;
        if (typeof checked === 'boolean') {
          nextDetail = checked
            ? detail
            : (r.selectedDetail === detail ? undefined : r.selectedDetail);
        } else {
          nextDetail = r.selectedDetail === detail ? undefined : detail;
        }

        // If switching to a "Coveralls" detail, and both Top & Pants sizes are currently selected,
        // clear those size selections (keep others, including a Coveralls type, intact).
        let nextSelectedSizesByType = r.selectedSizesByType;
        if (nextDetail && /coveralls/i.test(nextDetail) && r.types && r.types.length) {
          const topKey = r.types.find(t => /top/i.test(t));
          const pantsKey = r.types.find(t => /pants/i.test(t));

          const topSel = topKey ? r.selectedSizesByType?.[topKey] : undefined;
          const pantsSel = pantsKey ? r.selectedSizesByType?.[pantsKey] : undefined;

          const hasTop = !!(topSel && String(topSel).trim());
          const hasPants = !!(pantsSel && String(pantsSel).trim());

          if (topKey && pantsKey && hasTop && hasPants) {
            nextSelectedSizesByType = { ...(r.selectedSizesByType || {}) };
            // nextSelectedSizesByType[topKey] = undefined;
            nextSelectedSizesByType[pantsKey] = undefined;
          }
        }

        return {
          ...r,
          selectedDetail: nextDetail,
          selectedSizesByType: nextSelectedSizesByType
        };
      })
    );
  }, []);

  const toggleBrand = useCallback((rowIndex: number, brandVal?: string, checked?: boolean) => {
    setItemRows(prev =>
      prev.map((r, idx) => {
        if (idx !== rowIndex) return r;
        if (!brandVal) return r;
        if (!r.required) return r;

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

  const toggleSize = useCallback((rowIndex: number, sizeVal?: string, checked?: boolean) => {
    setItemRows(prev =>
      prev.map((r, idx) => {
        if (idx !== rowIndex) return r;
        if (!sizeVal) return r;
        if (!r.required) return r;
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

  const updateOtherPurpose = useCallback((rowIndex: number, value?: string) => {
    setItemRows(prev =>
      prev.map((r, idx) => {
        if (idx !== rowIndex) return r;
        if (!r.required) return r; // keep it locked unless the row is marked Required
        const next = (value ?? '').trim();
        return { ...r, otherPurpose: next.length ? next : undefined };
      })
    );
  }, []);

  const toggleSizeType = useCallback(
    (rowIndex: number, sizeVal?: string, checked?: boolean, typeKey?: string, id?: string) => {
      setItemRows(prev =>
        prev.map((r, idx) => {
          if (idx !== rowIndex) return r;
          if (!sizeVal) return r;
          if (!r.required) return r;
          // If the item has types and a typeKey is provided, maintain one size per type
          if (typeKey && r.types && r.types.length) {
            // Block changes to Pants sizes when a "Coveralls" detail is selected
            const isCoverallsDetail = /coveralls/i.test(r.selectedDetail || '');
            if (isCoverallsDetail && /pants/i.test(typeKey)) {
              return r; // disabled -> do nothing
            }

            const byType = { ...(r.selectedSizesByType || {}) };

            if (typeof checked === 'boolean') {
              if (checked) {
                byType[typeKey] = sizeVal;               // select this size for this type
              } else {
                // Only clear if we're unchecking the currently selected size for this type
                if (byType[typeKey] === sizeVal) {
                  byType[typeKey] = undefined;
                }
              }
            } else {
              // Fallback toggle: toggle the same size for this type
              byType[typeKey] = byType[typeKey] === sizeVal ? undefined : sizeVal;
            }
            return {
              ...r,
              selectedSizesByType: byType,
              // Keep legacy fields untouched for typed items
            };
          }

          // Non-typed fallback
          return {
            ...r,
            itemSizeSelected: r.itemSizeSelected === sizeVal ? undefined : sizeVal,
            selectedType: undefined,
          };
        })
      );
    },
    []
  );

  const updateItemQty = useCallback((rowIndex: number, qty?: string) => {
    setItemRows(prev => prev.map((r, i) => i === rowIndex ? { ...r, qty: qty } : r));
  }, []);

  // ---------------------------
  // Handlers
  // ---------------------------

  const handleEmployeeChange = useCallback(async (items?: IPersonaProps[], selectedOption?: string) => {

    if (items && items.length > 0) {
      const selected = items[0];
      // First try to find in employees list by FullName (fullName -> persona.text)
      const emp = employees.find(e => (e.fullName || '').toLowerCase() === (selected?.text || '').toLowerCase());
      // Fallback to users (Graph) if not found
      const user = users.find(u => u.displayName?.toLowerCase() === (selected?.text || '').toLowerCase() || u.id === selected?.id);

      const jobTitle: ICommon = emp?.jobTitle
        ? { id: emp.jobTitle.id ? String(emp.jobTitle.id) : undefined, title: emp.jobTitle.title || '' }
        : { id: undefined, title: user?.jobTitle || '' };

      const department: ICommon = emp?.department
        ? { id: emp.department.id ? String(emp.department.id) : undefined, title: emp.department.title || '' }
        : { id: undefined, title: user?.department || '' };

      const division: ICommon | undefined = emp?.division
        ? { id: emp.division.id ? String(emp.division.id) : undefined, title: emp.division.title || '' }
        : { id: undefined, title: '' };

      const company: ICommon = emp?.company
        ? { id: emp.company.id ? String(emp.company.id) : undefined, title: emp.company.title || '' }
        : { id: undefined, title: user?.company || '' };

      setEmployee([selected]);
      setEmployeeId(emp?.employeeID);
      setJobTitleId(jobTitle);
      setDepartmentId(department);
      setDivisionId(division);
      setCompanyId(company);
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
            selectedSizesByType: {},
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
      setJobTitleId({ id: '', title: '' });
      setDepartmentId({ id: '', title: '' });
      setDivisionId({ id: '', title: '' });
      setCompanyId({ id: '', title: '' });
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
    // const employeeMatches = employees
    //   .filter(e => (e.fullName || '').toLowerCase().includes(lower))
    //   .map(e => ({ text: e.fullName || '', secondaryText: e.jobTitle?.title, id: (e.Id ? String(e.Id) : e.fullName) }) as IPersonaProps);
    const userMatches = users
      .filter(u => (u.displayName || '').toLowerCase().includes(lower))
      .map(u => ({ text: u.displayName || '', secondaryText: u.jobTitle, id: u.id }) as IPersonaProps);
    // const combined = employeeMatches.concat(userMatches);
    // const deduped: IPersonaProps[] = [];
    // const seen = new Set<string>();
    // combined.forEach(p => { const key = (p.text || '').toLowerCase(); if (!seen.has(key)) { seen.add(key); deduped.push(p); } });
    // return limitResults ? deduped.slice(0, limitResults) : deduped;
    return userMatches
  }, [users]);

  const handleRequesterChange = useCallback(async (items?: IPersonaProps[], selectedOption?: string) => {
    if (items && items.length) setRequester([items[0]]); else setRequester([]);
  }, []);

  const handleNewRequestChange = useCallback((ev: React.FormEvent<HTMLElement>, checked?: boolean) => {
    if (checked) setIsReplacementChecked(false);
  }, []);

  const handleReplacementChange = useCallback((ev: React.FormEvent<HTMLElement>, checked?: boolean) => {
    setIsReplacementChecked(!!checked);
  }, []);

  const handleApprovalChange = useCallback((id: number | string, field: string, value: any) => {
    setFormsApprovalWorkflow(prev => {
      if (!prev || prev.length === 0) return prev;

      const i = prev.findIndex(r => String(r.Id ?? '') === String(id));
      if (i < 0) return prev;

      // Block edits unless the logged-in user is the Department Manager for this row
      if (!canEditApprovalRow(prev[i])) return prev;

      const next = [...prev];
      const row: any = { ...next[i] };

      switch (field) {
        case 'Status':
          row.Status = value ? String(value.key) : '';
          break;
        case 'Reason':
          row.Reason = value ?? '';
          break;
        case 'Date':
          row.Date = value ? new Date(value) : undefined;
          break;
        default:
          row[field] = value;
      }
      row.__index = i;
      next[i] = row;

      return next;
    });
  },
    [canEditApprovalRow]
  );

  const showBanner = useCallback((text: string) => {
    setBannerText(text);
    setBannerTick(t => t + 1);
  }, []);

  // const handleSave = useCallback(async () => {
  //   try {
  //     showBanner(undefined);
  //     setIsSaving(true);
  //     const payload = formPayload('Draft');
  //     // TODO: Wire to SharePoint persistence here.
  //     // console.log as a placeholder so you can see the shape:
  //     console.log('Save payload (Draft):', payload);
  //     showBanner('Draft saved (demo). Hook this up to SharePoint to persist.');
  //   } catch (e) {
  //     // console.error(e);
  //     showBanner('Failed to save draft.');
  //   } finally {
  //     setIsSaving(false);
  //   }
  // }, [formPayload]);

  const handleSubmit = useCallback(async () => {
    try {
      const validationError = validateBeforeSubmit();
      if (validationError) {
        showBanner(validationError);
        return;
      }

      setIsSubmitting(true);
      const payload = formPayload('Submitted');
      _createPPEForm(payload).then(async (newId) => {
        await _createPPEItemDetailsRows(newId, payload);
        // await _createPPEApprovalsRows(newId, payload);
        showBanner('PPE Form is submitted Successfully.');

      }).catch(err => {
        showBanner('Submit info Error:' + err.message + '. Please try again.');
      });
    } catch (e) {
      showBanner('Failed to submit. Please try again.');
    } finally {
      setIsSubmitting(false);
    }
  }, [formPayload, validateBeforeSubmit, showBanner]);


  //   // Create parent PPEForm item and return its Id
  const _createPPEForm = useCallback(async (payload: ReturnType<typeof formPayload>): Promise<number> => {
    const requesterEmail = emailFromPersona(_requester?.[0]) || loggedInUser?.email;
    const submitterEmail = emailFromPersona(_submitter?.[0]) || loggedInUser?.email;
    const requesterId = await ensureUserId(requesterEmail);
    const submitterId = await ensureUserId(submitterEmail);

    const _employeeSPId = _employee ? Number(_employee[0]?.id) : undefined;
    if (_employeeSPId == null) throw new Error('Employee is required');

    const body = {
      EmployeeRecordId: _employeeSPId,
      // EmployeeNameId: _employeeSPId, // lookup to Employee list
      SubmitterNameId: submitterId ?? null, // SharePoint person field
      RequesterNameId: requesterId ?? null, // SharePoint person field
      JobTitleRecordId: _jobTitle?.id ? Number(_jobTitle.id) : null,
      CompanyRecordId: _company?.id ? Number(_company.id) : null,
      DivisionRecordId: _division?.id ? Number(_division.id) : null,
      DepartmentRecordId: _department?.id ? Number(_department.id) : null,
      ReasonForRequest: payload.requestType ?? null,
      ReplacementReason: payload.replacementReason ?? null,
      EmployeeID: payload.employeeId ?? null,
    };
    const listGuid = sharePointLists.PPEForm?.value;
    spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl,
      listGuid, '');
    const data = await spCrudRef.current._insertItem(body);
    if (!data) throw new Error('Failed to create PPE Form');
    return data as number;
  }, [sharePointLists.PPEForm, emailFromPersona, ensureUserId, formPayload, _requester, _submitter, loggedInUser, props.context.spHttpClient]);

  // // Create detail rows for each required item
  const _createPPEItemDetailsRows = useCallback(async (parentId: number, payload: ReturnType<typeof formPayload>) => {
    const listGuid = sharePointLists.PPEFormItems?.value;
    const requiredItems = (payload.items || []).filter(i => i.required);
    if (requiredItems.length === 0) return;
    const posts = requiredItems.map(item => {
      const itemId = item?.itemId != null ? Number(item.itemId) : undefined;
      const detailId = item?.selectedDetailId != null ? Number(item.selectedDetailId) : undefined;

      // Map fields to your PPEItemDetails list’s internal names
      const body = {
        PPEFormIDId: parentId,
        ItemId: itemId ?? null,
        IsRequiredRecord: item.required ?? null,
        Brands: item.brand ?? null,
        Quantity: item.qty ?? null,
        Size: item.size ?? null,
        PPEFormItemDetailId: detailId ?? null,
        OthersPurpose: item.othersText ?? null,
      };
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, listGuid, '');
      const data = spCrudRef.current._insertItem(body);
      if (!data) throw new Error('Failed to create PPE Item Details');
      return data;
    });
    await Promise.all(posts);
  }, [sharePointLists.PPEFormItems, props.context.spHttpClient]);

  // const _createBulkPPEApprovalsRows = useCallback(async (parentId: number, payload: ReturnType<typeof formPayload>) => {
  //   const listGuid = sharePointLists.PPEFormApprovalWorkflow?.value;
  //   if (!listGuid) return; // not configured

  //   const rows = (payload.approvals || []);
  //   if (rows.length === 0) return;

  //   const posts = rows.map(async row => {
  //     const dmEmail = emailFromPersona(row.DepartmentManager);
  //     const dmId = await ensureUserId(dmEmail);
  //     const statusLookup = lKPWorkflowStatus.find(s => (s.Title || '').toLowerCase() === String(row.Status || '').toLowerCase());

  //     const body: any = {
  //       PPEFormId: Number(parentId) || null,
  //       ApproverId: dmId || null, // SharePoint Person field
  //       FormApprovalsWorkflowRecordId: row?.Id || null, // SharePoint Lookup field
  //       StatusRecordId: Number(statusLookup?.Id) || null, // SharePoint Lookup field
  //       Reason: row.Reason || null,
  //     };
  //     spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, listGuid, '');
  //     const data = spCrudRef.current._insertItem(body);
  //     if (!data) throw new Error('Failed to create PPE Item Details');
  //     return data;
  //   });

  //   await Promise.all(posts);
  // }, [sharePointLists.PPEFormApprovalWorkflow, ensureUserId, emailFromPersona, lKPWorkflowStatus, props.context.spHttpClient]);

  // const _createPPEApprovalsRows = useCallback(async (parentId: number, payload: ReturnType<typeof formPayload>) => {
  //   const listGuid = sharePointLists.PPEFormApprovalWorkflow?.value;
  //   if (!listGuid) return; // not configured

  //   const rows = (payload.approvals || []);
  //   if (rows.length === 0) return;
  //   const firstLevelApproval = rows.filter(r => r.Order === 1);
  //   if (firstLevelApproval.length === 0) return;
  //   else if (firstLevelApproval.length >= 1) {

  //     const dmEmail = emailFromPersona(firstLevelApproval[0].DepartmentManager);
  //     const dmId = await ensureUserId(dmEmail);
  //     const statusLookup = lKPWorkflowStatus.find(s => (s.Title || '').toLowerCase() === String(firstLevelApproval[0].Status || '').toLowerCase());

  //     const body: any = {
  //       PPEFormId: Number(parentId) || null,
  //       ApproverId: dmId || null, // SharePoint Person field
  //       FormApprovalsWorkflowRecordId: firstLevelApproval[0]?.Id || null, // SharePoint Lookup field
  //       StatusRecordId: Number(statusLookup?.Id) || null, // SharePoint Lookup field
  //       Reason: firstLevelApproval[0].Reason || null,
  //     };
  //     spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, listGuid, '');
  //     const data = spCrudRef.current._insertItem(body);

  //     if (!data) throw new Error('Failed to create PPE Form Workflow');
  //     return data
  //   }

  // }, [sharePointLists.PPEFormApprovalWorkflow, ensureUserId, emailFromPersona, lKPWorkflowStatus, props.context.spHttpClient]);

  // Save everything to SharePoint (parent + details + approvals)
  // const saveToSharePoint = useCallback(async () => {
  //   const payload = formPayload('Submitted');
  //   const parentId = await createPPEForm(payload);
  //   await createPPEItemDetailsRows(parentId, payload);
  //   await createPPEApprovalsRows(parentId, payload); // optional
  //   return parentId;
  // }, [formPayload, createPPEForm, createPPEItemDetailsRows, createPPEApprovalsRows]);

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
    <div className={styles.ppeFormBackground} ref={containerRef}>
      <div ref={bannerTopRef} />
      {bannerText && <MessageBar styles={{ root: { marginBottom: 8, color: 'red' } }}>{bannerText}</MessageBar>}
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
                selectedItems={_employee}
                onChange={(items) => {
                  const selectedText = items?.[0]?.text || '';
                  const empId = employees.find(e => (e.fullName || '').toLowerCase() === selectedText.toLowerCase())?.Id;
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
              <TextField label="Job Title" value={_jobTitle?.title} disabled={true} />
            </div>
            <div className="form-group col-md-6">
              <TextField label="Department" value={_department?.title} disabled={true} />
            </div>
          </div>

          <div className="row">
            <div className="form-group col-md-6"><TextField label="Division" value={_division?.title} disabled={true} /></div>
            <div className="form-group col-md-6"><TextField label="Company" value={_company?.title} disabled={true} /></div>
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
                selectedItems={_requester}
              />
            </div>

            <div className="form-group col-md-6">
              <NormalPeoplePicker label={"Submitter Name"} itemLimit={1} onResolveSuggestions={onFilterChanged} className={'ms-PeoplePicker'} key={'normal'} removeButtonAriaLabel={'Remove'} inputProps={{ onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'), onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'), 'aria-label': 'People Picker' }} onInputChange={onInputChange} resolveDelay={300} disabled={true} selectedItems={_submitter} />
            </div>
          </div>

          <div className={`row  ${styles.mt10}`}>
            <div className="form-group col-md-12 d-flex justify-content-between" >
              <Label htmlFor={""}>Reason for Request</Label>

              <Checkbox label="New Request" className="align-items-center" checked={!_isReplacementChecked} onChange={handleNewRequestChange} />

              <Checkbox label="Replacement" className="align-items-center" checked={_isReplacementChecked} onChange={handleReplacementChange} />

              <TextField placeholder="Reason" disabled={!_isReplacementChecked} value={_replacementReason}
                onChange={(_e, v) => setReplacementReason(v || '')} />
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
                items={itemRows.sort((a, b) => (a.order ? a.order : 0) - (b.order ? b.order : 0))}
                setKey="ppeAggregatedItemsList"
                selectionMode={SelectionMode.none}
                layoutMode={DetailsListLayoutMode.fixedColumns}
                columns={[
                  {
                    key: 'colItem', name: 'Item', fieldName: 'item', minWidth: 80, isResizable: true,
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
                                  <TextField placeholder={detail} multiline autoAdjustHeight resizable
                                    scrollContainerRef={containerRef} styles={{ root: { width: '100%' } }}
                                    value={r.otherPurpose ?? undefined}
                                    disabled={!r.required}
                                    key={`purpose-${r.itemId}-${r.required ? 'on' : 'off'}`}
                                    // eslint-disable-next-line react/jsx-no-bind
                                    onChange={(ev, newValue) => updateOtherPurpose(itemRows.indexOf(r), newValue ?? '')}
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

                  // ...existing code...
                  {
                    key: 'colSizes', name: 'Size', fieldName: 'size', minWidth: 200, isResizable: true,
                    onRender: (r: ItemRowState) => {
                      if (r.item.toLowerCase() === 'others') {
                        // Show Sizes only if Required is checked
                        if (!r.required) return <span />;

                        const sizes = Array.from(new Set((r.itemSizes || []).map(s => String(s).trim()).filter(Boolean)))
                          .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));

                        return (
                          <div key={r.item} style={{ display: 'flex', alignItems: 'center', marginBottom: 4 }}>
                            <ComboBox
                              placeholder={sizes.length ? 'Size' : 'No sizes'}
                              selectedKey={r.itemSizeSelected || undefined}
                              options={sizes.map(s => ({ key: s, text: s }))}
                              styles={{ root: { width: 140 } }}
                              disabled={!sizes.length}
                              onChange={(_e, opt) => {
                                const val = opt?.key ? String(opt.key) : undefined;
                                // If cleared, consider it as unchecked
                                toggleSize(itemRows.indexOf(r), val, !!val);
                              }}
                            />
                          </div>
                        );
                      }

                      // If Types exist, render types next to each other (horizontally) with a vertical separator.
                      // Under each type label, stack the sizes vertically (one per line).
                      const hasTypes = r.types && r.types.length > 0;
                      if (hasTypes) {
                        return (
                          <div style={{ display: 'flex', gap: 12, alignItems: 'flex-start' }}>

                            {(r.types || []).map((type, idx) => {
                              const sizesForType = (r.typeSizesMap && r.typeSizesMap[type]) || r.itemSizes || [];
                              const sizes = Array.from(new Set(sizesForType.map(s => String(s).trim()).filter(Boolean)))
                                .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));

                              return (
                                <div key={type}
                                  style={{ display: 'flex', flexDirection: 'column', gap: 4, paddingLeft: idx === 0 ? 0 : 12, marginLeft: idx === 0 ? 0 : 12, borderLeft: idx === 0 ? 'none' : '1px solid #ddd' }}>
                                  <Label styles={{ root: { marginBottom: 4, fontWeight: 600 } }}>{type}</Label>

                                  {sizes.length === 0 ? (<span>N/A</span>) :
                                    (
                                      <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                                        {sizes.map(size => {
                                          const sizeChecked = r.selectedSizesByType?.[type] === size;
                                          const id = `${r.itemId}-${type}-${size}`;
                                          return (
                                            <div key={`${type}-${size}`} style={{ display: 'flex', alignItems: 'center' }}>
                                              <Checkbox
                                                id={id}
                                                label={size}
                                                checked={sizeChecked}
                                                onChange={(_e, ch) => toggleSizeType(itemRows.indexOf(r), size, !!ch, type, id)}
                                                styles={{
                                                  root: { alignItems: 'flex-start' },
                                                  label: {
                                                    whiteSpace: 'normal', wordWrap: 'break-word', overflowWrap: 'anywhere', lineHeight: '1.3'
                                                  }
                                                }}
                                              />
                                            </div>
                                          );
                                        })}
                                      </div>
                                    )
                                  }
                                </div>
                              );
                            })
                            }
                          </div>
                        );
                      }

                      // No types: original sizes grid
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
                                <Checkbox
                                  label={size}
                                  checked={sizeChecked}
                                  onChange={(_e, ch) => toggleSize(itemRows.indexOf(r), size, !!ch)}
                                  styles={{
                                    root: { alignItems: 'flex-start' },
                                    label: { whiteSpace: 'normal', wordWrap: 'break-word', overflowWrap: 'anywhere', lineHeight: '1.3' }
                                  }}
                                />
                              </div>
                            );
                          })}
                        </div>
                      );
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
              <div style={{ backgroundColor: "#f3f2f1", padding: 10, borderRadius: 4 }}>
                {itemInstructionsForUse.map((instr: ILKPItemInstructionsForUse, idx: number) => (
                  <MessageBar key={instr.Id ?? instr.Order} isMultiline styles={{ root: { marginBottom: 6 } }}>
                    <strong>{`${idx + 1}. `}</strong>
                    {instr.Description}
                  </MessageBar>
                )
                )}
              </div>
            </div>
          )}
        </Stack>

        <Separator />
        {/* Approvals sign-off table */}
        <Stack horizontal styles={stackStyles} className="mt-3 mb-3">
          <div>
            <Label>Approvals / Sign-off</Label>
            <DetailsList
              items={formsApprovalWorkflow}
              columns={[
                {
                  key: 'colSignOff', name: 'Sign off', fieldName: 'SignOffName', minWidth: 120, isResizable: true,
                  onRender: (item: any) => (
                    <div>
                      <span>{item.SignOffName}</span>
                    </div>
                  )
                },
                {
                  key: 'colDepartmentManager', name: 'Department Manager', fieldName: 'DepartmentManager', minWidth: 180, isResizable: true,
                  onRender: (item: any, idx?: number) => {
                    return (
                      <div style={{ minWidth: 130 }}>
                        <NormalPeoplePicker
                          itemLimit={1}
                          required={true}
                          onResolveSuggestions={onFilterChanged}
                          disabled={item.DepartmentManager !== undefined}
                          selectedItems={item.DepartmentManager ? [item.DepartmentManager] : []}
                          resolveDelay={300}
                          inputProps={{ 'aria-label': 'Approvee' }}
                        />
                      </div>
                    );
                  }
                },
                {
                  key: 'colStatus', name: 'Status', fieldName: 'Status', minWidth: 130, isResizable: true,
                  onRender: (item: any, idx?: number) => {
                    const sorted = (lKPWorkflowStatus || []).slice()
                      .sort((a, b) => {
                        const ao = a?.Order ?? Number.POSITIVE_INFINITY;
                        const bo = b?.Order ?? Number.POSITIVE_INFINITY;
                        return Number(ao) - Number(bo);
                      });


                    const isFinalApprover = !!item.IsFinalFormApprover;
                    const closedId = sorted.find(s => (s.Title || '').toLowerCase() === 'closed')?.Id;

                    const options = sorted.map(s => {
                      const id = String(s.Id);
                      const title = String(s.Title ?? '').trim();
                      const isClosed = s.Id === closedId || title.toLowerCase() === 'closed';
                      return { key: id, text: title, disabled: !isFinalApprover && isClosed, };
                    });
                    const selectedKey = item.Status ? String(item.Status) : undefined;

                    return (
                      <ComboBox
                        placeholder={options.length ? 'Select status' : 'No status'}
                        selectedKey={selectedKey}
                        options={options}
                        useComboBoxAsMenuWidth={true}
                        disabled={!canEditApprovalRow(item)}
                        onChange={(_, option) => handleApprovalChange(item.Id!, 'Status', option)}
                      />
                    );
                  }
                },
                {
                  key: 'colReason', name: 'Reason', fieldName: 'Reason', minWidth: 160, isResizable: true,
                  onRender: (item: any, idx?: number) => (
                    <TextField value={item.Reason || ''}
                      disabled={!canEditApprovalRow(item)}
                      onChange={(ev, newValue) => handleApprovalChange(item.Id!, 'Reason', newValue || '')}
                    />)
                },
                {
                  key: 'colDate', name: 'Date', fieldName: 'Date', minWidth: 140, isResizable: true,
                  onRender: (item: any, idx?: number) => (
                    <DatePicker value={item.Date ? new Date(item.Date) : new Date()}
                      disabled={!canEditApprovalRow(item)}
                      onSelectDate={(date) => handleApprovalChange(item.Id!, 'Date', date || undefined)}
                      strings={defaultDatePickerStrings}
                    />)

                }
              ]}
              selectionMode={SelectionMode.none}
              setKey="approvalsList"
              layoutMode={DetailsListLayoutMode.fixedColumns}
              styles={{
                // target cells and rows
                contentWrapper: {
                  selectors: {
                    '.ms-DetailsRow-fields': {
                      alignItems: 'center'  // stretch to max height of tallest cell in the row
                    },
                    '.ms-DetailsRow-cell': {
                      padding: '8px 0px 8px 8px !important', // top-bottom left-right
                    },
                  }
                }
              }}
            />
          </div>
        </Stack>
        <Separator />

        <DocumentMetaBanner docCode="COR-HSE-01-FOR-001" version="V03" effectiveDate="16-SEP-2020" page={1} />

        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 8 }}>
          {/* <DefaultButton
            text={isSaving ? 'Saving…' : 'Save as Draft'}
            onClick={handleSave}
            disabled={isSaving || isSubmitting}
          /> */}
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
