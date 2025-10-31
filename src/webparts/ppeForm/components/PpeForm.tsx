import * as React from "react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { ISPHttpClientOptions, MSGraphClientV3, SPHttpClient } from "@microsoft/sp-http";
import { ICommon, IGraphResponse, IGraphUserResponse, ILKPItemInstructionsForUse, ISPListItem } from "../../../Interfaces/Common/ICommon";

// Components
import type { IPpeFormWebPartProps } from "./IPpeFormProps";
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';
import { TextField, ITextFieldStyles } from '@fluentui/react/lib/TextField';
import { Stack, IStackStyles } from '@fluentui/react/lib/Stack';
import {
  DatePicker, defaultDatePickerStrings, ConstrainMode,
  IBasePickerStyles, IComboBoxStyles, IDatePickerStyles
} from '@fluentui/react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar } from '@fluentui/react/lib/MessageBar';
import { PrimaryButton, DefaultButton } from '@fluentui/react';
import ExportPdfControls from './ExportPdfControls';
import {
  DetailsList, DetailsListLayoutMode, SelectionMode, Label, Separator,
  ComboBox, DefaultPalette, Checkbox
} from '@fluentui/react';

// Styles
import "bootstrap/dist/css/bootstrap.min.css";
import styles from "./PpeForm.module.scss";

// Classes
import { SPCrudOperations } from "../../../Classes/SPCrudOperations";
import { SPHelpers } from "../../../Classes/SPHelpers";
import { ICoralFormsList } from "../../../Interfaces/Common/ICoralFormsList";
import { IUser } from "../../../Interfaces/Common/IUser";
import { IPPEItemDetails } from "../../../Interfaces/PpeForm/IPPEItemDetails";
import { IEmployeeProps, IEmployeesPPEItemsCriteria } from "../../../Interfaces/PpeForm/IEmployeeProps";
import { IFormsApprovalWorkflow } from "../../../Interfaces/PpeForm/IFormsApprovalWorkflow";
import { IPPEItem } from "../../../Interfaces/PpeForm/IPPEItem";
import { DocumentMetaBanner } from "./DocumentMetaBanner";
import BannerComponent, { BannerKind } from "./BannerComponent";


const textFieldBlackStyles: Partial<ITextFieldStyles> = {
  // Applies to both input and textarea
  field: {
    color: '#000', // <-- main text
    selectors: {
      '&::placeholder': { color: '#666', fontWeight: 500, },        // optional: darker placeholder
      '&:disabled': { color: '#000', fontWeight: 500, }             // ensure disabled still renders black
    },
    subComponentStyles: {
      label: { root: { color: '#000', fontWeight: 500, } }
    }
  }
};

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

const peoplePickerBlackStyles: Partial<IBasePickerStyles> = {
  text: {
    selectors: {
      '.primaryText': { color: '#000 !important', fontWeight: '500 !important', },
      '.ms-Persona-primaryText': { color: '#000 !important', fontWeight: '500 !important', },
      '.ms-BasePicker-input': { color: '#000 !important', fontWeight: '500 !important', },
      '&.is-disabled .ms-BasePicker-input': { color: '#000 !important', fontWeight: '500 !important', }
    }
  },
  input: { color: '#000 !important', fontWeight: '500 !important', }
};

const datePickerBlackStyles: Partial<IDatePickerStyles> = {
  root: { width: '100%', selectors: { '> *': { marginBottom: 15 } } },
  readOnlyTextField: {
    selectors: {
      '&.is-disabled .ms-TextField-field': { color: '#000 !important', fontWeight: 500, '-webkit-text-fill-color': '#000 !important' },
      '.field': { color: '#000 !important', fontWeight: 500, },
    }
  },
  textField: {
    selectors: {
      '&.is-disabled .ms-TextField-field': { color: '#000 !important', fontWeight: 500, '-webkit-text-fill-color': '#000 !important' },
      '.field': { color: '#000 !important', fontWeight: 500, },
    },
    field: { color: '#000 !important', fontWeight: 500, },
    root: { color: '#000  !important' },
    suffix: { color: '#000' },
    description: { color: '#000  !important' },
    fieldGroup: {
      // keep disabled background clean
      selectors: { '&.is-disabled': { background: 'transparent' } }
    }
  },
  icon: { color: '#000  !important' }
};


export default function PpeForm(props: IPpeFormWebPartProps) {
  // Helpers and refs
  const formName = "PERSONAL PROTECTIVE EQUIPMENT";
  const spHelpers = useMemo(() => new SPHelpers(), []);
  const spCrudRef = useRef<SPCrudOperations | undefined>(undefined);
  const containerRef = React.useRef<HTMLDivElement>(null);
  const bannerTopRef = useRef<HTMLDivElement>(null);
  const overlayRef = useRef<HTMLDivElement>(null);
  const [_jobTitle, setJobTitleId] = useState<ICommon>({ id: '', title: '' });
  const [_department, setDepartmentId] = useState<ICommon>({ id: '', title: '' });
  const [_company, setCompanyId] = useState<ICommon>({ id: '', title: '' });
  const [_employee, setEmployee] = useState<IPersonaProps[]>([]);
  const [_SPEmployeeId, setSPEmployeeId] = useState<number>();
  const [_coralEmployeeId, setCoralEmployeeId] = useState<string | undefined>(undefined);
  const [_submitter, setSubmitter] = useState<IPersonaProps[]>([]);
  const [_requester, setRequester] = useState<IPersonaProps[]>([]);
  const [_isReplacementChecked, setIsReplacementChecked] = useState(false);
  const [_isAccidentalChecked, setIsAccidentalChecked] = useState(false);
  const [_replacementReason, setReplacementReason] = useState<string>('');
  const [_coralReferenceNumber, setCoralReferenceNumber] = useState<string>('');
  const [users, setUsers] = useState<IUser[]>([]);
  const [employees, setEmployees] = useState<IEmployeeProps[]>([]);
  const [employeePPEItemsCriteria, setEmployeePPEItemsCriteria] = useState<IEmployeesPPEItemsCriteria>({ Id: '' });
  const [ppeItems, setPpeItems] = useState<IPPEItem[]>([]);
  const [ppeItemDetails, setPpeItemDetails] = useState<IPPEItemDetails[]>([]);
  const [itemInstructionsForUse, setItemInstructionsForUse] = useState<ILKPItemInstructionsForUse[]>([]);
  const [lKPWorkflowStatus, setLKPWorkflowStatus] = useState<ISPListItem[]>([]);
  const [formsApprovalWorkflow, setFormsApprovalWorkflow] = useState<IFormsApprovalWorkflow[]>([]);
  const [_coralFormsList, setCoralFormsList] = useState<ICoralFormsList>({ Id: "" });
  const [loading, setLoading] = useState<boolean>(true);
  const [isSubmitting, setIsSubmitting] = useState<boolean>(false); // Submit button state
  const [bannerText, setBannerText] = useState<string>();
  const [bannerTick, setBannerTick] = useState(0);
  const [prefilledFormId, setPrefilledFormId] = useState<number | undefined>(undefined);
  const [, setIsHSEApprovalLevel] = React.useState<boolean>(false);
  const [, setIsWarehouseApprovalLevel] = React.useState<boolean>(false);
  const [IsHSEgroupMembership, setHSEGroupMembership] = useState<boolean>(false);
  const [IsWarehouseGroupMembership, setWarehouseGroupMembership] = useState<boolean>(false);
  const [editableRows, setEditableRows] = useState<Record<number, boolean>>({});
  const [canChangeApprovalRows, setCanChangeApprovalRows] = useState<boolean>(false);
  const [IsEligibleToSubmitForm, setIsEligibleToSubmitForm] = useState<boolean>(true);
  const [groupMembers, setGroupMembers] = useState<Record<string, IPersonaProps[]>>({});
  const [, setLockedApprovalRowIds] = useState<Record<string, boolean>>({});
  const [itemRows, setItemRows] = useState<ItemRowState[]>([]);
  const [criteriaAppliedForEmployeeId, setCriteriaAppliedForEmployeeId] = useState<string | undefined>(undefined);
  const [bannerOpts, setBannerOpts] = React.useState<{ autoHideMs?: number; fade?: boolean; kind?: BannerKind } | undefined>();
  const [exportMode, setExportMode] = React.useState(false);
  const [isExportingPdf, setIsExportingPdf] = React.useState(false); // NEW
  const webUrl = props.context.pageContext.web.absoluteUrl;
  interface ItemRowState {
    itemId: number | undefined;  // unique key per row
    item: string;
    order?: number | undefined;             // original order for sorting
    brands: string[];            // all available brands for item
    brandSelected?: string;      // chosen brand
    requiredRecord: boolean | undefined;           // required flag per item
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

  // Build a consolidated Required Items summary for export in-place
  type ItemSummary = { item: string; detail?: string; quantity?: string; size?: string; brand?: string };
  const itemsSummary: ItemSummary[] = React.useMemo(() => {
    const rows = (itemRows || []).filter(r => !!r.requiredRecord);
    return rows.map(r => {
      const hasTypes = Array.isArray(r.types) && r.types.length > 0;
      const size = hasTypes
        ? Object.entries(r.selectedSizesByType || {})
          .filter(([, v]) => !!v && String(v).trim().length > 0)
          .map(([k, v]) => `${k}: ${v}`)
          .join('; ')
        : (r.itemSizeSelected || '');
      const detail = r.selectedDetail || (r.item.toLowerCase() === 'others' ? (r.otherPurpose || '') : '');
      return { item: r.item, detail, quantity: r.qty || '', size: size || '', brand: r.brandSelected || '' };
    });
  }, [itemRows]);

  const uiDisabled = React.useCallback((normalDisabled: boolean) => (exportMode ? false : normalDisabled), [exportMode]);
  const stackStyles: IStackStyles = React.useMemo(() => ({
    root: {
      display: 'inline',
      // Blue normally, transparent when exporting
      background: exportMode ? 'transparent' : DefaultPalette.themeTertiary,
    },
  }), [exportMode]);

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
      const query: string = `?$select=Id,CoralEmployeeID,FullName,EmailAddressRecord/EMail,Company/Id,Company/Title,EmploymentStatus,JobTitle/Id,JobTitle/Title,` +
        `Department/Id,Department/Title,Created,Author/EMail,DirectManager/Id,DirectManager/Title,DirectManager/EMail` +
        `&$expand=Author,Company,JobTitle,Department,Author,DirectManager,EmailAddressRecord` +
        `&$filter=substringof('${employeeFullName}', FullName)&$orderby=Order asc`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'Employee', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: IEmployeeProps[] = [];
      const usersToUse = usersArg && usersArg.length ? usersArg : users;
      data.forEach((obj: any) => {
        if (obj) {
          const createdBy = usersToUse && usersToUse.length ? usersToUse.filter(u => u.email?.toString() === obj.Author?.EMail?.toString())[0] : undefined;
          let created: Date | undefined;
          if (obj.Created !== undefined) created = new Date(spHelpers.adjustDateForGMTOffset(obj.Created));
          const DirectManager: IEmployeeProps | undefined = obj.DirectManager !== undefined && obj.DirectManager !== null ?
            { Id: obj.DirectManager.Id, fullName: obj.DirectManager.Title } as IEmployeeProps : undefined;

          const temp: IEmployeeProps = {
            Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            coralEmployeeID: obj.CoralEmployeeID !== undefined && obj.CoralEmployeeID !== null ? obj.CoralEmployeeID : 0,
            fullName: obj.FullName !== undefined && obj.FullName !== null ? obj.FullName : undefined,
            jobTitle: obj.JobTitle !== undefined && obj.JobTitle !== null ? { id: obj.JobTitle.Id, title: obj.JobTitle.Title } : undefined,
            company: obj.Company !== undefined && obj.Company !== null ? { id: obj.Company.Id, title: obj.Company.Title } : undefined,
            department: obj.Department !== undefined && obj.Department !== null ? { id: obj.Department.Id, title: obj.Department.Title } : undefined,
            manager: obj.DirectManager !== undefined && obj.DirectManager !== null ? DirectManager : undefined,
            employmentStatus: obj.EmploymentStatus !== undefined && obj.EmploymentStatus !== null ? obj.EmploymentStatus : undefined,
            Created: created !== undefined ? created : undefined,
            CreatedBy: createdBy !== undefined ? createdBy : undefined,
            EMailAddress: obj.EmailAddressRecord !== undefined && obj.EmailAddressRecord !== null ? obj.EmailAddressRecord.EMail : undefined,
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

  const _getEmployeesPPEItemsCriteria = useCallback(async (usersArg?: IUser[], employeeID?: number | undefined) => {
    try {
      const query: string = `?$select=Id,Employee/ID,Employee/FullName,Employee/CoralEmployeeID,Created,SafetyHelmet,ReflectiveVest,SafetyShoes,` +
        `RainSuit/Id,RainSuit/DisplayText,UniformCoveralls/Id,UniformCoveralls/DisplayText,UniformTop/Id,UniformTop/DisplayText,` +
        `UniformPants/Id,UniformPants/DisplayText,WinterJacket/Id,WinterJacket/DisplayText,Author/EMail,AdditionalPPEItems` +
        `&$expand=Author,Employee,RainSuit,UniformCoveralls,UniformTop,UniformPants,WinterJacket` +
        `&$filter=Employee/ID eq ${employeeID}&$orderby=Order asc`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'Employee_PPE_Items_Criteria', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      let result: IEmployeesPPEItemsCriteria;

      if (data && data.length > 0) {
        const obj = data[0]; // Get the first object
        result = {
          Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
          employeeID: obj.Employee !== undefined && obj.Employee !== null ? obj.Employee.ID : undefined,
          coralEmployeeID: obj.Employee !== undefined && obj.Employee !== null ? obj.Employee.CoralEmployeeID : undefined,
          fullName: obj.Employee !== undefined && obj.Employee !== null ? obj.Employee.FullName : undefined,
          reflectiveVest: obj.ReflectiveVest !== undefined && obj.ReflectiveVest !== null ? obj.ReflectiveVest : undefined,
          safetyHelmet: obj.SafetyHelmet !== undefined && obj.SafetyHelmet !== null ? obj.SafetyHelmet : undefined,
          safetyShoes: obj.SafetyShoes !== undefined && obj.SafetyShoes !== null ? obj.SafetyShoes : undefined,
          rainSuit: obj.RainSuit !== undefined && obj.RainSuit !== null ? obj.RainSuit.DisplayText : undefined,
          uniformCoveralls: obj.UniformCoveralls !== undefined && obj.UniformCoveralls !== null ? obj.UniformCoveralls.DisplayText : undefined,
          uniformTop: obj.UniformTop !== undefined && obj.UniformTop !== null ? obj.UniformTop.DisplayText : undefined,
          uniformPants: obj.UniformPants !== undefined && obj.UniformPants !== null ? obj.UniformPants.DisplayText : undefined,
          winterJacket: obj.WinterJacket !== undefined && obj.WinterJacket !== null ? obj.WinterJacket.DisplayText : undefined,
          additionalPPEItems: obj.AdditionalPPEItems !== undefined && obj.AdditionalPPEItems !== null ? obj.AdditionalPPEItems : undefined,
          Created: undefined, CreatedBy: undefined,
        };
        setEmployeePPEItemsCriteria(result);
      }

    } catch (error) {
      setEmployeePPEItemsCriteria({ Id: '' });
    }
  }, [props.context, spHelpers]);

  const _getCoralFormsList = useCallback(async (usersArg?: IUser[]): Promise<ICoralFormsList | undefined> => {
    try {

      const searchEscaped = formName.replace(/'/g, "''");
      const query: string = `?$select=Id,Title,hasInstructionForUse,hasWorkflow,Created,Author/EMail,SubmissionRangeInterval&$expand=Author&$filter=substringof('${searchEscaped}', Title)`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'Coral_Forms_List', query);
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
          SubmissionRangeInterval: ppeform.SubmissionRangeInterval ?? undefined,
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
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PPE_Items', query);
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
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PPE_Items_Details', query);
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
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'LKP_Item_Instructions_For_Use', query);
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

  const _getLKPWorkflowStatus = useCallback(async (usersArg?: IUser[], formName?: string) => {
    try {
      const query: string = `?$select=Id,Title,RecordOrder,Created,Author/EMail&$expand=Author&$orderby=RecordOrder asc`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'LKP_Workflow_Status', query);
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
            Order: obj.RecordOrder !== undefined && obj.RecordOrder !== null ? obj.Order : undefined,
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
      setLKPWorkflowStatus(result);
    } catch (error) {
      setLKPWorkflowStatus([]);
      // console.error('An error has occurred while retrieving items!', error);
    }
  }, [props.context, spHelpers]);

  const _getPPEFormApprovalWorkflows = useCallback(async (usersArg?: IUser[], formId?: number) => {
    try {
      const query: string = `?$select=Id,SignOffName,FinalLevel,Approver/Id,Approver/EMail,Approver/Title,Author/EMail,PPEForm/Id,ApproversName/Id,ApproversName/Title,ApproversName/EMail,` +
        `PPEForm/Title,IsFinalApprover,OrderRecord,Created,StatusRecord/Id,StatusRecord/Title,Reason,Modified,Editor/Id,Editor/EMail,Editor/Title` +
        `&$expand=Author,Editor,PPEForm,StatusRecord,Approver,ApproversName` +
        (formId && formId > 0 ? `&$filter=PPEForm/Id eq ${formId}` : '') +
        `&$orderby=OrderRecord asc`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PPE_Form_Approval_Workflow', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: IFormsApprovalWorkflow[] = [];
      const usersToUse = usersArg && usersArg.length ? usersArg : users;
      // Collect group member fetches and set state once at the end
      const membersAccumulator: Record<string, IPersonaProps[]> = {};

      for (const obj of data) {
        if (!obj) continue;

        const createdBy = usersToUse && usersToUse.length ? usersToUse.filter(u => u.email?.toString() === obj.Author?.EMail?.toString())[0] : undefined;
        let created: Date | undefined;
        const approverEmail = obj?.Editor?.EMail;
        const approverTitle = obj?.Editor?.Title;
        const match = (approverEmail && usersToUse.find(u => (u.email || '').toLowerCase() === String(approverEmail).toLowerCase()));

        const deptApproverPersona: IPersonaProps | undefined = match
          ? { text: match.displayName || approverTitle || '', secondaryText: match.email || match.jobTitle || '', id: match.id }
          : (approverTitle ? { text: approverTitle, secondaryText: approverEmail || '', id: String(obj.Editor?.Id ?? approverTitle) } as IPersonaProps : undefined);
        const deptApproverGroupPersona: IPersonaProps | undefined = { id: String(obj?.Approver?.Id), text: obj?.Approver?.Title, secondaryText: '' };

        let approvalDate: Date | undefined = undefined;
        if (obj.Created) {
          approvalDate = new Date(spHelpers.adjustDateForGMTOffset(obj.Created));
        } else if (obj.Modified) {
          approvalDate = new Date(spHelpers.adjustDateForGMTOffset(obj.Modified));
        }

        // Helper to normalize SharePoint multi-person fields to a simple array
        const toPeopleArray = (field: any): any[] => {
          if (!field) return [];
          if (Array.isArray(field)) return field;
          if (Array.isArray(field?.results)) return field.results;
          if (Array.isArray(field?.value)) return field.value;
          return [];
        };

        // Build IPersonaProps[] from ApproversName
        const approversPeople = toPeopleArray(obj.ApproversName);
        const approversPersonas: IPersonaProps[] = approversPeople.map((u: any) => ({
          text: u?.Title || u?.Email || u?.LoginName || '',
          secondaryText: u?.EMail || u?.Email || '',
          id: u?.Id != null ? String(u.Id) : (u?.LoginName || u?.Title || '')
        }) as IPersonaProps);

        // Use the group title from ApproverGroup (already mapped above as a Persona)
        const approverGroupTitle = (deptApproverGroupPersona?.text || '').trim();

        // Create the record keyed by group name
        const approversNamesList: Record<string, IPersonaProps[]> = {};
        if (approverGroupTitle) {
          approversNamesList[approverGroupTitle] = approversPersonas;
        }

        const temp: IFormsApprovalWorkflow = {
          Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
          FormName: obj.FormName !== undefined && obj.FormName !== null ? { title: obj.FormName.Title, id: obj.FormName.Id } : undefined,
          Order: obj.OrderRecord !== undefined && obj.OrderRecord !== null ? obj.OrderRecord : undefined,
          SignOffName: obj.SignOffName !== undefined && obj.SignOffName !== null ? obj.SignOffName : undefined,
          EmployeeId: obj.ManagerName !== undefined && obj.ManagerName !== null ? obj.ManagerName.Id : undefined,
          DepartmentManagerApprover: deptApproverPersona,
          ApproverGroup: deptApproverGroupPersona,
          FinalLevel: obj.FinalLevel !== undefined && obj.FinalLevel !== null ? obj.FinalLevel : false,
          IsFinalFormApprover: obj.IsFinalApprover !== undefined && obj.IsFinalApprover !== null ? obj.IsFinalApprover : false,
          Status: obj.StatusRecord !== undefined && obj.StatusRecord !== null ? { id: obj.StatusRecord.Id?.toString(), title: obj.StatusRecord.Title } : undefined,
          Reason: obj.Reason !== undefined && obj.Reason !== null ? obj.Reason : undefined,
          Date: approvalDate,
          Created: created !== undefined ? created : undefined,
          CreatedBy: createdBy !== undefined ? createdBy : undefined,
          ModifiedByPersona: obj.Editor !== undefined && obj.Editor !== null ? obj.Editor : undefined,
          ApproversNamesList: approversNamesList,
        };

        if (approverGroupTitle) {
          const key = approverGroupTitle.toLowerCase();
          if (!membersAccumulator[key]) {
            const members = await spCrudRef.current._getSharePointGroupMembers(approverGroupTitle);
            membersAccumulator[key] = members;
          }
          // setGroupMembers(prev => ({ ...prev, [approverGroupTitle.toLowerCase()]: approversPersonas }));
        }
        result.push(temp);
      }

      // sort by Order (ascending). If Order is missing, place those items at the end.
      result.sort((a, b) => {
        const aOrder = (a && a.Order !== undefined && a.Order !== null) ? Number(a.Order) : Number.POSITIVE_INFINITY;
        const bOrder = (b && b.Order !== undefined && b.Order !== null) ? Number(b.Order) : Number.POSITIVE_INFINITY;
        return aOrder - bOrder;
      });
      setFormsApprovalWorkflow(result);

      if (Object.keys(membersAccumulator).length) {
        setGroupMembers(prev => ({ ...prev, ...membersAccumulator }));
      }

    } catch (error) {
      setFormsApprovalWorkflow([]);
      // console.error('An error has occurred while retrieving items!', error);
    }
  }, [props.context, spHelpers]);

  const resolveGroupUserForItemRow = useCallback((row: IFormsApprovalWorkflow): string | undefined => {
    const fromGroup = row?.ApproverGroup?.text;
    if (!fromGroup) return undefined;
    const name = String(fromGroup).trim();
    return name.length ? name : undefined;
  }, []);
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

  const loggedInUser = useMemo(() => users.find(u => u.email === props.context.pageContext?.user?.email), [users]);

  // Determine if current user is Requester or Submitter (identity match by email/id/name)
  const isCurrentRequester = useMemo(() => {
    const currentEmail = (props.context.pageContext?.user?.email || '').toLowerCase();
    const requesterEmail = (emailFromPersona(_requester?.[0]) || '').toLowerCase();
    if (currentEmail && (currentEmail === requesterEmail)) return true;

    return false;
  }, [_requester, emailFromPersona]);

  const isCurrentSubmitter = useMemo(() => {
    const currentEmail = (props.context.pageContext?.user?.email || '').toLowerCase();
    const submitterEmail = (emailFromPersona(_submitter?.[0]) || '').toLowerCase();

    if (currentEmail && (currentEmail === submitterEmail)) return true;

    return false;
  }, [_submitter, emailFromPersona]);

  // Any approval row already marked as Approved?
  const anyApproved: boolean = useMemo(() => {
    return !!formsApprovalWorkflow?.find(r =>
      (r?.Status?.title?.trim().toLowerCase() === 'approved' && Number(r?.Order) >= 1)
    );
  }, [formsApprovalWorkflow]);

  const hasTopPendingForm: boolean = useMemo(() => {
    return !!formsApprovalWorkflow?.find(r => r?.Status?.title?.trim().toLowerCase() === 'pending' && r?.Order === 1);
  }, [formsApprovalWorkflow]);

  const isProcessedForm: boolean = useMemo(() => {
    return !!formsApprovalWorkflow && formsApprovalWorkflow.length === 0;
  }, [formsApprovalWorkflow]);

  const isProcessingHSEDepartment: boolean = useMemo(() => {
    if (!formsApprovalWorkflow || formsApprovalWorkflow.length < 3 || formsApprovalWorkflow.length > 3) return false;
    const hseLevel = formsApprovalWorkflow.find(r => r?.Order === 3);
    if (!hseLevel) return false;
    return hseLevel?.Status?.title?.trim().toLowerCase() === 'pending';
  }, [formsApprovalWorkflow]);

  const isProcessingWareHouseDepartment: boolean = useMemo(() => {
    if (!formsApprovalWorkflow || formsApprovalWorkflow.length < 4 || formsApprovalWorkflow.length > 4) return false;
    const wareHouseLevel = formsApprovalWorkflow.find(r => r?.Order === 4);
    if (!wareHouseLevel) return false;
    return wareHouseLevel?.Status?.title?.trim().toLowerCase() === 'pending';
  }, [formsApprovalWorkflow]);

  // Whether the form is in edit mode (has a valid formId)
  const isEditMode = useMemo(() => {
    const editFormId = props.formId ? Number(props.formId) : undefined;
    return !!(editFormId && editFormId > 0);
  }, [props.formId]);

  // New canEditForm: Requester/Submitter can edit header only if no approval is yet approved
  const canEditFormHeader = useMemo(() => {
    if (!isCurrentRequester && !isCurrentSubmitter) return false;
    if ((isCurrentRequester || isCurrentSubmitter) && !isEditMode) return true; // new form
    if ((isCurrentRequester || isCurrentSubmitter) && (hasTopPendingForm || isProcessedForm)) return true;
  }, [isCurrentRequester, isCurrentSubmitter, hasTopPendingForm, isProcessedForm]);

  // Derived permission: can edit items grid
  const canEditItems = useMemo(() => {
    if (IsHSEgroupMembership && isProcessingHSEDepartment) return true;
    if (IsWarehouseGroupMembership && isProcessingWareHouseDepartment) return true;
    if (isCurrentRequester && isCurrentSubmitter && anyApproved) return false; // Rule 1
    return !!canEditFormHeader;
  }, [IsHSEgroupMembership, IsWarehouseGroupMembership, isCurrentRequester, isCurrentSubmitter, anyApproved, canEditFormHeader]);

  // Determine which approval rows can be edited by current user
  const canEditApprovalRow = useCallback((item: IFormsApprovalWorkflow): boolean => {

    if (!item) {
      setCanChangeApprovalRows(false);
      return false;
    }

    // Always allow editing if already dirty (pending save)
    if ((item as any).__dirty) {
      setCanChangeApprovalRows(true);
      return true;
    }

    // Member + Pending gate (from async check)
    const isEditableByGroup = editableRows[Number(item.Id!)] === true;
    if (!isEditableByGroup) {
      setCanChangeApprovalRows(false);
      return false;
    }

    // Otherwise rely on cached group membership check
    return editableRows[Number(item.Id!)] === true;
  }, [editableRows]);

  const hasApprovalChanges = useMemo(() => {
    return (formsApprovalWorkflow || []).some(r => (r as any)?.__dirty === true);
  }, [formsApprovalWorkflow]);

  const formPayload = useCallback((status: 'Draft' | 'Submitted') => {

    const requestType = _isAccidentalChecked ? 'Accidental' : _isReplacementChecked ? 'Replacement' : 'New Request';
    const replacementReason = (_isReplacementChecked || _isAccidentalChecked) ? (_replacementReason?.trim() || undefined) : undefined;

    return {
      formName,
      status,
      employeeId: _SPEmployeeId,
      employeeName: _employee?.[0]?.text,
      _jobTitle,
      _department,
      _company,
      requestType: requestType,
      replacementReason: replacementReason,
      items: itemRows.map(r => {
        const hasTypes = r.types && r.types.length > 0;
        const sizeCsv = hasTypes ? r.types!.map(t => (r.selectedSizesByType?.[t] ?? '')).join(',') : (r.itemSizeSelected || '');
        const typeCsv = hasTypes ? r.types!.join(',') : (r.selectedType || '');
        return {
          itemId: r.itemId,
          item: r.item,
          requiredRecord: !!r.requiredRecord,
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
  }, [_employee, _SPEmployeeId, _jobTitle, _department, _company, _isReplacementChecked, _isAccidentalChecked, _replacementReason, itemRows, formsApprovalWorkflow, formName]);

  const validateBeforeSubmit = useCallback((): string | undefined => {
    const missing: string[] = [];
    if (!_employee?.[0]?.text?.trim()) missing.push('Employee Name');
    if (_requester.length === 0) missing.push('Requester');

    if (missing.length) {
      return `Please fill in the required fields: ${missing.join(', ')}.`;
    }

    // Example: if Replacement, require a reason
    if ((_isReplacementChecked || _isAccidentalChecked) && !(_replacementReason && _replacementReason.trim().length)) {
      return 'Please provide a reason for this request.';
    }

    // Ensure at least one item is required or has any selection
    const anyRequired = itemRows.some(r => r.requiredRecord);
    if (!anyRequired) return 'Please select at least one item or mark one as Required.';

    if (anyRequired) {
      const othersMissingPurpose = itemRows.some(r => r.item.toLowerCase() === 'others' && r.requiredRecord && (r.otherPurpose === undefined || !r.otherPurpose.trim()));
      if (othersMissingPurpose) return 'Please fill in the Purpose field for "Others" since it is marked Required.';

      const othersMissingSize = itemRows.some(r => r.item.toLowerCase() === 'others' && r.requiredRecord && (!r.itemSizeSelected || !r.itemSizeSelected.trim()));
      if (othersMissingSize) return 'Please choose a size for "Others" since it is marked Required.';

      // Validate each required item individually and stop on first failure
      for (const r of itemRows.filter(r => r.requiredRecord)) {
        if (!r.requiredRecord) continue;

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

    const nonApprovedForm = formsApprovalWorkflow.filter(item => item.DepartmentManagerApprover?.id === loggedInUser?.id && item.Status === undefined);
    if (nonApprovedForm && nonApprovedForm.length >= 1) { return 'Please update your approval status before submitting the form.'; }
    const rejectedWorkflowStatusId = lKPWorkflowStatus.find(p => p.Title?.toLowerCase().includes("rejected"));
    const rejectedForm = formsApprovalWorkflow.filter(item => item.DepartmentManagerApprover?.id === loggedInUser?.id && item.Status === rejectedWorkflowStatusId?.Id?.toString());
    if (rejectedForm && rejectedForm.length > 0 && rejectedForm[0]?.Reason === undefined) { return 'Please provide a reason for rejection before submitting the form.' };

    return undefined;
  }, [_employee, _jobTitle, _department, _company, _requester, itemRows, _isReplacementChecked, _isAccidentalChecked, _replacementReason, formsApprovalWorkflow]);

  // const _getGroupMembers = useCallback(async (goupName: string): Promise<IPersonaProps[]> => {
  //   const members: IPersonaProps[] = [];
  //   if (!goupName) return members;
  //   const name = String(goupName).trim();
  //   const webUrl = props.context.pageContext.web.absoluteUrl;
  //   const esc = (s: string) => s.replace(/'/g, "''");

  //   try {
  //     const url = `${webUrl}/_api/web/sitegroups/getbyname('${esc(name)}')/users?$select=Id,Title,Email,LoginName`;
  //     const resp: any = await (props.context as any).spHttpClient.get(url, SPHttpClient.configurations.v1);
  //     if (!resp || resp.status !== 200) {
  //       members;
  //     }
  //     const json = await resp.json();
  //     const personas: IPersonaProps[] = Array.isArray(json?.value) ? json.value.map((u: any) => ({
  //       text: u?.Title || u?.Email || u?.LoginName || '',
  //       secondaryText: u?.Email || '',
  //       id: (u?.Id != null ? String(u.Id) : (u?.LoginName || u?.Title || '')),
  //     } as IPersonaProps)) : [];
  //     return personas;
  //   }
  //   catch (ex) {
  //     return members;
  //   }
  // }, [props.context]);

  // Initial load of users, PPE items, Coral form config, etc.
  useEffect(() => {
    let cancelled = false;
    const load = async () => {
      setLoading(true);
      const fetchedUsers = await _getUsers();
      const coralListResult = await _getCoralFormsList(fetchedUsers);
      await _getPPEItems(fetchedUsers);
      await _getPPEItemsDetails(fetchedUsers);

      // Use the returned result from _getCoralFormsList instead of the (possibly stale) coralFormsList state
      if (coralListResult && coralListResult.hasInstructionForUse) {
        if (coralListResult.hasInstructionForUse) await _getLKPItemInstructionsForUse(fetchedUsers, formName);
        if (coralListResult.hasWorkflow) {
          await _getLKPWorkflowStatus(fetchedUsers);
          const editFormId = props.formId ? Number(props.formId) : undefined;
          if (editFormId && editFormId > 0) {
            await _getPPEFormApprovalWorkflows(fetchedUsers, editFormId);
          } else {
            setFormsApprovalWorkflow([]);
          }
        }
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
  }, [_getEmployees, _getUsers, _getPPEItems, _getPPEItemsDetails, _getCoralFormsList, _getLKPItemInstructionsForUse, _getPPEFormApprovalWorkflows, props.context, props.formId]);

  useEffect(() => {
    if (prefilledFormId) {
      _getPPEFormApprovalWorkflows(users, prefilledFormId);
    }
  }, [prefilledFormId, users, _getPPEFormApprovalWorkflows]);


  // Check when in edit mode and the 4th approval level is Warehouse Approval and set setIsWarehouseApprovalLevel
  useEffect(() => {

    const wareHouseLevel = (formsApprovalWorkflow && formsApprovalWorkflow.length === 4) ? formsApprovalWorkflow.
      find(wareHouseLevel => wareHouseLevel.Order === 4 && wareHouseLevel.SignOffName?.toLowerCase().includes('warehouse')) : undefined;

    if (!isEditMode || !wareHouseLevel) {
      setIsWarehouseApprovalLevel(false);
      return;
    }

    const groupTitle: string = (wareHouseLevel && wareHouseLevel.ApproverGroup && wareHouseLevel.ApproverGroup.text) ?
      String(wareHouseLevel.ApproverGroup.text) : 'WarehouseApproverGroup';

    let cancelled = false;
    const checkMembership = async () => {
      try {

        const spCrud = new SPCrudOperations((props.context as any).spHttpClient, webUrl, '', '');
        const isMember: boolean | undefined = await spCrud._IsSPGroup(groupTitle);
        if (!cancelled) setIsWarehouseApprovalLevel(isMember === true);
      } catch {
        if (!cancelled) setIsWarehouseApprovalLevel(false);
      }
    };

    checkMembership();

    return () => { cancelled = true; };
  }, [props.context, formsApprovalWorkflow]);

  // Check if current user is in WarehouseApproverGroup when the form has Warehouse Approval level
  useEffect(() => {
    const wareHouseLevel = (formsApprovalWorkflow && formsApprovalWorkflow.length === 4) ? formsApprovalWorkflow.
      find(wareHouseLevel => wareHouseLevel.Order === 4 && wareHouseLevel.SignOffName?.toLowerCase().includes('warehouse')) : undefined

    if (!wareHouseLevel) {
      return;
    }

    const groupTitle: string = (wareHouseLevel && wareHouseLevel.ApproverGroup && wareHouseLevel.ApproverGroup.text)
      ? String(wareHouseLevel.ApproverGroup.text) : 'Warehouse Approver Group';

    let cancelled = false;
    const checkInGroupMembership = async () => {
      try {
        const spCrud = new SPCrudOperations((props.context as any).spHttpClient, webUrl, '', '');
        const loggesUserEmail = props.context.pageContext?.user?.email || '';
        const isMember: boolean | undefined = await spCrud._IsUserInSPGroup(groupTitle, loggesUserEmail);
        if (!cancelled) setWarehouseGroupMembership(isMember === true);
      } catch {
        if (!cancelled) setWarehouseGroupMembership(false);
      }
    };

    checkInGroupMembership();

    return () => { cancelled = true; };
  }, [props.context, formsApprovalWorkflow]);

  // Check when in edit mode and the 3rd approval level is HSE Approval and set isHSEApprovalLevel
  useEffect(() => {

    const hseLEvel = (formsApprovalWorkflow && formsApprovalWorkflow.length === 3) ? formsApprovalWorkflow.
      find(hseLevel => hseLevel.Order === 3 && hseLevel.SignOffName?.toLowerCase().includes('hse')) : undefined;

    if (!isEditMode || !hseLEvel) {
      setIsHSEApprovalLevel(false);
      return;
    }

    const groupTitle: string = (hseLEvel && hseLEvel.ApproverGroup && hseLEvel.ApproverGroup.text) ? String(hseLEvel.ApproverGroup.text) : 'HSE Approvers Group';

    let cancelled = false;
    const checkMembership = async () => {
      try {
        const spCrud = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, '', '');
        const isMember: boolean | undefined = await spCrud._IsSPGroup(groupTitle);
        if (!cancelled) setIsHSEApprovalLevel(isMember === true);
      } catch {
        if (!cancelled) setIsHSEApprovalLevel(false);
      }
    };

    checkMembership();

    return () => { cancelled = true; };
  }, [props.context, formsApprovalWorkflow]);

  // Check if current user is in HSEApproverGroup when the form has HSE Approval level
  useEffect(() => {
    const hseLEvel = (formsApprovalWorkflow && formsApprovalWorkflow.length === 3) ? formsApprovalWorkflow.
      find(hseLevel => hseLevel.Order === 3 && hseLevel.SignOffName?.toLowerCase().includes('hse')) : undefined

    if (!hseLEvel) {
      return;
    }

    const groupTitle: string = (hseLEvel && hseLEvel.ApproverGroup && hseLEvel.ApproverGroup.text)
      ? String(hseLEvel.ApproverGroup.text) : 'HSEApproverGroup';

    let cancelled = false;
    const checkInGroupMembership = async () => {
      try {
        const spCrud = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, '', '');
        const loggesUserEmail = props.context.pageContext?.user?.email || '';
        const isMember: boolean | undefined = await spCrud._IsUserInSPGroup(groupTitle, loggesUserEmail);
        if (!cancelled) setHSEGroupMembership(isMember === true);
      } catch {
        if (!cancelled) setHSEGroupMembership(false);
      }
    };

    checkInGroupMembership();

    return () => { cancelled = true; };
  }, [props.context, formsApprovalWorkflow, loggedInUser]);

  // Check for each approval level if current user can edit it
  useEffect(() => {
    let cancelled = false;

    const checkMembership = async () => {
      if (!formsApprovalWorkflow || formsApprovalWorkflow.length === 0) {
        setEditableRows({});
        return;
      };

      const results: Record<number, boolean> = {};
      // let anyEditable = false;

      for (const item of formsApprovalWorkflow) {
        if (!item?.ApproverGroup?.text) {
          results[Number(item.Id!)] = false;
          continue;
        }

        try {
          const spCrud = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, '', '');
          const userEmail = props.context.pageContext?.user?.email ?? '';
          const isMember: boolean | undefined = await spCrud._IsUserInSPGroup(item.ApproverGroup.text, userEmail);
          results[Number(item.Id!)] = isMember === true && item.Status?.title?.toLowerCase() === 'pending';
        } catch {
          results[Number(item.Id!)] = false;
        }
      }

      if (!cancelled) {
        setEditableRows(results);
      }
    };

    checkMembership();

    return () => {
      cancelled = true;
    };
  }, [formsApprovalWorkflow, props.context, loggedInUser]);

  // Default the approver to the logged-in user if they’re in the row’s group and the row is Pending
  useEffect(() => {
    if (!formsApprovalWorkflow?.length) return;
    if (!props.context.pageContext?.user?.email) return;

    setFormsApprovalWorkflow(prev => {
      let changed = false;
      const next = (prev || []).map(row => {
        const grpName = resolveGroupUserForItemRow(row);
        if (!grpName) return row;
        const members = groupMembers[grpName.toLowerCase()] || [];
        if (!members.length) return row;

        const isPending = String(row?.Status?.title || '').toLowerCase() === 'pending';
        const isMember = members.some(m => (String(m.secondaryText || '').toLowerCase()) === props.context.pageContext?.user?.email);
        const hasApprover = !!row?.DepartmentManagerApprover?.secondaryText;

        if (isPending && isMember && !hasApprover) {
          const persona: IPersonaProps = {
            text: loggedInUser?.displayName || props.context.pageContext?.user?.email,
            secondaryText: props.context.pageContext?.user?.email,
            id: loggedInUser?.id || props.context.pageContext?.user?.email
          };
          changed = true;
          return { ...row, DepartmentManagerApprover: persona };
        }
        return row;
      });

      return changed ? next : prev;
    });
  }, [formsApprovalWorkflow, groupMembers, loggedInUser, resolveGroupUserForItemRow]);

  // Check if the Submitter can submit a new form within a 3 month period of time from the last submitted form
  const isEligibleToSubmit = useCallback(async (employeeId?: number, submissionDate?: Date): Promise<boolean> => {
    try {
      // If we can't determine the employee or date, don't block.
      if (!employeeId || !submissionDate) {
        return true;
      }

      // Query the latest PPE_Form created for this employee
      // Assumption: PPE_Form has a numeric "EmployeeID" column (same value you store in _employeeId)
      // If your list uses a lookup instead, adjust the filter to EmployeeRecord/Id eq {id} and add &$expand=EmployeeRecord
      const query = `?$select=Id,Created,EmployeeRecord/Id&$expand=EmployeeRecord` +
        `&$filter=EmployeeRecord/Id eq ${employeeId} and ReasonForRequest ne 'Accidental'` +
        `&$orderby=Created desc&$top=1`;
      const spCrud = new SPCrudOperations((props.context as any).spHttpClient, webUrl, 'PPE_Form', query);
      const items = await spCrud._getItemsWithQuery();
      if (!Array.isArray(items) || items.length === 0) {
        // No previous forms -> then allow
        setIsEligibleToSubmitForm(true);
        return true;
      }

      const createdRaw = items[0]?.Created;
      if (!createdRaw) {
        setIsEligibleToSubmitForm(true);
        return true;
      }

      // Normalize dates and compare in days
      const lastDate = new Date(spHelpers.adjustDateForGMTOffset(createdRaw));
      const intervalDays = _coralFormsList?.SubmissionRangeInterval ? _coralFormsList?.SubmissionRangeInterval : 90;
      const msPerDay = 1000 * 60 * 60 * 24;
      const diffDays = Math.floor((submissionDate.getTime() - lastDate.getTime()) / msPerDay);
      const eligible = diffDays >= intervalDays;
      setIsEligibleToSubmitForm(eligible);
      return eligible;

    } catch {
      // On any error, don't block the user
      setIsEligibleToSubmitForm(true);
      return true
    }
  }, [props.context, spHelpers, _coralFormsList]);

  useEffect(() => {
    let mounted = true;
    (async () => {
      if (!_SPEmployeeId || isEditMode) return;
      const ok = await isEligibleToSubmit(_SPEmployeeId, new Date());
      if (mounted) setIsEligibleToSubmitForm(ok); // true = allowed, false = blocked
    })();
    return () => { mounted = false; };
  }, [_SPEmployeeId, isEditMode]);

  // Prefill when editing an existing form
  useEffect(() => {
    const formId = props.formId;
    if (!formId || prefilledFormId === formId) return;
    // Wait until base items are loaded and itemRows initialized
    if (loading || itemRows.length === 0) return;

    let cancelled = false;

    const toPersona = (obj?: { Id?: any; Title?: string; EMail?: string; FullName?: string }): IPersonaProps | undefined => {
      if (!obj) return undefined;
      const text = obj.FullName || obj.Title || '';
      const email = obj.EMail || '';
      const id = obj.Id != null ? String(obj.Id) : text;
      return { text, secondaryText: email, id } as IPersonaProps;
    };

    const load = async () => {
      try {
        // Load PPEForm header by Id
        const headerQuery = `?$select=Id,ReasonForRequest,ReasonRecord,Created,CoralReferenceNumber,` +
          `EmployeeRecord/Id,EmployeeRecord/FullName,EmployeeRecord/CoralEmployeeID,` +
          `JobTitleRecord/Id,JobTitleRecord/Title,` +
          `DepartmentRecord/Id,DepartmentRecord/Title,` +
          `CompanyRecord/Id,CompanyRecord/Title,` +
          `RequesterName/Id,RequesterName/Title,RequesterName/EMail,` +
          `SubmitterName/Id,SubmitterName/Title,SubmitterName/EMail` +
          `&$expand=EmployeeRecord,JobTitleRecord,DepartmentRecord,CompanyRecord,RequesterName,SubmitterName` +
          `&$filter=Id eq ${formId}`;

        const formCrud = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PPE_Form', headerQuery);
        const headerItems = await formCrud._getItemsWithQuery();
        const header = Array.isArray(headerItems) ? headerItems[0] : undefined;

        if (header && !cancelled) {
          // Top-level fields prefill
          const employeePersona = toPersona({ Id: header?.EmployeeRecord?.Id, FullName: header?.EmployeeRecord?.FullName });
          setEmployee(employeePersona ? [employeePersona] : []);
          setSPEmployeeId(header?.EmployeeRecord?.Id != null ? Number(header.EmployeeRecord?.Id) : undefined);
          setCoralEmployeeId(header?.EmployeeRecord?.CoralEmployeeID != null ? String(header.EmployeeRecord?.CoralEmployeeID) : undefined);

          const jt = header?.JobTitleRecord ? { id: header.JobTitleRecord.Id ? String(header.JobTitleRecord.Id) : undefined, title: header.JobTitleRecord.Title || '' } : { id: undefined, title: '' };
          const dept = header?.DepartmentRecord ? { id: header.DepartmentRecord.Id ? String(header.DepartmentRecord.Id) : undefined, title: header.DepartmentRecord.Title || '' } : { id: undefined, title: '' };
          const comp = header?.CompanyRecord ? { id: header.CompanyRecord.Id ? String(header.CompanyRecord.Id) : undefined, title: header.CompanyRecord.Title || '' } : { id: undefined, title: '' };
          setJobTitleId(jt);
          setDepartmentId(dept);
          setCompanyId(comp);
          setCoralReferenceNumber(header?.CoralReferenceNumber || '');

          const requesterPersona = toPersona({ Id: header?.RequesterName?.Id, Title: header?.RequesterName?.Title, EMail: header?.RequesterName?.EMail });
          setRequester(requesterPersona ? [requesterPersona] : []);
          const submitterPersona = toPersona({ Id: header?.SubmitterName?.Id, Title: header?.SubmitterName?.Title, EMail: header?.SubmitterName?.EMail });
          setSubmitter(submitterPersona ? [submitterPersona] : []);

          const reason: string = header?.ReasonForRequest || '';
          setIsReplacementChecked(/replacement/i.test(reason));
          setReplacementReason(header?.ReasonRecord || '');
          setIsAccidentalChecked(/accidental/i.test(reason));
        }

        // Load child PPEFormItems rows for this form
        const itemsQuery = `?$select=Id,Brands,Quantity,Size,OthersPurpose,IsRequiredRecord,` +
          `PPEFormID/Id,Item/Id,Item/Title,PPEFormItemDetail/Id,PPEFormItemDetail/Title` +
          `&$expand=PPEFormID,Item,PPEFormItemDetail` +
          `&$filter=PPEFormID/Id eq ${formId}`;

        const itemsCrud = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, "PPE_Form_Items", itemsQuery);
        const childRows = await itemsCrud._getItemsWithQuery();

        if (!cancelled && Array.isArray(childRows)) {
          setItemRows(prev => prev.map(r => {
            const match = childRows.find((cr: any) => Number(cr?.Item?.Id) === Number(r.itemId));
            if (!match) return r;

            const next = { ...r } as any;
            // Presence of a child row means the item was required when saved.
            // Prefer the explicit flag if present; otherwise mark as required.
            next.requiredRecord = (typeof match.IsRequiredRecord !== 'undefined') ? !!match.IsRequiredRecord : true;
            next.brandSelected = match.Brands || undefined;
            next.qty = match.Quantity != null && match.Quantity !== '' ? String(match.Quantity) : undefined;

            // Detail title
            const detailTitle = match?.PPEFormItemDetail?.Title || (match?.PPEFormItemDetail?.Id ? (ppeItemDetails.find(d => Number(d.Id) === Number(match.PPEFormItemDetail.Id))?.Title) : undefined);
            if (detailTitle) next.selectedDetail = detailTitle;

            // Size mapping
            const sizeStr: string = match.Size || '';
            if (Array.isArray(r.types) && r.types.length > 0) {
              const parts = sizeStr.split(',');
              const byType: Record<string, string | undefined> = { ...(r.selectedSizesByType || {}) };
              (r.types || []).forEach((t, i) => {
                const val = (parts[i] || '').trim();
                byType[t] = val || undefined;
              });
              next.selectedSizesByType = byType;
              next.itemSizeSelected = undefined;
            } else {
              next.itemSizeSelected = sizeStr || undefined;
            }

            // Others purpose
            if ((r.item || '').toLowerCase() === 'others') {
              next.otherPurpose = match.OthersPurpose || undefined;
            }

            return next as typeof r;
          }));
        }

        if (!cancelled) setPrefilledFormId(formId);
      } catch (e) {
        // swallow prefill errors, show minimal message
        setBannerText('Failed to load the selected form for editing.');
        setBannerTick(t => t + 1);
      }
    };

    load();

    return () => { cancelled = true; };
  }, [props.formId, prefilledFormId, loading, itemRows.length, props.context, ppeItemDetails]);

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

  // When we start submitting/updating, scroll to where the loader overlay is rendered
  useEffect(() => {
    if (!isSubmitting) return;
    // Wait for overlay to render, then scroll it into view
    requestAnimationFrame(() => {
      if (overlayRef.current && overlayRef.current.scrollIntoView) {
        try { overlayRef.current.scrollIntoView({ behavior: 'smooth', block: 'center' }); } catch { /* ignore */ }
      } else if (containerRef.current) {
        try { containerRef.current.scrollTo({ top: 0, behavior: 'smooth' }); } catch { /* ignore */ }
      } else {
        try { window.scrollTo({ top: 0, behavior: 'smooth' }); } catch { /* ignore */ }
      }
    });
  }, [isSubmitting]);
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
        requiredRecord: undefined,
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
  useEffect(() => {
    const c = employeePPEItemsCriteria;
    // Only apply when we are creating a new form (not editing) and we have an employee criteria loaded
    if (isEditMode) return;
    if (!c || !c.employeeID) return;
    if (!itemRows || itemRows.length === 0) return;
    // ---------------------------
    // HSE approver group membership (for item edit permission)
    // Allow editing items if: canEditForm OR (3rd approval row is HSE Approval AND user is member of the assigned group)
    // ---------------------------
    const empKey = String(c.employeeID);
    if (criteriaAppliedForEmployeeId && criteriaAppliedForEmployeeId === empKey) return;

    const normalize = (v?: string) => (v || '').trim().toLowerCase();
    const contains = (text: string, needle: string) => normalize(text).includes(normalize(needle));

    const nextRows = itemRows.map(r => {
      // Don't override if the row was already interacted with
      if (typeof r.requiredRecord !== 'undefined' || r.selectedDetail) return r;

      const name = normalize(r.item || '');
      const details = (r.details || []);
      const itemSizes = (r.itemSizes || []);
      const detailsLower = details.map(d => normalize(d));
      const hasCoverallsDetail = detailsLower.some(d => /coveralls/.test(d));

      const findDetail = (label?: string): string | undefined => {
        if (!label) return undefined;
        const l = normalize(label);
        const exact = details.find(d => normalize(d) === l);
        if (exact) return exact;
        const partial = details.find(d => contains(d, l) || contains(l, d));
        return partial;
      };

      const findSizedDetail = (label?: string): string | undefined => {
        if (!label) return undefined;
        const l = normalize(label);
        const exact = itemSizes.find(d => normalize(d) === l);
        if (exact) return exact;
        const partial = itemSizes.find(d => contains(d, l) || contains(l, d));
        return partial;
      };

      const setReq = (selectedDetail?: string) => ({ ...r, requiredRecord: true, selectedDetail });
      const setReqSizedDetail = (selectedDetail?: string, size?: string) =>
        ({ ...r, requiredRecord: true, selectedDetail, itemSizeSelected: size });
      const setAdditionalItemDetail = (selectedDetail?: string, othersText?: string) =>
        ({ ...r, requiredRecord: true, selectedDetail, otherPurpose: othersText, itemSizeSelected: 'N/A' });
      const setReqSizedDetailByTypes = (
        selectedDetail?: string,
        sizes?: { coveralls?: string; top?: string; pants?: string }
      ) => {
        // Use existing type keys; find best matches
        const typeKeys = r.types || [];
        const topKey =
          typeKeys.find(t => /coverall\/?top|top/i.test(String(t))) ||
          typeKeys.find(t => /coverall/i.test(String(t))); // fallback if only "Coveralls" exists
        const pantsKey = typeKeys.find(t => /pants/i.test(String(t)));

        const nextMap: Record<string, string | undefined> = { ...(r.selectedSizesByType || {}) };

        // If coveralls size is given, set it to Top-equivalent key and clear Pants (business rule)
        if (sizes?.coveralls && !/^\s*n\/a\s*$/i.test(sizes.coveralls)) {
          if (topKey) nextMap[topKey] = sizes.coveralls;
          if (pantsKey) nextMap[pantsKey] = undefined;
        } else {
          // Otherwise set the specific Top and/or Pants sizes if provided and not N/A
          if (sizes?.top && !/^\s*n\/a\s*$/i.test(sizes.top) && topKey) nextMap[topKey] = sizes.top;
          if (sizes?.pants && !/^\s*n\/a\s*$/i.test(sizes.pants) && pantsKey) nextMap[pantsKey] = sizes.pants;
        }

        return {
          ...r,
          requiredRecord: true,
          selectedDetail,
          selectedSizesByType: nextMap,
        };
      };

      // Rain Suit
      if (name === 'rain suit') {
        const match = findSizedDetail(c.rainSuit);
        return match ? setReqSizedDetail(undefined, match) : r;
        // return setReq(match);
      }
      // Winter Jacket: special UI, just mark required
      if (name === 'winter jacket') {
        const match = findSizedDetail(c.winterJacket);
        return match ? setReqSizedDetail(undefined, match) : r;
        // return setReqSizedDetail(undefined, match);
      }
      // Uniform / body cover: choose Coveralls vs non-Coveralls
      if (name.includes('uniform') || hasCoverallsDetail || name.includes('coveralls') || name.includes('body')) {
        if (c.uniformCoveralls && !/^\s*n\/a\s*$/i.test(c.uniformCoveralls)) {
          const cv = details.find(d => /coveralls/i.test(d));
          return setReqSizedDetailByTypes(cv, { coveralls: c.uniformCoveralls });
        }
        const topSize = c.uniformTop && !/^\s*n\/a\s*$/i.test(c.uniformTop) ? c.uniformTop : undefined;
        const pantsSize = c.uniformPants && !/^\s*n\/a\s*$/i.test(c.uniformPants) ? c.uniformPants : undefined;

        if (topSize || pantsSize) {
          // Pick a reasonable non-coveralls detail so the row is "selected"
          const nonCoverallsDetail = details.find(d => /sweatshirt|pants/i.test(d)) || r.selectedDetail;
          return setReqSizedDetailByTypes(nonCoverallsDetail, { top: topSize, pants: pantsSize });
        }

      }
      // Reflective Vest
      if (name === 'reflective vest') {
        const match = findDetail(c.reflectiveVest);
        return match ? setReq(match) : r;
      }
      // Safety Helmet
      if (name === 'safety helmet') {
        const match = findDetail(c.safetyHelmet);
        return match ? setReq(match) : r;
      }
      // Safety Shoes
      if (name === 'safety shoes') {
        const size = c.safetyShoes;
        if (!size) return r;
        return setReqSizedDetail(undefined, size);
      }

      // Safety Shoes
      if (name === 'others') {
        const additionaItems = c.additionalPPEItems;
        if (!additionaItems) return r;
        return setAdditionalItemDetail(undefined, additionaItems || undefined);
      }

      return r;
    });

    // Only update if something actually changed
    const changed = nextRows.some((nr, i) => nr !== itemRows[i]);
    if (changed) setItemRows(nextRows);
    setCriteriaAppliedForEmployeeId(empKey);
  }, [employeePPEItemsCriteria, itemRows, isEditMode, criteriaAppliedForEmployeeId, _SPEmployeeId]);

  const toggleRequired = useCallback((rowIndex: number, checked?: boolean) => {
    setItemRows(prev => prev.map((r, i) => {
      if (i !== rowIndex) return r;
      if (checked) return { ...r, requiredRecord: true };
      // when unchecking, clear selections
      return {
        ...r,
        requiredRecord: false,
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
        if (!r.requiredRecord) return r;

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
        if (!r.requiredRecord) return r;

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
        if (!r.requiredRecord) return r;
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
        if (!r.requiredRecord) return r; // keep it locked unless the row is marked Required
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
          if (!r.requiredRecord) return r;
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

  // Sanitize quantity input to allow only 0-99 digits
  const sanitizeQty = useCallback((value?: string): string => {
    const digits = (value || '').replace(/\D/g, '');
    if (!digits) return '';
    let num = parseInt(digits, 10);
    if (!Number.isFinite(num) || num < 0) num = 0;
    if (num > 99) num = 99;
    return String(num);
  }, []);

  // ---------------------------
  // Handlers
  // ---------------------------

  // Block non-numeric key entries except navigation and clipboard combos
  const handleQtyKeyDown = useCallback((e: React.KeyboardEvent<HTMLInputElement>) => {
    const allowedKeys = ['Backspace', 'Delete', 'Tab', 'ArrowLeft', 'ArrowRight', 'Home', 'End'];
    if (allowedKeys.includes(e.key)) return;
    if ((e.ctrlKey || e.metaKey) && ['a', 'c', 'v', 'x'].includes(e.key.toLowerCase())) return;
    if (!/^\d$/.test(e.key)) {
      e.preventDefault();
    }
  }, []);

  const hideBanner = useCallback(() => {
    showBanner(``);
    setBannerText(undefined);
    setBannerOpts(undefined);
  }, []);

  const handleEmployeeChange = useCallback(async (items?: IPersonaProps[], selectedOption?: string) => {
    if (items && items.length > 0) {
      const selected = items[0];
      const emp = employees.find(e => Number(e.Id) === Number(selected?.id));
      setEmployee([selected]);
      setSPEmployeeId(Number(emp?.Id));
      setCoralEmployeeId(emp?.coralEmployeeID ? String(emp.coralEmployeeID) : undefined);
      // First try to find in employees list by FullName (fullName -> persona.text)

      const jobTitle: ICommon = emp?.jobTitle ? { id: emp.jobTitle.id ? String(emp.jobTitle.id) : undefined, title: emp.jobTitle.title || '' } : { id: undefined, title: '' };
      const department: ICommon = emp?.department ? { id: emp.department.id ? String(emp.department.id) : undefined, title: emp.department.title || '' } : { id: undefined, title: '' };
      const company: ICommon = emp?.company ? { id: emp.company.id ? String(emp.company.id) : undefined, title: emp.company.title || '' } : { id: undefined, title: '' };

      setJobTitleId(jobTitle);
      setDepartmentId(department);
      setCompanyId(company);
      // Auto-set requester ONLY if Employee list record has a manager; otherwise leave empty
      if (emp?.manager?.fullName) {
        setRequester([{ text: emp.manager.fullName, id: emp.manager.Id ? String(emp.manager.Id) : emp.manager.fullName }]);
      } else {
        setRequester([]);
      }

      // Reset the one-time-apply guard so criteria can be applied for this selection
      setCriteriaAppliedForEmployeeId(undefined);

      try {

        if (!isEditMode) {
          const eligible = await isEligibleToSubmit(Number(selected?.id), new Date());
          if (!eligible) {
            setIsAccidentalChecked(true);
            showBanner(`A Registered PPE Request for this employee was submitted within the last ${_coralFormsList?.SubmissionRangeInterval || 90} days.`
              // , { autoHideMs: 60000, fade: true, kind: 'error' }
              , { kind: 'error' });
            return;
          }
          else {
            setIsAccidentalChecked(false);
            hideBanner();
          }
        }

        // Fetch PPE items criteria for this employee ID
        await _getEmployeesPPEItemsCriteria(users, selected?.id ? Number(selected?.id) : undefined);

        if (employeePPEItemsCriteria && employeePPEItemsCriteria.employeeID !== selected?.id) {
          setItemRows(prev => prev.map(r => ({
            ...r,
            brandSelected: undefined,
            itemSizeSelected: undefined,
            selectedSizesByType: {},
            qty: undefined,
            requiredRecord: undefined,
            selectedDetails: [],            // added: clear details too
            othersItemdetailsText: {}
          })));
        }

      } catch (e) {
        console.warn('Failed to load PPE items criteria for employee', e);
      }

    } else {
      hideBanner();
      setIsEligibleToSubmitForm(true);
      setIsAccidentalChecked(false);
      setEmployee([]);
      setSPEmployeeId(undefined);
      setCoralEmployeeId(undefined);
      setJobTitleId({ id: '', title: '' });
      setDepartmentId({ id: '', title: '' });
      setCompanyId({ id: '', title: '' });
      setRequester([]);
      setEmployeePPEItemsCriteria({ Id: '' });
      setCriteriaAppliedForEmployeeId(undefined);
      setItemRows(prev =>
        prev.map(r => ({
          ...r,
          requiredRecord: undefined,
          brandSelected: undefined,
          selectedDetail: undefined,
          selectedDetails: [],           // for multi-select details
          itemSizeSelected: undefined,   // single-size path
          selectedType: undefined,       // if present in your state
          selectedSizesByType: {},       // typed sizes (Top/Pants etc.)
          qty: undefined,
          otherPurpose: undefined,
          othersItemdetailsText: {}
        }))
      );
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
        .map(e => ({ text: e.fullName || '', secondaryText: e.jobTitle?.title, id: (e.Id ? String(e.Id) : e.fullName), tertiaryText: (e.Id ? String(e.Id) : e.Id) }) as IPersonaProps);
      const deduped: IPersonaProps[] = [];
      const seen = new Set<string>();
      matches.forEach(m => { const key = (m.text || '').toLowerCase(); if (!seen.has(key)) { seen.add(key); deduped.push(m); } });
      return limitResults ? deduped.slice(0, limitResults) : deduped;
    });
  }, [_getEmployees, employees]);

  // Requester resolver (merge employees and Graph users for broader search)
  const requesterOnFilterChanged = useCallback((filterText: string, currentPersonas: IPersonaProps[], limitResults?: number): IPersonaProps[] | Promise<IPersonaProps[]> => {
    if (!filterText || filterText.trim().length === 0) return [];
    const userMatches = users.filter(u => (u.displayName || '').toLowerCase().includes(filterText.toLowerCase()))
      .map(u => ({ text: u.displayName || '', secondaryText: u.jobTitle, id: u.id }) as IPersonaProps);
    return userMatches
  }, [users]);

  const handleRequesterChange = useCallback(async (items?: IPersonaProps[], selectedOption?: string) => {
    if (items && items.length) setRequester([items[0]]); else setRequester([]);
  }, []);

  const handleNewRequestChange = useCallback((ev: React.FormEvent<HTMLElement>, checked?: boolean) => {
    if (!checked) return;
    setIsReplacementChecked(false);
    setIsAccidentalChecked(false);
    setReplacementReason('');

    setItemRows(prev =>
      prev.map(r => ({
        ...r,
        requiredRecord: undefined,
        brandSelected: undefined,
        selectedDetail: undefined,
        selectedDetails: [],
        itemSizeSelected: undefined,
        selectedType: undefined,
        selectedSizesByType: {},
        qty: undefined,
        otherPurpose: undefined,
        othersItemdetailsText: {}
      }))
    );

  }, []);

  const handleReplacementChange = useCallback((ev: React.FormEvent<HTMLElement>, checked?: boolean) => {

    const next = !!checked;
    setIsReplacementChecked(next);
    if (next) {
      setIsAccidentalChecked(false);
    } else {
      // If both are off, clear reason (returns to "New Request")
      setReplacementReason('');
    }

    setItemRows(prev =>
      prev.map(r => ({
        ...r,
        requiredRecord: undefined,
        brandSelected: undefined,
        selectedDetail: undefined,
        selectedDetails: [],           // for multi-select details
        itemSizeSelected: undefined,   // single-size path
        selectedType: undefined,       // if present in your state
        selectedSizesByType: {},       // typed sizes (Top/Pants etc.)
        qty: undefined,
        otherPurpose: undefined,
        othersItemdetailsText: {}
      }))
    );
  }, []);

  const handleApprovalApproverChange = useCallback((id: number | string, persona?: IPersonaProps) => {
    setFormsApprovalWorkflow(prev => {
      if (!prev || !prev.length) return prev;

      const idx = prev.findIndex(r => String(r.Id ?? '') === String(id));
      if (idx < 0) return prev;

      const next = [...prev];
      const row: any = { ...next[idx] };

      // Only allow assigning self
      const selEmail = String(persona?.secondaryText || '').toLowerCase();
      if (selEmail && selEmail !== props.context.pageContext?.user?.email) {
        return prev; // ignore picking someone else
      }

      row.DepartmentManagerApprover = persona || undefined;
      // no need to mark dirty if you don't persist approver; add row.__dirty=true if you plan to save it
      next[idx] = row;
      return next;
    });
  }, [editableRows]);

  //  Handles reason text change
  const handleApprovalReasonChange = useCallback((id: number | string, reason: string) => {
    setFormsApprovalWorkflow(prev => {
      if (!prev || prev.length === 0) return prev;

      const idx = prev.findIndex(r => String(r.Id ?? '') === String(id));
      if (idx < 0) return prev;

      const rowIdNum = Number(prev[idx].Id!);

      // Allow if the row is currently editable OR already dirty (mid-edit)
      if (!editableRows[rowIdNum] && !(prev[idx] as any)?.__dirty) return prev;

      // Only allow editing Reason when status is Rejected
      const isRejected = /reject/i.test(String(prev[idx]?.Status?.title || ''));
      if (!isRejected) return prev;

      const next = [...prev];
      const row: any = { ...next[idx] };

      row.Reason = (reason ?? '').toString();
      row.__index = idx;
      row.__dirty = true;

      next[idx] = row;
      return next;
    });
  }, [editableRows]);

  // Status-only change handler for approval rows
  const handleApprovalStatusChange = useCallback(
    (id: number | string, option?: { key?: string | number; text?: string }) => {
      if (!option) return; // no change selected

      setFormsApprovalWorkflow(prev => {
        if (!prev || prev.length === 0) return prev;

        const idx = prev.findIndex(r => String(r.Id ?? '') === String(id));
        if (idx < 0) return prev;

        const rowIdNum = Number(prev[idx].Id!);
        // Allow change only if row is currently editable or already dirty (mid-edit session)
        if (!editableRows[rowIdNum] && !(prev[idx] as any)?.__dirty) return prev;

        const next = [...prev];
        const row: any = { ...next[idx] };

        row.Status = { id: String(option.key ?? ''), title: String(option.text ?? '') };

        const nowRejected = /reject/i.test(String(option.text || ''));
        // Clear reason if no longer rejected (prevents stale reasons sticking around)
        if (!nowRejected) {
          row.Reason = undefined;
        }

        row.__index = idx;
        row.__dirty = true;

        next[idx] = row;
        return next;
      });
    },
    [editableRows]
  );

  const showBanner = useCallback((text: string, opts?: { autoHideMs?: number; fade?: boolean, kind?: BannerKind }) => {
    setBannerText(text);
    setBannerTick(t => t + 1);
    setBannerOpts(opts);
  }, []);

  // Navigate back to host list view (via callback or URL params)
  const goBackToHost = useCallback(() => {
    if (typeof props.onClose === 'function') {
      props.onClose();
      return;
    }
    const url = new URL(window.location.href);
    url.searchParams.delete('mode');
    url.searchParams.delete('formId');
    window.location.href = url.toString();
  }, [props.onClose]);

  const handleCancel = useCallback(() => {
    goBackToHost();
  }, [goBackToHost]);

  const handleSubmit = useCallback(async (withapprovalflag: boolean): Promise<boolean> => {
    try {
      const validationError = validateBeforeSubmit();
      if (validationError) {
        showBanner(validationError);
        return false;
      }

      if ((_isReplacementChecked || _isAccidentalChecked) && !(_replacementReason && _replacementReason.trim().length)) {
        showBanner('Please provide a reason for this request.');
        return false;
      }
      const editFormId = props.formId ? Number(props.formId) : undefined;

      // If the user cannot edit the header: allow approvals-only and/or HSE items-only update
      if (!canEditFormHeader && editFormId && editFormId > 0) {
        setIsSubmitting(true);
        try {
          let savedSomething = false;
          const payload = formPayload('Submitted');
          await _replacePPEItemDetailsRows(editFormId, payload);
          if (withapprovalflag) {
            await _saveApprovalWorkflowChanges(editFormId);
          }

          savedSomething = true;

          if (savedSomething) {
            try { window.alert('Changes saved.'); } catch { /* ignore */ }
            if (typeof props.onSubmitted === 'function') props.onSubmitted(editFormId);
            else goBackToHost();
            return true;
          } else {
            showBanner('Nothing to save.');
            return false;
          }
        } finally {
          setIsSubmitting(false);
        }
      }

      setIsSubmitting(true);
      const payload = formPayload('Submitted');

      if (editFormId && editFormId > 0) {
        // Update existing parent + replace child rows
        await _updatePPEForm(editFormId, payload);
        await _replacePPEItemDetailsRows(editFormId, payload);
        try { window.alert('PPE Form updated successfully.'); } catch { /* ignore */ }
        if (typeof props.onSubmitted === 'function') props.onSubmitted(editFormId);
        else goBackToHost();
        return true;
      } else {
        // Create new parent and children
        const newId = await _createPPEForm(payload);
        await _createPPEItemDetailsRows(newId, payload);
        try { window.alert('Your PPE Form is submitted successfully and it is now under processing.'); } catch { /* ignore */ }
        if (typeof props.onSubmitted === 'function') props.onSubmitted(newId);
        else goBackToHost();
        return true;
      }
    } catch (err: any) {
      showBanner('Submit info Error: ' + (err?.message || err) + '. Please try again.');
      return false;
    } finally {
      setIsSubmitting(false);
    }
  }, [formPayload, validateBeforeSubmit, showBanner, props.onSubmitted, goBackToHost, canEditFormHeader, formsApprovalWorkflow
  ]);
  // Persist approval changes for only rows the user is allowed to edit and that were modified
  const _saveApprovalWorkflowChanges = useCallback(async (formId: number): Promise<number> => {
    // Only persist rows that were actually changed
    const rows = (formsApprovalWorkflow || []).filter(r => (r as any)?.__dirty === true);
    if (!rows.length) return 0;

    // Guard: cannot reject without a reason
    const invalid = rows.filter(r => ((r.Status?.title || '').toLowerCase().includes('reject')) && !(r.Reason && String(r.Reason).trim().length));
    if (invalid.length) {
      const names = invalid.map(r => r.SignOffName || r.Id).join(', ');
      throw new Error(`Reason is required when rejecting for: ${names}`);
    }
    const ops = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PPE_Form_Approval_Workflow', '');

    const updates = rows.map(async (row) => {
      const body: any = {
        StatusRecordId: row.Status?.id ? Number(row.Status.id) : null,
        Reason: row.Reason ?? null,
      };

      if (row.Status?.title && row.Status.title.toLowerCase().includes('rejected')) {
        const RejectionReason = row.Reason || undefined;
        const WorkflowStatus = `${row.Status?.title} by ` + (loggedInUser?.displayName || 'Approver');
        _updatePPEFormStatus(formId, RejectionReason, WorkflowStatus);
      }
      else {
        const WorkflowStatus = `${row.Status?.title} by ` + (loggedInUser?.displayName || 'Approver');
        _updatePPEFormStatus(formId, '', WorkflowStatus);
      }

      return ops._updateItem(String(row.Id), body);
    });

    const res = await Promise.all(updates);
    return res.length;
  }, [formsApprovalWorkflow, props.context]);

  // Create parent PPEForm item and return its Id
  const _createPPEForm = useCallback(async (payload: ReturnType<typeof formPayload>): Promise<number> => {
    const requesterEmail = emailFromPersona(_requester?.[0]) || loggedInUser?.email;
    const submitterEmail = emailFromPersona(_submitter?.[0]) || loggedInUser?.email;
    const requesterId = await ensureUserId(requesterEmail);
    const submitterId = await ensureUserId(submitterEmail);

    const _employeeSPId = _employee ? Number(_employee[0]?.id) : undefined;
    if (_employeeSPId == null) throw new Error('Employee is required');

    const body = {
      EmployeeRecordId: _employeeSPId,
      SubmitterNameId: submitterId ?? null, // SharePoint person field
      RequesterNameId: requesterId ?? null, // SharePoint person field
      JobTitleRecordId: _jobTitle?.id ? Number(_jobTitle.id) : null,
      CompanyRecordId: _company?.id ? Number(_company.id) : null,
      DepartmentRecordId: _department?.id ? Number(_department.id) : null,
      ReasonForRequest: payload.requestType ?? null,
      ReasonRecord: payload.replacementReason ?? null,
      // ReplacementReason: payload.replacementReason ?? null,
      // EmployeeID: payload.employeeId ?? null,
      WorkflowStatus: 'In Process',
      RejectionReason: null,
    };
    spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PPE_Form', '');
    const newId = await spCrudRef.current._insertItem(body);
    if (!newId) throw new Error('Failed to create PPE Form');

    try {
      const coralReferenceNumber = await spHelpers.assignCoralReferenceNumber(props.context.spHttpClient,
        props.context.pageContext.web.absoluteUrl, 'PPE_Form', { Id: Number(newId) }, _company?.title , 'PPE');

      setCoralReferenceNumber(coralReferenceNumber);
    } catch (e) {
      // Optional: log/show a non-blocking message; the form is created even if reference assignment fails
      console.warn('Failed to set CoralReferenceNumber', e);
    }

    return newId as number;
  }, [emailFromPersona, ensureUserId, formPayload, _requester, _submitter, loggedInUser, props.context.spHttpClient]);

  // Update existing PPEForm item
  const _updatePPEForm = useCallback(async (formId: number, payload: ReturnType<typeof formPayload>): Promise<void> => {
    const requesterEmail = emailFromPersona(_requester?.[0]) || loggedInUser?.email;
    const submitterEmail = emailFromPersona(_submitter?.[0]) || loggedInUser?.email;
    const requesterId = await ensureUserId(requesterEmail);
    const submitterId = await ensureUserId(submitterEmail);

    const _employeeSPId = _employee ? Number(_employee[0]?.id) : undefined;
    if (_employeeSPId == null) throw new Error('Employee is required');

    const lastApproval = (payload.approvals || []).reduce<IFormsApprovalWorkflow | undefined>((acc, cur) => {
      if (!acc) return cur;
      const a = Number(acc?.Order || 0);
      const b = Number(cur?.Order || 0);
      return b >= a ? cur : acc;
    }, undefined);

    const rejectionReason = lastApproval?.Reason ? String(lastApproval.Reason).trim() : '';
    // Get the status text (fallback to 'Pending' when empty)
    const lastStatusTitle = (lastApproval?.Status?.title || '').trim() || 'Pending';
    const statusLower = lastStatusTitle.toLowerCase();
    let workflowStatusFinal: string;
    const modifiedUserName = props.context.pageContext?.user?.displayName;

    if (statusLower === 'pending') {
      // Pending: do not show "by <user>"
      workflowStatusFinal = lastStatusTitle;
    } else if (statusLower.includes('rejected') || statusLower.includes('closed')) {
      // Rejected/Closed: include actor
      workflowStatusFinal = `${lastStatusTitle} by ${modifiedUserName || 'Approver'}`;
    } else {
      // Other statuses (e.g., Approved): include actor as well
      workflowStatusFinal = `${lastStatusTitle} by ${modifiedUserName || 'Approver'}`;
    }

    const body = {
      EmployeeRecordId: _employeeSPId,
      SubmitterNameId: submitterId ?? null,
      RequesterNameId: requesterId ?? null,
      JobTitleRecordId: _jobTitle?.id ? Number(_jobTitle.id) : null,
      CompanyRecordId: _company?.id ? Number(_company.id) : null,
      DepartmentRecordId: _department?.id ? Number(_department.id) : null,
      ReasonForRequest: payload.requestType ?? null,
      ReasonRecord: payload.replacementReason,
      // EmployeeID: payload.employeeId ?? null,
      RejectionReason: rejectionReason,
      WorkflowStatus: workflowStatusFinal,
    };

    spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PPE_Form', '');
    await spCrudRef.current._updateItem(String(formId), body);
  }, [emailFromPersona, ensureUserId, _requester, _submitter, loggedInUser, _employee, _jobTitle, _company, _department, props.context.spHttpClient]);

  // Update existing PPEForm item workflow status only
  const _updatePPEFormStatus = useCallback(async (formId: number, RejectionReason?: string, WorkflowStatus?: string): Promise<void> => {
    const body = {
      RejectionReason: RejectionReason ?? null,
      WorkflowStatus: WorkflowStatus ?? null,
    };

    spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PPE_Form', '');
    await spCrudRef.current._updateItem(String(formId), body);
  }, [emailFromPersona, ensureUserId, _requester, _submitter, loggedInUser, _employee, _jobTitle, _company, _department, props.context.spHttpClient]);

  // // Create detail rows for each required item
  const _createPPEItemDetailsRows = useCallback(async (parentId: number, payload: ReturnType<typeof formPayload>) => {
    const requiredItems = (payload.items || []).filter(i => i.requiredRecord);
    if (requiredItems.length === 0) return;
    const posts = requiredItems.map(item => {
      const itemId = item?.itemId != null ? Number(item.itemId) : undefined;
      const detailId = item?.selectedDetailId != null ? Number(item.selectedDetailId) : undefined;

      // Map fields to your PPEItemDetails list’s internal names
      const body = {
        PPEFormIDId: parentId,
        ItemId: itemId ?? null,
        IsRequiredRecord: item.requiredRecord ?? null,
        Brands: item.brand ?? null,
        Quantity: item.qty ?? null,
        Size: item.size ?? null,
        PPEFormItemDetailId: detailId ?? null,
        OthersPurpose: item.othersText ?? null,
      };

      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PPE_Form_Items', '');
      const data = spCrudRef.current._insertItem(body);
      if (!data) throw new Error('Failed to create PPE Item Details');
      return data;
    });
    await Promise.all(posts);
  }, [props.context.spHttpClient]);

  // Replace child rows: delete existing detail rows then insert current required ones
  const _replacePPEItemDetailsRows = useCallback(async (parentId: number, payload: ReturnType<typeof formPayload>) => {
    // First, fetch existing children for this parent form
    const query = `?$select=Id&$filter=PPEFormID/Id eq ${parentId}`;
    const itemsOps = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PPE_Form_Items', query);
    const existing = await itemsOps._getItemsWithQuery();
    if (Array.isArray(existing) && existing.length) {
      // Delete all existing children
      const delOps = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PPE_Form_Items', '');
      await Promise.all(existing.map((row: any) => delOps._deleteItem(Number(row.Id))));
    }
    // Insert current selection
    await _createPPEItemDetailsRows(parentId, payload);
  }, [props.context.spHttpClient, _createPPEItemDetailsRows]);

  // ---------------------------
  // Render
  // ---------------------------
  if (loading) {
    return (
      <div className={styles.loadingContainer}>
        <Spinner label={"Preparing PPE form.. "} size={SpinnerSize.large} />
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
    <div className={styles.ppeFormBackground} ref={containerRef} style={{ position: 'relative' }} data-export-mode={exportMode ? 'true' : 'false'}>
      <div>
        <div ref={bannerTopRef} />
        {isSubmitting && !exportMode && (
          <div
            ref={overlayRef}
            aria-busy="true"
            role="dialog"
            aria-modal="true"
            className="no-pdf"
            data-html2canvas-ignore="true"
            aria-label={props.formId ? 'Updating form' : 'Submitting form'}
            style={{
              position: 'absolute',
              inset: 0,
              background: 'rgba(255,255,255,0.6)',
              zIndex: 1000,
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              pointerEvents: 'all'
            }}>
            <Spinner label={props.formId ? 'Updating form…' : 'Submitting form…'} size={SpinnerSize.large} />
          </div>
        )}

        {/* Screen-blocking overlay while preparing the PDF */}
        {isExportingPdf && (
          <div
            aria-busy="true"
            role="dialog"
            aria-modal="true"
            aria-label="Preparing PDF"
            className="no-pdf"
            data-html2canvas-ignore="true"
            style={{
              position: 'absolute',
              inset: 0,
              background: 'rgba(255,255,255,0.75)',
              zIndex: 1500,
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              pointerEvents: 'all'
            }}
          >
            <Spinner label="Preparing PDF…" />
          </div>
        )}

        <form >
          <div id="PdfEmployeeInfoSegment">
            <div className={styles.formHeader} >
              <img src={logoUrl} alt="Logo" className={styles.formLogo} />
              <span className={styles.formTitle}>PERSONAL PROTECTIVE EQUIPMENT (PPE) REQUISITION FORM</span>
            </div>
            <BannerComponent
              text={bannerText}
              kind={bannerOpts?.kind || 'error'}
              autoHideMs={bannerOpts?.autoHideMs}
              fade={bannerOpts?.fade}
              onDismiss={() => {
                setBannerText(undefined);
                setBannerOpts(undefined);
              }}
            />

            {/* {bannerText && <MessageBar styles={{ root: { marginBottom: 8, color: 'red' } }}>{bannerText}</MessageBar>} */}
            <Stack horizontal styles={stackStyles} id="EmployeeInfoStack">
              {isEditMode && _coralReferenceNumber && _coralReferenceNumber.trim().length > 0 && (
                <div className="row" style={{ marginBottom: 8 }}>
                  <Label styles={{ root: { color: '#000', fontWeight: 500 } }}>
                    Reference No. <span style={{ fontWeight: 500 }}>{_coralReferenceNumber}</span>
                  </Label>
                </div>
              )}

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
                    styles={peoplePickerBlackStyles}
                    disabled={uiDisabled(!canEditFormHeader || isEditMode)}
                    // disabled={!canEditFormHeader || isEditMode} // cannot change employee in edit mode
                    selectedItems={_employee}
                    onChange={(items) => {
                      const selectedText = items?.[0]?.text || '';
                      const empId = employees.find(e => (e.fullName || '').toLowerCase() === selectedText.toLowerCase())?.Id;
                      return handleEmployeeChange(items, empId ? String(empId) : undefined);
                    }}
                  />
                </div>
                <div className="form-group col-md-6">
                  <TextField label="Employee ID" styles={textFieldBlackStyles} value={_coralEmployeeId?.toString()} disabled={true} /></div>
              </div>

              <div className="row">
                <div className="form-group col-md-6">
                  <TextField label="Job Title" styles={textFieldBlackStyles} value={_jobTitle?.title} disabled={true} />
                </div>
                <div className="form-group col-md-6">
                  <TextField label="Department" styles={textFieldBlackStyles} value={_department?.title} disabled={true} />
                </div>
              </div>

              <div className="row">
                <div className="form-group col-md-6">
                  <TextField label="Company" styles={textFieldBlackStyles} value={_company?.title} disabled={true} /></div>
                <div className="form-group col-md-6">
                  <DatePicker disabled value={new Date(Date.now())} label="Date Requested"
                    strings={defaultDatePickerStrings}
                    style={{ maxWidth: "100%", color: 'black !important' }}
                    styles={datePickerBlackStyles}
                  />
                </div>
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
                    styles={peoplePickerBlackStyles}
                    disabled={uiDisabled(!canEditFormHeader)}
                    onChange={handleRequesterChange}
                    selectedItems={_requester}
                  />
                </div>

                <div className="form-group col-md-6">
                  <NormalPeoplePicker label={"Submitter Name"}
                    itemLimit={1} onResolveSuggestions={onFilterChanged}
                    className={'ms-PeoplePicker'}
                    key={'normal'}
                    removeButtonAriaLabel={'Remove'}
                    styles={peoplePickerBlackStyles}
                    inputProps={{ onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'), onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'), 'aria-label': 'People Picker' }} onInputChange={onInputChange} resolveDelay={300} disabled={true} selectedItems={_submitter} />
                </div>
              </div>

              <div className={`row  ${styles.mt10}`}>
                <div className="form-group col-md-12" >
                  <div className={styles.requestReasonRow}>
                    <Label className={styles.requestReasonLabel}>Request Reason</Label>
                    <Checkbox label="New Request" checked={!_isReplacementChecked && !_isAccidentalChecked} onChange={handleNewRequestChange}
                      disabled={uiDisabled(!canEditFormHeader || !IsEligibleToSubmitForm || isEditMode)} />

                    <Checkbox label="Replacement" checked={_isReplacementChecked} onChange={handleReplacementChange}
                      disabled={uiDisabled(!canEditFormHeader || !IsEligibleToSubmitForm || isEditMode)} />

                    <Checkbox label="Accidental" checked={_isAccidentalChecked}
                      disabled={uiDisabled(IsEligibleToSubmitForm)} />


                    <div className={styles.requestReasonField}>
                      <TextField placeholder="Reason" multiline autoAdjustHeight resizable
                        styles={textFieldBlackStyles}
                        disabled={uiDisabled(!(_isReplacementChecked || _isAccidentalChecked) || !canEditFormHeader)}
                        value={_replacementReason}
                        onChange={(_e, v) => setReplacementReason(v || '')} />
                    </div>
                  </div>
                </div>
              </div>
            </Stack>
            <Separator />
          </div>  {/* end PdfEmployeeInfoSegment */}

          <div className="text-center">
            <small className="text-muted" style={{ fontStyle: 'italic', fontSize: '1.0rem' }}>Please complete the table below in the blank spaces; grey spaces are for administrative use only.</small>
          </div>

          <div id="PdfItemsSegment" style={{ pageBreakAfter: exportMode ? 'always' : 'auto' }}>
            <div >
              {exportMode ? (
                // During export: show summary where the items grid usually is
                <Stack horizontal styles={stackStyles} id="ItemsSummaryStack">
                  <div className="form-group col-md-12" style={{ width: '100%' }}>
                    <Label>PPE Required Items</Label>
                    <DetailsList
                      items={itemsSummary}
                      columns={[
                        {
                          key: 'colItem', name: 'Item', fieldName: 'item', minWidth: 120, maxWidth: 130, isResizable: true,
                          styles: { root: { color: 'black !important', fontWeight: "500 !important" } },
                          onRender: (r: ItemSummary) => <span style={{
                            display: 'block', whiteSpace: 'normal',
                            wordWrap: 'break-word', overflowWrap: 'anywhere', lineHeight: 1.3,
                            color: "black !important", fontWeight: "500 !important"
                          }}>{r.item}</span>
                        },
                        {
                          key: 'colDetail', name: 'Detail/Purpose', fieldName: 'detail', minWidth: 230, isResizable: true,
                          onRender: (r: ItemSummary) => <span style={{

                            display: 'block', whiteSpace: 'normal',
                            wordWrap: 'break-word', overflowWrap: 'anywhere', lineHeight: 1.3,
                            color: "black !important", fontWeight: "500 !important"
                          }}>{r.detail}</span>
                        },
                        {
                          key: 'colQty', name: 'Qty', fieldName: 'quantity', minWidth: 50, isResizable: true,
                          onRender: (r: ItemSummary) => <span style={{
                            display: 'block', whiteSpace: 'normal', color: "black !important", fontWeight: "500 !important",
                            wordWrap: 'break-word', overflowWrap: 'anywhere', lineHeight: 1.3
                          }}>{r.quantity}</span>
                        },
                        {
                          key: 'colBrand', name: 'Brand', fieldName: 'brand', minWidth: 150, isResizable: true,
                          onRender: (r: ItemSummary) => <span style={{
                            display: 'block', whiteSpace: 'normal',
                            wordWrap: 'break-word', overflowWrap: 'anywhere', lineHeight: 1.3, color: "black !important", fontWeight: "500 !important",
                          }}>{r.brand}</span>
                        },
                        {
                          key: 'colSize', name: 'Size(s)', fieldName: 'size', minWidth: 150, isResizable: true,
                          onRender: (r: ItemSummary) => <span style={{
                            display: 'block', whiteSpace: 'normal',
                            wordWrap: 'break-word', overflowWrap: 'anywhere', lineHeight: 1.3, color: "black !important", fontWeight: "500 !important",
                          }}>{r.size}</span>
                        },
                      ]}
                      selectionMode={SelectionMode.none}
                      layoutMode={DetailsListLayoutMode.justified}
                      className={styles.detailsListHeaderCenter}
                      styles={{
                        root: { width: '100%' },
                        // target cells and rows
                        contentWrapper: {
                          selectors: {
                            '.ms-DetailsRow-fields': {
                              alignItems: 'center'  // stretch to max height of tallest cell in the row
                            },
                            '.ms-DetailsRow-cell': {
                              color: 'black !important',
                              fontWeight: '600 !important',
                              padding: '8px 0px 8px 8px !important', // top-bottom left-right
                            },
                            '&': { overflowX: 'visible', overflowY: 'visible' }
                          }
                        }
                      }}
                    />
                  </div>
                </Stack>
              ) : (
                <Stack horizontal styles={stackStyles} id="ItemsStack">
                  <div className="row">
                    <div className="form-group col-md-12">
                      <DetailsList
                        items={itemRows.sort((a, b) => (a.order ? a.order : 0) - (b.order ? b.order : 0))}
                        setKey="ppeAggregatedItemsList"
                        selectionMode={SelectionMode.none}
                        layoutMode={DetailsListLayoutMode.justified}          // <-- responsive fill
                        constrainMode={ConstrainMode.horizontalConstrained}
                        styles={{
                          root: { width: '100%' },
                          // target cells and rows
                          contentWrapper: {
                            selectors: {
                              '.ms-DetailsRow-fields': {
                                alignItems: 'center'  // stretch to max height of tallest cell in the row
                              },
                              '.ms-DetailsRow-cell': {
                                color: 'black !important',
                                fontWeight: '500 !important',
                                padding: '8px 0px 8px 8px !important', // top-bottom left-right
                              },
                              '&': { overflowX: 'visible', overflowY: 'visible' }
                            }
                          }
                        }}

                        columns={[
                          {
                            key: 'colItem', name: 'Item', fieldName: 'item', minWidth: 90, isResizable: true,
                            onRender: (r: ItemRowState) => <span style={{
                              display: 'block', whiteSpace: 'normal',
                              wordWrap: 'break-word', overflowWrap: 'anywhere', lineHeight: 1.3
                            }}>{r.item}</span>
                          },
                          {
                            key: 'colRequired', name: 'Required', fieldName: 'requiredRecord', minWidth: 70, maxWidth: 70,
                            onRender: (r: ItemRowState) => (
                              <Checkbox
                                checked={!!r.requiredRecord}
                                ariaLabel="Required"
                                id={r.item}
                                onChange={(_e, ch) => toggleRequired(itemRows.indexOf(r), ch)}
                                // disabled={!canEditItems}
                                disabled={uiDisabled(!canEditItems)}
                                styles={{ root: { display: 'flex', justifyContent: 'center', alignItems: 'center', width: '100%' } }}
                              />
                            )
                          },
                          {
                            key: 'colDetails', name: 'Specific Detail', fieldName: 'itemDetails', minWidth: 320, isResizable: true, onRender: (r: ItemRowState) => (
                              <div>
                                {r.details.map(detail => {
                                  // ...inside the onRender of colDetails...
                                  {
                                    const itemLabel = r.item.toLowerCase() === 'others';
                                    if (itemLabel) {
                                      return (
                                        <div key={detail} style={{ display: 'flex', flexDirection: 'column', marginBottom: 8 }}>
                                          <TextField placeholder={detail} multiline autoAdjustHeight resizable
                                            scrollContainerRef={containerRef} styles={{
                                              root: {
                                                width: '100%',
                                                field: {
                                                  color: '#000', // <-- main text
                                                  selectors: {
                                                    '&::placeholder': { color: '#666' },        // optional: darker placeholder
                                                    '&:disabled': { color: '#000' }             // ensure disabled still renders black
                                                  },
                                                  subComponentStyles: {
                                                    label: { root: { color: '#000' } }
                                                  }
                                                }
                                              }
                                            }}
                                            value={r.otherPurpose ?? undefined}

                                            disabled={!r.requiredRecord || !canEditItems}
                                            key={`purpose-${r.itemId}-${r.requiredRecord ? 'on' : 'off'}`}
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
                                      <div key={detail} style={{ display: 'flex', alignItems: 'center', marginBottom: 2 }}>
                                        <Checkbox
                                          label={detail}
                                          checked={checked}
                                          onChange={(_e, ch) => toggleItemDetail(itemRows.indexOf(r), detail, !!ch)}
                                          // disabled={!canEditItems || !r.requiredRecord}
                                          disabled={uiDisabled(!canEditItems || !r.requiredRecord)}
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
                            key: 'colBrand', name: 'Brand', fieldName: 'brand', minWidth: 180, isResizable: false,
                            onRender: (r: ItemRowState) => {
                              return (
                                <>
                                  {r.brands.length === 0 && <span>N/A</span>}
                                  {
                                    r.brands.map(brand => {
                                      const brandChecked = r.brandSelected === brand;
                                      return (
                                        <div key={brand} style={{ display: 'flex', alignItems: 'center', marginBottom: 2 }}>
                                          <Checkbox label={brand} checked={brandChecked}
                                            onChange={(_e, ch) => toggleBrand(itemRows.indexOf(r), brand, !!ch)}
                                            // disabled={!canEditItems || !r.requiredRecord}
                                            disabled={uiDisabled(!canEditItems || !r.requiredRecord)}

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
                            key: 'colQty', name: 'Qty', fieldName: 'qty', minWidth: 50, maxWidth: 60, onRender: (r: ItemRowState) => (
                              <TextField
                                value={r.qty || ''}
                                type='text'

                                onChange={(_e, v) => {
                                  const next = sanitizeQty(v);
                                  updateItemQty(itemRows.indexOf(r), next);
                                }}
                                onKeyDown={handleQtyKeyDown}
                                // disabled={!canEditItems || !r.requiredRecord}
                                disabled={uiDisabled(!canEditItems || !r.requiredRecord)}
                                styles={{
                                  field: {
                                    color: '#000', // <-- main text
                                    selectors: {
                                      '&::placeholder': { color: '#666' },        // optional: darker placeholder
                                      '&:disabled': { color: '#000' }             // ensure disabled still renders black
                                    },
                                    subComponentStyles: {
                                      label: { root: { color: '#000' } }
                                    }
                                  },
                                  root: { display: 'flex', justifyContent: 'center', alignItems: 'center', width: '100%' },
                                }}
                              />
                            )
                          },

                          // ...existing code...
                          {
                            key: 'colSizes', name: 'Size', fieldName: 'size', minWidth: 280, isResizable: true,
                            onRender: (r: ItemRowState) => {
                              if (r.item.toLowerCase() === 'others') {
                                // Show Sizes only if Required is checked
                                if (!r.requiredRecord) return <span />;

                                const sizes = Array.from(new Set((r.itemSizes || []).map(s => String(s).trim()).filter(Boolean)))
                                  .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));

                                return (
                                  <div key={r.item} style={{ display: 'flex', alignItems: 'center', marginBottom: 2 }}>
                                    <ComboBox
                                      placeholder={sizes.length ? 'Size' : 'No sizes'}
                                      selectedKey={r.itemSizeSelected || undefined}
                                      options={sizes.map(s => ({ key: s, text: s }))}
                                      styles={comboBoxBlackStyles}
                                      dropdownMaxWidth={140}
                                      // disabled={!sizes.length || !canEditItems || !r.requiredRecord}
                                      disabled={uiDisabled(!sizes.length || !canEditItems || !r.requiredRecord)}
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
                                          style={{ display: 'flex', flexDirection: 'column', gap: 2, paddingLeft: idx === 0 ? 0 : 12, marginLeft: idx === 0 ? 0 : 12, borderLeft: idx === 0 ? 'none' : '1px solid #ddd' }}>
                                          <Label styles={{ root: { marginBottom: 4, fontWeight: 600 } }}>{type}</Label>

                                          {sizes.length === 0 ? (<span>N/A</span>) :
                                            (
                                              <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
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
                                                        // disabled={!canEditItems || !r.requiredRecord}
                                                        disabled={uiDisabled(!canEditItems || !r.requiredRecord)}
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
                                <div style={{ display: 'grid', gridTemplateColumns: `repeat(${cols}, minmax(0, 1fr))`, gap: 2 }}>
                                  {sizes.map(size => {
                                    const sizeChecked = r.itemSizeSelected === size;
                                    return (
                                      <div key={size} style={{ display: 'flex', alignItems: 'center' }}>
                                        <Checkbox
                                          label={size}
                                          checked={sizeChecked}
                                          onChange={(_e, ch) => toggleSize(itemRows.indexOf(r), size, !!ch)}
                                          // disabled={!canEditItems || !r.requiredRecord}
                                          disabled={uiDisabled(!canEditItems || !r.requiredRecord)}
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
              )}
            </div>
          </div> {/* end PdfItemsSegment */}

          {/* PdfBottomSegment: everything below items */}
          <div id="PdfInstructionsSegment" style={{ pageBreakAfter: exportMode ? 'always' : 'auto' }}>
            {/* <Separator /> */}
            {/* Instructions For Use */}
            <Stack horizontal styles={stackStyles} id="InstructionsStack">
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
          </div> {/* end PdfInstructionsSegment */}

          {/* Approvals (edit mode) */}
          {isEditMode && <Separator />}

          <div id="PdfApprovalsSegment" style={{ pageBreakAfter: exportMode ? 'always' : 'auto' }}>
            {isEditMode && (
              <Stack horizontal styles={stackStyles} className="mt-3 mb-3" id="approvalsSection" style={{ width: '100%' }}>
                <div style={{ width: '100%' }}>
                  <Label>Approvals / Sign-off</Label>
                  <DetailsList
                    items={formsApprovalWorkflow}
                    columns={[
                      {
                        key: 'colSignOff', name: 'Sign off', fieldName: 'SignOffName', minWidth: !!exportMode ? 130 : 160, isResizable: true,
                        onRender: (item: any) => (<div> <span style={{ color: "black !important", fontWeight: "400 !important" }}>{item.SignOffName}</span></div>)
                      },
                      {
                        key: 'colDepartmentManager', name: 'Name', fieldName: 'DepartmentManager', minWidth: !!exportMode ? 180 : 220, isResizable: true,
                        onRender: (item: any) => {
                          const grpName = resolveGroupUserForItemRow(item as IFormsApprovalWorkflow);
                          const key = (grpName || '').toLowerCase();
                          const members = key ? (groupMembers[key] || []) : [];
                          // const members: IPersonaProps[] = key ? ((item.ApproversNamesList?.[key] as IPersonaProps[]) ?? []) : [];
                          const isMember = members.find(m => (String(m.secondaryText).toLowerCase()) === (props.context.pageContext?.user?.email || '').toLowerCase());
                          const isPending = item?.Status?.title && (String(item.Status.title).toLowerCase() === 'pending');
                          let selectedKey = '';
                          let placeHolderDisplay = '';
                          if (isMember && isPending) {
                            selectedKey = (props.context.pageContext?.user?.displayName || '');
                            placeHolderDisplay = selectedKey;
                          }
                          else if (isMember && !isPending) {
                            selectedKey = isMember?.text || '';
                            // selectedKey = item?.DepartmentManagerApprover?.text || '';
                            placeHolderDisplay = selectedKey;
                          }

                          if (!isMember && item?.Status?.title && (String(item.Status.title).toLowerCase() !== 'pending')) {
                            return (
                              <TextField value={item?.DepartmentManagerApprover?.text || ''} disabled={true} styles={textFieldBlackStyles} />
                            )
                          }
                          else {
                            return (
                              <ComboBox
                                placeholder={placeHolderDisplay || (members.length ? '' : 'No members Assigned to this role.')}
                                selectedKey={selectedKey}
                                options={members.map(m => ({ key: String(m.secondaryText), text: m.text || (m.secondaryText || ''), data: m }))}
                                useComboBoxAsMenuWidth
                                styles={comboBoxBlackStyles}
                                disabled={!members.length || !canEditApprovalRow(item)}
                                onChange={(_, opt) => {
                                  const persona = (opt?.data as IPersonaProps) || (opt ? { id: String(opt.key), text: String(opt.text || ''), secondaryText: String((opt as any).secondaryText || '') } as IPersonaProps : undefined);
                                  if (persona) {
                                    // Only allow selecting yourself; ignore picking others
                                    const selEmail = (persona.secondaryText || '').toLowerCase();
                                    if (selEmail !== props.context.pageContext?.user?.email) return;
                                    handleApprovalApproverChange(item.Id!, persona);
                                    const rid = String(item.Id ?? '');
                                    if (rid) setLockedApprovalRowIds(prev => ({ ...prev, [rid]: true }));
                                  }
                                }}
                              />
                            );
                          }
                        }
                      },
                      {
                        key: 'colStatus', name: 'Status', fieldName: 'Status', minWidth: !!exportMode ? 130 : 160, isResizable: true,
                        onRender: (item: any, idx?: number) => {
                          const sorted = (lKPWorkflowStatus || []).slice()
                            .sort((a, b) => {
                              const ao = a?.Order ?? Number.POSITIVE_INFINITY;
                              const bo = b?.Order ?? Number.POSITIVE_INFINITY;
                              return Number(ao) - Number(bo);
                            });
                          const options = sorted.filter(s => String(s.Title ?? '').trim().toLowerCase() !== 'closed')
                            .map(s => ({
                              key: String(s.Id),
                              text: String(s.Title ?? '').trim(),
                            }));
                          // item.Status is ICommon { id, title }
                          const selectedKey = item.Status?.id ? String(item.Status.id) : undefined;

                          if (item?.Status?.title && (String(item.Status.title).toLowerCase() === 'closed')) {
                            return (
                              <TextField value={item?.Status?.title || ''} disabled={true}
                                styles={textFieldBlackStyles} />
                            )
                          } else {
                            return (
                              <ComboBox
                                placeholder={options.length ? 'Select status' : 'No status'}
                                selectedKey={selectedKey}
                                styles={comboBoxBlackStyles}
                                options={options}
                                useComboBoxAsMenuWidth={true}
                                disabled={!canEditApprovalRow(item)}
                                onChange={(_, option) => handleApprovalStatusChange(item.Id!, option as any)} />
                            );
                          }
                        }
                      },
                      {
                        key: 'colReason', name: 'Reason', fieldName: 'Reason', minWidth: !!exportMode ? 160 : 280, isResizable: true,
                        onRender: (item: any, idx?: number) => {
                          const canEdit = canEditApprovalRow(item);
                          const isRejected = /reject/i.test(String(item?.Status?.title || ''));
                          const canEditReason = canEdit && isRejected;
                          return (
                            <TextField value={item.Reason || ''}
                              placeholder={canEditReason ? 'Enter rejection reason' : ''}
                              disabled={!canEditReason}
                              styles={textFieldBlackStyles} rows={1}
                              multiline autoAdjustHeight
                              onChange={(ev, newValue) => handleApprovalReasonChange(item.Id!, newValue || '')}
                            />);
                        }
                      },
                      {
                        key: 'colDate', name: 'Date', fieldName: 'Date', minWidth: !!exportMode ? 180 : 200, isResizable: true,
                        onRender: (item: any) => {
                          const isApproved = String(item?.Status?.title || '').toLowerCase() === 'approved';
                          const dateValue = isApproved ? (item.Date ? new Date(item.Date) : undefined) : new Date();
                          return (
                            <DatePicker
                              value={dateValue}
                              disabled={prefilledFormId ? true : false}
                              strings={defaultDatePickerStrings}
                              styles={datePickerBlackStyles}
                            />
                          );
                        }
                      }
                    ]}
                    selectionMode={SelectionMode.none}
                    setKey="approvalsList"
                    layoutMode={DetailsListLayoutMode.justified}          // <-- responsive fill
                    constrainMode={ConstrainMode.horizontalConstrained}
                    onShouldVirtualize={() => false}
                    styles={{
                      root: { width: '100%' },
                      // target cells and rows
                      contentWrapper: {
                        selectors: {
                          '.ms-DetailsRow-fields': {
                            alignItems: 'center'  // stretch to max height of tallest cell in the row
                          },
                          '.ms-DetailsRow-cell': {
                            color: 'black !important',
                            fontWeight: '500 !important',
                            padding: '8px 0px 8px 8px !important', // top-bottom left-right
                          },
                          '&': { overflowX: 'visible', overflowY: 'visible' }
                        }
                      }
                    }}
                  />
                </div>
              </Stack>
            )}
            <DocumentMetaBanner docCode={'COR-HSE-01-FOR-001'} version="V03" effectiveDate="16-OCT-2025" page={1} />
          </div> {/* end PdfApprovalsSegment */}

          <div id="PdfbuttonsSection">
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 8 }}
              className="no-pdf" data-html2canvas-ignore="true">
              <DefaultButton text="Close" onClick={handleCancel} disabled={isSubmitting} />
              <ExportPdfControls
                targetRef={containerRef}
                coralReferenceNumber={_coralReferenceNumber}
                employeeName={_employee?.[0]?.text}
                exportMode={exportMode}
                onExportModeChange={setExportMode}
                onBusyChange={setIsExportingPdf}
                isClosedBySystem={(formsApprovalWorkflow || []).some(r => String(r?.Status?.title || '').toLowerCase().includes('approved') && r.FinalLevel === r.Order)}
                onError={(m) => showBanner(m)}
              />

              {isEditMode && !canEditFormHeader ? (
                // Approval-phase: show approvals-only button
                <PrimaryButton
                  text={isSubmitting ? 'Saving approvals…' : 'Save approvals'}
                  onClick={() => handleSubmit(true)}
                  disabled={isSubmitting || !canChangeApprovalRows || !hasApprovalChanges}
                />
              ) : (
                // Normal create/update
                <PrimaryButton
                  text={isSubmitting ? (props.formId ? 'Updating…' : 'Submitting…') : (props.formId ? 'Update' : 'Submit')}
                  onClick={() => handleSubmit(false)}
                  disabled={isSubmitting ||
                    (!canEditFormHeader && !canEditItems && !canChangeApprovalRows)}
                />
              )}
            </div>
          </div>

        </form>
      </div>
    </div>
  );
}