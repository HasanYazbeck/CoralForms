import * as React from 'react';
import type { IPTWFormProps } from './IPTWFormProps';
import PermitSchedule from './PermitSchedule';
import { IPermitScheduleRow } from '../../../Interfaces/PtwForm/IPermitSchedule';

// Styles
import "bootstrap/dist/css/bootstrap.min.css";
import styles from './PtwForm.module.scss';
import {
  IPersonaProps, Spinner, SpinnerSize,
  TextField,
  Label,
  IDropdownOption,
  ComboBox,
  Checkbox,
  IComboBoxStyles,
  IComboBox,
  Stack,
  MessageBar,
  IconButton,
  DefaultButton,
  PrimaryButton,
  Separator,
  DatePicker
} from '@fluentui/react';
import { NormalPeoplePicker, IBasePickerSuggestionsProps, IBasePickerStyles } from '@fluentui/react/lib/Pickers';
import { ILKPItemInstructionsForUse } from '../../../Interfaces/Common/ICommon';
import { MSGraphClientV3 } from '@microsoft/sp-http-msgraph';
import { SPHttpClient } from '@microsoft/sp-http';
import { IUser } from '../../../Interfaces/Common/IUser';
import { SPCrudOperations } from "../../../Classes/SPCrudOperations";
import { SPHelpers } from "../../../Classes/SPHelpers";
import { IAssetCategoryDetails, IAssetsDetails, ICoralForm, IEmployeePeronellePassport, ILookupItem, IPTWForm, IPTWWorkflow, ISagefaurdsItem, IWorkCategory } from '../../../Interfaces/PtwForm/IPTWForm';
import { CheckBoxDistributerComponent } from './CheckBoxDistributerComponent';
import RiskAssessmentList, { IRiskTaskRow } from './RiskAssessmentList';
import { CheckBoxDistributerOnlyComponent } from './CheckBoxDistributerOnlyComponent';
import { DocumentMetaBanner } from '../../../Components/DocumentMetaBanner';
import { ICoralFormsList } from '../../../Interfaces/Common/ICoralFormsList';
import ExportPdfControls from '../../ppeForm/components/ExportPdfControls';
import BannerComponent, { BannerKind } from '../../ppeForm/components/BannerComponent';

interface IRiskAssessmentResult {
  rows: IRiskTaskRow[];
  overallRisk?: string;
  l2Required: boolean;
  l2Ref?: string;
}

export default function PTWForm(props: IPTWFormProps) {
  // Helpers and refs
  const formName = "Permit To Work";
  const containerRef = React.useRef<HTMLDivElement>(null);
  const overlayRef = React.useRef<HTMLDivElement>(null);
  const spCrudRef = React.useRef<SPCrudOperations | undefined>(undefined);
  const spHelpers = React.useMemo(() => new SPHelpers(), []);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [mode, setMode] = React.useState<'saved' | 'submitted' | 'approved' | 'new'>('new');

  const [isExportingPdf, setIsExportingPdf] = React.useState(false); // NEW
  const [exportMode, setExportMode] = React.useState(false);
  const [bannerText, setBannerText] = React.useState<string>();
  const [bannerTick, setBannerTick] = React.useState(0);
  const [bannerOpts, setBannerOpts] = React.useState<{ autoHideMs?: number; fade?: boolean; kind?: BannerKind } | undefined>();
  const bannerTopRef = React.useRef<HTMLDivElement>(null);

  const [_users, setUsers] = React.useState<IUser[]>([]);
  const [, setCoralFormsList] = React.useState<ICoralFormsList>({ Id: "" });
  const [ptwFormStructure, setPTWFormStructure] = React.useState<IPTWForm>({ issuanceInstrunctions: [], personnalInvolved: [] });
  const [itemInstructionsForUse, setItemInstructionsForUse] = React.useState<ILKPItemInstructionsForUse[]>([]);
  const [personnelInvolved, setPersonnelInvolved] = React.useState<IEmployeePeronellePassport[]>([]);
  const [, setAssetDetails] = React.useState<IAssetCategoryDetails[]>([]);
  const [safeguards, setSafeguards] = React.useState<ISagefaurdsItem[]>([]);
  const [filteredSafeguards, setFilteredSafeguards] = React.useState<ISagefaurdsItem[]>([]);
  const webUrl = props.context.pageContext.web.absoluteUrl;
  const [assetCategoriesList, setAssetCategoriesList] = React.useState<ILookupItem[] | undefined>([]);
  const [assetCategoriesDetailsList, setAssetCategoriesDetailsList] = React.useState<IAssetsDetails[] | undefined>([]);
  const [, setPtwApprovalWorkflow] = React.useState<IPTWWorkflow | undefined>(undefined);

  // Form State to used on update or submit
  const [_coralReferenceNumber, setCoralReferenceNumber] = React.useState<string>('');
  const [_PermitOriginator, setPermitOriginator] = React.useState<IPersonaProps[]>([]);
  const [_assetId, setAssetId] = React.useState<string>('');
  const [_selectedCompany, setSelectedCompany] = React.useState<ILookupItem | undefined>(undefined);
  const [_selectedAssetCategory, setSelectedAssetCategory] = React.useState<number | undefined>(undefined);
  const [_selectedAssetDetails, setSelectedAssetDetails] = React.useState<number | undefined>(0);
  const [_projectTitle, setProjectTitle] = React.useState<string>('');
  const [_selectedPermitTypeList, setSelectedPermitTypeList] = React.useState<IWorkCategory[]>([]);
  const [_permitPayload, setPermitPayload] = React.useState<IPermitScheduleRow[]>([]);
  const [_selectedHacWorkAreaId, setSelectedHacWorkAreaId] = React.useState<number | undefined>(undefined);
  const [_selectedWorkHazardIds, setSelectedWorkHazardIds] = React.useState<Set<number>>(new Set());
  const [_workHazardsOtherText, setWorkHazardsOtherText] = React.useState<string>('');

  const [_riskAssessmentsTasks, setRiskAssessmentsTasks] = React.useState<IRiskTaskRow[] | undefined>(undefined);
  const [_overAllRiskAssessment, setOverAllRiskAssessment] = React.useState<string>('');
  const [_detailedRiskAssessment, setDetailedRiskAssessment] = React.useState<Boolean | undefined>(undefined);
  const [_riskAssessmentReferenceNumber, setRiskAssessmentReferenceNumber] = React.useState<string>('');

  const [_selectedPrecautionIds, setSelectedPrecautionIds] = React.useState<Set<number>>(new Set());
  const [_precautionsOtherText, setPrecautionsOtherText] = React.useState<string>('');

  const [_gasTestValue, setGasTestValue] = React.useState('');
  const [_gasTestResult, setGasTestResult] = React.useState('');
  const [_fireWatchValue, setFireWatchValue] = React.useState('');
  const [_fireWatchAssigned, setFireWatchAssigned] = React.useState('');
  const [_attachmentsValue, setAttachmentsValue] = React.useState('');
  const [_attachmentsResult, setAttachmentsResult] = React.useState('');
  const [_selectedProtectiveEquipmentIds, setSelectedProtectiveEquipmentIds] = React.useState<Set<number>>(new Set());
  const [_protectiveEquipmentsOtherText, setProtectiveEquipmentsOtherText] = React.useState<string>('');

  const [_selectedMachineryIds, setSelectedMachineryIds] = React.useState<number[] | undefined>(undefined);
  const [_selectedPersonnelIds, setSelectedPersonnelIds] = React.useState<number[] | undefined>(undefined);

  const [_selectedToolboxTalk, setToolboxTalk] = React.useState<Boolean | undefined>(undefined);
  const [_toolboxHSEReference, setToolboxHSEReference] = React.useState<String>('');
  const [_selectedToolboxTalkDate, setToolboxTalkDate] = React.useState<String | undefined>(undefined);
  const [_selectedToolboxConductedBy, setToolboxConductedBy] = React.useState<IPersonaProps[] | undefined>(undefined);

  // Busy overlay and notifications
  const [isBusy, setIsBusy] = React.useState<boolean>(false);
  const [busyLabel, setBusyLabel] = React.useState<string>('Processing…');

  // Current user role
  const [isPermitOriginator, setIsPermitOriginator] = React.useState<boolean>(false);
  const [isPerformingAuthority, setIsPerformingAuthority] = React.useState<boolean>(false);
  const [isPermitIssuer, setIsPermitIssuer] = React.useState<boolean>(false);
  const [isAssetDirector, setIsAssetDirector] = React.useState<boolean>(false);
  const [isAssetManager, setIsAssetManager] = React.useState<boolean>(false);
  const [isHSEDirector, setIsHSEDirector] = React.useState<boolean>(false);

  // Add status type and options
  type SignOffStatus = 'Pending' | 'Approved' | 'Rejected' | 'Closed';

  // SharePoint group members cache
  type SPGroupUser = { id: number; title: string; email: string };
  const [groupMembers, setGroupMembers] = React.useState<Record<string, SPGroupUser[]>>({});

  // Sign-off state
  const [_poDate, setPoDate] = React.useState<string | undefined>(new Date().toISOString());
  const [_poStatus, setPoStatus] = React.useState<SignOffStatus>('Approved');

  const [_paPicker, setPaPicker] = React.useState<IPersonaProps[]>([]);
  const [_paDate, setPaDate] = React.useState<string | undefined>(undefined);
  const [_paStatus, setPaStatus] = React.useState<SignOffStatus>('Pending');
  const [_paStatusEnabled, setPaStatusEnabled] = React.useState<boolean>(false);

  const [_piPicker, setPiPicker] = React.useState<IPersonaProps[]>([]);
  const [_piDate, setPiDate] = React.useState<string | undefined>(undefined);
  const [_piStatus, setPiStatus] = React.useState<SignOffStatus>('Pending');
  const [_piStatusEnabled, setPiStatusEnabled] = React.useState<boolean>(false);
  const [_piUnlockedByPA, setPiUnlockedByPA] = React.useState<boolean>(false);

  const [_assetDirPicker, setAssetDirPicker] = React.useState<IPersonaProps[]>([]);
  const [_assetDirDate, setAssetDirDate] = React.useState<string | undefined>(undefined);
  const [_assetDirStatus, setAssetDirStatus] = React.useState<SignOffStatus>('Pending');
  const [_assetDirStatusEnabled, setAssetDirStatusEnabled] = React.useState<boolean>(false);

  const [_hseDirPicker, setHseDirPicker] = React.useState<IPersonaProps[]>([]);
  const [_hseDirDate, setHseDirDate] = React.useState<string | undefined>(undefined);
  const [_hseDirStatus, setHseDirStatus] = React.useState<SignOffStatus>('Pending');
  const [_hseDirStatusEnabled, setHseDirStatusEnabled] = React.useState<boolean>(false);

  // PTW Closure state
  const [_closurePoDate, setClosurePoDate] = React.useState<string | undefined>(undefined);
  const [_closurePoStatus, setClosurePoStatus] = React.useState<SignOffStatus>('Pending');

  const [_closureAssetManagerPicker, setClosureAssetManagerPicker] = React.useState<IPersonaProps[]>([]);
  const [_closureAssetManagerDate, setClosureAssetManagerDate] = React.useState<string | undefined>(undefined);
  const [_closureAssetManagerStatus, setClosureAssetManagerStatus] = React.useState<SignOffStatus>('Pending');
  const [_closureAssetManagerStatusEnabled, setClosureAssetManagerStatusEnabled] = React.useState<boolean>(false);

  // const isSubmitted = mode === 'submitted';
  const isHighRisk = String(_overAllRiskAssessment || '').toLowerCase().includes('high');

  // Determine if current user is the Permit Originator
  const currentUserEmail = (props.context?.pageContext?.user?.email || '').toLowerCase();
  const permitOriginatorEmail = (_PermitOriginator?.[0]?.secondaryText || '').toLowerCase();

  // Resolve eligibility from SP group membership for Permit Originator Group Logged In Users
  React.useEffect(() => {
    let disposed = false;
    async function PermitOriginatorGroup() {
      try {
        const spOps = spCrudRef.current ?? new SPCrudOperations((props.context as any).spHttpClient, webUrl, '', '');
        const isEligibleGroup = await spOps._IsUserInSPGroup('PermitOriginatorGroup', currentUserEmail);
        if (!isEligibleGroup) { if (!disposed) setIsPermitOriginator(false); return; }
        if (!disposed) setIsPermitOriginator(isEligibleGroup);
      } catch {
        if (!disposed) setIsPermitOriginator(false);
      }
    }

    if (currentUserEmail) PermitOriginatorGroup();
    return () => { disposed = true; };
  }, [props.context.spHttpClient, webUrl, currentUserEmail]);

  // Resolve eligibility from SP group membership for Performing Authority Group Logged In Users
  React.useEffect(() => {
    let disposed = false;
    async function PerformingAuthorityGroup() {
      try {
        const spOps = spCrudRef.current ?? new SPCrudOperations((props.context as any).spHttpClient, webUrl, '', '');
        const isEligibleGroup = await spOps._IsUserInSPGroup('PerformingAuthorityGroup', currentUserEmail);
        if (!isEligibleGroup) { if (!disposed) setIsPerformingAuthority(false); return; }
        if (!disposed) setIsPerformingAuthority(isEligibleGroup);
      } catch {
        if (!disposed) setIsPerformingAuthority(false);
      }
    }

    if (currentUserEmail) PerformingAuthorityGroup();
    return () => { disposed = true; };
  }, [props.context.spHttpClient, webUrl, currentUserEmail]);

  // Resolve eligibility from SP group membership for Permit Issuer Group Logged In Users
  React.useEffect(() => {
    let disposed = false;
    async function PermitIssuerGroup() {
      try {
        const spOps = spCrudRef.current ?? new SPCrudOperations((props.context as any).spHttpClient, webUrl, '', '');
        const isEligibleGroup = await spOps._IsUserInSPGroup('PermitIssuerGroup', currentUserEmail);
        if (!isEligibleGroup) { if (!disposed) setIsPermitIssuer(false); return; }
        if (!disposed) setIsPermitIssuer(isEligibleGroup);
      } catch {
        if (!disposed) setIsPermitIssuer(false);
      }
    }

    if (currentUserEmail) PermitIssuerGroup();
    return () => { disposed = true; };
  }, [props.context.spHttpClient, webUrl, currentUserEmail]);

  // Resolve eligibility from SP group membership for Asset Director Group Logged In Users
  React.useEffect(() => {
    let disposed = false;
    async function AssetDirectorsGroup() {
      try {
        const spOps = spCrudRef.current ?? new SPCrudOperations((props.context as any).spHttpClient, webUrl, '', '');
        const isEligibleGroup = await spOps._IsUserInSPGroup('AssetDirectorsGroup', currentUserEmail);
        if (!isEligibleGroup) { if (!disposed) setIsAssetDirector(false); return; }
        if (!disposed) setIsAssetDirector(isEligibleGroup);
      } catch {
        if (!disposed) setIsAssetDirector(false);
      }
    }

    if (currentUserEmail) AssetDirectorsGroup();
    return () => { disposed = true; };
  }, [props.context.spHttpClient, webUrl, currentUserEmail]);

  // Resolve eligibility from SP group membership for HSE Director Group Logged In Users
  React.useEffect(() => {
    let disposed = false;
    async function HSEDirectorGroup() {
      try {
        const spOps = spCrudRef.current ?? new SPCrudOperations((props.context as any).spHttpClient, webUrl, '', '');
        const isEligibleGroup = await spOps._IsUserInSPGroup('HSEDirectorGroup', currentUserEmail);
        if (!isEligibleGroup) { if (!disposed) setIsHSEDirector(false); return; }
        if (!disposed) setIsHSEDirector(isEligibleGroup);
      } catch {
        if (!disposed) setIsHSEDirector(false);
      }
    }

    if (currentUserEmail) HSEDirectorGroup();
    return () => { disposed = true; };
  }, [props.context.spHttpClient, webUrl, currentUserEmail]);

  // Resolve eligibility from SP group membership for Asset Managers Group Logged In Users
  React.useEffect(() => {
    let disposed = false;
    async function AssetManagersGroup() {
      try {
        const spOps = spCrudRef.current ?? new SPCrudOperations((props.context as any).spHttpClient, webUrl, '', '');
        const isEligibleGroup = await spOps._IsUserInSPGroup('AssetManagersGroup', currentUserEmail);
        if (!isEligibleGroup) { if (!disposed) setIsAssetManager(false); return; }
        if (!disposed) setIsAssetManager(isEligibleGroup);
      } catch {
        if (!disposed) setIsAssetManager(false);
      }
    }

    if (currentUserEmail) AssetManagersGroup();
    return () => { disposed = true; };
  }, [props.context.spHttpClient, webUrl, currentUserEmail]);

  const getGroupMembers = React.useCallback(async (groupName: string): Promise<SPGroupUser[]> => {
    const url = `${webUrl}/_api/web/sitegroups/getbyname('${encodeURIComponent(groupName)}')/users?$select=Id,Title,Email`;
    const res = await props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!res.ok) return [];
    const json: any = await res.json();
    const users: SPGroupUser[] = Array.isArray(json?.value)
      ? json.value.map((u: any) => ({ id: u.Id, title: u.Title, email: u.Email }))
      : [];
    return users;
  }, [props.context.spHttpClient, webUrl]);

  React.useEffect(() => {
    let disposed = false;
    (async () => {
      try {
        const names = [
          'PermitOriginatorGroup',
          'PerformingAuthorityGroup',
          'PermitIssuerGroup',
          'AssetDirectorsGroup',
          'HSEDirectorGroup',
          'AssetManagersGroup'
        ];
        const entries = await Promise.all(
          names.map(async n => [n, await getGroupMembers(n)] as const)
        );
        if (!disposed) {
          const map: Record<string, SPGroupUser[]> = {};
          entries.forEach(([name, users]) => { map[name] = users; });
          setGroupMembers(map);
        }
      } catch {
        if (!disposed) setGroupMembers({});
      }
    })();
    return () => { disposed = true; };
  }, [getGroupMembers]);

  // Build dropdown options from a group
  const getOptionsForGroup = React.useCallback(
    (groupName: string): IDropdownOption[] =>
      (groupMembers[groupName] || []).map(m => ({
        key: m.email || String(m.id),
        text: m.title || m.email || String(m.id)
      })),
    [groupMembers]
  );

  // Current selected email from a PeoplePicker-like state
  const selectedKeyFromPicker = React.useCallback(
    (picker?: IPersonaProps[]) => picker?.[0]?.secondaryText || undefined,
    []
  );
  // Handler for single-approver ComboBox change
  const onSingleApproverChange = React.useCallback((groupName: string, setPicker: (items: IPersonaProps[]) => void, setStatusEnabled?: (enabled: boolean) => void) =>
    (_: React.FormEvent<IComboBox>, opt?: IDropdownOption) => {
      if (!opt) {
        setPicker([]);
        setStatusEnabled?.(false);
        if (groupName === 'PerformingAuthorityGroup') setPiUnlockedByPA(false)
        return;
      }
      const emailKey = String(opt.key);
      const u = (groupMembers[groupName] || []).find(x => (x.email || String(x.id)) === emailKey);
      const selectedEmail = (u?.email || '').toLowerCase();
      const isCurrentUser = !!selectedEmail && selectedEmail === currentUserEmail;
      setStatusEnabled?.(isCurrentUser);
      // If PA equals the logged-in user, unlock the PI section (ComboBox + Status gating)
      if (groupName === 'PerformingAuthorityGroup') {
        setPiUnlockedByPA(isCurrentUser);
      }

      setPicker(u ? [{
        text: u.title || '',
        secondaryText: u.email || '',
        id: String(u.id)
      }] : []);
    },
    [groupMembers, currentUserEmail]
  );

  const stageEnabled = React.useMemo(() => {
    const poEnabled = isPermitOriginator; // Originator signs first
    const paEnabled = isPerformingAuthority && _poStatus === 'Approved';
    const piEnabled = (isPermitIssuer && _paStatus === 'Approved') || _piUnlockedByPA;
    // High risk requires AD then HSE; otherwise skip to closure after PI
    const assetDirEnabled = isHighRisk && isAssetDirector && _piStatus === 'Approved';
    const hseDirEnabled = isHighRisk && isHSEDirector && _assetDirStatus === 'Approved';
    const closureEnabled = isAssetManager && (
      (isHighRisk ? _hseDirStatus === 'Approved' : _piStatus === 'Approved')
    );
    return { poEnabled, paEnabled, piEnabled, assetDirEnabled, hseDirEnabled, closureEnabled };
  }, [
    isPermitOriginator, isPerformingAuthority, isPermitIssuer, isAssetDirector, isHSEDirector, isAssetManager,
    _poStatus, _paStatus, _piStatus, _assetDirStatus, _hseDirStatus, isHighRisk, _piUnlockedByPA
  ]);

  // State for controlling conditional rendering of sections
  const [workPermitRequired, setWorkPermitRequired] = React.useState<boolean>(false);

  const statusOptions: IDropdownOption[] = React.useMemo(() => ([
    { key: 'Pending', text: 'Pending' },
    { key: 'Approved', text: 'Approved' },
    { key: 'Rejected', text: 'Rejected' },
    { key: 'Closed', text: 'Closed' }
  ]), []);

  // Styling Components
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

  const toPersona = (obj?: { Id?: any; Title?: string; EMail?: string; FullName?: string }): IPersonaProps | undefined => {
    if (!obj) return undefined;
    const text = obj.FullName || obj.Title || '';
    const email = obj.EMail || '';
    const id = obj.Id != null ? String(obj.Id) : text;
    return { text, secondaryText: email, id } as IPersonaProps;
  };

  // ---------------------------
  // Data-loading functions (ported)
  // ---------------------------

  // const isEditMode = React.useMemo(() => {
  //   const editFormId = props.formId ? Number(props.formId) : undefined;
  //   return !!(editFormId && editFormId > 0);
  // }, [props.formId]);

  const [formStatus, setFormStatus] = React.useState<string>(() => {
    try {
      const raw = localStorage.getItem('FormStatusRecord');
      return raw ? String(JSON.parse(raw)?.value || '').toLowerCase() : '';
    } catch { return ''; }
  });

  React.useEffect(() => {

    const onStorage = (e: StorageEvent) => {
      if (e.key === 'FormStatusRecord') {
        try {
          const v = e.newValue ? String(JSON.parse(e.newValue)?.value || '').toLowerCase() : '';
          setFormStatus(v);
        } catch { /* no-op */ }
      }
    };

    window.addEventListener('storage', onStorage);
    return () => window.removeEventListener('storage', onStorage);
  }, []);

  const isSubmitted = (formStatus || mode) === 'submitted';
  // const isSaved = (formStatus || mode) === 'saved';

  // const isPermitsAllClosed = React.useMemo(() => {
  //   const formStatus = JSON.parse(localStorage.getItem("FormStatusRecord") || '{value: ""}').value
  //   return formStatus === "Closed";
  // },[localStorage]);

  const ptwStructureSelect = React.useMemo(() => (
    `?$select=Id,AttachmentsProvided,InitialRisk,ResidualRisk,OverallRiskAssessment,FireWatchNeeded,GasTestRequired,` +
    `CoralFormId/Title,CoralFormId/ArabicTitle,` +
    `CompanyRecord/Id,CompanyRecord/Title,CompanyRecord/RecordOrder,` +
    `WorkCategory/Id,WorkCategory/Title,WorkCategory/OrderRecord,WorkCategory/RenewalValidity,` +
    `HACWorkArea/Id,HACWorkArea/Title,HACWorkArea/OrderRecord,` +
    `WorkHazards/Id,WorkHazards/Title,WorkHazards/OrderRecord,` +
    `Machinery/Id,Machinery/Title,Machinery/OrderRecord,` +
    `PrecuationItems/Id,PrecuationItems/Title,PrecuationItems/OrderRecord,` +
    `ProtectiveSafetyEquiment/Id,ProtectiveSafetyEquiment/Title,ProtectiveSafetyEquiment/OrderRecord` +
    `&$expand=CoralFormId,CompanyRecord,WorkCategory,HACWorkArea,WorkHazards,Machinery,PrecuationItems,` +
    `ProtectiveSafetyEquiment`
  ), []);

  // const _getUsers = React.useCallback(async (EMail?: string, displayName?: string): Promise<IUser[]> => {
  //   let fetched: IUser[] = [];
  //   let endpoint: string | null = "/users?$select=id,displayName,mail,department,jobTitle,mobilePhone,officeLocation&$expand=manager($select=id,displayName)";

  //   try {
  //     do {
  //       const client: MSGraphClientV3 = await (props.context as any).msGraphClientFactory.getClient("3");
  //       const response: IGraphResponse = await client.api(endpoint).get();
  //       if (response?.value && Array.isArray(response.value)) {
  //         const seenIds = new Set<string>();
  //         const mappedUsers = response.value
  //           .filter((u: IGraphUserResponse) => u.mail)
  //           .filter((user) => user.mail && !user.mail?.toLowerCase().includes("healthmailbox") && !user.mail?.toLowerCase().includes("softflow-intl.com") && !user.mail?.toLowerCase().includes("sync"))
  //           .filter(user => {
  //             if (seenIds.has(user.id)) return false;
  //             seenIds.add(user.id);
  //             return true;
  //           })
  //           .map((user: IGraphUserResponse) => ({
  //             id: user.id,
  //             displayName: user.displayName,
  //             email: user.mail,
  //             jobTitle: user.jobTitle,
  //             department: user.department,
  //             officeLocation: user.officeLocation,
  //             mobilePhone: user.mobilePhone,
  //             profileImageUrl: undefined,
  //             isSelected: false,
  //             manager: user.manager ? { id: user.manager.id, displayName: user.manager.displayName } : undefined,
  //           } as IUser));
  //         fetched.push(...mappedUsers);
  //         endpoint = (response as any)["@odata.nextLink"] || null;
  //       } else {
  //         endpoint = null;
  //       }
  //     } while (endpoint);
  //     if (fetched.length > 0) setUsers(fetched);
  //     return fetched;
  //   } catch (error) {
  //     // console.error("Error fetching users:", error);
  //     setUsers([]);
  //     return [];
  //   }
  // }, [props.context]);

  const _getUsers = React.useCallback(async (EMail?: string, displayName?: string, top: number = 25): Promise<IUser[]> => {
    const termRaw = (displayName || EMail || '').trim();
    if (!termRaw) return [];
    const term = termRaw.replace(/"/g, '');

    try {
      const client: MSGraphClientV3 = await (props.context as any).msGraphClientFactory.getClient("3");
      let res: any;
      // Try ranked search first (needs ConsistencyLevel: eventual)
      try {
        res = await client
          .api('/users')
          .header('ConsistencyLevel', 'eventual')
          .search(`"displayName:${term}" OR "mail:${term}"`)
          .select('id,displayName,mail,department,jobTitle,mobilePhone,officeLocation')
          .top(top)
          .get();
      } catch {
        // Fallback to $filter startswith
        const t = term.toLowerCase();
        const filter = `startswith(tolower(displayName),'${t}') or startswith(tolower(mail),'${t}') or startswith(tolower(userPrincipalName),'${t}')`;
        res = await client
          .api(`/users?$select=id,displayName,mail,department,jobTitle,mobilePhone,officeLocation&$filter=${encodeURIComponent(filter)}&$top=${top}`)
          .get();
      }

      const seen = new Set<string>();
      const mapped: IUser[] = (res?.value || [])
        .filter((u: any) => u.mail)
        .filter((u: any) => {
          const m = (u.mail || '').toLowerCase();
          return !m.includes('healthmailbox') && !m.includes('softflow-intl.com') && !m.includes('sync');
        })
        .filter((u: any) => !seen.has(u.id) && seen.add(u.id))
        .map((u: any) => ({
          id: u.id,
          displayName: u.displayName,
          email: u.mail,
          jobTitle: u.jobTitle,
          department: u.department,
          officeLocation: u.officeLocation,
          mobilePhone: u.mobilePhone,
          profileImageUrl: undefined,
          isSelected: false,
          manager: undefined
        } as IUser));
      setUsers(mapped);
      return mapped;
    } catch (error) {
      // console.error("Error fetching users:", error);
      setUsers([]);
      return [];
    }
  }, [props.context]);

  const _getCoralFormsList = React.useCallback(async (): Promise<ICoralFormsList | undefined> => {
    try {

      const searchEscaped = formName.replace(/'/g, "''");
      const query: string = `?$select=Id,Title,hasInstructionForUse,hasWorkflow,SubmissionRangeInterval&$filter=substringof('${searchEscaped}', Title)`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'Coral_Forms_List', query);
      const data = await spCrudRef.current._getItemsWithQuery();

      const ppeform = data.find((obj: any) => obj !== null);
      let result: ICoralFormsList = { Id: "" };

      if (ppeform) {

        result = {
          Id: ppeform.Id ?? undefined,
          Title: ppeform.Title ?? undefined,
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

  const _getPTWFormStructure = React.useCallback(async () => {
    try {
      const query: string = ptwStructureSelect;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Items', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      let result: IPTWForm;

      if (data && data.length > 0) {
        const obj = data[0];
        const coralForm: ICoralForm = obj.CoralFormId ? {
          id: obj.CoralFormId.Id !== undefined && obj.CoralFormId.Id !== null ? obj.CoralFormId.Id : undefined,
          title: obj.CoralFormId.Title !== undefined && obj.CoralFormId.Title !== null ? obj.CoralFormId.Title : undefined,
          arTitle: obj.CoralFormId.ArabicTitle !== undefined && obj.CoralFormId.ArabicTitle !== null ? obj.CoralFormId.ArabicTitle : undefined,
          hasInstructionsForUse: obj.CoralFormId.hasInstructionsForUse !== undefined && obj.CoralFormId.hasInstructionsForUse !== null ? obj.CoralFormId.hasInstructionsForUse : undefined,
        } : '{}' as ICoralForm;

        const _companies: ILookupItem[] = [];
        if (obj.CompanyRecord !== undefined && obj.CompanyRecord !== null && Array.isArray(obj.CompanyRecord)) {
          obj.CompanyRecord.forEach((item: any) => {
            if (item) {
              _companies.push({
                id: item.Id,
                title: item.Title,
                orderRecord: item.OrderRecord || 0,
              });
            }
          });
        }

        const _workCategories: IWorkCategory[] = [];
        if (obj.WorkCategory !== undefined && obj.WorkCategory !== null && Array.isArray(obj.WorkCategory)) {
          obj.WorkCategory.forEach((item: any) => {
            if (item) {
              _workCategories.push({
                id: item.Id,
                title: item.Title,
                orderRecord: item.OrderRecord || 0,
                renewalValidity: item.RenewalValidity || 0,
                isChecked: false,
              });
            }
          });
        }

        const _hacWorkAreas: ILookupItem[] = [];
        if (obj.HACWorkArea !== undefined && obj.HACWorkArea !== null && Array.isArray(obj.HACWorkArea)) {
          obj.HACWorkArea.forEach((item: any) => {
            if (item) {
              _hacWorkAreas.push({
                id: item.Id,
                title: item.Title,
                orderRecord: item.OrderRecord || 0
              });
            }
          });
        }

        const _workHazardosList: ILookupItem[] = [];
        if (obj.WorkHazards !== undefined && obj.WorkHazards !== null && Array.isArray(obj.WorkHazards)) {
          obj.WorkHazards.forEach((item: any) => {
            if (item) {
              _workHazardosList.push({
                id: item.Id,
                title: item.Title,
                orderRecord: item.OrderRecord || 0,
              });
            }
          });
        }

        const _machineryList: ILookupItem[] = [];
        if (obj.Machinery !== undefined && obj.Machinery !== null && Array.isArray(obj.Machinery)) {
          obj.Machinery.forEach((item: any) => {
            if (item) {
              _machineryList.push({
                id: item.Id,
                title: item.Title,
                orderRecord: item.OrderRecord || 0,
              });
            }
          });
        }

        const _precuationsItemsList: ILookupItem[] = [];
        if (obj.PrecuationItems !== undefined && obj.PrecuationItems !== null && Array.isArray(obj.PrecuationItems)) {
          obj.PrecuationItems.forEach((item: any) => {
            if (item) {
              _precuationsItemsList.push({
                id: item.Id,
                title: item.Title,
                orderRecord: item.OrderRecord || 0,
              });
            }
          });
        }

        const _protectiveSafetyEquipmentsList: ILookupItem[] = [];
        if (obj.ProtectiveSafetyEquiment !== undefined && obj.ProtectiveSafetyEquiment !== null && Array.isArray(obj.ProtectiveSafetyEquiment)) {
          obj.ProtectiveSafetyEquiment.forEach((item: any) => {
            if (item) {
              _protectiveSafetyEquipmentsList.push({
                id: item.Id,
                title: item.Title,
                orderRecord: item.OrderRecord || 0,
              });
            }
          });
        }

        result = {
          id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
          coralForm: coralForm, companies: _companies,
          workCategories: _workCategories, hacWorkAreas: _hacWorkAreas,
          workHazardosList: _workHazardosList, machinaries: _machineryList,
          precuationsItems: _precuationsItemsList,
          protectiveSafetyEquipments: _protectiveSafetyEquipmentsList,
          attachmentsProvided: obj.AttachmentsProvided !== undefined && obj.AttachmentsProvided !== null ? obj.AttachmentsProvided : undefined,
          gasTestRequired: obj.GasTestRequired !== undefined && obj.GasTestRequired !== null ? obj.GasTestRequired : undefined,
          fireWatchNeeded: obj.FireWatchNeeded !== undefined && obj.FireWatchNeeded !== null ? obj.FireWatchNeeded : undefined,
          overallRiskAssessment: obj.OverallRiskAssessment !== undefined && obj.OverallRiskAssessment !== null ? obj.OverallRiskAssessment : undefined,
          initialRisk: obj.InitialRisk !== undefined && obj.InitialRisk !== null ? obj.InitialRisk : undefined,
          residualRisk: obj.ResidualRisk !== undefined && obj.ResidualRisk !== null ? obj.ResidualRisk : undefined,
          personnalInvolved: [],
          issuanceInstrunctions: [],
          assetsCategories: [],
          assetsDetails: []
        };
        setPTWFormStructure(result);
      }

    } catch (error) {
      setPTWFormStructure({ issuanceInstrunctions: [], personnalInvolved: [] });
    }
  }, [props.context, spHelpers, spCrudRef, ptwStructureSelect]);

  const _getLKPItemInstructionsForUse = React.useCallback(async (formName?: string) => {
    try {
      const query: string = `?$select=Id,FormName,RecordOrder,Description&$filter=substringof('${formName}', FormName)&$orderby=RecordOrder asc`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'LKP_Item_Instructions_For_Use', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: ILKPItemInstructionsForUse[] = [];
      data.forEach((obj: any) => {
        if (obj) {
          const temp: ILKPItemInstructionsForUse = {
            Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            FormName: obj.FormName !== undefined && obj.FormName !== null ? obj.FormName : undefined,
            Order: obj.RecordOrder !== undefined && obj.RecordOrder !== null ? obj.Order : undefined,
            Description: obj.Description !== undefined && obj.Description !== null ? obj.Description : undefined,
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

  const _getPersonnelInvolved = React.useCallback(async () => {
    try {
      const query: string = `?$select=Id,EmployeeRecord/Id,EmployeeRecord/FullName,IsHSEInductionCompleted,IsFireFightingTrained` +
        `&$expand=EmployeeRecord` +
        `&$filter=IsHSEInductionCompleted eq 1`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'Employee_Personelle_Passport', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: IEmployeePeronellePassport[] = [];
      data.forEach((obj: any) => {
        if (obj) {
          const temp: IEmployeePeronellePassport = {
            Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            fullName: obj.EmployeeRecord?.FullName !== undefined && obj.EmployeeRecord?.FullName !== null ? obj.EmployeeRecord.FullName : undefined,
            isHSEInductionCompleted: obj.IsHSEInductionCompleted !== undefined && obj.IsHSEInductionCompleted !== null ? obj.IsHSEInductionCompleted : undefined,
            isFireFightingTrained: obj.IsFireFightingTrained !== undefined && obj.IsFireFightingTrained !== null ? obj.IsFireFightingTrained : undefined,
          };
          result.push(temp);
        }
      });
      setPersonnelInvolved(result);
    } catch (error) {
      setPersonnelInvolved([]);
      // console.error('An error has occurred while retrieving items!', error);
    }
  }, [props.context, spHelpers]);

  const _getAssetCategories = React.useCallback(async (): Promise<ILookupItem[]> => {
    try {
      const query: string = `?$select=Id,Title,OrderRecord&$orderby=OrderRecord asc`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'LKP_Asset_Category', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: ILookupItem[] = [];
      data.forEach((obj: any) => {
        if (obj) {
          const temp: ILookupItem = {
            id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            title: obj.Title !== undefined && obj.Title !== null ? obj.Title : undefined,
            orderRecord: obj.OrderRecord !== undefined && obj.OrderRecord !== null ? obj.OrderRecord : undefined,
          };
          result.push(temp);
        }
      });
      setAssetCategoriesList(result);
      return result;
    } catch (error) {
      setAssetCategoriesList([]);
      return [];
    }
  }, [props.context]);

  // Modified _getAssetDetails function
  const _getAssetDetails = React.useCallback(async () => {
    try {
      const query: string = `?$select=Id,Title,OrderRecord,` +
        `AssetDirector/Id,AssetDirector/EMail,` +
        `AssetManager/Id,AssetManager/EMail,` +
        `HSEPartner/Id,HSEPartner/EMail,` +
        `AssetCategoryRecord/Id,AssetCategoryRecord/Title,AssetCategoryRecord/OrderRecord` +
        `&$expand=AssetCategoryRecord,AssetDirector,AssetManager,HSEPartner`;

      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'LKP_Asset_Details', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const detailsByCategory = new Map<number, IAssetsDetails[]>();
      // Process asset details and group them by category
      data.forEach((obj: any) => {
        if (obj && obj.AssetCategoryRecord?.Id) {
          const categoryId = obj.AssetCategoryRecord.Id;

          const assetDetail: IAssetsDetails = {
            id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            title: obj.Title !== undefined && obj.Title !== null ? obj.Title : undefined,
            orderRecord: obj.OrderRecord !== undefined && obj.OrderRecord !== null ? obj.OrderRecord : undefined,
            assetCategoryId: categoryId,
            assetDirector: obj.AssetDirector !== undefined && obj.AssetDirector !== null ? toPersona(obj.AssetDirector) : undefined,
            assetManager: obj.AssetManager !== undefined && obj.AssetManager !== null ? toPersona(obj.AssetManager) : undefined,
            hsePartner: obj.HSEPartner !== undefined && obj.HSEPartner !== null ? toPersona(obj.HSEPartner) : undefined,
          };

          if (!detailsByCategory.has(categoryId)) {
            detailsByCategory.set(categoryId, []);
          }
          detailsByCategory.get(categoryId)!.push(assetDetail);
        }
      });

      setAssetCategoriesDetailsList(data.map((obj: any) => ({
        id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
        title: obj.Title !== undefined && obj.Title !== null ? obj.Title : undefined,
        orderRecord: obj.OrderRecord !== undefined && obj.OrderRecord !== null ? obj.OrderRecord : undefined,
        assetCategoryId: obj.AssetCategoryRecord?.Id !== undefined && obj.AssetCategoryRecord?.Id !== null ? obj.AssetCategoryRecord.Id : undefined,
        assetDirector: obj.AssetDirector !== undefined && obj.AssetDirector !== null ? toPersona(obj.AssetDirector) : undefined,
        assetManager: obj.AssetManager !== undefined && obj.AssetManager !== null ? toPersona(obj.AssetManager) : undefined,
        hsePartner: obj.HSEPartner !== undefined && obj.HSEPartner !== null ? toPersona(obj.HSEPartner) : undefined,
      })));

    } catch (error) {
      setAssetDetails([]);
      setAssetCategoriesDetailsList([]);
      setPTWFormStructure(prev => ({
        ...prev,
        assetsCategories: [],
        assetsDetails: []
      }));
    }
  }, [props.context]);

  const _getWorkSafeguards = React.useCallback(async (): Promise<ISagefaurdsItem[]> => {
    try {
      const query: string = `?$select=Id,Title,OrderRecord,WorkCatetegoryRecord/Id,WorkCatetegoryRecord/Title` +
        `&$expand=WorkCatetegoryRecord` +
        `&$orderby=OrderRecord asc`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'LKP_Safegaurds', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: ISagefaurdsItem[] = [];
      data.forEach((obj: any) => {
        if (obj) {
          const temp: ISagefaurdsItem = {
            id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            title: obj.Title !== undefined && obj.Title !== null ? obj.Title : undefined,
            orderRecord: obj.OrderRecord !== undefined && obj.OrderRecord !== null ? obj.OrderRecord : undefined,
            workCategoryId: obj.WorkCatetegoryRecord?.Id !== undefined && obj.WorkCatetegoryRecord?.Id !== null ? obj.WorkCatetegoryRecord.Id : undefined,
            workCategoryTitle: obj.WorkCatetegoryRecord?.Title !== undefined && obj.WorkCatetegoryRecord?.Title !== null ? obj.WorkCatetegoryRecord.Title : undefined,
          };
          result.push(temp);
        }
      });
      setSafeguards(result);
      setFilteredSafeguards(result);
      return result;
    } catch (error) {
      setSafeguards([]);
      setFilteredSafeguards([]);
      return [];
    }
  }, [props.context]);

  const _getPTWWorkflow = React.useCallback(async (): Promise<IPTWWorkflow> => {
    try {
      const query: string = `?$select=Id,PTWForm/Id,PTWForm/CoralReferenceNumber,POStatus,PAStatus,PIStatus,AssetDirectorStatus,HSEDirectorStatus,POClosureStatus,AssetManagerStatus,` +
        `POApprovalDate,PAApprovalDate,PIApprovalDate,AssetDirectorApprovalDate,HSEDirectorApprovalDate,POClosureDate,AssetManagerDate,` +
        `POApprover/Id,POApprover/EMail,` +
        `PAApprover/Id,PAApprover/EMail,` +
        `PIApprover/Id,PIApprover/EMail,` +
        `AssetDirectorApprover/Id,AssetDirectorApprover/EMail,` +
        `HSEDirectorApprover/Id,HSEDirectorApprover/EMail,` +
        `POClosureApprover/Id,POClosureApprover/EMail,` +
        `AssetManagerApprover/Id,AssetManagerApprover/EMail,` +
        `&$expand=PTWForm,POApprover,PAApprover,PIApprover,AssetDirectorApprover,HSEDirectorApprover,POClosureApprover,AssetManagerApprover`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form_Approval_Workflow', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      let result: IPTWWorkflow = {} as IPTWWorkflow;

      if (data && data.length > 0) {
        const obj = data[0];
        result = {
          id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
          PTWFormId: obj.PTWForm?.Id !== undefined && obj.PTWForm?.Id !== null ? obj.PTWForm.Id : undefined,
          CoralReferenceNumber: obj.PTWForm?.CoralReferenceNumber !== undefined && obj.PTWForm?.CoralReferenceNumber !== null ? obj.PTWForm.CoralReferenceNumber : undefined,
          POApprover: obj.POApprover !== undefined && obj.POApprover !== null ? toPersona(obj.POApprover) : undefined,
          POApprovalDate: obj.POApprovalDate !== undefined && obj.POApprovalDate !== null ? obj.POApprovalDate : undefined,
          POStatus: obj.POStatus !== undefined && obj.POStatus !== null ? obj.POStatus : undefined,

          PAApprover: obj.PAApprover !== undefined && obj.PAApprover !== null ? toPersona(obj.PAApprover) : undefined,
          PAApprovalDate: obj.PAApprovalDate !== undefined && obj.PAApprovalDate !== null ? obj.PAApprovalDate : undefined,
          PAStatus: obj.PAStatus !== undefined && obj.PAStatus !== null ? obj.PAStatus : undefined,

          PIApprover: obj.PIApprover !== undefined && obj.PIApprover !== null ? toPersona(obj.PIApprover) : undefined,
          PIApprovalDate: obj.PIApprovalDate !== undefined && obj.PIApprovalDate !== null ? obj.PIApprovalDate : undefined,
          PIStatus: obj.PIStatus !== undefined && obj.PIStatus !== null ? obj.PIStatus : undefined,

          AssetDirectorApprover: obj.AssetDirectorApprover !== undefined && obj.AssetDirectorApprover !== null ? toPersona(obj.AssetDirectorApprover) : undefined,
          AssetDirectorApprovalDate: obj.AssetDirectorApprovalDate !== undefined && obj.AssetDirectorApprovalDate !== null ? obj.AssetDirectorApprovalDate : undefined,
          AssetDirectorStatus: obj.AssetDirectorStatus !== undefined && obj.AssetDirectorStatus !== null ? obj.AssetDirectorStatus : undefined,

          HSEDirectorApprover: obj.HSEDirectorApprover !== undefined && obj.HSEDirectorApprover !== null ? toPersona(obj.HSEDirectorApprover) : undefined,
          HSEDirectorApprovalDate: obj.HSEDirectorApprovalDate !== undefined && obj.HSEDirectorApprovalDate !== null ? obj.HSEDirectorApprovalDate : undefined,
          HSEDirectorStatus: obj.HSEDirectorStatus !== undefined && obj.HSEDirectorStatus !== null ? obj.HSEDirectorStatus : undefined,

          POClosureApprover: obj.POClosureApprover !== undefined && obj.POClosureApprover !== null ? toPersona(obj.POClosureApprover) : undefined,
          POClosureDate: obj.POClosureDate !== undefined && obj.POClosureDate !== null ? obj.POClosureDate : undefined,
          POClosureStatus: obj.POClosureStatus !== undefined && obj.POClosureStatus !== null ? obj.POClosureStatus : undefined,

          AssetManagerApprover: obj.AssetManagerApprover !== undefined && obj.AssetManagerApprover !== null ? toPersona(obj.AssetManagerApprover) : undefined,
          AssetManagerApprovalDate: obj.AssetManagerApprovalDate !== undefined && obj.AssetManagerApprovalDate !== null ? obj.AssetManagerApprovalDate : undefined,
          AssetManagerStatus: obj.AssetManagerStatus !== undefined && obj.AssetManagerStatus !== null ? obj.AssetManagerStatus : undefined,
        };

        setPermitOriginator(result.POApprover ? [{ text: result.POApprover.text || '', secondaryText: result.POApprover.secondaryText || '', id: result.POApprover.id || '' }] : []);
        setPoDate(result.POApprovalDate ? new Date(result.POApprovalDate).toISOString() : undefined);
        setPoStatus((result.POStatus as SignOffStatus) ?? undefined);

        setPaPicker(result.PAApprover ? [{ text: result.PAApprover.text || '', secondaryText: result.PAApprover.secondaryText || '', id: result.PAApprover.id || '' }] : []);
        setPaDate(result.PAApprovalDate ? new Date(result.PAApprovalDate).toISOString() : undefined);
        setPaStatus((result.PAStatus as SignOffStatus) ?? undefined);

        setPiPicker(result.PIApprover ? [{ text: result.PIApprover.text || '', secondaryText: result.PIApprover.secondaryText || '', id: result.PIApprover.id || '' }] : []);
        setPiDate(result.PIApprovalDate ? new Date(result.PIApprovalDate).toISOString() : undefined);
        setPiStatus((result.PIStatus as SignOffStatus) ?? undefined);

        setAssetDirPicker(result.AssetDirectorApprover ? [{ text: result.AssetDirectorApprover.text || '', secondaryText: result.AssetDirectorApprover.secondaryText || '', id: result.AssetDirectorApprover.id || '' }] : []);
        setAssetDirDate(result.AssetDirectorApprovalDate ? new Date(result.AssetDirectorApprovalDate).toISOString() : undefined);
        setAssetDirStatus((result.AssetDirectorStatus as SignOffStatus) ?? undefined);

        setHseDirPicker(result.HSEDirectorApprover ? [{ text: result.HSEDirectorApprover.text || '', secondaryText: result.HSEDirectorApprover.secondaryText || '', id: result.HSEDirectorApprover.id || '' }] : []);
        setHseDirDate(result.HSEDirectorApprovalDate ? new Date(result.HSEDirectorApprovalDate).toISOString() : undefined);
        setHseDirStatus((result.HSEDirectorStatus as SignOffStatus) ?? undefined);

        setClosurePoStatus((result.POClosureStatus as SignOffStatus) ?? undefined);
        setClosurePoDate(result.POClosureDate ? new Date(result.POClosureDate).toISOString() : undefined);

        setClosureAssetManagerPicker(result.AssetManagerApprover ? [{ text: result.AssetManagerApprover.text || '', secondaryText: result.AssetManagerApprover.secondaryText || '', id: result.AssetManagerApprover.id || '' }] : []);
        setClosureAssetManagerDate(result.AssetManagerApprovalDate ? new Date(result.AssetManagerApprovalDate).toISOString() : undefined);
        setClosureAssetManagerStatus((result.AssetManagerStatus as SignOffStatus) ?? undefined);
      }
      setPtwApprovalWorkflow(result);
      return result;
    } catch (error) {
      setPtwApprovalWorkflow({} as IPTWWorkflow);
      return {} as IPTWWorkflow;
    }
  }, [props.context]);

  // Initial load of users
  React.useEffect(() => {
    let cancelled = false;
    const load = async () => {
      setLoading(true);
      // const fetchedUsers = await _getUsers();
      const meEmail = props.context?.pageContext?.user?.email;
      const fetchedUsers = meEmail ? await _getUsers(meEmail) : [];
      const coralListResult = await _getCoralFormsList();
      await _getPTWFormStructure();
      await _getAssetCategories();
      await _getAssetDetails();
      await _getWorkSafeguards();
      await _getPersonnelInvolved();

      if (coralListResult && coralListResult?.hasInstructionForUse) {
        await _getLKPItemInstructionsForUse(formName);
      }

      if (!cancelled) {
        try {
          // const currentUserEmail = props.context.pageContext.user.email;
          const current = meEmail ? fetchedUsers.find(u => (u.email || '').toLowerCase() === meEmail.toLowerCase()) : undefined;
          if (current) setPermitOriginator([{ text: current.displayName || '', secondaryText: current.email || '', id: current.id }]);
        } catch (e) {
          // ignore if context not available
        }
        setLoading(false);
      }
    };
    load();
    return () => { cancelled = true; };
  }, [props.context, props.formId]);

  // People picker configuration
  const suggestionProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: 'Suggested People',
    mostRecentlyUsedHeaderText: 'Suggested Contacts',
    noResultsFoundText: 'No results found',
    loadingText: 'Loading',
    showRemoveButtons: true,
    suggestionsAvailableAlertText: 'People Picker Suggestions available',
    suggestionsContainerAriaLabel: 'Suggested contacts'
  };

  const _onFilterChanged = (filterText: string, currentPersonas: IPersonaProps[]): Promise<IPersonaProps[]> => {
    const term = (filterText || '').trim();
    if (term.length < 2) return Promise.resolve([]);
    return _getUsers(undefined, term, 25).then(users =>
      users
        .map(u => ({ text: u.displayName || '', secondaryText: u.email || '', id: u.id } as IPersonaProps))
        .filter(p => !currentPersonas.some(cp => cp.id === p.id))
    );
  };

  // Handle asset category change
  const onAssetCategoryChange = (event: React.FormEvent<IComboBox>, item: IDropdownOption | undefined): void => {
    setSelectedAssetCategory(item ? Number(item.key) : undefined);
    setSelectedAssetDetails(0);
    setPiPicker([]);
  };

  // Handle asset details change
  const onAssetDetailsChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption | undefined): void => {
    const selectedId = item ? Number(item.key) : undefined;
    setSelectedAssetDetails(selectedId);
    if (selectedId) {
      // Find the selected asset detail and use its HSEPartner as the Permit Issuer
      const detail = (assetCategoriesDetailsList || []).find(d => Number(d.id) === selectedId);
      const hsePartnerPersona = detail?.hsePartner; // IPersonaProps set in _getAssetDetails

      if (hsePartnerPersona) {
        setPiPicker([hsePartnerPersona]);
        // Enable PI status only if selected PI is the logged-in user
        setPiStatusEnabled(((hsePartnerPersona.secondaryText || '').toLowerCase() === currentUserEmail));
      } else {
        setPiPicker([]);
        setPiStatusEnabled(false);
      }
    } else {
      // Cleared asset details; clear PI selection/status
      setPiPicker([]);
      setPiStatusEnabled(false);
    }
  };

  // Asset details options (filtered by selected category)
  const assetDetailsOptions: IDropdownOption[] = React.useMemo(() => {
    if (!assetCategoriesDetailsList) return [];

    const catId = _selectedAssetCategory !== undefined && _selectedAssetCategory !== null
      ? Number(_selectedAssetCategory)
      : undefined;

    const filtered = catId ? assetCategoriesDetailsList.filter(item => Number(item.assetCategoryId) === catId)
      : assetCategoriesDetailsList;

    return filtered
      .sort((a, b) => (a.orderRecord || 0) - (b.orderRecord || 0))
      .map(item => ({ key: item.id, text: item.title || '' }));

  }, [assetCategoriesDetailsList, _selectedAssetCategory]);

  // Machinery/Tools - multi-select ComboBox wiring
  const machineryOptions = React.useMemo(() => {
    const items = ptwFormStructure?.machinaries || [];
    return items.sort((a, b) => (a.orderRecord || 0) - (b.orderRecord || 0))
      .map(m => ({ key: m.id, text: m.title, selected: _selectedMachineryIds?.includes(Number(m.id)) }));
  }, [ptwFormStructure?.machinaries, _selectedMachineryIds]);

  const onMachineryChange = React.useCallback((_: React.FormEvent<IComboBox>, option?: any) => {
    if (!option) return;
    const idNum = Number(option.key);
    setSelectedMachineryIds(prev => {
      const set = new Set(prev);
      if (option.selected) set.add(idNum); else set.delete(idNum);
      return Array.from(set);
    });
  }, []);

  const selectedMachinery = React.useMemo(() => {
    const list = ptwFormStructure?.machinaries || [];
    const byId = new Map<number, ILookupItem>(list.map(m => [m.id, m]));
    return _selectedMachineryIds?.length ? _selectedMachineryIds
      .map(id => byId.get(Number(id)))
      .filter((m): m is ILookupItem => !!m) : undefined;
  }, [ptwFormStructure?.machinaries, _selectedMachineryIds]);

  const removeMachinery = React.useCallback((id: number) => {
    setSelectedMachineryIds(prev => prev?.filter(x => x !== id));
  }, []);

  // Personnel Involved - multi-select ComboBox wiring
  const personnelOptions = React.useMemo(() => {
    return (personnelInvolved || []).map(p => ({
      key: p.Id,
      text: p.fullName || '',
      selected: _selectedPersonnelIds?.includes(Number(p.Id))
    }));
  }, [personnelInvolved, _selectedPersonnelIds]);

  const onPersonnelChange = React.useCallback((_: React.FormEvent<IComboBox>, option?: any) => {
    if (!option) return;
    const idNum = Number(option.key);
    setSelectedPersonnelIds(prev => {
      const set = new Set(prev);
      if (option.selected) set.add(idNum); else set.delete(idNum);
      return Array.from(set);
    });
  }, []);

  const selectedPersonnel = React.useMemo(() => {
    const byId = new Map<number, IEmployeePeronellePassport>((personnelInvolved || []).map(p => [Number(p.Id), p]));
    return _selectedPersonnelIds?.length ? _selectedPersonnelIds
      .map(id => byId.get(Number(id)))
      .filter((p): p is IEmployeePeronellePassport => !!p) : undefined;
  }, [personnelInvolved, _selectedPersonnelIds]);

  const removePersonnel = React.useCallback((id: number) => {
    setSelectedPersonnelIds(prev => prev?.filter(x => x !== id));
  }, []);

  // Add these handler functions
  const handlePermitTypeChange = React.useCallback((checked?: boolean, workCategory?: IWorkCategory) => {
    // Support multi-select and derive permit rows by the minimum renewal validity across selected categories
    if (!workCategory) {
      setSelectedPermitTypeList([]);
      setPermitPayload([]);
      return;
    }

    setPTWFormStructure(prev => {
      const nextWorkCategories: IWorkCategory[] = (prev.workCategories || []).map(cat =>
        cat.id === workCategory.id ? { ...cat, isChecked: !!checked } : cat
      );

      // Checks if any work category is selected then show all other form sections
      const checkedWorkPermitCount = nextWorkCategories?.filter(item => item.isChecked == true).length;
      setWorkPermitRequired(checkedWorkPermitCount > 0);

      // Compute selected list after this toggle
      const selectedItems = nextWorkCategories.filter(cat => cat.isChecked);
      setSelectedPermitTypeList(selectedItems);
      // Filter safeguards list based on selected work categories
      if (selectedItems.length > 0) {
        const selectedIds = new Set(selectedItems.map(s => s.id));
        setFilteredSafeguards((safeguards || []).filter(s => s.workCategoryId !== undefined && selectedIds.has(s.workCategoryId)));
      } else {
        setFilteredSafeguards([]);
      }

      if (selectedItems.length === 0) {
        setPermitPayload([]);
        setSelectedMachineryIds([]);
        setSelectedPersonnelIds([]);
        setSelectedPrecautionIds(new Set<number>());
        setSelectedProtectiveEquipmentIds(new Set<number>());
        setGasTestResult('');
        setGasTestValue('');
        setFireWatchAssigned('');
        setFireWatchValue('');
        setAttachmentsResult('');
        setAttachmentsValue('');
        setSelectedWorkHazardIds(new Set<number>());
        setSelectedHacWorkAreaId(0);
        setProtectiveEquipmentsOtherText('');
        setPrecautionsOtherText('');
        setProtectiveEquipmentsOtherText('');
        setWorkHazardsOtherText('');
      } else {
        // Minimum number of renewals among selected categories
        const minRenewals = Math.min(...selectedItems.map(cat => (cat.renewalValidity ?? 0)));

        // Preserve any existing row values when possible
        const existingById = new Map(_permitPayload.map(r => [r.id, r] as const));

        const rows: IPermitScheduleRow[] = [];
        // Always include the New Permit row
        rows.push(
          existingById.get('permit-row-0') ?? {
            id: 'permit-row-0', type: 'new', date: '', startTime: '', endTime: '', isChecked: false, orderRecord: 1
          }
        );

        // If renewalValidity indicates "renewable N times", render N renewal rows (1..N)
        for (let i = 1; i < minRenewals; i++) {
          const id = `permit-row-${i}`;
          rows.push(
            existingById.get(id) ?? {
              id, type: 'renewal', date: '', startTime: '', endTime: '', isChecked: false, orderRecord: i + 1
            }
          );
        }

        setPermitPayload(rows);
      }

      return { ...prev, workCategories: nextWorkCategories } as IPTWForm;
    });
  }, [_permitPayload, safeguards]);

  // Keep filtered safeguards in sync if safeguards or selected categories change elsewhere
  React.useEffect(() => {
    if (_selectedPermitTypeList.length > 0) {
      const ids = new Set(_selectedPermitTypeList.map(s => s.id));
      setFilteredSafeguards((safeguards || []).filter(s => s.workCategoryId !== undefined && ids.has(s.workCategoryId)));
    } else {
      setFilteredSafeguards(safeguards || []);
    }
  }, [safeguards, _selectedPermitTypeList]);

  const updatePermitRow = React.useCallback((rowId: string, field: string, value: string, checked: boolean) => {

    setPermitPayload(prevItems => {
      // Helper to compare date-only in UTC
      const toDayUtc = (iso?: string): number => {
        if (!iso) return NaN;
        const d = new Date(iso);
        if (isNaN(d.getTime())) return NaN;
        return Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
      };

      // Find latest previous selected date (by array order) for this row
      const currIndex = prevItems.findIndex(r => r.id === rowId);
      const latestPrevDay = currIndex > 0
        ? Math.max(
          ...prevItems
            .slice(0, currIndex)
            .filter(r => r.isChecked && r.date)
            .map(r => toDayUtc(r.date))
            .filter(n => !isNaN(n)),
          Number.NEGATIVE_INFINITY
        )
        : Number.NEGATIVE_INFINITY;

      return prevItems.map(item => {
        if (item.id !== rowId) return item;

        // Block invalid date chronologically (must be strictly after any previous selected dates)
        if (field === 'date') {
          const newDay = toDayUtc(value);
          if (!isNaN(newDay) && latestPrevDay !== Number.NEGATIVE_INFINITY && newDay <= latestPrevDay) {
            showBanner('Permit date must be after previous selected permit dates.', { autoHideMs: 4000, fade: true, kind: 'error' });
            return item; // block invalid date change
          }
        }
        // Base update for the edited field and selection state
        let next = { ...item, [field]: value, isChecked: !!checked } as IPermitScheduleRow;

        // Determine current start/end after this edit
        const start = field === 'startTime' ? value : item.startTime;
        const end = field === 'endTime' ? value : item.endTime;
        const startMins = spHelpers.parseTimeToMinutes(start);
        const endMins = spHelpers.parseTimeToMinutes(end);

        // Show MessageBar if invalid and handle accordingly
        if (!isNaN(startMins) && !isNaN(endMins) && startMins > endMins) {
          if (field === 'startTime') {
            showBanner('Start time cannot be later than end time.', { autoHideMs: 4000, fade: true, kind: 'error' });
            next = { ...next, endTime: '' }; // clear invalid end time
          } else if (field === 'endTime') {
            showBanner('End time must be after start time.', { autoHideMs: 4000, fade: true, kind: 'error' });
            return item; // block invalid end time change
          }
        }

        // If the row was just deselected via the checkbox, clear the other inputs
        if (field === 'type' && !checked) {
          return { ...next, date: '', startTime: '', endTime: '' };
        }
        return next;
      });
    });
  }, []);

  const handleHACChange = React.useCallback((checked?: boolean, hacArea?: ILookupItem) => {
    if (!hacArea || hacArea.id === undefined || hacArea.id === null) return;
    if (checked) {
      // Single selection: pick this id and deselect others
      setSelectedHacWorkAreaId(hacArea.id);
    } else {
      // If the same item is being unchecked, clear selection
      setSelectedHacWorkAreaId(prev => (prev === hacArea.id ? undefined : prev));
    }
  }, []);

  const showBanner = React.useCallback((text: string, opts?: { autoHideMs?: number; fade?: boolean, kind?: BannerKind }) => {
    setBannerText(text);
    setBannerTick(t => t + 1);
    setBannerOpts(opts);
  }, []);

  const hideBanner = React.useCallback(() => {
    showBanner(``);
    setBannerText(undefined);
    setBannerOpts(undefined);
  }, []);

  // Navigate back to host list view (via callback or URL params)
  const goBackToHost = React.useCallback(() => {
    if (typeof props.onClose === 'function') {
      props.onClose();
      return;
    }
    const url = new URL(window.location.href);
    url.searchParams.delete('mode');
    url.searchParams.delete('formId');
    window.location.href = url.toString();
  }, [props.onClose]);

  const handleCancel = React.useCallback(() => {
    goBackToHost();
  }, [goBackToHost]);

  React.useEffect(() => {
    const selectedEmail = (_paPicker?.[0]?.secondaryText || '').toLowerCase();
    setPaStatusEnabled(!!selectedEmail && selectedEmail === currentUserEmail);
    setPaStatus(!!selectedEmail && selectedEmail === currentUserEmail ? 'Approved' : 'Pending');
  }, [_paPicker, currentUserEmail]);

  React.useEffect(() => {
    const selectedEmail = (_piPicker?.[0]?.secondaryText || '').toLowerCase();
    setPiStatusEnabled(!!selectedEmail && selectedEmail === currentUserEmail);
  }, [_piPicker, currentUserEmail]);

  React.useEffect(() => {
    const selectedEmail = (_assetDirPicker?.[0]?.secondaryText || '').toLowerCase();
    setAssetDirStatusEnabled(!!selectedEmail && selectedEmail === currentUserEmail);
  }, [_assetDirPicker, currentUserEmail]);

  React.useEffect(() => {
    const selectedEmail = (_hseDirPicker?.[0]?.secondaryText || '').toLowerCase();
    setHseDirStatusEnabled(!!selectedEmail && selectedEmail === currentUserEmail);
  }, [_hseDirPicker, currentUserEmail]);

  React.useEffect(() => {
    const selectedEmail = (_closureAssetManagerPicker?.[0]?.secondaryText || '').toLowerCase();
    setClosureAssetManagerStatusEnabled(!!selectedEmail && selectedEmail === currentUserEmail);
  }, [_closureAssetManagerPicker, currentUserEmail]);

  React.useEffect(() => {
    // If selected permit types is unchecked, clear the permit payload
    if (!workPermitRequired) {
      setPermitPayload([]);
      setSelectedHacWorkAreaId(undefined);
      setSelectedWorkHazardIds(new Set<number>());
      setSelectedPrecautionIds(new Set<number>());
      setSelectedProtectiveEquipmentIds(new Set<number>());
      setGasTestValue('');
      setGasTestResult('');
      setFireWatchValue('');
      setGasTestResult('');
      setAttachmentsResult('');
      setAttachmentsValue('');
    }
  }, [workPermitRequired]);

  React.useEffect(() => {
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
  React.useEffect(() => {
    if (!isBusy) return;
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
  }, [isBusy]);

  const mergeRiskRows = (prev?: IRiskTaskRow[], next?: IRiskTaskRow[]): IRiskTaskRow[] => {
    if (!next || next.length === 0) return prev ? prev.slice() : [];
    if (!prev || prev.length === 0) return next.slice();
    const byId = new Map(prev.map(r => [r.id, r]));
    return next.map(n => ({ ...byId.get(n.id), ...n }));
  };

  const handleRiskTasksChange = React.useCallback((tasks?: IRiskAssessmentResult) => {
    if (!tasks) {
      setRiskAssessmentsTasks(undefined);
      setRiskAssessmentReferenceNumber('');
      setOverAllRiskAssessment('');
      setDetailedRiskAssessment(false);
      return;
    }

    setRiskAssessmentReferenceNumber(tasks?.l2Ref || '');
    setOverAllRiskAssessment(tasks?.overallRisk || '');
    setDetailedRiskAssessment(!!tasks?.l2Required);
    setRiskAssessmentsTasks(prev => mergeRiskRows(prev, tasks?.rows || []));
  }, []);

  // Minimal payload builder (adjust to your save schema)
  const buildPayload = React.useCallback(() => {
    return {
      reference: _coralReferenceNumber,
      assetId: _assetId,
      assetCategoryId: _selectedAssetCategory,
      assetDetailsId: _selectedAssetDetails,
      company: _selectedCompany,
      projectTitle: _projectTitle,
      permitTypes: _selectedPermitTypeList?.map(x => x.id),
      permitRows: _permitPayload,
      hacWorkAreaId: _selectedHacWorkAreaId,
      workHazardIds: Array.from(_selectedWorkHazardIds || []),
      workHazardsOtherText: _workHazardsOtherText,
      workTaskLists: _riskAssessmentsTasks || [],
      overallRiskAssessment: _overAllRiskAssessment || '',
      detailedRiskAssessment: _detailedRiskAssessment || '',
      detailedRiskAssessmentRef: _riskAssessmentReferenceNumber || '',
      precautionsIds: Array.from(_selectedPrecautionIds || []),
      precautionsOtherText: _precautionsOtherText,
      protectiveEquipmentIds: Array.from(_selectedProtectiveEquipmentIds || []),
      protectiveEquipmentsOtherText: _protectiveEquipmentsOtherText,
      gasTestRequired: _gasTestValue,
      gasTestResult: _gasTestResult,
      fireWatchNeeded: _fireWatchValue,
      fireWatchAssigned: _fireWatchAssigned,
      attachmentsProvided: _attachmentsValue,
      attachmentsDetails: _attachmentsResult,
      machineryIds: _selectedMachineryIds || [],
      personnelIds: _selectedPersonnelIds || [],
      originator: _PermitOriginator?.[0]?.secondaryText || '',
    };
  }, [
    _coralReferenceNumber, _assetId, _selectedAssetCategory, _selectedAssetDetails, _projectTitle,
    _selectedPermitTypeList, _permitPayload, _selectedHacWorkAreaId,
    _selectedWorkHazardIds, _selectedPrecautionIds, _selectedProtectiveEquipmentIds,
    _gasTestValue, _gasTestResult, _fireWatchValue, _fireWatchAssigned, _protectiveEquipmentsOtherText, _precautionsOtherText,
    _attachmentsValue, _attachmentsResult, _selectedMachineryIds, _selectedPersonnelIds, permitOriginatorEmail,
    _workHazardsOtherText, _riskAssessmentsTasks, _riskAssessmentReferenceNumber, _overAllRiskAssessment, _detailedRiskAssessment
  ]);

  const validateBeforeSubmit = React.useCallback((mode: 'save' | 'submit' | 'approve'): string | undefined => {
    const missing: string[] = [];
    const payload = buildPayload();

    if (!payload.originator.trim()) {
      missing.push('Permit Originator');
      return `Please fill in the required fields: ${missing.join(', ')}.`;
    };

    if (isPermitOriginator && (mode === 'submit')) {
      if (!payload?.assetId?.trim()) missing.push('Asset ID');
      if (!payload.assetCategoryId?.toString().trim()) missing.push('Asset Category');
      if (!payload.assetDetailsId?.toString().trim()) missing.push('Asset Details');
      if (!payload.projectTitle?.trim()) missing.push('Project Title');
      if (!payload.company?.id?.toString().trim()) missing.push('Company');
      if (!payload.permitTypes || payload.permitTypes.length === 0) missing.push('At least one Permit Type');
      if (!payload.permitRows || payload.permitRows.length === 0) {
        missing.push('At least one Permit Row in Permit Schedule');
      } else {
        const selectedNewPermitRows = payload.permitRows.filter(r => r.isChecked && r.type === 'new');

        if (selectedNewPermitRows.length >= 1) {
          const newRowDateIso = selectedNewPermitRows[0].date;

          if (!newRowDateIso) {
            missing.push('New Permit Row Date');
          } else {
            const permitDate = new Date(newRowDateIso);
            if (isNaN(permitDate.getTime())) {
              missing.push('New Permit Row Date (invalid)');
            } else {
              const now = new Date();
              const ms24h = 24 * 60 * 60 * 1000;
              const isAtLeast24hAfterNow = (permitDate.getTime() - now.getTime()) >= ms24h;

              if (!isAtLeast24hAfterNow) {
                missing.push('New Permit Row Date must be at least 24 hours after the current submission date/time.');
              }
            }
          }
        }
      }

      if (!payload.hacWorkAreaId?.toString().trim()) missing.push('HAC Work Area');

      // Tasks required when 3+ hazards: list rows missing a task
      const hazardsCount = Array.isArray(payload.workHazardIds) ? payload.workHazardIds.length : 0;
      if (hazardsCount >= 3) {
        const rows = Array.isArray(payload.workTaskLists) ? payload.workTaskLists : [];
        if (rows.length === 0) {
          missing.push('At least one Task / Job Description');
        } else {
          const missingTaskRows = rows
            .map((row, idx) => ({ idx, hasTask: !!String(row?.task || '').trim() }))
            .filter(x => !x.hasTask)
            .map(x => x.idx + 1); // 1-based
          if (missingTaskRows.length) {
            missing.push(`Task / Job Description missing for row(s): ${missingTaskRows.join(', ')}`);
          }
        }
      }

      const otherHazard = (ptwFormStructure?.workHazardosList || [])
        .find(h => (h.title || '').toLowerCase().includes('other'));
      const othersSelected = !!otherHazard && payload.workHazardIds?.includes(Number(otherHazard.id));
      if (othersSelected && !String(payload.workHazardsOtherText || '').trim()) {
        missing.push('Work Hazard "Others" details');
      }

      // Ensure at least one Precaution selected
      if (!payload.precautionsIds || payload.precautionsIds.length === 0) {
        missing.push('At least one Precaution');
      }

      // NEW: Attachments validation
      // if (!payload.attachmentsProvided || String(payload.attachmentsProvided).trim() === '') {
      //   missing.push('Attachment(s) provided');
      // }
      // const isAttachmentYes = String(payload.attachmentsProvided || '').toLowerCase() === 'yes';
      // if (isAttachmentYes && !String(payload.attachmentsDetails || '').trim()) {
      //   missing.push('Attachment(s) details');
      // }

      // NEW: Ensure at least one Protective & Safety Equipment selected
      if (!payload.protectiveEquipmentIds || payload.protectiveEquipmentIds.length === 0) {
        missing.push('At least one Protective & Safety Equipment');
      }

      // NEW: Ensure at least one Machinery/Tool selected
      if (!payload.machineryIds || payload.machineryIds.length === 0) {
        missing.push('At least one Machinery/Tool');
      }

      // NEW: Ensure at least one Personnel Involved selected
      if (!payload.personnelIds || payload.personnelIds.length === 0) {
        missing.push('At least one Personnel Involved');
      }

      if (missing.length) {
        return `Please fill in the required fields: ${missing.join(', ')}.`;
      }
    }

    // TODO: Check if is PermitAuthority and mode is approve
    // if (isPermitAuthority && (mode === 'approve')) {
    //   // Add any approval-specific validations here
    // }

    // TODO: Check if is PermistIssuer and mode is approve , validate for 
    // if (isPermitIssuer && (mode === 'approve')) {
    //   // Add any approval-specific validations here

    // Tasks required when 3+ hazards: list rows missing a task
    // const hazardsCount = Array.isArray(payload.workHazardIds) ? payload.workHazardIds.length : 0;
    // if (hazardsCount >= 3) {
    //   const rows = Array.isArray(payload.workTaskLists) ? payload.workTaskLists : [];
    //   if (rows.length >= 1) {
    //     // Initial Risk required per row
    //     const missingInitialRiskRows = rows
    //       .map((row, idx) => ({ idx, ok: !!String((row as any)?.initialRisk || '').trim() }))
    //       .filter(x => !x.ok)
    //       .map(x => x.idx + 1);
    //     if (missingInitialRiskRows.length) {
    //       missing.push(`Initial Risk missing for row(s): ${missingInitialRiskRows.join(', ')}`);
    //     }

    //     // Residual Risk required per row
    //     const missingResidualRiskRows = rows
    //       .map((row, idx) => ({ idx, ok: !!String((row as any)?.residualRisk || '').trim() }))
    //       .filter(x => !x.ok)
    //       .map(x => x.idx + 1);
    //     if (missingResidualRiskRows.length) {
    //       missing.push(`Residual Risk missing for row(s): ${missingResidualRiskRows.join(', ')}`);
    //     }
    //   }

    //   // Overall Risk Assessment required
    //   if (!String(payload.overallRiskAssessment || '').trim()) {
    //     missing.push('Overall Risk Assessment');
    //   }

    //   // Detailed L2: if checked, require Ref Number
    //   const l2Required = !!payload.detailedRiskAssessment;
    //   if (l2Required && !String(payload.detailedRiskAssessmentRef || '').trim()) {
    //     missing.push('Risk Assessment Ref Number (Detailed L2)');
    //   }
    // }

    // // Gas Test: if Yes, require result
    // if (String(payload.gasTestRequired || '').toLowerCase() === 'yes' &&
    //   !String(payload.gasTestResult || '').trim()) {
    //   missing.push('Gas Test Result');
    // }

    // // Fire Watch: if Yes, require assigned
    // if (String(payload.fireWatchNeeded || '').toLowerCase() === 'yes' &&
    //   !String(payload.fireWatchAssigned || '').trim()) {
    //   missing.push('Firewatch Assigned');
    // }

    // // Attachments
    // if (!payload.attachmentsProvided || String(payload.attachmentsProvided).trim() === '') {
    //   missing.push('Attachment(s) provided');
    // }
    // const isAttachmentYes = String(payload.attachmentsProvided || '').toLowerCase() === 'yes';
    // if (isAttachmentYes && !String(payload.attachmentsDetails || '').trim()) {
    //   missing.push('Attachment(s) details');
    // }
    // }

    return undefined;
  }, [buildPayload, ptwFormStructure?.workHazardosList]);

  const approveForm = React.useCallback(async (mode: 'approve') => {

  }, []);

  const submitForm = React.useCallback(async (mode: 'save' | 'submit'): Promise<boolean> => {
    if (!isPermitOriginator) {
      showBanner('Only the Permit Originator can save or submit this form.',
        { autoHideMs: 5000, fade: true, kind: 'error' });
      return false;
    } else {
      hideBanner();
    }

    setIsBusy(true);
    if (mode === 'submit') {
      setPoStatus('Approved');
    }
    else {
      setPoStatus('Pending');
    }

    setBusyLabel(mode === 'save' ? 'Saving form…' : 'Submitting form…');
    try {
      // const payload = buildPayload();
      const validationError = validateBeforeSubmit(mode);
      if (validationError) {
        showBanner(validationError);
        return false;
      } else {
        const editFormId = props.formId ? Number(props.formId) : undefined;
        const formStatusRecord = JSON.parse(localStorage.getItem('FormStatusRecord') || '{}');

        if (editFormId === undefined) {
          const savedId = await _createPTWForm(mode);
          if (savedId) {
            // await new Promise(res => setTimeout(res, 1000));
            if (mode === 'save') {
              showBanner('Form saved successfully.', { autoHideMs: 5000, fade: true, kind: 'success' });
            }
            else if (mode === 'submit') {
              showBanner('Form submitted successfully.', { autoHideMs: 5000, fade: true, kind: 'success' });
            }
          }
        }

        if (editFormId && editFormId > 0 && formStatusRecord.value.toLowerCase() === 'saved') {
          const updated = await _updatePTWForm(editFormId, mode);
          if (updated) {
            if (mode === 'save') {
              showBanner('Form updated successfully.', { autoHideMs: 5000, fade: true, kind: 'success' });
            }
            else if (mode === 'submit') {
              showBanner('Form submitted successfully.', { autoHideMs: 5000, fade: true, kind: 'success' });
            }
          }
        }
      }
      goBackToHost();
      return true;
    } catch (e) {
      showBanner('An error occurred while processing the form.', { autoHideMs: 5000, fade: true, kind: 'error' });
      return false;
    } finally {
      setIsBusy(false);
    }
  }, [isPermitOriginator, buildPayload]);

  // Create parent PTWForm item and return its Id
  const _createPTWForm = React.useCallback(async (mode: 'save' | 'submit'): Promise<number> => {
    const payload = buildPayload();
    if (!payload) throw new Error('Form payload is not available');

    const spOps = spCrudRef.current ?? new SPCrudOperations((props.context as any).spHttpClient, webUrl, '', '');
    const originatorId = await spOps.ensureUserId(payload.originator);

    const body: any = {
      PermitOriginatorId: originatorId ?? null,
      Title: 'PTW Form' + (originatorId ? ` - ${payload.originator}` : ''),
      AssetID: payload.assetId ?? null,
      AssetCategoryId: payload.assetCategoryId ? Number(payload.assetCategoryId) : null,
      AssetDetailsId: payload.assetDetailsId ? Number(payload.assetDetailsId) : null,
      CompanyRecordId: payload.company?.id ? Number(payload.company.id) : null,
      ProjectTitle: payload.projectTitle ?? null,
      HACClassificationWorkAreaId: payload.hacWorkAreaId ? Number(payload.hacWorkAreaId) : null,
      WorkHazardsOthers: payload.workHazardsOtherText ?? null,
      ProtectiveSafetyEquipmentsOthers: payload.protectiveEquipmentsOtherText ?? null,
      PrecautionsOthers: payload.precautionsOtherText ?? null,
      FormStatusRecord: mode === 'submit' ? 'Submitted' : 'Saved',
      WorkflowStatus: mode === 'submit' ? 'New' : '',
      AttachmentsProvided: payload.attachmentsProvided.toLowerCase() === "yes" ? true : false,
      AttachmentsProvidedDetails: payload.attachmentsDetails ?? '',
    };
    debugger;
    // OData v4 style for multi-lookup fields: array + @odata.type
    if (payload.permitTypes?.length) {
      body['WorkCategoryId@odata.type'] = 'Collection(Edm.Int32)';
      body['WorkCategoryId'] = payload.permitTypes.map(Number);
    }
    if (payload.workHazardIds?.length) {
      body['WorkHazardsId@odata.type'] = 'Collection(Edm.Int32)';
      body['WorkHazardsId'] = payload.workHazardIds.map(Number);
    }
    if (payload.precautionsIds?.length) {
      body['PrecuationsId@odata.type'] = 'Collection(Edm.Int32)';
      body['PrecuationsId'] = payload.precautionsIds.map(Number);
    }
    if (payload.protectiveEquipmentIds?.length) {
      body['ProtectiveSafetyEquipmentsId@odata.type'] = 'Collection(Edm.Int32)';
      body['ProtectiveSafetyEquipmentsId'] = payload.protectiveEquipmentIds.map(Number);
    }
    if (payload.machineryIds?.length) {
      body['MachineryInvolvedId@odata.type'] = 'Collection(Edm.Int32)';
      body['MachineryInvolvedId'] = payload.machineryIds.map(Number);
    }
    if (payload.personnelIds?.length) {
      body['PersonnelInvolvedId'] = _selectedPersonnelIds?.map(Number);
    }

    spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form', '');
    const newId = await spCrudRef.current._insertItem(body);
    if (!newId) throw new Error('Failed to create PTW Form');

    try {
      const coralReferenceNumber = await spHelpers.assignCoralReferenceNumber(props.context.spHttpClient,
        webUrl, 'PTW_Form', { Id: Number(newId) }, payload.company?.title);
      if (!coralReferenceNumber) throw new Error('Failed to generate Coral Reference Number. Please try again later.');

      setCoralReferenceNumber(coralReferenceNumber);

      if (payload.permitRows?.length && payload.permitRows.some(r => r.isChecked)) {
        const _createdPermits = await _createPTWWorkPermits(Number(newId), payload.permitRows);

        if (!_createdPermits?.length) {
          throw new Error('Failed to create PTW Work Permits');
        }
      }
      debugger;
      if (mode === 'submit') {
        const _createdWorkflow = await _createPTWFormApprovalWorkflow(Number(newId), originatorId);

        if (!_createdWorkflow) {
          throw new Error('Failed to create PTW Form Approval Workflow');
        }
      }

      if (payload.workTaskLists?.length) {
        const _createdTask = await _createPTWTasksJobsDescriptions(Number(newId), payload.workTaskLists);

        if (!_createdTask?.length) {
          throw new Error('Failed to create PTW Tasks and Job Descriptions');
        }
      }

    } catch (e) {
      console.warn('Failed to create PTW Form:', e);
    }

    return newId as number;
  }, [buildPayload, props.context.spHttpClient]);

  const _createPTWWorkPermits = React.useCallback(async (parentId: number, permitRows: IPermitScheduleRow[]) => {
    const opsDelete = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form_Work_Permits', '');
    await Promise.all(permitRows.map(async (item) => {
      await opsDelete._deleteLookUPItems(Number(parentId), "PTWForm");
    }));

    const requiredItems = permitRows.filter((row) => row.isChecked);
    if (requiredItems.length === 0) return [];
    const ops = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form_Work_Permits', '');
    const posts = requiredItems.map((item, index) => {
      const body = {
        PTWFormId: parentId,
        PermitType: item.type ?? null,
        PermitDate: item.date,
        PermitStartTime: spHelpers.combineDateAndTime(item.date.toString(), item.startTime),
        PermitEndTime: spHelpers.combineDateAndTime(item.date.toString(), item.endTime),
        RecordOrder: index + 1,
        Title: item.type === 'new' ? 'New Permit' : 'Renewal Permit'
      };

      const data = ops._insertItem(body);

      if (!data) throw new Error('Failed to create PTW Work Permits.');
      return typeof data === 'number' ? data : (data);
    });
    const results = await Promise.all(posts);
    return results;
  }, [props.context.spHttpClient, spHelpers]);

  const _createPTWTasksJobsDescriptions = React.useCallback(async (parentId: number, workTaskLists: IRiskTaskRow[]) => {
    const opsDelete = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form_Job_Descriptions', '');
    await Promise.all(workTaskLists.map(async (item) => {
      await opsDelete._deleteLookUPItems(Number(parentId), "PTWForm");
    }));

    const requiredItems = workTaskLists.filter((row) => row.disabledFields !== true);
    if (requiredItems.length === 0) return [];
    const ops = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form_Job_Descriptions', '');
    const posts = requiredItems.map((item, index) => {
      const body: any = {
        PTWFormId: parentId,
        JobDescription: item.task ?? null,
        InitialRisk: item.initialRisk ?? null,
        ResidualRisk: item.residualRisk ?? null,
        Title: item.task,
        OrderRecord: index + 1
      };

      if (item.safeguardIds?.length) {
        body['SafeguardsId@odata.type'] = 'Collection(Edm.Int32)';
        body['SafeguardsId'] = item.safeguardIds.map(Number);
      } else {
        body['SafeguardsId'] = { results: [] };
      }

      const data = ops._insertItem(body);

      if (!data) throw new Error('Failed to create PTW Tasks Descriptions.');
      return typeof data === 'number' ? data : (data);
    });
    const results = await Promise.all(posts);
    return results;
  }, [props.context.spHttpClient]);

  const _createPTWFormApprovalWorkflow = React.useCallback(async (parentId: number, originatorId: number | undefined) => {

    if (originatorId === undefined) return;
    const ops = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form_Approval_Workflow', '');
    let _paStatusForPOPA = '';
    debugger;
    if (originatorId === _paPicker?.length ? await ops.ensureUserId(_paPicker[0].secondaryText) : null) {
      _paStatusForPOPA = _paStatus === 'Approved' ? _paStatus : 'Approved';
    } else {
      _paStatusForPOPA = _paStatus || 'Pending';
    }

    try {
      let body: any = {
        PTWFormId: parentId,
        StatusRecord: 'New',
        IsFinalApprover: false,

        POApproverId: originatorId ?? null,
        POApprovalDate: _poDate || null,
        POStatus: _poStatus || 'Approved',

        PAApproverId: _paPicker?.length ? await ops.ensureUserId(_paPicker[0].secondaryText) : null,
        PAStatus: _paStatusForPOPA,
        PAApprovalDate: _paDate || null,

        PIApproverId: _piPicker?.length ? await ops.ensureUserId(_piPicker[0].secondaryText) : null,
        PIStatus: _piStatus || 'Pending',
        PIApprovalDate: _piDate || null,
      };

      if (_overAllRiskAssessment && _overAllRiskAssessment.toLowerCase() === 'high') {
        body.AssetDirectorApproverId = _assetDirPicker?.length ? await ops.ensureUserId(_assetDirPicker[0].secondaryText) : null;
        body.AssetDirectorStatus = _assetDirStatus || 'Pending';
        body.AssetDirectorApprovalDate = _assetDirDate || null;

        body.HSEDirectorApproverId = _hseDirPicker?.length ? await ops.ensureUserId(_hseDirPicker[0].secondaryText) : null;
        body.HSEDirectorStatus = _hseDirStatus || 'Pending';
        body.HSEDirectorApprovalDate = _hseDirDate || null;
      }

      const data = ops._insertItem(body);
      if (!data) throw new Error('Failed to create PTW Form Approval Workflow.');
      return typeof data === 'number' ? data : (data);;
    }
    catch (e) {
      console.warn('Failed to create PTW Form Approval Workflow', e);
    }
  }, [props.context.spHttpClient]);

  const _updatePTWForm = React.useCallback(async (id: number, mode: 'save' | 'submit'): Promise<boolean> => {
    const payload = buildPayload();
    if (!payload) throw new Error('Form payload is not available');

    const spOps = spCrudRef.current ?? new SPCrudOperations((props.context as any).spHttpClient, webUrl, '', '');
    const originatorId = await spOps.ensureUserId(payload.originator);

    const body: any = {
      PermitOriginatorId: originatorId ?? null,
      AssetID: payload.assetId ?? null,
      AssetCategoryId: payload.assetCategoryId ? Number(payload.assetCategoryId) : null,
      AssetDetailsId: payload.assetDetailsId ? Number(payload.assetDetailsId) : null,
      CompanyRecordId: payload.company?.id ? Number(payload.company.id) : null,
      ProjectTitle: payload.projectTitle ?? null,
      HACClassificationWorkAreaId: payload.hacWorkAreaId ? Number(payload.hacWorkAreaId) : null,
      WorkHazardsOthers: payload.workHazardsOtherText ?? null,
      ProtectiveSafetyEquipmentsOthers: payload.protectiveEquipmentsOtherText ?? null,
      PrecautionsOthers: payload.precautionsOtherText ?? null,
      FormStatusRecord: mode === 'submit' ? 'Submitted' : 'Saved',
      WorkflowStatus: mode === 'submit' ? 'New' : '',
    };

    if (payload.permitTypes?.length) {
      body['WorkCategoryId@odata.type'] = 'Collection(Edm.Int32)';
      body['WorkCategoryId'] = payload.permitTypes.map(Number);
    } else {
      body['WorkCategoryId'] = { results: [] };
    }
    if (payload.workHazardIds?.length) {
      body['WorkHazardsId@odata.type'] = 'Collection(Edm.Int32)';
      body['WorkHazardsId'] = payload.workHazardIds.map(Number);
    } else {
      body['WorkHazardsId'] = { results: [] };
    }
    if (payload.precautionsIds?.length) {
      body['PrecuationsId@odata.type'] = 'Collection(Edm.Int32)';
      body['PrecuationsId'] = payload.precautionsIds.map(Number);
    } else {
      body['PrecuationsId'] = { results: [] };
    }
    if (payload.protectiveEquipmentIds?.length) {
      body['ProtectiveSafetyEquipmentsId@odata.type'] = 'Collection(Edm.Int32)';
      body['ProtectiveSafetyEquipmentsId'] = payload.protectiveEquipmentIds.map(Number);
    } else {
      body['ProtectiveSafetyEquipmentsId'] = { results: [] };
    }
    if (payload.machineryIds?.length) {
      body['MachineryInvolvedId@odata.type'] = 'Collection(Edm.Int32)';
      body['MachineryInvolvedId'] = payload.machineryIds.map(Number);
    } else {
      body['MachineryInvolvedId'] = { results: [] };
    }
    if (payload.personnelIds?.length) {
      body['PersonnelInvolvedId@odata.type'] = 'Collection(Edm.Int32)';
      body['PersonnelInvolvedId'] = payload.personnelIds.map(Number);
    } else {
      body['PersonnelInvolvedId'] = { results: [] };
    }

    spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form', '');
    const response = await spCrudRef.current._updateItem(String(id), body);
    if (!response.ok) {
      showBanner('Failed to update PTW Form.', { autoHideMs: 5000, fade: true, kind: 'error' });
      return false;
    }

    if (payload.permitRows?.length && payload.permitRows.some(r => r.isChecked)) {
      const _createdPermits = await _createPTWWorkPermits(Number(id), payload.permitRows);

      if (!_createdPermits?.length) {
        throw new Error('Failed to create PTW Work Permits');
      }
    }

    if (mode === 'submit') {
      const _createdWorkflow = await _createPTWFormApprovalWorkflow(Number(id), originatorId);

      if (!_createdWorkflow) {
        throw new Error('Failed to create PTW Form Approval Workflow');
      }
    }

    if (payload.workTaskLists?.length) {
      const _createdTask = await _createPTWTasksJobsDescriptions(Number(id), payload.workTaskLists);

      if (!_createdTask?.length) {
        throw new Error('Failed to create PTW Tasks and Job Descriptions');
      }
    }

    return true;
  }, [buildPayload, props.context.spHttpClient]);

  // ---------------------------
  // Render
  // ---------------------------

  const [prefilledFormId, setPrefilledFormId] = React.useState<number | undefined>(undefined);

  // Prefill when editing an existing form
  React.useEffect(() => {
    const formId = props.formId;
    if (!formId || prefilledFormId === formId) return;

    // Wait until base items are loaded and itemRows initialized
    if (loading) return;

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

        const ptwFirstSelect = `?$select=Id,CoralReferenceNumber,AssetID,ProjectTitle,Created,FormStatusRecord,IsDetailedRiskAssessmentRequired,RiskAssessmentRefNumber,WorkHazardsOthers,` +
          `OverallRiskAssessment,GasTestRequired,GasTestResult,WorkflowStatus,` +
          `PermitOriginator/Id,PermitOriginator/Title,PermitOriginator/EMail,` +
          `AssetCategory/Id,AssetCategory/Title,` +
          `AssetDetails/Id,AssetDetails/Title,` +
          `CompanyRecord/Id,CompanyRecord/Title,` +
          `WorkCategory/Id,WorkCategory/Title,WorkCategory/OrderRecord,WorkCategory/RenewalValidity,` +
          `HACClassificationWorkArea/Id,HACClassificationWorkArea/Title,` +
          `WorkHazards/Id,WorkHazards/Title` +
          `&$expand=PermitOriginator,AssetCategory,AssetDetails,CompanyRecord,WorkCategory,` +
          `HACClassificationWorkArea,WorkHazards` +
          `&$filter=Id eq ${formId}`;

        const ptwSecondSelect = `?$select=Id,FireWatchNeeded,AttachmentsProvided,AttachmentsProvidedDetails,ToolboxTalk,` +
          `ToolboxTalkHSEReference,ToolBoxTalkDate,ProtectiveSafetyEquipmentsOthers,PrecautionsOthers,` +
          `Precuations/Id,Precuations/Title,` +
          `ProtectiveSafetyEquipments/Id,ProtectiveSafetyEquipments/Title,` +
          `MachineryInvolved/Id,MachineryInvolved/Title,` +
          `FireWatchAssigned/Id,FireWatchAssigned/FullName,` +
          `PersonnelInvolved/Id,PersonnelInvolved/Title,` +
          `ToolboxConductedBy/Id,ToolboxConductedBy/Title,ToolboxConductedBy/EMail` +
          `&$expand=Precuations,ProtectiveSafetyEquipments,MachineryInvolved,FireWatchAssigned,` +
          `PersonnelInvolved,ToolboxConductedBy` +
          `&$filter=Id eq ${formId}`;

        const ptwWorkPermits = `?$select=Id,PermitType,PermitDate,PermitStartTime,PermitEndTime,RecordOrder,StatusRecord,` +
          `PTWForm/Id,PTWForm/CoralReferenceNumber` +
          `&$expand=PTWForm` +
          `&$filter=PTWForm/Id eq ${formId}`;

        const ptwTaskDescription = `?$select=Id,JobDescription,InitialRisk,ResidualRisk,OrderRecord,OtherSafeguards,` +
          `PTWForm/Id,PTWForm/CoralReferenceNumber,` +
          `Safeguards/Id,Safeguards/Title` +
          `&$expand=PTWForm,Safeguards` +
          `&$filter=PTWForm/Id eq '${formId}'`;

        const formCrudFirstSelect = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form', ptwFirstSelect);
        const headerItemsFirstSelect = await formCrudFirstSelect._getItemsWithQuery();
        const headerFirstSelect = Array.isArray(headerItemsFirstSelect) ? headerItemsFirstSelect[0] : undefined;

        const formCrudSecondSelect = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form', ptwSecondSelect);
        const headerItemsSecondSelect = await formCrudSecondSelect._getItemsWithQuery();
        const headerSecondSelect = Array.isArray(headerItemsSecondSelect) ? headerItemsSecondSelect[0] : undefined;

        const formCrudWorkPermits = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form_Work_Permits', ptwWorkPermits);
        const headerItemsWorkPermits = await formCrudWorkPermits._getItemsWithQuery();
        const headerWorkPermits = Array.isArray(headerItemsWorkPermits) ? headerItemsWorkPermits : undefined;

        const formCrudTaskDescription = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form_Job_Descriptions', ptwTaskDescription);
        const headerItemsTaskDescription = await formCrudTaskDescription._getItemsWithQuery();
        const headerTaskDescription = Array.isArray(headerItemsTaskDescription) ? headerItemsTaskDescription : undefined;

        if (headerFirstSelect && !cancelled && headerSecondSelect) {
          // Top-level fields prefill
          if (headerFirstSelect?.FormStatusRecord) {
            setMode(headerFirstSelect?.FormStatusRecord.toLowerCase());
          }

          const permitOriginator = toPersona({ Id: headerFirstSelect?.PermitOriginator?.Id, FullName: headerFirstSelect?.PermitOriginator?.Title, EMail: headerFirstSelect?.PermitOriginator?.EMail });
          // setPermitOriginator([{ text: permitOriginator?.Title || '', secondaryText: permitOriginator.email || '', id: permitOriginator.id }]);
          setPermitOriginator(permitOriginator ? [permitOriginator] : []);
          setCoralReferenceNumber(headerFirstSelect?.CoralReferenceNumber || '');
          setAssetId(headerFirstSelect?.AssetID);
          setSelectedCompany(headerFirstSelect?.CompanyRecord ? { id: headerFirstSelect.CompanyRecord.Id, title: headerFirstSelect.CompanyRecord.Title || '', orderRecord: headerFirstSelect.CompanyRecord.OrderRecord || 0 } : undefined);
          setProjectTitle(headerFirstSelect?.ProjectTitle || '');
          setSelectedAssetCategory(headerFirstSelect?.AssetCategory ? Number(headerFirstSelect.AssetCategory.Id) : undefined);
          setSelectedAssetDetails(headerFirstSelect?.AssetDetails ? Number(headerFirstSelect.AssetDetails.Id) : undefined);
          setSelectedHacWorkAreaId(headerFirstSelect?.HACClassificationWorkArea?.Id != null ? Number(headerFirstSelect.HACClassificationWorkArea.Id) : undefined);
          setSelectedHacWorkAreaId(headerFirstSelect?.HACClassificationWorkArea?.Id != null ? Number(headerFirstSelect.HACClassificationWorkArea.Id) : undefined);
          setSelectedWorkHazardIds(new Set(Array.isArray(headerFirstSelect.WorkHazards) ? headerFirstSelect.WorkHazards.map((wh: any) => Number(wh.Id)) : []));
          setWorkHazardsOtherText(headerFirstSelect?.WorkHazardsOthers || '');
          setOverAllRiskAssessment(headerFirstSelect?.OverallRiskAssessment || '');
          setDetailedRiskAssessment(!!headerFirstSelect?.IsDetailedRiskAssessmentRequired);
          setRiskAssessmentReferenceNumber(headerFirstSelect?.RiskAssessmentRefNumber || '');
          setSelectedPrecautionIds(new Set(Array.isArray(headerSecondSelect.Precuations) ? headerSecondSelect.Precuations.map((pc: any) => Number(pc.Id)) : []));
          setPrecautionsOtherText(headerSecondSelect?.PrecautionsOthers || '');
          setGasTestValue(headerFirstSelect?.GasTestRequired || '');
          setGasTestResult(headerFirstSelect?.GasTestResult || '');
          setFireWatchValue(headerSecondSelect?.FireWatchNeeded || '');
          setFireWatchAssigned(headerSecondSelect?.FireWatchAssigned ? String(headerSecondSelect.FireWatchAssigned.FullName) : '');
          setAttachmentsValue(headerSecondSelect?.AttachmentsProvided ? (headerSecondSelect.AttachmentsProvided ? 'Yes' : 'No') : '');
          setAttachmentsResult(headerSecondSelect?.AttachmentsProvidedDetails || '');

          if (headerSecondSelect.ProtectiveSafetyEquipments.length > 0) {
            setSelectedProtectiveEquipmentIds(headerSecondSelect.ProtectiveSafetyEquipments.map(
              (item: any) => {
                if (item.Title.toLowerCase().includes('other')) {
                  setProtectiveEquipmentsOtherText(headerSecondSelect?.ProtectiveSafetyEquipmentsOthers || '');
                }
                return Number(item.Id);
              }));
          }
          if (headerSecondSelect?.MachineryInvolved.length > 0) {
            setSelectedMachineryIds(headerSecondSelect?.MachineryInvolved.map((item: any) => Number(item.Id)) || []);
          }
          if (headerSecondSelect?.PersonnelInvolved.length > 0) {
            setSelectedPersonnelIds(headerSecondSelect?.PersonnelInvolved.map((item: any) => Number(item.Id)) || []);
          }

          setToolboxTalk(headerSecondSelect?.ToolboxTalk || '');
          setToolboxHSEReference(headerSecondSelect?.ToolboxTalkHSEReference || '');
          setToolboxTalkDate(headerSecondSelect?.ToolBoxTalkDate ? spHelpers.toISO(headerSecondSelect?.ToolBoxTalkDate) : undefined);
          const toolboxConductedBy = toPersona({ Id: headerSecondSelect?.ToolboxConductedBy?.Id, Title: headerSecondSelect?.ToolboxConductedBy?.Title, EMail: headerSecondSelect?.ToolboxConductedBy?.EMail });
          setToolboxConductedBy(toolboxConductedBy ? [toolboxConductedBy] : undefined);

          if (headerTaskDescription && headerTaskDescription.length > 0) {
            const tasksList: IRiskTaskRow[] = [];
            headerTaskDescription.forEach((item: any, index) => {
              if (item) {
                tasksList.push({
                  id: item.Id,
                  task: item.JobDescription || '',
                  initialRisk: item.InitialRisk || '',
                  residualRisk: item.ResidualRisk || '',
                  safeguardsNote: item.OtherSafeguards || '',
                  disabledFields: false,
                  orderRecord: item.OrderRecord || 0,
                  safeguardIds: Array.isArray(item.Safeguards) ? item.Safeguards
                    .map((sg: any) => Number(sg.Id)) : [],
                })
              }
            });
            setRiskAssessmentsTasks(tasksList.sort((a, b) => a.orderRecord - b.orderRecord));
          } else {
            setRiskAssessmentsTasks([]);
          }

          const _workCategories: IWorkCategory[] = [];
          if (headerFirstSelect.WorkCategory !== undefined && headerFirstSelect.WorkCategory !== null && Array.isArray(headerFirstSelect.WorkCategory)) {
            headerFirstSelect.WorkCategory.forEach((item: any) => {
              if (item) {
                _workCategories.push({
                  id: item.Id,
                  title: item.Title,
                  orderRecord: item.OrderRecord || 0,
                  renewalValidity: item.RenewalValidity || 0,
                  isChecked: true,
                });
              }
            });
          }
          setSelectedPermitTypeList(_workCategories);
          setWorkPermitRequired(_workCategories.length > 0);
          if (headerWorkPermits && headerWorkPermits.length > 0) {
            const permitsList: IPermitScheduleRow[] = [];
            headerWorkPermits.sort((a: any, b: any) => a.OrderRecord - b.OrderRecord).forEach((item: any, index) => {
              if (item) {
                const startTime = item?.PermitStartTime ? spHelpers.toHHmm(item.PermitStartTime) : '';
                const endTime = item?.PermitEndTime ? spHelpers.toHHmm(item.PermitEndTime) : '';
                permitsList.push({
                  id: String(item.Id),
                  type: item.PermitType,
                  date: item.PermitDate ? new Date(item.PermitDate).toISOString() : '',
                  startTime: startTime,
                  endTime: endTime,
                  orderRecord: item.RecordOrder,
                  isChecked: true,
                });
              }
            });
            setPermitPayload(permitsList.sort((a, b) => a.orderRecord - b.orderRecord));
          } else {
            setPermitPayload([]);
          }
        }

        await _getPTWWorkflow();
        if (!cancelled) setPrefilledFormId(formId);
      } catch (e) {
        showBanner('An error occurred while loading the form for editing.', { autoHideMs: 5000, fade: true, kind: 'error' });
      }
    };

    load();

    return () => { cancelled = true; };
  }, [props.formId, prefilledFormId, loading, props.context, spHelpers]);

  if (loading) {
    return (
      <div className={styles.loadingContainer}>
        <Spinner label={"Preparing PTW form.. "} size={SpinnerSize.large} />
      </div>
    );
  }

  // const delayResults = false;
  const logoUrl = `${props.context.pageContext.web.absoluteUrl}/SiteAssets/coral-logo.png`;
  // const logoPTWUrl = `${props.context.pageContext.web.absoluteUrl}/SiteAssets/ptw-logo.png`;
  // const peopleList: IPersonaProps[] = users.map(user => ({ text: user.displayName || '', secondaryText: user.email || '', id: user.id }));
  function onInputChange(input: string): string { const outlookRegEx = /<.*>/g; const emailAddress = outlookRegEx.exec(input); if (emailAddress && emailAddress[0]) return emailAddress[0].substring(1, emailAddress[0].length - 1); return input; }

  return (

    <div style={{ position: 'relative' }} ref={containerRef}>
      <div ref={bannerTopRef} />
      {isBusy && (
        <div
          ref={overlayRef}
          aria-busy="true"
          role="dialog"
          aria-modal="true"
          className="no-pdf"
          data-html2canvas-ignore="true"
          aria-label={busyLabel}
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
          <Spinner label={busyLabel} size={SpinnerSize.large} />
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

      <form id="ptwFormMain">
        {/* Top action bar removed; Save/Submit moved to bottom */}
        <div id="formTitleSection">
          <div className={styles.ptwformHeader} >
            <div>
              <img src={logoUrl} alt="Logo" className={styles.formLogo} />
            </div>
            <div className={styles.ptwFormTitleLogo}>
              {/* <div>
                <img src={logoPTWUrl} alt="PTWLogo" className={styles.ptwformLogo} />
              </div> */}
              <div className={styles.ptwTitles}>
                <span className={styles.formArTitle}>{ptwFormStructure?.coralForm?.arTitle}</span>
                <span className={styles.formTitle}>{ptwFormStructure?.coralForm?.title}</span>
              </div>
            </div>
          </div>
        </div>

        <BannerComponent text={bannerText} kind={bannerOpts?.kind || 'error'}
          autoHideMs={bannerOpts?.autoHideMs} fade={bannerOpts?.fade}
          onDismiss={() => { setBannerText(undefined); setBannerOpts(undefined); }}
        />

        <div id="formHeaderInfo" className={styles.formBody}>
          {/* Administrative Note */}
          <div className={`row mb-1`} id="administrativeNoteDiv">
            <div className={`form-group col-md-6`}>
              <NormalPeoplePicker label={"Permit Originator"} onResolveSuggestions={_onFilterChanged} itemLimit={1}
                className={'ms-PeoplePicker'} key={'permitOriginator'} removeButtonAriaLabel={'Remove'}
                onInputChange={onInputChange} resolveDelay={150}
                styles={peoplePickerBlackStyles}
                selectedItems={_PermitOriginator}
                inputProps={{ placeholder: 'Enter name or email' }}
                pickerSuggestionsProps={suggestionProps}
                disabled={true}
              />
            </div>

            <div className={`form-group col-md-6`}>
              <TextField label="PTW Ref #"
                readOnly
                value={_coralReferenceNumber}
              // styles={{ root: { color: '#000', fontWeight: 500, backgroundColor: '#f4f4f4' } }}
              // onChange={(_, newValue) => setCoralReferenceNumber(newValue || '')}
              />
            </div>
          </div>


          <div className='row' id="permitOriginatorDiv">
            <div className={`form-group col-md-6`}>
              <ComboBox
                label="Company"
                placeholder="Select a company"
                options={ptwFormStructure?.companies?.sort((a, b) => (a.orderRecord || 0) - (b.orderRecord || 0))
                  .map(c => ({ key: c.id, text: c.title || '' })) || []}
                selectedKey={_selectedCompany?.id}
                onChange={(_e, item) => setSelectedCompany(item ? { id: Number(item.key), title: item.text, orderRecord: 0 } : undefined)}
                styles={comboBoxBlackStyles}
                useComboBoxAsMenuWidth={true}
              />
            </div>
            <div className={`form-group col-md-6`}>
              <TextField
                label="Asset ID"
                value={_assetId}
                onChange={(_, newValue) => setAssetId(newValue || '')} />
            </div>
          </div>

          <div className={`row`} id="assetCategoryDetailsDiv">
            <div className={`form-group col-md-6`}>
              <ComboBox
                label="Asset Category"
                placeholder="Select an asset category"
                options={assetCategoriesList?.map(c => ({ key: c.id, text: c.title || '' })) || []}
                selectedKey={_selectedAssetCategory}
                onChange={(_e, ch) => onAssetCategoryChange(_e, ch)}
                styles={comboBoxBlackStyles}
                useComboBoxAsMenuWidth={true}
              />
            </div>
            <div className={`form-group col-md-6`}>
              <ComboBox
                label="Asset Details"
                placeholder="Select asset details"
                options={assetDetailsOptions}
                selectedKey={_selectedAssetDetails}
                onChange={(_e, ch) => onAssetDetailsChange(_e as any, ch as any)}
                disabled={!_selectedAssetCategory}
                styles={comboBoxBlackStyles}
                useComboBoxAsMenuWidth={true}
              />
            </div>
          </div>

          <div className={`row`} id="projectTitleDiv">
            <div className={`form-group col-md-12`}>
              <TextField
                label="Project Title / Description"
                value={_projectTitle}
                onChange={(_, newValue) => setProjectTitle(newValue || '')}
                multiline
                rows={2}
                styles={{
                  fieldGroup: { backgroundColor: '#f6f6f7ff' }
                }}
              />
            </div>
          </div>
        </div>

        <div id="formContentSection">
          <div className='row pb-3' id="permitScheduleSection">
            <PermitSchedule
              workCategories={ptwFormStructure?.workCategories?.sort((a, b) => a.orderRecord - b.orderRecord) || []}
              selectedPermitTypeList={_selectedPermitTypeList.sort((a, b) => a.orderRecord - b.orderRecord)}
              permitRows={_permitPayload}
              onPermitTypeChange={handlePermitTypeChange}
              onPermitRowUpdate={updatePermitRow}
              styles={styles}
            />
          </div>

          {workPermitRequired && (
            <div id="ptwFormsSections">

              <div className="row pb-3" id="hacClassificationWorkAreaSection">
                <div>
                  <Label className={styles.ptwLabel}>HAC Classification of Work Area</Label>
                </div>
                <CheckBoxDistributerOnlyComponent id="hacClassificationWorkAreaComponent"
                  optionList={ptwFormStructure?.hacWorkAreas || []}
                  colSpacing='col-2'
                  onChange={(checked, item) => handleHACChange(checked, item)}
                  selectedIds={_selectedHacWorkAreaId !== undefined ? [_selectedHacWorkAreaId] : []}
                />
              </div>

              <div className="row pb-3" id="workHazardSection" >
                <div>
                  <Label className={styles.ptwLabel}>Work Hazards</Label>
                  <div className="text-center pb-3">
                    <small className="text-muted" style={{ fontStyle: 'italic', fontSize: '0.8rem' }}>
                      if 3 or more working hazards, detailed job description/tasks shall be provided below.
                    </small>
                  </div>
                </div>

                <CheckBoxDistributerComponent id="workHazardsComponent"
                  optionList={ptwFormStructure?.workHazardosList || []}
                  selectedIds={Array.from(_selectedWorkHazardIds)}
                  onChange={(ids) => setSelectedWorkHazardIds(new Set(ids))}
                  othersTextValue={_workHazardsOtherText}
                  onOthersChange={(checked, othersText) => setWorkHazardsOtherText(othersText)}
                />
              </div>

              {_selectedWorkHazardIds.size >= 3 && (
                <div className="row pb-2" id="riskAssessmentListSection">
                  <div className="form-group col-md-12">
                    <RiskAssessmentList
                      initialRiskOptions={ptwFormStructure?.initialRisk || []}
                      residualRiskOptions={ptwFormStructure?.residualRisk || []}
                      safeguards={filteredSafeguards || []}
                      overallRiskOptions={ptwFormStructure?.overallRiskAssessment || []}
                      disableRiskControls={isPermitOriginator}
                      defaultRows={_riskAssessmentsTasks?.sort((a, b) => a.orderRecord - b.orderRecord) || []}
                      onChange={handleRiskTasksChange}
                    />
                  </div>
                </div>
              )}

              <div className="row pb-3" id="precautionsSection" >
                <div>
                  <Label className={styles.ptwLabel}>Precautions Required</Label>
                </div>

                <div className="form-group col-md-12">
                  <div className={styles.checkboxContainer}>
                    <CheckBoxDistributerComponent id="precautionsComponent"
                      optionList={ptwFormStructure?.precuationsItems || []}
                      selectedIds={Array.from(_selectedPrecautionIds)}
                      onChange={(ids) => setSelectedPrecautionIds(new Set(ids))}
                      othersTextValue={_precautionsOtherText}
                      onOthersChange={(checked, othersText) => setPrecautionsOtherText(othersText)}
                    />
                  </div>
                </div>
              </div>

              <Separator />
              <div className='row pb-3' id="gasTestFireWatchAttachmentsSection">
                {/* Gas Test Required Section */}
                <div className='form-group col-md-12' style={{ display: "flex", alignItems: "center" }}>
                  <div className='col-md-3'><Label>Gas Test Required</Label></div>
                  <div className="col-md-9" style={{ display: "flex", alignItems: "center" }}>
                    <div style={{ display: "flex", gap: "30px" }}>
                      {ptwFormStructure?.gasTestRequired?.map((gas, i) => (
                        <div key={i}>
                          <Checkbox
                            label={gas}
                            checked={_gasTestValue === gas}
                            // onChange={() => setGasTestValue(gas)}
                            onChange={() => {
                              setGasTestValue(prev => (prev === gas ? '' : gas));
                              setGasTestResult('');
                            }
                            }
                            disabled={!isPermitIssuer}
                          />
                        </div>
                      ))}

                      <Label style={{ paddingRight: '10px' }}>Gas Test Result:</Label>
                    </div>
                    <div style={{ flex: '1' }}>
                      <TextField
                        type="text" style={{ padding: '4px 6px', border: '1px solid #ccc', borderRadius: '4px' }}
                        placeholder="Enter result"
                        disabled={!isPermitIssuer || _gasTestValue !== 'Yes'}
                        value={_gasTestResult}
                        onChange={(e, newValue) => setGasTestResult(newValue || '')}
                      />
                    </div>
                  </div>
                </div>

                {/* Fire Watch Needed Section */}
                <div className='form-group col-md-12 mt-3' style={{ display: "flex", alignItems: "center" }}>
                  <div className='col-md-3'><Label>Fire Watch Needed</Label></div>
                  <div className="col-md-9" style={{ display: "flex", alignItems: "center" }}>
                    <div style={{ display: "flex", gap: "30px" }}>
                      {ptwFormStructure?.fireWatchNeeded?.map((item, i) => (
                        <div key={i}>
                          <Checkbox
                            label={item}
                            checked={_fireWatchValue === item}
                            // onChange={() => setFireWatchValue(item)}
                            onChange={() => {
                              setFireWatchValue(prev => (prev === item ? '' : item));
                              setFireWatchAssigned('');
                            }}
                            disabled={!isPermitIssuer}
                          />
                        </div>
                      ))}
                      <Label style={{ paddingRight: '10px' }}>Firewatch Assigned:</Label>
                    </div>
                    <div style={{ flex: '1' }}>
                      <TextField type="text" style={{ padding: '4px 6px', border: '1px solid #ccc', borderRadius: '4px' }}
                        placeholder="Enter name"
                        disabled={!isPermitIssuer || _fireWatchValue !== 'Yes'}
                        value={_fireWatchAssigned}
                        onChange={(e, newValue) => setFireWatchAssigned(newValue || '')}
                      />
                    </div>
                  </div>
                </div>

                {/* Attachments Required */}
                <div className='form-group col-md-12 mt-3' style={{ display: "flex", alignItems: "center" }}>
                  <div className='col-md-3'><Label>Attachment(s) provided</Label></div>
                  <div className="" style={{ display: "flex", alignItems: "center" }}></div>
                  <div style={{ display: "flex", gap: "30px" }}>
                    {ptwFormStructure?.attachmentsProvided?.map((attachment, i) => (
                      <div key={i}>
                        <Checkbox
                          label={attachment}
                          checked={_attachmentsValue.toLowerCase() == attachment.toLowerCase() ? true : false}
                          onChange={() => {
                            setAttachmentsValue(prev => (prev === attachment ? '' : attachment))
                            setAttachmentsResult('');
                          }}
                        />
                      </div>
                    ))}
                    <Label style={{ paddingRight: '10px' }}>Details:</Label>
                  </div>
                  <div style={{ flex: '1' }}>
                    <TextField type="text" style={{ padding: '4px 6px', border: '1px solid #ccc', borderRadius: '4px' }}
                      placeholder="Enter detail"
                      disabled={_attachmentsValue !== 'Yes'}
                      value={_attachmentsResult}
                      onChange={(e, newValue) => setAttachmentsResult(newValue || '')}
                    />
                  </div>
                </div>
              </div>
              <Separator />

              <div className="row pb-3" id="protectiveSafetyEquipmentSection" >
                <div>
                  <Label className={styles.ptwLabel}>Protective & Safety Equipment</Label>
                </div>

                <div className="form-group col-md-12">
                  <div className={styles.checkboxContainer}>
                    <CheckBoxDistributerComponent id="protectiveSafetyEquipmentComponent"
                      optionList={ptwFormStructure?.protectiveSafetyEquipments || []}
                      selectedIds={Array.from(_selectedProtectiveEquipmentIds)}
                      onChange={(ids) => setSelectedProtectiveEquipmentIds(new Set(ids))}
                      othersTextValue={_protectiveEquipmentsOtherText}
                      onOthersChange={(checked, othersText) => setProtectiveEquipmentsOtherText(othersText)}
                    />
                  </div>
                </div>
              </div>

              <div className='row pb-3' id="machineryToolsSection">
                <div>
                  <Label className={styles.ptwLabel}>Machinery Involved / Tools</Label>
                </div>
                <div className="form-group col-md-12">
                  <div className='col-md-12'>
                    <ComboBox
                      key={`machinery-${_selectedMachineryIds?.slice().sort((a, b) => a - b).join('_')}`}
                      placeholder="Select machinery/tools"
                      options={machineryOptions as any}
                      selectedKey={_selectedMachineryIds}
                      onChange={onMachineryChange}
                      multiSelect
                      useComboBoxAsMenuWidth
                      styles={comboBoxBlackStyles}
                    />
                  </div>
                  <div className='col-md-12'>
                    <div style={{ borderRadius: 4, padding: 8, marginTop: 8, width: '100%' }}>
                      {selectedMachinery?.length === 0 ? (
                        <span style={{ color: '#605e5c', fontStyle: 'italic' }}>No machines selected</span>
                      ) : (
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                          {selectedMachinery?.map(m => (
                            <span key={m.id}
                              style={{
                                background: '#f3f2f1',
                                border: '1px solid #c8c6c4',
                                borderRadius: 12,
                                padding: '2px 6px',
                                display: 'inline-flex',
                                alignItems: 'center',
                                gap: 6
                              }}>
                              <span style={{ color: '#323130' }}>{m.title}</span>
                              <IconButton
                                iconProps={{ iconName: 'Cancel' }}
                                ariaLabel={`Remove ${m.title}`}
                                title={`Remove ${m.title}`}
                                onClick={() => removeMachinery(m.id)}
                                styles={{ root: { height: 20, width: 20, minWidth: 20 }, icon: { fontSize: 12 } }}
                              />
                            </span>
                          ))}
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              </div>

              {/* Personnel Involved - placed under Attachments section */}
              <div className='row pb-3' id="personnelInvolvedSection">
                <div>
                  <Label className={styles.ptwLabel}>Personnel Involved</Label>
                </div>
                <div className="form-group col-md-12">
                  <ComboBox
                    key={`personnel-${_selectedPersonnelIds?.slice().sort((a, b) => a - b).join('_')}`}
                    placeholder="Select personnel"
                    options={personnelOptions as any}
                    onChange={onPersonnelChange}
                    selectedKey={_selectedPersonnelIds}
                    multiSelect
                    useComboBoxAsMenuWidth
                    styles={comboBoxBlackStyles}
                  />
                  <div style={{ borderRadius: 4, padding: 8, marginTop: 8, width: '100%' }}>
                    {selectedPersonnel?.length === 0 ? (
                      <span style={{ color: '#605e5c', fontStyle: 'italic' }}>No personnel selected</span>
                    ) : (
                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                        {selectedPersonnel?.map(p => (
                          <span key={p.Id}
                            style={{
                              background: '#f3f2f1',
                              border: '1px solid #c8c6c4',
                              borderRadius: 12,
                              padding: '2px 6px',
                              display: 'inline-flex',
                              alignItems: 'center',
                              gap: 6
                            }}>
                            <span style={{ color: '#323130' }}>{p.fullName}</span>
                            <IconButton
                              iconProps={{ iconName: 'Cancel' }}
                              ariaLabel={`Remove ${p.fullName}`}
                              title={`Remove ${p.fullName}`}
                              onClick={() => removePersonnel(Number(p.Id))}
                              styles={{ root: { height: 20, width: 20, minWidth: 20 }, icon: { fontSize: 12 } }}
                            />
                          </span>
                        ))}
                      </div>
                    )}
                  </div>
                </div>
              </div>

              <div className="row pb-3" id="InstructionsSection">
                {/* Instructions For Use */}
                <Stack horizontal id="InstructionsStack">
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
              </div>

              {/* Toolbox Talk (TBT) */}
              {
                isSubmitted && (
                  <div className="row pb-3" id="toolboxTalkSection" style={{ alignItems: 'center' }}>
                    <div className="col-md-3" style={{ display: 'flex', alignItems: 'center' }}>
                      <Checkbox
                        label="Toolbox Talk (TBT); complete details if applicable"
                        checked={!!_selectedToolboxTalk}
                        onChange={(_, chk) => {
                          const isChecked = !!chk;
                          setToolboxTalk(isChecked);
                          if (!isChecked) {
                            setToolboxConductedBy([]);
                            setToolboxHSEReference('');
                            setToolboxTalkDate(undefined);
                          }
                        }}
                        disabled={!isPermitOriginator && !isPerformingAuthority}
                      />
                    </div>

                    <div className="col-md-4">
                      <Label>Conducted By (Title)</Label>
                      <NormalPeoplePicker
                        onResolveSuggestions={_onFilterChanged}
                        itemLimit={1}
                        className={'ms-PeoplePicker'}
                        key={'toolboxConductedBy'}
                        removeButtonAriaLabel={'Remove'}
                        onInputChange={onInputChange}
                        resolveDelay={150}
                        styles={peoplePickerBlackStyles}
                        selectedItems={_selectedToolboxConductedBy}
                        onChange={(items) => setToolboxConductedBy(items || [])}
                        inputProps={{ placeholder: 'Enter name or email' }}
                        pickerSuggestionsProps={suggestionProps}
                        disabled={(!isPermitOriginator && !isPerformingAuthority) || !_selectedToolboxTalk}
                      />
                    </div>

                    <div className="col-md-3">
                      <Label>HSE TBT Reference</Label>
                      <TextField
                        placeholder="Enter reference"
                        value={String(_toolboxHSEReference || '')}
                        onChange={(_, v) => setToolboxHSEReference(v || '')}
                        disabled={(!isPermitOriginator && !isPerformingAuthority) || !_selectedToolboxTalk}
                      />
                    </div>

                    <div className="col-md-2">
                      <Label>Date</Label>
                      <DatePicker
                        placeholder="Select date"
                        value={_selectedToolboxTalkDate ? new Date(String(_selectedToolboxTalkDate)) : new Date()}
                        onSelectDate={(d) => setToolboxTalkDate(d ? d.toISOString() : undefined)}
                        disabled={(!isPermitOriginator && !isPerformingAuthority)}
                      />
                    </div>
                  </div>
                )
              }

              {/* PTW Sign Off and Approval - visible when submitted */}
              {/* {isSubmitted &&
                (  */}
              <div className="row pb-3" id="ptwSignOffSection" style={{ border: '1px solid #c8c6c4', borderRadius: 4, background: '#e9edf7' }}>
                <div className="col-md-12" style={{ paddingTop: 8 }}>
                  <Label style={{ fontWeight: 600 }}>PTW Sign Off and Approval</Label>
                </div>

                {/* Permit Originator (PO) */}
                <div className="col-md-4" style={{ padding: 8 }}>
                  <Label style={{ fontWeight: 600 }}>Permit Originator (PO)</Label>
                  <TextField className='pb-1'
                    value={_PermitOriginator?.[0]?.text || ''}
                    readOnly
                  // disabled={!isPermitOriginator && isPermitOriginator} 
                  />
                  <DatePicker
                    disabled={true}
                    placeholder="Select date"
                    value={_poDate ? new Date(_poDate) : new Date()}
                    onSelectDate={d => setPoDate(d ? d.toISOString() : undefined)}
                  />
                  {/* <ComboBox
                    placeholder="Status"
                    options={statusOptions}
                    selectedKey={_poStatus}
                    onChange={(_, opt) => setPoStatus((opt?.key as SignOffStatus) ?? 'Pending')}
                    useComboBoxAsMenuWidth
                    disabled={isPermitOriginator}
                  /> */}
                </div>

                {/* Performing Authority (PA) */}
                <div className="col-md-4" style={{ padding: 8 }}>
                  <Label style={{ fontWeight: 600 }}>Performing Authority (PA)</Label>
                  <ComboBox
                    placeholder="Select Performing Authority"
                    disabled={!stageEnabled.paEnabled}
                    options={getOptionsForGroup('PerformingAuthorityGroup')}
                    selectedKey={selectedKeyFromPicker(_paPicker)}
                    onChange={onSingleApproverChange('PerformingAuthorityGroup', (items) => setPaPicker(items), setPaStatusEnabled)}
                    useComboBoxAsMenuWidth
                    styles={comboBoxBlackStyles}
                    className={'pb-1'}
                  />
                  <DatePicker
                    disabled={true}
                    placeholder="Select date"
                    value={_paDate ? new Date(_paDate) : new Date()}
                    onSelectDate={d => setPaDate(d ? d.toISOString() : undefined)}
                  />
                  <ComboBox
                    // disabled={!isPerformingAuthority}
                    disabled={!stageEnabled.paEnabled || !_paStatusEnabled}
                    placeholder="Status"
                    options={statusOptions}
                    selectedKey={_paStatus}
                    onChange={(_, opt) => setPaStatus((opt?.key as SignOffStatus) ?? 'Pending')}
                    useComboBoxAsMenuWidth
                  />
                </div>

                {/* Permit Issuer (PI) */}
                <div className="col-md-4" style={{ padding: 8 }}>
                  <Label style={{ fontWeight: 600 }}>Permit Issuer (PI)</Label>
                  <ComboBox
                    placeholder="Select Permit Issuer"
                    disabled={!stageEnabled.piEnabled}
                    options={getOptionsForGroup('PermitIssuerGroup')}
                    selectedKey={selectedKeyFromPicker(_piPicker)}
                    onChange={onSingleApproverChange('PermitIssuerGroup', (items) => setPiPicker(items), setPiStatusEnabled)}
                    useComboBoxAsMenuWidth
                    styles={comboBoxBlackStyles}
                    className={'pb-1'}
                  />
                  <DatePicker
                    disabled={true}
                    placeholder="Select date"
                    value={_piDate ? new Date(_piDate) : new Date()}
                    onSelectDate={d => setPiDate(d ? d.toISOString() : undefined)}
                  />
                  <ComboBox
                    disabled={!stageEnabled.piEnabled || !_piStatusEnabled}
                    placeholder="Status"
                    options={statusOptions}
                    selectedKey={_piStatus}
                    onChange={(_, opt) => setPiStatus((opt?.key as SignOffStatus) ?? 'Pending')}
                    useComboBoxAsMenuWidth
                  />
                </div>
              </div>
              {/* )} */}

              {/* HIGH RISK PTW Approval (if applicable) - visible when submitted and overall risk is High */}
              {isSubmitted && isHighRisk && (
                <div className="row pb-3" id="highRiskApprovalSection" style={{ border: '1px solid #c8c6c4', borderRadius: 4, background: '#e9edf7' }}>
                  <div className="col-md-12" style={{ paddingTop: 8 }}>
                    <Label style={{ fontWeight: 600 }}>
                      HIGH RISK PTW Approval <span style={{ fontStyle: 'italic', fontWeight: 400 }}>(if applicable)</span>
                    </Label>
                  </div>

                  <div className="col-md-6" style={{ padding: 8 }}>
                    <Label style={{ fontWeight: 600 }}>Asset Director</Label>
                    <ComboBox
                      placeholder="Select Asset Director"
                      disabled={!stageEnabled.assetDirEnabled}
                      options={getOptionsForGroup('AssetDirectorsGroup')}
                      selectedKey={selectedKeyFromPicker(_assetDirPicker)}
                      onChange={onSingleApproverChange('AssetDirectorsGroup', (items) => setAssetDirPicker(items), setAssetDirStatusEnabled)}
                      useComboBoxAsMenuWidth
                      styles={comboBoxBlackStyles}
                      className={'pb-1'}
                    />
                    <DatePicker
                      disabled={true}
                      placeholder="Select date"
                      value={_assetDirDate ? new Date(_assetDirDate) : new Date()}
                      onSelectDate={d => setAssetDirDate(d ? d.toISOString() : undefined)}
                    />
                    <ComboBox
                      disabled={!stageEnabled.assetDirEnabled || !_assetDirStatusEnabled}
                      placeholder="Status"
                      options={statusOptions}
                      selectedKey={_assetDirStatus}
                      onChange={(_, opt) => setAssetDirStatus((opt?.key as SignOffStatus) ?? 'Pending')}
                      useComboBoxAsMenuWidth
                    />
                  </div>

                  <div className="col-md-6" style={{ padding: 8 }}>
                    <Label style={{ fontWeight: 600 }}>HSE Director</Label>
                    <ComboBox
                      placeholder="Select HSE Director"
                      disabled={!stageEnabled.hseDirEnabled}
                      options={getOptionsForGroup('HSEDirectorGroup')}
                      selectedKey={selectedKeyFromPicker(_hseDirPicker)}
                      onChange={onSingleApproverChange('HSEDirectorGroup', (items) => setHseDirPicker(items), setHseDirStatusEnabled)}
                      useComboBoxAsMenuWidth
                      styles={comboBoxBlackStyles}
                      className={'pb-1'}
                    />
                    <DatePicker
                      disabled={true}
                      placeholder="Select date"
                      value={_hseDirDate ? new Date(_hseDirDate) : new Date()}
                      onSelectDate={d => setHseDirDate(d ? d.toISOString() : undefined)}
                    />
                    <ComboBox
                      // disabled={!isHSEDirector}
                      disabled={!stageEnabled.hseDirEnabled || !_hseDirStatusEnabled}
                      placeholder="Status"
                      options={statusOptions}
                      selectedKey={_hseDirStatus}
                      onChange={(_, opt) => setHseDirStatus((opt?.key as SignOffStatus) ?? 'Pending')}
                      useComboBoxAsMenuWidth
                    />
                  </div>
                </div>
              )}

              {/* PTW Closure */}
              {isSubmitted && (
                <div className="row pb-3" id="ptwClosureSection" style={{ border: '1px solid #c8c6c4', borderRadius: 4, background: '#e9edf7' }}>
                  <div className="col-md-12" style={{ paddingTop: 8 }}>
                    <Label style={{ fontWeight: 600 }}>PTW Closure</Label>
                    <div style={{ fontStyle: 'italic', color: '#323130', marginTop: 2, fontSize: 'smaller' }}>
                      I declare that the jobs stated in this PTW have been completed, the precautions stated above can be removed and normal operations can be resumed.
                    </div>
                  </div>

                  <div className="col-md-6" style={{ padding: 8 }}>
                    <Label style={{ fontWeight: 600 }}>Permit Originator (PO)</Label>
                    <TextField className='pb-1'
                      value={_PermitOriginator?.[0]?.text || ''}
                      disabled={true}
                    />
                    <DatePicker
                      placeholder="Select date"
                      value={_closurePoDate ? new Date(_closurePoDate) : new Date()}
                      onSelectDate={d => setClosurePoDate(d ? d.toISOString() : undefined)}
                      disabled={true}
                    />
                    <ComboBox
                      placeholder='Status'
                      options={statusOptions}
                      selectedKey={_closurePoStatus}
                      disabled={!isPermitOriginator}
                      onChange={(_, opt) => setClosurePoStatus((opt?.key as SignOffStatus) ?? 'Pending')}
                      useComboBoxAsMenuWidth
                    />
                  </div>

                  <div className="col-md-6" style={{ padding: 8 }}>
                    <Label style={{ fontWeight: 600 }}>Asset Manager</Label>
                    <ComboBox
                      placeholder="Select Asset Manager"
                      disabled={!stageEnabled.closureEnabled}
                      options={getOptionsForGroup('AssetManagersGroup')}
                      selectedKey={selectedKeyFromPicker(_closureAssetManagerPicker)}
                      // onChange={onSingleApproverChange('AssetManagersGroup', (items) => setClosureAssetManagerPicker(items))}
                      onChange={onSingleApproverChange('AssetManagersGroup', (items) => setClosureAssetManagerPicker(items), setClosureAssetManagerStatusEnabled)}
                      useComboBoxAsMenuWidth
                      styles={comboBoxBlackStyles}
                      className={'pb-1'}
                    />
                    <DatePicker
                      disabled={true}
                      placeholder="Select date"
                      value={_closureAssetManagerDate ? new Date(_closureAssetManagerDate) : new Date()}
                      onSelectDate={d => setClosureAssetManagerDate(d ? d.toISOString() : undefined)}
                    />
                    <ComboBox
                      // disabled={!isAssetManager}
                      disabled={!stageEnabled.closureEnabled || !_closureAssetManagerStatusEnabled}
                      placeholder='Status'
                      options={statusOptions}
                      selectedKey={_closureAssetManagerStatus}
                      onChange={(_, opt) => setClosureAssetManagerStatus((opt?.key as SignOffStatus) ?? 'Pending')}
                      useComboBoxAsMenuWidth
                    />
                  </div>
                </div>
              )}

            </div>
          )}
        </div>

      </form>

      <Separator />

      {/* Bottom action buttons */}
      <div id="formButtonsSection" className="no-pdf" style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 8, marginBottom: 8 }}>
        <DefaultButton text="Close" onClick={handleCancel} />

        <ExportPdfControls targetRef={containerRef} coralReferenceNumber={_coralReferenceNumber}
          employeeName={_PermitOriginator?.[0]?.text}
          exportMode={exportMode}
          onExportModeChange={setExportMode}
          onBusyChange={setIsExportingPdf}
          // isClosedBySystem={(formsApprovalWorkflow || []).some(r => String(r?.Status?.title || '').toLowerCase().includes('approved') && r.FinalLevel === r.Order)}
          onError={(m) => showBanner(m)}
        />

        {(mode === "new" || mode === "saved") &&
          <>
            <DefaultButton text="Save"
              onClick={() => submitForm('save')}
              disabled={!isPermitOriginator || isBusy}
            />

            <DefaultButton text="Submit"
              onClick={() => submitForm('submit')}
              disabled={!isPermitOriginator || isBusy}
            />
          </>
        }

        {(mode === "submitted") && (
          <PrimaryButton text="Approve"
            onClick={() => approveForm('approve')}
            disabled={!isPermitOriginator || isBusy}
          />
        )}
      </div>

      <div id="formFooterSection" className='row'>
        <div className='col-md-12 col-lg-12 col-xl-12 col-sm-12'>
          <DocumentMetaBanner docCode='COR-HSE-21-FOR-005' version='V04' effectiveDate='06-AUG-2024' />
        </div>
      </div>

    </div>
  );
}