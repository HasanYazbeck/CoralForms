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
  DatePicker,
  Toggle,
  DialogFooter,
  Dialog,
  DialogType,
  DialogContent,
  defaultDatePickerStrings,
  IDatePickerStyles,
  TeachingBubble,
  ITextFieldStyles,
  IToggleStyles
} from '@fluentui/react';
import { NormalPeoplePicker, IBasePickerSuggestionsProps, IBasePickerStyles } from '@fluentui/react/lib/Pickers';
import { ICompany, ILKPItemInstructionsForUse } from '../../../Interfaces/Common/ICommon';
import { MSGraphClientV3 } from '@microsoft/sp-http-msgraph';
import { SPHttpClient } from '@microsoft/sp-http';
import { IUser } from '../../../Interfaces/Common/IUser';
import { SPCrudOperations } from "../../../Classes/SPCrudOperations";
import { SPHelpers } from "../../../Classes/SPHelpers";
import { IAssetCategoryDetails, IAssetsDetails, ICoralForm, IEmployeePeronellePassport, ILookupItem, IPTWForm, IPTWWorkflow, ISagefaurdsItem, IWorkCategory, WorkflowStages } from '../../../Interfaces/PtwForm/IPTWForm';
import { CheckBoxDistributerComponent } from './CheckBoxDistributerComponent';
import RiskAssessmentList, { IRiskTaskRow } from './RiskAssessmentList';
import { CheckBoxDistributerOnlyComponent } from './CheckBoxDistributerOnlyComponent';
import { DocumentMetaBanner } from '../../../Components/DocumentMetaBanner';
import { ICoralFormsList } from '../../../Interfaces/Common/ICoralFormsList';
import ExportPdfControls from '../../ptwForm/components/ExportPdfControls';
import BannerComponent, { BannerKind } from '../../ppeForm/components/BannerComponent';
import { PTWWorkflowStatus } from '../../../Enums/enums';

interface IRiskAssessmentResult {
  rows: IRiskTaskRow[];
  overallRisk?: string;
  l2Required?: boolean;
  l2Ref?: string;
}

export default function PTWForm(props: IPTWFormProps) {

  // Add status type and options
  type SignOffStatus = 'Pending' | 'Approved' | 'Rejected' | 'Closed';

  const statusOptions: IDropdownOption[] = React.useMemo(() => ([
    { key: 'Pending', text: 'Pending' },
    { key: 'Approved', text: 'Approved' },
    { key: 'Rejected', text: 'Rejected' },
    { key: 'Closed', text: 'Closed' }
  ]), []);

  // Helpers and refs
  const formName = "Permit To Work";
  const containerRef = React.useRef<HTMLDivElement>(null);
  const overlayRef = React.useRef<HTMLDivElement>(null);
  const spCrudRef = React.useRef<SPCrudOperations | undefined>(undefined);
  const payloadRef = React.useRef<any>(null);
  const spHelpers = React.useMemo(() => new SPHelpers(), []);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [mode, setMode] = React.useState<'saved' | 'submitted' | 'approved' | 'new' | 'rejected'>('new');

  const [isExportingPdf, setIsExportingPdf] = React.useState(false); // NEW
  const [exportMode, setExportMode] = React.useState(false);
  const [bannerText, setBannerText] = React.useState<string>();
  const [bannerTick, setBannerTick] = React.useState(0);
  const [bannerOpts, setBannerOpts] = React.useState<{ autoHideMs?: number; fade?: boolean; kind?: BannerKind } | undefined>();
  const bannerTopRef = React.useRef<HTMLDivElement>(null);

  const [_users, setUsers] = React.useState<IUser[]>([]);
  const [_coralFormList, setCoralFormsList] = React.useState<ICoralFormsList>({ Id: "" });
  const [ptwFormStructure, setPTWFormStructure] = React.useState<IPTWForm>({ issuanceInstrunctions: [], personnalInvolved: [] });
  const [itemInstructionsForUse, setItemInstructionsForUse] = React.useState<ILKPItemInstructionsForUse[]>([]);
  const [personnelInvolved, setPersonnelInvolved] = React.useState<IEmployeePeronellePassport[]>([]);
  const [, setAssetDetails] = React.useState<IAssetCategoryDetails[]>([]);
  const [safeguards, setSafeguards] = React.useState<ISagefaurdsItem[]>([]);
  const [filteredSafeguards, setFilteredSafeguards] = React.useState<ISagefaurdsItem[]>([]);

  const webUrl = props.context.pageContext.web.absoluteUrl;
  // Header logo and doc code derived from selected company
  const initialLogoUrl = `${webUrl}/SiteAssets/coral-logo.png`;
  const [companyLogoUrl, setCompanyLogoUrl] = React.useState<string>(initialLogoUrl);
  const [docCode, setDocCode] = React.useState<string>('COR-HSE-21-FOR-005');

  const [assetCategoriesList, setAssetCategoriesList] = React.useState<ILookupItem[] | undefined>([]);
  const [assetCategoriesDetailsList, setAssetCategoriesDetailsList] = React.useState<IAssetsDetails[] | undefined>([]);

  // Form State to used on update or submit
  const [_coralReferenceNumber, setCoralReferenceNumber] = React.useState<string>('');
  const [_previousPtwRef, setPreviousPtwRef] = React.useState<string>('');
  const [_PermitOriginator, setPermitOriginator] = React.useState<IPersonaProps[]>([]);
  const [_assetId, setAssetId] = React.useState<string>('');
  const [_selectedCompany, setSelectedCompany] = React.useState<ICompany | undefined>(undefined);
  const [_selectedAssetCategory, setSelectedAssetCategory] = React.useState<number | undefined>(undefined);
  const [_selectedAssetDetails, setSelectedAssetDetails] = React.useState<number | undefined>(0);
  const [_projectTitle, setProjectTitle] = React.useState<string>('');
  const [_selectedPermitTypeList, setSelectedPermitTypeList] = React.useState<IWorkCategory[]>([]);
  const [_permitPayload, setPermitPayload] = React.useState<IPermitScheduleRow[]>([]);
  const [_permitPayloadValidityDays, setPermitPayloadValidityDays] = React.useState<number>(0);
  const [_selectedHacWorkAreaId, setSelectedHacWorkAreaId] = React.useState<number | undefined>(undefined);
  const [_selectedWorkHazardIds, setSelectedWorkHazardIds] = React.useState<Set<number>>(new Set());
  const [_workHazardsOtherText, setWorkHazardsOtherText] = React.useState<string>('');

  const [_riskAssessmentsTasks, setRiskAssessmentsTasks] = React.useState<IRiskTaskRow[] | undefined>(undefined);
  const [_overAllRiskAssessment, setOverAllRiskAssessment] = React.useState<string | undefined>(undefined);
  const [_detailedRiskAssessment, setDetailedRiskAssessment] = React.useState<boolean>(false);
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
  const [_selectedToolboxTalkDate, setToolboxTalkDate] = React.useState<Date | undefined>(new Date());
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
  const [isIssued, setIsIssued] = React.useState<boolean>(false);
  const [_isAssetDirReplacer, setIsAssetDirectorReplacer] = React.useState<boolean>(false);
  const [_isHseDirReplacer, setIsHseDirectorReplacer] = React.useState<boolean>(false);

  // SharePoint group members cache
  type SPGroupUser = { id: number; title: string; email: string };
  const [groupMembers, setGroupMembers] = React.useState<Record<string, SPGroupUser[]>>({});

  // Sign-off state
  const [_poDate, setPoDate] = React.useState<Date | undefined>(new Date());
  const [_poStatus, setPoStatus] = React.useState<SignOffStatus>('Approved');

  const [_paPicker, setPaPicker] = React.useState<IPersonaProps[]>([]);
  const [_paDate, setPaDate] = React.useState<Date | undefined>(new Date());
  const [_paStatus, setPaStatus] = React.useState<SignOffStatus>('Pending');
  const [_paRejectionReason, setPaRejectionReason] = React.useState<string>('');
  const [_paStatusEnabled, setPaStatusEnabled] = React.useState<boolean>(false);

  const [_piHsePartnerFilteredByCategory, setPiHsePartnerFilteredByCategory] = React.useState<IPersonaProps[]>([]);
  const [_assetDirFilteredByCategory, setAssetDirFilteredByCategory] = React.useState<IPersonaProps[]>([]);
  const [_assetManagerFilteredByCategory, setAssetManagerFilteredByCategory] = React.useState<IPersonaProps[]>([]);

  const [_piPicker, setPiPicker] = React.useState<IPersonaProps[]>([]);
  const [_piDate, setPiDate] = React.useState<Date | undefined>(new Date());
  const [_piStatus, setPiStatus] = React.useState<SignOffStatus>('Pending');
  const [_piRejectionReason, setPiRejectionReason] = React.useState<string>('');
  const [_piStatusEnabled, setPiStatusEnabled] = React.useState<boolean>(false);
  const [_piUnlockedByPA, setPiUnlockedByPA] = React.useState<boolean>(false);

  const [_assetDirPicker, setAssetDirPicker] = React.useState<IPersonaProps[]>([]);
  const [_assetDirReplacerPicker, setAssetDirReplacerPicker] = React.useState<IPersonaProps[]>([]);

  const [_assetDirDate, setAssetDirDate] = React.useState<Date | undefined>(new Date());
  const [_assetDirStatus, setAssetDirStatus] = React.useState<SignOffStatus>('Pending');
  const [_assetDirRejectionReason, setAssetDirRejectionReason] = React.useState<string>('');
  const [_assetDirStatusEnabled, setAssetDirStatusEnabled] = React.useState<boolean>(false);

  const [_urgentAssetDirDate, setUrgentAssetDirDate] = React.useState<Date | undefined>(new Date());
  const [_urgentAssetDirStatus, setUrgentAssetDirStatus] = React.useState<SignOffStatus>('Pending');
  const [_urgentAssetDirRejectionReas, setUrgentAssetDirRejectionReas] = React.useState<string>('');

  const [_hseDirPicker, setHseDirPicker] = React.useState<IPersonaProps[]>([]);
  const [_hseDirReplacerPicker, setHseDirReplacerPicker] = React.useState<IPersonaProps[]>([]);

  const [_hseDirDate, setHseDirDate] = React.useState<Date | undefined>(new Date());
  const [_hseDirStatus, setHseDirStatus] = React.useState<SignOffStatus>('Pending');
  const [_hseDirRejectionReason, setHseDirRejectionReason] = React.useState<string>('');
  const [_hseDirStatusEnabled, setHseDirStatusEnabled] = React.useState<boolean>(false);

  // PTW Closure state
  const [_closurePoDate, setClosurePoDate] = React.useState<Date | undefined>(new Date());
  const [_closurePoStatus, setClosurePoStatus] = React.useState<SignOffStatus>('Pending');
  const [_poRejectionReason, setPORejectionReason] = React.useState<string>('');

  const [_closureAssetManagerPicker, setClosureAssetManagerPicker] = React.useState<IPersonaProps[]>([]);
  const [_closureAssetManagerDate, setClosureAssetManagerDate] = React.useState<Date | undefined>(new Date());
  const [_closureAssetManagerStatus, setClosureAssetManagerStatus] = React.useState<SignOffStatus>('Pending');
  const [_closureAssetManagerStatusEnabled, setClosureAssetManagerStatusEnabled] = React.useState<boolean>(false);
  const [_asssetManagerRejectionReason, setAssetManagerRejectionReason] = React.useState<string>('');
  const [_workflowStage, setWorkflowStage] = React.useState<WorkflowStages>(undefined);

  // State for controlling conditional rendering of sections
  const [workPermitRequired, setWorkPermitRequired] = React.useState<boolean>(false);

  // Urgent submission: bypass Submission Range Interval validation on submit
  const [_isUrgentSubmission, setIsUrgentSubmission] = React.useState<boolean>(false);
  const [prefilledFormId, setPrefilledFormId] = React.useState<number | undefined>(undefined);
  const [_canPOResubmitAfterRejection, setCanPOResubmitAfterRejection] = React.useState<boolean>(false);
  const [suppressAutoPrefill, setSuppressAutoPrefill] = React.useState<boolean>(false);
  const [showExtendDialog, setShowExtendDialog] = React.useState(false);

  // States for Work Extension Dialog
  const rejectionResetDoneRef = React.useRef(false);
  const [selectedDate, setSelectedDate] = React.useState<Date | undefined>(undefined);
  const [startTime, setStartTime] = React.useState('');
  const [endTime, setEndTime] = React.useState('');
  const [selectedApprover, setSelectedApprover] = React.useState<string | undefined>(undefined);
  const [errors, setErrors] = React.useState<string[]>([]);
  const [isTeachingBubbleVisible, setIsTeachingBubbleVisible] = React.useState(false);
  const buttonRef = React.useRef<HTMLDivElement>(null);

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

  const customToggleStyles: Partial<IToggleStyles> = {
    label: {
      color: '#000',       // Black color
      fontWeight: 500
    },
    // thumb: {
    //   backgroundColor: '#000' // Black color for the thumb
    // }

  };

  // End Styling Components
  const uiDisabled = React.useCallback((normalDisabled: boolean) => (exportMode ? false : normalDisabled), [exportMode]);

  const isHighRisk = React.useMemo(() => {
    return String(_overAllRiskAssessment || '').toLowerCase().includes('high');
  }, [_overAllRiskAssessment]);

  const currentUserEmail = (props.context?.pageContext?.user?.email || '').toLowerCase();
  const filterOutCurrentUser = React.useCallback((people?: IPersonaProps[]) => {
    if (!people?.length) return [];
    return people.filter(p => (p.secondaryText || '').toLowerCase() !== currentUserEmail);
  }, [currentUserEmail]);

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
        if (!disposed) setIsPerformingAuthority(isEligibleGroup && (_workflowStage?.toLowerCase() === "ApprovedFromPOToPA".toLowerCase()));
      } catch {
        if (!disposed) setIsPerformingAuthority(false);
      }
    }

    if (currentUserEmail) PerformingAuthorityGroup();
    return () => { disposed = true; };
  }, [props.context.spHttpClient, webUrl, currentUserEmail, _workflowStage]);

  // Determine eligibility for Permit Issuer / HSE Partner based on selected asset details (people field)
  React.useEffect(() => {
    let disposed = false;
    try {
      const selId = _selectedAssetDetails != null ? Number(_selectedAssetDetails) : NaN;
      const detail = (assetCategoriesDetailsList || []).find(d => Number(d.id) === selId);
      const hsePartners: IPersonaProps[] = detail?.hsePartner || [];

      const isPI = (hsePartners || []).some(p => (p.secondaryText || '').toLowerCase() === currentUserEmail);
      const isPIIssuer = (isPI && (_workflowStage?.toLowerCase() === "ApprovedFromPAToPI".toLowerCase() || _workflowStage?.toLowerCase() === "ApprovedFromPOToPI".toLowerCase()
        || _workflowStage?.toLowerCase() === "Issued".toLowerCase()));
      if (!disposed) setIsPermitIssuer(isPIIssuer);
    } catch {
      if (!disposed) setIsPermitIssuer(false);
    }
    return () => { disposed = true; };
  }, [assetCategoriesDetailsList, _selectedAssetDetails, currentUserEmail, _workflowStage]);

  // Determine eligibility for Asset Director based on selected asset details (people field)
  React.useEffect(() => {
    let disposed = false;
    try {
      const selId = _selectedAssetDetails != null ? Number(_selectedAssetDetails) : NaN;
      // Prefer the selected asset details; fall back to the filtered list if needed
      const detail = (assetCategoriesDetailsList || []).find(d => Number(d.id) === selId);
      const director: IPersonaProps | undefined = detail?.assetDirector?.[0];
      const directorReplacer: IPersonaProps | undefined = detail?.assetDirectorReplacer?.[0];

      let isMember: boolean = false;
      if (director || directorReplacer) {
        if (director?.secondaryText?.toLowerCase() == currentUserEmail) {
          isMember = true;
        }
        else if (directorReplacer?.secondaryText?.toLowerCase() == currentUserEmail) {
          isMember = true;
        }
      }

      if (!disposed) setIsAssetDirector(isMember);
    } catch {
      if (!disposed) setIsAssetDirector(false);
    }
    return () => { disposed = true; };
  }, [assetCategoriesDetailsList, _selectedAssetDetails, currentUserEmail]);

  // Determine eligibility for HSE Director based on selected asset details (people field)
  React.useEffect(() => {
    let disposed = false;
    try {
      const selId = _selectedAssetDetails != null ? Number(_selectedAssetDetails) : NaN;
      // Prefer the selected asset details; fall back to the filtered list if needed
      const detail = (assetCategoriesDetailsList || []).find(d => Number(d.id) === selId);
      const director: IPersonaProps | undefined = detail?.hseDirector?.[0];
      const directorReplacer: IPersonaProps | undefined = detail?.hseDirectorReplacer?.[0];

      let isMember: boolean = false;
      if (director || directorReplacer) {
        if (director?.secondaryText?.toLowerCase() == currentUserEmail) {
          isMember = true;
        }
        else if (directorReplacer?.secondaryText?.toLowerCase() == currentUserEmail) {
          isMember = true;
        }
      }

      if (!disposed) setIsHSEDirector(isMember);
    } catch {
      if (!disposed) setIsHSEDirector(false);
    }
    return () => { disposed = true; };
  }, [assetCategoriesDetailsList, _selectedAssetDetails, currentUserEmail]);

  // Determine eligibility for Asset Manager based on selected asset details (people field)
  React.useEffect(() => {
    let disposed = false;
    try {
      const selId = _selectedAssetDetails != null ? Number(_selectedAssetDetails) : NaN;
      const detail = (assetCategoriesDetailsList || []).find(d => Number(d.id) === selId);
      const managers: IPersonaProps[] = detail?.assetManager || [];
      const isMember = (managers || []).some(p => (p.secondaryText || '').toLowerCase() === currentUserEmail);
      if (!disposed) setIsAssetManager(isMember);
    } catch {
      if (!disposed) setIsAssetManager(false);
    }
    return () => { disposed = true; };
  }, [assetCategoriesDetailsList, _selectedAssetDetails, currentUserEmail]);

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
          'PerformingAuthorityGroup'
        ];
        const entries = await Promise.all(
          names.map(async n => [n, await getGroupMembers(n)] as const)
        );
        if (!disposed) {
          const map: Record<string, SPGroupUser[]> = {};
          entries.forEach(([name, users]) => {
            map[name] = users;
          });
          setGroupMembers(map);
        }
      } catch {
        if (!disposed) setGroupMembers({});
      }
    })();
    return () => { disposed = true; };
  }, [getGroupMembers]);

  // Build dropdown options from a group
  const getOptionsForGroup = React.useCallback((groupName: string): IDropdownOption[] =>
    (groupMembers[groupName] || []).map(m => ({
      key: String(m.id) || m.email,
      text: m.title || m.email
    })),
    [groupMembers]
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
      const idKey = opt.key;
      const u = (groupMembers[groupName] || []).find(x => x.id === Number(idKey));
      const selectedEmail = (u?.email || '').toLowerCase();
      const isCurrentUser = !!selectedEmail && selectedEmail === currentUserEmail;
      setStatusEnabled?.(isCurrentUser);

      if (!isCurrentUser && groupName === 'PerformingAuthorityGroup') {
        setPiPicker([]);
        setPiStatus('Pending');
      }

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

  const onPermitIssuerChange = React.useCallback((setPicker: (items: IPersonaProps[]) => void, setStatusEnabled?: (enabled: boolean) => void) =>
    (_: React.FormEvent<IComboBox>, opt?: IDropdownOption) => {
      if (!opt) {
        setPicker([]);
        setStatusEnabled?.(false);
        return;
      }
      const idKey = opt.key;
      const u = _piHsePartnerFilteredByCategory.find(x => x.id == idKey);
      const selectedEmail = (u?.secondaryText || '').toLowerCase();
      const isCurrentUser = !!selectedEmail && selectedEmail === currentUserEmail;
      setStatusEnabled?.(isCurrentUser);
      // If PA equals the logged-in user, unlock the PI section (ComboBox + Status gating)
      if (_paPicker?.[0].id == u?.id) {
        setPiUnlockedByPA(isCurrentUser);
      }

      setPicker(u ? [{
        text: u.title || '',
        secondaryText: u.secondaryText || '',
        id: String(u.id)
      }] : []);
    },
    [_piHsePartnerFilteredByCategory, currentUserEmail, _paPicker]
  );

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

  // const isSubmitted = (formStatus || mode) === 'submitted';
  const isSubmitted = React.useMemo(() => {
    return (formStatus || mode) === 'submitted';
  }, [formStatus, mode]);

  const ptwStructureSelect = React.useMemo(() => (
    `?$select=Id,AttachmentsProvided,InitialRisk,ResidualRisk,OverallRiskAssessment,FireWatchNeeded,GasTestRequired,` +
    `CoralFormId/Title,CoralFormId/ArabicTitle,` +
    `WorkCategory/Id,WorkCategory/Title,WorkCategory/OrderRecord,WorkCategory/RenewalValidity,` +
    `HACWorkArea/Id,HACWorkArea/Title,HACWorkArea/OrderRecord,` +
    `WorkHazards/Id,WorkHazards/Title,WorkHazards/OrderRecord,` +
    `Machinery/Id,Machinery/Title,Machinery/OrderRecord,` +
    `PrecuationItems/Id,PrecuationItems/Title,PrecuationItems/OrderRecord,` +
    `ProtectiveSafetyEquiment/Id,ProtectiveSafetyEquiment/Title,ProtectiveSafetyEquiment/OrderRecord` +
    `&$expand=CoralFormId,WorkCategory,HACWorkArea,WorkHazards,Machinery,PrecuationItems,` +
    `ProtectiveSafetyEquiment`
  ), []);

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

      const ptwform = data.find((obj: any) => obj !== null);
      let result: ICoralFormsList = { Id: "" };

      if (ptwform) {
        result = {
          Id: ptwform.Id ?? undefined,
          Title: ptwform.Title ?? undefined,
          hasInstructionForUse: ptwform.hasInstructionForUse ?? undefined,
          hasWorkflow: ptwform.hasWorkflow ?? undefined,
          SubmissionRangeInterval: ptwform.SubmissionRangeInterval ?? undefined,
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

        // const _companies: ILookupItem[] = [];
        // if (obj.CompanyRecord !== undefined && obj.CompanyRecord !== null && Array.isArray(obj.CompanyRecord)) {
        //   obj.CompanyRecord.forEach((item: any) => {
        //     if (item) {
        //       _companies.push({
        //         id: item.Id,
        //         title: item.Title,
        //         orderRecord: item.OrderRecord || 0,
        //       });
        //     }
        //   });
        // }

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
          coralForm: coralForm,
          companies: [],
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

  const _getCompany = React.useCallback(async () => {
    try {
      const query: string = `?$select=Id,Title,RecordOrder,LogoPath,FullName&$orderby=RecordOrder asc`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'LKP_company', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: ICompany[] = [];
      data.forEach((obj: any) => {
        if (obj) {
          const temp: ICompany = {
            id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            title: obj.Title !== undefined && obj.Title !== null ? obj.Title : undefined,
            orderRecord: obj.RecordOrder !== undefined && obj.RecordOrder !== null ? obj.RecordOrder : undefined,
            logoUrl: obj.LogoPath !== undefined && obj.LogoPath !== null ? `${webUrl}` + `${obj.LogoPath}`.toString() : '',
            fullName: obj.FullName !== undefined && obj.FullName !== null ? obj.FullName : undefined,
          };
          result.push(temp);
        }
      });
      // sort by Order (ascending). If Order is missing, place those items at the end.
      result.sort((a, b) => {
        const aOrder = (a && a.orderRecord !== undefined && a.orderRecord !== null) ? Number(a.orderRecord) : Number.POSITIVE_INFINITY;
        const bOrder = (b && b.orderRecord !== undefined && b.orderRecord !== null) ? Number(b.orderRecord) : Number.POSITIVE_INFINITY;
        return aOrder - bOrder;
      });
      setPTWFormStructure(prev => ({ ...prev, companies: result }));
    } catch (error) {
      setPTWFormStructure(prev => ({ ...prev, companies: [] }));
      // console.error('An error has occurred while retrieving items!', error);
    }
  }, [props.context, spHelpers]);

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
        `AssetDirector/Id,AssetDirector/EMail,AssetDirector/Title,` +
        `AssetDirectorReplacer/Id,AssetDirectorReplacer/EMail,AssetDirectorReplacer/Title,` +
        `AssetManager/Id,AssetManager/EMail,AssetManager/Title,` +
        `HSEPartner/Id,HSEPartner/EMail,HSEPartner/Title,` +
        `HSEDirector/Id,HSEDirector/EMail,HSEDirector/Title,` +
        `HSEDirectorReplacer/Id,HSEDirectorReplacer/EMail,HSEDirectorReplacer/Title,` +
        `AssetCategoryRecord/Id,AssetCategoryRecord/Title,AssetCategoryRecord/OrderRecord` +
        `&$expand=AssetCategoryRecord,AssetDirector,AssetDirectorReplacer,AssetManager,HSEPartner,HSEDirectorReplacer,HSEDirector`;

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
            assetDirector: spHelpers.toPersonaArray(obj.AssetDirector),
            assetManager: spHelpers.toPersonaArray(obj.AssetManager),
            hsePartner: spHelpers.toPersonaArray(obj.HSEPartner),
            assetDirectorReplacer: spHelpers.toPersonaArray(obj.AssetDirectorReplacer),
            hseDirector: spHelpers.toPersonaArray(obj.HSEDirector),
            hseDirectorReplacer: spHelpers.toPersonaArray(obj.HSEDirectorReplacer),
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
        assetDirector: spHelpers.toPersonaArray(obj.AssetDirector),
        assetManager: spHelpers.toPersonaArray(obj.AssetManager),
        hsePartner: spHelpers.toPersonaArray(obj.HSEPartner),
        assetDirectorReplacer: spHelpers.toPersonaArray(obj.AssetDirectorReplacer),
        hseDirector: spHelpers.toPersonaArray(obj.HSEDirector),
        hseDirectorReplacer: spHelpers.toPersonaArray(obj.HSEDirectorReplacer),
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
      await _getCompany();
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
    setPiHsePartnerFilteredByCategory([]);
    setAssetDirFilteredByCategory([]);
    setAssetManagerFilteredByCategory([]);
    setPiPicker([]);
  };

  // Handle asset details change: always update lists; only set selections if not suppressed or currently empty
  const onAssetDetailsChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption | undefined): void => {
    const selectedId = item ? Number(item.key) : undefined;
    setSelectedAssetDetails(selectedId);
    if (selectedId) {
      const detail = (assetCategoriesDetailsList || []).find(d => Number(d.id) === selectedId);

      setAssetDirFilteredByCategory(detail?.assetDirector || []);
      setAssetManagerFilteredByCategory(detail?.assetManager || []);
      setPiHsePartnerFilteredByCategory(detail?.hsePartner || []);

      // Apply selections safely (won't overwrite user-cleared values while suppressed)
      safeSetPicker(_assetDirPicker, setAssetDirPicker, detail?.assetDirector);
      safeSetPicker(_assetDirReplacerPicker, setAssetDirReplacerPicker, detail?.assetDirectorReplacer);
      safeSetPicker(_hseDirPicker, setHseDirPicker, detail?.hseDirector);
      safeSetPicker(_hseDirReplacerPicker, setHseDirReplacerPicker, detail?.hseDirectorReplacer);
      safeSetPicker(_closureAssetManagerPicker, setClosureAssetManagerPicker, detail?.assetManager);

      setPiStatus('Pending');
      setAssetDirStatus('Pending');
      setClosureAssetManagerStatus('Pending');
      setHseDirStatus('Pending');
    } else {
      setPiHsePartnerFilteredByCategory([]);
      setAssetDirFilteredByCategory([]);
      setAssetManagerFilteredByCategory([]);
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
      .map(m => ({
        key: m.id, text: m.title,
        // selected: _selectedMachineryIds?.includes(Number(m.id))
      }));
  }, [ptwFormStructure?.machinaries]);


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
      // selected: _selectedPersonnelIds?.includes(Number(p.Id))
    }));
  }, [personnelInvolved]);

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
      if (!isSubmitted) {
        // Allow full reset only before submission
        setSelectedPermitTypeList([]);
        setPermitPayload([]);
        setPermitPayloadValidityDays(0);
      } else {
        // Block clearing all after submission
        showBanner('At least one Work Category must remain selected after submission.', {
          kind: 'warning', autoHideMs: 4000, fade: true
        });
      }
      return;
    }

    setPTWFormStructure(prev => {
      const existing = prev.workCategories || [];
      const beforeSelected = existing.filter(c => c.isChecked);

      // Attempting to uncheck the last remaining selected category after submission?
      const isTryingToUncheckLast =
        isSubmitted &&
        !checked &&
        beforeSelected.length === 1 &&
        beforeSelected[0].id === workCategory.id;

      if (isTryingToUncheckLast) {
        showBanner('At least one Work Category must remain selected after submission.', {
          kind: 'warning', autoHideMs: 4000, fade: true
        });
      }

      const nextWorkCategories: IWorkCategory[] = (prev.workCategories || []).map(cat =>
        cat.id === workCategory?.id ? { ...cat, isChecked: !!checked } : cat
      );

      // Compute selected list after this toggle
      const selectedItems = nextWorkCategories.filter(cat => cat.isChecked);
      setSelectedPermitTypeList(selectedItems);
      setWorkPermitRequired(selectedItems.length > 0);

      if (selectedItems.length === 0) {
        setFilteredSafeguards([]);
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
        setPermitPayloadValidityDays(0);
      } else {
        const selectedIds = new Set(selectedItems.map(s => s.id));
        setFilteredSafeguards((safeguards || []).filter(s => s.workCategoryId !== undefined && selectedIds.has(s.workCategoryId)));

        // Minimum number of renewals among selected categories
        const minRenewals = Math.min(...selectedItems.map(cat => (cat.renewalValidity ?? 0)));
        setPermitPayloadValidityDays(minRenewals);

        // Preserve any existing row values when possible
        const existingById = new Map(_permitPayload.map(r => [r.id, r] as const));
        const baseRows: IPermitScheduleRow[] = [];
        // Always include the New Permit row
        baseRows.push(
          existingById.get('permit-row-0') ?? {
            id: 'permit-row-0',
            type: 'new',
            date: '',
            startTime: '',
            endTime: '',
            isChecked: false,
            orderRecord: 1,
            statusRecord: undefined,
            piApprover: undefined,
            piApproverList: _piHsePartnerFilteredByCategory,
            piApprovalDate: undefined,
            piStatus: undefined
          }
        );

        if (!isSubmitted || (_permitPayload.length === 0 && isSubmitted)) {
          setPermitPayload(baseRows);
        }
      }

      return { ...prev, workCategories: nextWorkCategories } as IPTWForm;
    });
  }, [_permitPayload, safeguards, isSubmitted, _permitPayloadValidityDays, showBanner, _piHsePartnerFilteredByCategory]);

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

    setPermitPayload((prevItems) => {
      // TODO: UnComment to enable date ordering validation
      // Helper to compare date-only in UTC
      const toDayUtc = (iso?: string): number => {
        if (!iso) return NaN;
        const d = new Date(iso);
        if (isNaN(d.getTime())) return NaN;
        return Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
      };

      return prevItems.map(item => {
        if (item.id !== rowId) return item;

        if (field === 'piStatus') {
          const status = String(value || '').trim();
          const approvalDate = undefined;

          const next: IPermitScheduleRow = {
            ...item,
            piStatus: status,
            piApprovalDate: approvalDate,
            isChecked: !!checked
          };
          return next;
        }

        // Special: selecting approver from combo -> set piApprover
        if (field === 'piApproverList') {
          const selectedPersona = (item.piApproverList || []).find(p => String(p.id) === value);
          if (selectedPersona) {
            setPiPicker([selectedPersona]);
          } else {
            setPiPicker([]);
          }
          return { ...item, piApprover: selectedPersona, isChecked: !!checked };
        }

        // Block invalid date chronologically (must be strictly after any previous selected dates)
        // TODO: UnComment to enable date ordering validation
        if (field === 'date') {
          const newDay = toDayUtc(value);

          // Max date among rows with smaller orderRecord (previous permits)
          const currentOrder = Number(item.orderRecord || 0);
          const maxPrevDay = prevItems
            .filter(r =>
              r.id !== item.id &&
              Number(r.orderRecord || 0) < currentOrder &&
              !!r.date
            )
            .map(r => toDayUtc(r.date))
            .filter(n => !isNaN(n))
            .reduce((m, n) => Math.max(m, n), Number.NEGATIVE_INFINITY);

          if (!isNaN(newDay) && maxPrevDay !== Number.NEGATIVE_INFINITY && newDay <= maxPrevDay) {
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
          return { ...next, date: '', startTime: '', endTime: '', piApprover: undefined, piApprovalDate: undefined, piStatus: undefined };
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

  const mergeRiskRows = (prev?: IRiskTaskRow[], next?: IRiskTaskRow[]): IRiskTaskRow[] => {
    if (!next || next.length === 0) return prev ? prev.slice() : [];
    if (!prev || prev.length === 0) return next.slice();

    const byId = new Map(prev.map(r => [r.id, r]));
    const uniq = (arr: (string | undefined | null)[]) =>
      Array.from(new Set(arr.map(s => (s ?? '').trim()).filter(Boolean))) as string[];

    return next.map(n => {
      const p = byId.get(n.id);
      const merged: IRiskTaskRow = { ...p, ...n };

      const nextCustom = Array.isArray(n.customSafeguards)
        ? n.customSafeguards
        : (p?.customSafeguards ?? []);
      merged.customSafeguards = uniq(nextCustom);

      return merged;
    });
  };

  const handleRiskTasksChange = React.useCallback((tasks?: IRiskAssessmentResult) => {
    if (!tasks) {
      setRiskAssessmentsTasks(undefined);
      setRiskAssessmentReferenceNumber('');
      setOverAllRiskAssessment('');
      setDetailedRiskAssessment(false);
      return;
    }

    setRiskAssessmentsTasks(prev => mergeRiskRows(prev, tasks?.rows || []));
  }, []);

  const handleOverallRiskChange = React.useCallback((riskKey?: string | number) => {
    if (riskKey == null || String(riskKey).trim() === '') return;
    const value = String(riskKey).trim();
    setOverAllRiskAssessment(value);

    if (isPermitIssuer && value.toLowerCase() === "high" && _isUrgentSubmission) {
      setAssetDirStatus('Pending');
    }
  }, [_isUrgentSubmission, isPermitIssuer]);

  const handleDetailedRiskChange = React.useCallback((required: boolean) => {
    setDetailedRiskAssessment(required);
    if (!required) {
      setRiskAssessmentReferenceNumber(''); // clear ref if no detailed assessment
    }
  }, []);

  // ADD: Detailed risk reference handler
  const handleDetailedRiskRefChange = React.useCallback((ref?: string) => {
    setRiskAssessmentReferenceNumber(ref?.trim() || '');
  }, []);

  const isUniquePermitOriginator = React.useMemo((): boolean => {
    const pOEMail = (_PermitOriginator?.[0]?.secondaryText || '').toLowerCase();
    const pIEMail = (_piPicker?.[0]?.secondaryText || '').toLowerCase();
    const assetDirectorEMail = (_assetDirPicker?.[0]?.secondaryText || '').toLowerCase();
    const hSEDirectorEMail = (_hseDirPicker?.[0]?.secondaryText || '').toLowerCase();
    const assetManagerEMail = (_closureAssetManagerPicker?.[0]?.secondaryText || '').toLowerCase();
    const loggedInUserIsPOEMail = currentUserEmail.toLowerCase() === pOEMail;
    const isUniquePO = loggedInUserIsPOEMail && (pOEMail !== (pIEMail || assetDirectorEMail || hSEDirectorEMail || assetManagerEMail));

    if (isPermitOriginator && isUniquePO) {
      return true;
    }
    return false;
  }, [_PermitOriginator, _piPicker, _assetDirPicker, _assetDirPicker, _hseDirPicker, _closureAssetManagerPicker, currentUserEmail, isPermitOriginator]);

  const isUniquePermitIssuer = React.useMemo((): boolean => {
    const pOEMail = (_PermitOriginator?.[0]?.secondaryText || '').toLowerCase();
    const pAEMail = (_paPicker?.[0]?.secondaryText || '').toLowerCase();
    const pIEMail = (_piPicker?.[0]?.secondaryText || '').toLowerCase();
    const assetDirectorEMail = (_assetDirPicker?.[0]?.secondaryText || '').toLowerCase();
    const hSEDirectorEMail = (_hseDirPicker?.[0]?.secondaryText || '').toLowerCase();
    const assetManagerEMail = (_closureAssetManagerPicker?.[0]?.secondaryText || '').toLowerCase();

    const loggedInUserIsPIEMail = currentUserEmail.toLowerCase() === pIEMail.toLowerCase();

    const isUniquePI = (loggedInUserIsPIEMail && (pIEMail !== (pOEMail || pAEMail || assetDirectorEMail || hSEDirectorEMail || assetManagerEMail)));

    if (isPermitIssuer && isUniquePI) {
      return true;
    }
    return false;
  }, [isPermitIssuer, _PermitOriginator, _paPicker, _piPicker, _assetDirPicker, _hseDirPicker, _closureAssetManagerPicker, currentUserEmail]);

  const isUniqueHSEDirector = React.useMemo((): boolean => {
    const pOEMail = (_PermitOriginator?.[0]?.secondaryText || '').toLowerCase();
    const pAEMail = (_paPicker?.[0]?.secondaryText || '').toLowerCase();
    const pIEMail = (_piPicker?.[0]?.secondaryText || '').toLowerCase();
    const assetDirectorEMail = (_assetDirPicker?.[0]?.secondaryText || '').toLowerCase();
    const hSEDirectorEMail = (_hseDirPicker?.[0]?.secondaryText || '').toLowerCase();
    const assetManagerEMail = (_closureAssetManagerPicker?.[0]?.secondaryText || '').toLowerCase();

    const loggedInUserIshseEMail = currentUserEmail.toLowerCase() === hSEDirectorEMail.toLowerCase();

    const isUniquehse = (loggedInUserIshseEMail && (hSEDirectorEMail !== (pOEMail || pIEMail || pAEMail || assetDirectorEMail || assetManagerEMail)));

    if (isHSEDirector && isUniquehse) {
      return true;
    }
    return false;
  }, [isHSEDirector, _PermitOriginator, _paPicker, _piPicker, _assetDirPicker, _hseDirPicker, _closureAssetManagerPicker, currentUserEmail]);

  const isUniqueAssetDirector = React.useMemo((): boolean => {
    const pOEMail = (_PermitOriginator?.[0]?.secondaryText || '').toLowerCase();
    const pAEMail = (_paPicker?.[0]?.secondaryText || '').toLowerCase();
    const pIEMail = (_piPicker?.[0]?.secondaryText || '').toLowerCase();
    const assetDirectorEMail = (_assetDirPicker?.[0]?.secondaryText || '').toLowerCase();
    const hSEDirectorEMail = (_hseDirPicker?.[0]?.secondaryText || '').toLowerCase();
    const assetManagerEMail = (_closureAssetManagerPicker?.[0]?.secondaryText || '').toLowerCase();

    const loggedInUserIshseEMail = currentUserEmail.toLowerCase() === assetDirectorEMail.toLowerCase();

    const isUniqueAssetDir = (loggedInUserIshseEMail && (assetDirectorEMail !== (pOEMail || pIEMail || pAEMail || hSEDirectorEMail || assetManagerEMail)));

    if (isAssetDirector && isUniqueAssetDir) {
      return true;
    }
    return false;
  }, [isAssetDirector, _PermitOriginator, _paPicker, _piPicker, _assetDirPicker, _hseDirPicker, _closureAssetManagerPicker, currentUserEmail]);

  const isUniqueAssetManager = React.useMemo((): boolean => {
    const pOEMail = (_PermitOriginator?.[0]?.secondaryText || '').toLowerCase();
    const pAEMail = (_paPicker?.[0]?.secondaryText || '').toLowerCase();
    const pIEMail = (_piPicker?.[0]?.secondaryText || '').toLowerCase();
    const assetDirectorEMail = (_assetDirPicker?.[0]?.secondaryText || '').toLowerCase();
    const hSEDirectorEMail = (_hseDirPicker?.[0]?.secondaryText || '').toLowerCase();
    const assetManagerEMail = (_closureAssetManagerPicker?.[0]?.secondaryText || '').toLowerCase();

    const loggedInUserIsAssetManagerEMail = currentUserEmail.toLowerCase() === assetManagerEMail.toLowerCase();

    const isUniqueAssetManager = (loggedInUserIsAssetManagerEMail && (assetManagerEMail !== (pOEMail || pIEMail || pAEMail || hSEDirectorEMail || assetDirectorEMail)));

    if (isAssetManager && isUniqueAssetManager) {
      return true;
    }
    return false;
  }, [isAssetManager, _PermitOriginator, _paPicker, _piPicker, _assetDirPicker, _hseDirPicker, _closureAssetManagerPicker, currentUserEmail]);

  // Minimal payload builder (adjust to your save schema)
  const buildPayload = React.useCallback(() => {
    return {
      reference: _coralReferenceNumber,
      previousReferenceNumber: _previousPtwRef || '',
      assetId: _assetId,
      assetCategoryId: _selectedAssetCategory,
      assetDetailsId: _selectedAssetDetails,
      company: _selectedCompany,
      projectTitle: _projectTitle,
      permitTypes: _selectedPermitTypeList?.map(x => x.id),
      permitRows: _permitPayload,
      permitPayloadValidityDays: _permitPayloadValidityDays,
      hacWorkAreaId: _selectedHacWorkAreaId,
      workHazardIds: Array.from(_selectedWorkHazardIds || []),
      workHazardsOtherText: _workHazardsOtherText,
      workTaskLists: _riskAssessmentsTasks || [],
      overallRiskAssessment: _overAllRiskAssessment || '',
      detailedRiskAssessment: _detailedRiskAssessment,
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
      originatorId: _PermitOriginator?.[0]?.id || '',
      originatorEMail: _PermitOriginator?.[0]?.secondaryText || '',
      toolboxTalk: _selectedToolboxTalk !== undefined && _selectedToolboxTalk ? true : false,
      toolboxTalkDate: _selectedToolboxTalkDate,
      toolboxTalkConductedById: _selectedToolboxConductedBy?.[0]?.secondaryText || '',
      toolboxHSEReference: _toolboxHSEReference || '',
      poApprovalDate: _poDate || '',
      poStatus: _poStatus || '',

      paPickerId: _paPicker?.[0]?.id || '',
      paApprovalDate: _paDate || '',
      paStatus: _paStatus || '',
      paRejectionReason: _paRejectionReason || '',

      piPickerId: _piPicker?.[0]?.id || '',
      piApprovalDate: _piDate || '',
      piStatus: _piStatus || '',
      piRejectionReason: _piRejectionReason || '',

      assetDirPickerId: _assetDirPicker?.[0]?.id || '',
      assetDirReplacerPickerId: _assetDirReplacerPicker?.[0]?.id || '',
      assetDirApprovalDate: _assetDirDate || '',
      assetDirStatus: _assetDirStatus || '',
      assetDirRejectionReason: _assetDirRejectionReason || '',

      urgentAssetDirDate: _urgentAssetDirDate || '',
      urgentAssetDirStatus: _urgentAssetDirStatus || '',
      urgentAssetDirRejectionReas: _urgentAssetDirRejectionReas || '',

      hseDirPickerId: _hseDirPicker?.[0]?.id || '',
      hseDirReplacerPickerId: _hseDirReplacerPicker?.[0]?.id || '',
      hseDirApprovalDate: _hseDirDate || '',
      hseDirStatus: _hseDirStatus || '',
      hseDirRejectionReason: _hseDirRejectionReason || '',

      closurePOPickerId: _PermitOriginator?.[0]?.id || '',
      closurePOApprovalDate: _closurePoDate || '',
      closurePOStatus: _closurePoStatus || '',
      closurePORejectionReason: _poRejectionReason || '',

      closureAssetManagerPickerId: _closureAssetManagerPicker?.[0]?.id || '',
      closureAssetManagerApprovalDate: _closureAssetManagerDate || '',
      closureAssetManagerStatus: _closureAssetManagerStatus || '',
      assetManagerRejectionReason: _asssetManagerRejectionReason || '',

      isAssetDirectorReplacer: _isAssetDirReplacer,
      isHSEDirectorReplacer: _isHseDirReplacer,
      isUrgentSubmission: _isUrgentSubmission,
    };
  }, [_coralReferenceNumber, _assetId, _selectedAssetCategory, _selectedAssetDetails, _projectTitle,
    _selectedPermitTypeList, _permitPayload, _selectedHacWorkAreaId,
    _selectedWorkHazardIds, _selectedPrecautionIds, _selectedProtectiveEquipmentIds,
    _gasTestValue, _gasTestResult, _fireWatchValue, _fireWatchAssigned, _protectiveEquipmentsOtherText, _precautionsOtherText,
    _attachmentsValue, _attachmentsResult, _selectedMachineryIds, _selectedPersonnelIds, _PermitOriginator,
    _workHazardsOtherText, _riskAssessmentsTasks, _riskAssessmentReferenceNumber, _overAllRiskAssessment, _detailedRiskAssessment,
    _poDate, _poStatus, _paPicker, _paDate, _paStatus, _piPicker, _piDate, _piStatus,
    _assetDirPicker, _assetDirDate, _assetDirStatus,
    _hseDirPicker, _hseDirDate, _hseDirStatus, _toolboxHSEReference, _selectedToolboxConductedBy, _selectedToolboxTalkDate, _selectedToolboxTalk,
    _closureAssetManagerPicker, _closureAssetManagerDate, _closureAssetManagerStatus,
    _closurePoDate, _closurePoStatus, _isUrgentSubmission, _previousPtwRef, _paRejectionReason, _piRejectionReason,
    _assetDirRejectionReason, _hseDirRejectionReason, _isAssetDirReplacer, _isHseDirReplacer, _permitPayloadValidityDays, _urgentAssetDirDate, _urgentAssetDirStatus,
    _urgentAssetDirRejectionReas, _poRejectionReason, _asssetManagerRejectionReason, _hseDirReplacerPicker, _assetDirReplacerPicker
  ]);

  const validateBeforeSubmit = React.useCallback((originatorId: number | undefined, mode: 'save' | 'submit' | 'approve' | 'approveWithoutUpdate'): string | undefined => {
    const missing: string[] = [];
    const payload = buildPayload();
    payloadRef.current = payload;

    if (!payload.originatorId.trim()) {
      missing.push('Permit Originator');
      return `Please fill in the required fields: ${missing.join(', ')}.`;
    };

    if ((mode === 'submit')) {
      if (!payload?.assetId?.trim()) missing.push('Asset ID');
      if (!payload.assetCategoryId?.toString().trim()) missing.push('Asset Category');
      if (!payload.assetDetailsId?.toString().trim()) missing.push('Asset Details');
      if (!payload.projectTitle?.trim()) missing.push('Project Title');
      if (!payload.company?.id?.toString().trim()) missing.push('Company');
      if (!payload.permitTypes || payload.permitTypes.length === 0) missing.push('At least one Permit Type');
      if (!payload.permitRows || payload.permitRows.length === 0) {
        missing.push('At least one Permit Row in Permit Schedule');
      } else {
        const selectedNewPermitRows = payload.permitRows.filter(r => r.isChecked && r.type.toLowerCase() === 'new');

        if (selectedNewPermitRows.length >= 1) {
          const newRowDateIso = selectedNewPermitRows[0].date;
          const permitStartTime = selectedNewPermitRows[0].startTime;
          const permitApprover = selectedNewPermitRows[0].piApprover;

          if (!newRowDateIso) {
            missing.push('New Permit Row Date');
          } else if (!permitStartTime) {
            missing.push('New Permit Row Start Time');
          } else if (!permitApprover) {
            missing.push('New Permit Row Permit Issuer');
          } else {
            const startDateTimeIso = spHelpers.combineDateAndTime(newRowDateIso, permitStartTime)?.toISOString();
            if (!startDateTimeIso) return `Please fill in the required fields: Invalid New Permit Row Date/Time.`;
            const permitDate = new Date(startDateTimeIso);

            if (isNaN(permitDate.getTime())) {
              missing.push('New Permit Row Date (invalid)');
            }
            else if (!payload.isUrgentSubmission) {
              const now = new Date();
              // Interpret SubmissionRangeInterval as hours (default 24)
              const intervalHours = Number(_coralFormList?.SubmissionRangeInterval) || 24;
              const diffMs = permitDate.getTime() - now.getTime();
              const diffHours = diffMs / (1000 * 60 * 60);
              const meetsInterval = diffHours >= intervalHours;
              if (!meetsInterval) {
                missing.push(`New Permit Row start must be at least ${intervalHours} hours after the current submission date/time.`);
              }
            }

          }
        }
        else if (selectedNewPermitRows.length === 0) {
          missing.push('At least one Permit Row in Permit Schedule required for New Permit.');
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
          const missingTaskRows = rows.map((row, idx) => ({ idx, hasTask: !!String(row?.task || '').trim() }))
            .filter(x => !x.hasTask)
            .map(x => x.idx + 1);
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
      const isAttachmentYes = String(payload.attachmentsProvided || '').toLowerCase() === 'yes';
      if (isAttachmentYes && !String(payload.attachmentsDetails || '').trim()) {
        missing.push('Attachment(s) details');
      }

      // NEW: Ensure at least one Protective & Safety Equipment selected
      if (!payload.protectiveEquipmentIds || payload.protectiveEquipmentIds.length === 0) {
        missing.push('At least one Protective & Safety Equipment');
      }

      // NEW: Ensure at least one Machinery/Tool selected
      if (!payload.machineryIds || payload.machineryIds.length === 0) {
        missing.push('At least one Machinery/Tool');
      }

      // NEW: Ensure at least one Personnel Involved selected
      // if (!payload.personnelIds || payload.personnelIds.length === 0) {
      //   missing.push('At least one Personnel Involved');
      // }

      if (!payload.isUrgentSubmission) {
        if (isPermitOriginator && !payload.paPickerId?.toString().trim()) missing.push('Performing Authority');
        if (isPermitOriginator && originatorId?.toString() == _paPicker?.[0]?.id && payload.paStatus.toLowerCase() == 'approved' && !_piPicker?.[0]?.id) missing.push('Permit Issuer');
      }

      if (missing.length) {
        return `Please fill in the required fields: ${missing.join(', ')}.`;
      }

    }
    return undefined;
  }, [buildPayload, ptwFormStructure?.workHazardosList, _isUrgentSubmission, _coralFormList]);

  const validateBeforeApprove = React.useCallback((mode: 'approve' | 'issuePermit' | 'approveRenewalPermit'): string | undefined => {
    const missing: string[] = [];
    const payload = buildPayload();
    payloadRef.current = payload;

    if (!payload) missing.push('Form data is missing. Please reload the page.');

    if (mode === 'approve') {
      if (isPerformingAuthority && !isIssued) {
        if (payload.paStatus === 'Pending') missing.push('Approval/Rejection Status.');
        if (payload.paStatus === 'Rejected' && (!payload.paPickerId || String(payload.paPickerId).trim() === '')) missing.push('Performing Authority');
        if (payload.paStatus === 'Approved' && (!payload.piPickerId || String(payload.piPickerId).trim() === '')) missing.push('Permit Issuer');
        if (payload.paStatus === 'Rejected' && !String(payload.paRejectionReason || '').trim()) missing.push('PA Rejection Reason');
      }

      if (isUniquePermitIssuer && !isIssued) {
        if (payload.piStatus === 'Pending') missing.push('Approval/Rejection Status.');
        if (payload.piStatus === 'Rejected' && (!payload.piPickerId || String(payload.piPickerId).trim() === '')) missing.push('Permit Issuer');
        // if (payload.piStatus === 'Approved' && (!payload.assetDirPickerId || String(payload.assetDirPickerId).trim() === '')) missing.push('Asset Director');
        if (payload.piStatus === 'Rejected' && !String(payload.piRejectionReason || '').trim()) missing.push('PI Rejection Reason');
        // Tasks required when 3 + hazards: list rows missing a task
        const hazardsCount = Array.isArray(payload.workHazardIds) ? payload.workHazardIds.length : 0;
        if (hazardsCount >= 3) {
          const rows = Array.isArray(payload.workTaskLists) ? payload.workTaskLists : [];
          if (rows.length >= 1) {
            // Initial Risk required per row
            const missingInitialRiskRows = rows
              .map((row: any, idx: number) => ({ idx, ok: !!String(row?.initialRisk || '').trim() }))
              .filter((x: { ok: any; }) => !x.ok)
              .map((x: { idx: number; }) => x.idx + 1);

            if (missingInitialRiskRows.length) {
              missing.push(`Initial Risk missing for row(s): ${missingInitialRiskRows.join(', ')}`);
            }

            const missingResidualRiskRows = rows
              .map((row: any, idx: number) => ({ idx, ok: !!String(row?.residualRisk || '').trim() }))
              .filter((x: { ok: any; }) => !x.ok)
              .map((x: { idx: number; }) => x.idx + 1);

            if (missingResidualRiskRows.length) {
              missing.push(`Residual Risk missing for row(s): ${missingResidualRiskRows.join(', ')}`);
            }
          }
        }

        // Overall Risk Assessment required
        if (!String(payload.overallRiskAssessment || '').trim()) {
          missing.push('Overall Risk Assessment');
        }

        // If Detailed (L2) is checked, require reference number
        const l2Required = !!payload.detailedRiskAssessment;
        if (l2Required && !String(payload.detailedRiskAssessmentRef || '').trim()) {
          missing.push('Risk Assessment Ref Number (Detailed L2)');
        }

        // NEW: Gas Test / Fire Watch / Attachments validations for renewal approval
        const gasYes = String(payload.gasTestRequired || '').toLowerCase() === 'yes';
        if (gasYes && !String(payload.gasTestResult || '').trim()) {
          missing.push('Gas Test Result');
        }

        const fireWatchYes = String(payload.fireWatchNeeded || '').toLowerCase() === 'yes';
        if (fireWatchYes && !String(payload.fireWatchAssigned || '').trim()) {
          missing.push('Firewatch Assigned');
        }

        const attachmentsYes = String(payload.attachmentsProvided || '').toLowerCase() === 'yes';
        if (attachmentsYes && !String(payload.attachmentsDetails || '').trim()) {
          missing.push('Attachment(s) Details');
        }

        // NEW: Toolbox Talk validations (if checked, all fields required)
        if (payload.toolboxTalk) {
          if (!String(payload.toolboxTalkConductedById || '').trim()) {
            missing.push('Toolbox Talk - Conducted By');
          }
          if (!String(payload.toolboxHSEReference || '').trim()) {
            missing.push('HSE TBT Reference');
          }
          const dt = payload.toolboxTalkDate instanceof Date
            ? payload.toolboxTalkDate
            : (payload.toolboxTalkDate ? new Date(payload.toolboxTalkDate) : undefined);
          if (!dt || isNaN(dt.getTime())) {
            missing.push('Toolbox Talk Date');
          }
        }
      }

      if (isUniqueAssetDirector && (isIssued || payload.isUrgentSubmission)) {
        if (payload.isUrgentSubmission) {
          if (payload.urgentAssetDirStatus === 'Pending') missing.push('Approval/Rejection Status.');
          if (payload.urgentAssetDirStatus === 'Rejected' && (!payload.assetDirPickerId || String(payload.assetDirPickerId).trim() === '')) missing.push('Asset Director');
          if (payload.urgentAssetDirStatus === 'Rejected' && !String(payload.urgentAssetDirRejectionReas || '').trim()) missing.push('Asset Director Rejection Reason');
        }
        else {
          if (payload.assetDirStatus === 'Pending') missing.push('Approval/Rejection Status.');
          if (payload.assetDirStatus === 'Rejected' && (!payload.assetDirPickerId || String(payload.assetDirPickerId).trim() === '')) missing.push('Asset Director');
          if (payload.assetDirStatus === 'Rejected' && !String(payload.assetDirRejectionReason || '').trim()) missing.push('Asset Director Rejection Reason');
        }
      }

      if (isUniqueHSEDirector && (isIssued || payload.isUrgentSubmission) && (_workflowStage?.toLowerCase() !== 'ApprovedFromAssetToHSE'.toLowerCase())) {
        if (payload.hseDirStatus === 'Pending') missing.push('Approval/Rejection Status.');
        if (payload.hseDirStatus === 'Rejected' && (!payload.hseDirPickerId || String(payload.hseDirPickerId).trim() === '')) missing.push('HSE Director');
        if (payload.hseDirStatus === 'Rejected' && !String(payload.hseDirRejectionReason || '').trim()) missing.push('HSE Director Rejection Reason');
      }

      if (isUniquePermitOriginator && _workflowStage?.toLowerCase() === 'Issued'.toLowerCase()) {
        if (payload.closurePOStatus === 'Pending') missing.push('Approval/Rejection Status.');
        if (payload.closurePOStatus === 'Rejected' && !String(payload.closurePORejectionReason || '').trim()) missing.push('Your Rejection Reason');
      }

      if (isUniqueAssetManager && _workflowStage?.toLowerCase() === 'ClosedByPO'.toLowerCase()) {
        if (payload.closureAssetManagerStatus === 'Pending') missing.push('Approval/Rejection Status.');
        if (payload.closureAssetManagerStatus === 'Rejected' && (!payload.closureAssetManagerPickerId || String(payload.closureAssetManagerPickerId).trim() === '')) missing.push('Asset Manager');
        if (payload.closureAssetManagerStatus === 'Rejected' && !String(payload.hseDirRejectionReason || '').trim()) missing.push('Asset Manager Rejection Reason');
      }
    }

    if (mode === 'approveRenewalPermit') {

      //TODO: Add validations if any on issue permit
      if (isUniquePermitIssuer) {
        // Tasks required when 3 + hazards: list rows missing a task
        const hazardsCount = Array.isArray(payload.workHazardIds) ? payload.workHazardIds.length : 0;
        if (hazardsCount >= 3) {
          const rows = Array.isArray(payload.workTaskLists) ? payload.workTaskLists : [];
          if (rows.length >= 1) {
            // Initial Risk required per row
            const missingInitialRiskRows = rows
              .map((row: any, idx: number) => ({ idx, ok: !!String(row?.initialRisk || '').trim() }))
              .filter((x: { ok: any; }) => !x.ok)
              .map((x: { idx: number; }) => x.idx + 1);

            if (missingInitialRiskRows.length) {
              missing.push(`Initial Risk missing for row(s): ${missingInitialRiskRows.join(', ')}`);
            }

            const missingResidualRiskRows = rows
              .map((row: any, idx: number) => ({ idx, ok: !!String(row?.residualRisk || '').trim() }))
              .filter((x: { ok: any; }) => !x.ok)
              .map((x: { idx: number; }) => x.idx + 1);

            if (missingResidualRiskRows.length) {
              missing.push(`Residual Risk missing for row(s): ${missingResidualRiskRows.join(', ')}`);
            }
          }

          // Overall Risk Assessment required
          if (!String(payload.overallRiskAssessment || '').trim()) {
            missing.push('Overall Risk Assessment');
          }

          // If Detailed (L2) is checked, require reference number
          const l2Required = !!payload.detailedRiskAssessment;
          if (l2Required && !String(payload.detailedRiskAssessmentRef || '').trim()) {
            missing.push('Risk Assessment Ref Number (Detailed L2)');
          }
        }

        // NEW: Gas Test / Fire Watch / Attachments validations for renewal approval
        const gasYes = String(payload.gasTestRequired || '').toLowerCase() === 'yes';
        if (gasYes && !String(payload.gasTestResult || '').trim()) {
          missing.push('Gas Test Result');
        }

        const fireWatchYes = String(payload.fireWatchNeeded || '').toLowerCase() === 'yes';
        if (fireWatchYes && !String(payload.fireWatchAssigned || '').trim()) {
          missing.push('Firewatch Assigned');
        }

        const attachmentsYes = String(payload.attachmentsProvided || '').toLowerCase() === 'yes';
        if (attachmentsYes && !String(payload.attachmentsDetails || '').trim()) {
          missing.push('Attachment(s) Details');
        }

        // NEW: Toolbox Talk validations (if checked, all fields required)
        if (payload.toolboxTalk) {
          if (!String(payload.toolboxTalkConductedById || '').trim()) {
            missing.push('Toolbox Talk - Conducted By');
          }
          if (!String(payload.toolboxHSEReference || '').trim()) {
            missing.push('HSE TBT Reference');
          }
          const dt = payload.toolboxTalkDate instanceof Date
            ? payload.toolboxTalkDate
            : (payload.toolboxTalkDate ? new Date(payload.toolboxTalkDate) : undefined);
          if (!dt || isNaN(dt.getTime())) {
            missing.push('Toolbox Talk Date');
          }
        }

        const isNumericId = (id: string) => /^[0-9]+$/.test(String(id || ''));
        let renewedPermit: IPermitScheduleRow | undefined;

        if (payload.permitRows && payload.permitRows.length) {
          // Validate there is at least one non-numeric renewal row fully filled
          renewedPermit = payload.permitRows
            .filter((r: IPermitScheduleRow) =>
              r.type === 'renewal' &&
              String(r.statusRecord || '').toLowerCase() === 'new' &&
              isNumericId(r.id) &&
              !!String(r.date || '').trim() &&
              !!String(r.startTime || '').trim() &&
              !!String(r.endTime || '').trim() &&
              !!r.piApprover).sort((a, b) => a.orderRecord - b.orderRecord)[0];

          const isApprovedOrRejected = !!String(renewedPermit.piStatus || '').trim();

          if (!isApprovedOrRejected) {
            missing.push('PI Status for the selected Permit Row to approve.');
          }
        }
      }
    }

    if (missing.length) {
      return `Please fill in the required fields: ${missing.join(', ')}.`;
    }
    return undefined;
  }, [buildPayload, isPerformingAuthority, isPermitIssuer, _assetDirPicker, isUniquePermitIssuer, isUniqueAssetManager, isUniquePermitOriginator, isUniqueHSEDirector]);

  const approveFormWWithUpdate = React.useCallback(async (mode: 'approve' | 'approveRenewalPermit'): Promise<boolean> => {
    payloadRef.current = buildPayload();
    const payload = payloadRef.current;
    const validationError = validateBeforeApprove(mode);
    if (validationError) {
      showBanner(validationError);
      return false;
    }
    else {
      setIsBusy(true);
      try {

        if (mode === 'approve') {
          // Find workflow item for this form
          const formId = Number(props.formId);
          const wfQuery = `?$select=Id&$filter=PTWForm/Id eq ${formId}`;
          const ops = new SPCrudOperations((props.context as any).spHttpClient, webUrl, 'PTW_Form_Approval_Workflow', wfQuery);
          const wfList = await ops._getItemsWithQuery();
          const wfItemId = Array.isArray(wfList) && wfList[0]?.Id;
          if (!wfItemId) throw new Error('Workflow item not found.');

          const nowIso = new Date().toISOString();
          let body = {};

          const rejectedReason = await _getPTWRejectionReason();
          if (rejectedReason && rejectedReason.trim().length > 0) {
            const ops = new SPCrudOperations((props.context as any).spHttpClient, webUrl, 'PTW_Form', '');
            const rejectedReasonUpdate = await ops._updateItem(formId.toString(), body = { RejectionReason: rejectedReason });

            if (!rejectedReasonUpdate.ok) throw new Error('Failed to update PTW Form with Rejection Reason.');
          }

          if (!payload.isUrgentSubmission) {
            if (isPerformingAuthority) {
              body = {
                PAStatus: payload.paStatus,
                PAApprovalDate: payload.paApprovalDate || nowIso,
                PIApproverId: payload.piPickerId ? Number(payload.piPickerId) : null,
                PARejectionReason: payload.paStatus === 'Rejected' ? payload.paRejectionReason : '',
              }

              const res = await ops._updateItem(String(wfItemId), body);
              if (!res.ok) throw new Error('Failed to update workflow status.');
              showBanner(`Approved Successfully.`, { autoHideMs: 3000, fade: true, kind: 'success' });
              goBackToHost();
              return true;
            }
          }

          const isHigh = String(payload.overallRiskAssessment || '').toLowerCase().includes('high');
          if (isPermitIssuer) {
            if (isHigh) {
              body = {
                PIStatus: payload.piStatus,
                PIApprovalDate: payload.piApprovalDate || nowIso,
                IsAssetDirectorReplacer: payload.isAssetDirectorReplacer,
                IsHSEDirectorReplacer: payload.isHSEDirectorReplacer,
                PIRejectionReason: payload.piStatus === 'Rejected' ? payload.piRejectionReason : '',
              }
            } else {
              // NOT High risk
              body = {
                PIStatus: payload.piStatus,
                PIApprovalDate: payload.piApprovalDate || nowIso,
                PIRejectionReason: payload.piStatus === 'Rejected' ? payload.piRejectionReason : '',
              }
            }

            const res1 = await _updatePTWForm(formId, 'approve');
            if (!res1) throw new Error('Failed to update PTW form.');

            const res = await ops._updateItem(String(wfItemId), body);
            if (!res.ok) throw new Error('Failed to update workflow status.');

            if (res.ok) {
              if (!isHigh) {
                const issueResult = await issuePermit('issuePermit');
                if (!issueResult) throw new Error('Failed to issue permit.');
              }
            }

            showBanner(`Approved Successfully.`, { autoHideMs: 3000, fade: true, kind: 'success' });
            goBackToHost();
            return true;
          }

          if (isAssetDirector && isHigh) {
            body = {
              AssetDirectorStatus: payload.assetDirStatus,
              AssetDirectorApprovalDate: payload.assetDirStatus !== 'Pending' ? nowIso : '',
              AssetDirectorRejectionReason: payload.assetDirStatus === 'Rejected' ? payload.assetDirRejectionReason : '',
            }
            const res1 = await _updatePTWForm(formId, 'approve');
            if (!res1) throw new Error('Failed to update PTW form.');

            const res = await ops._updateItem(String(wfItemId), body);
            if (!res.ok) throw new Error('Failed to update workflow status.');

            showBanner(`Approved Successfully.`, { autoHideMs: 3000, fade: true, kind: 'success' });
            goBackToHost();
            return true;
          }

          if (isAssetDirector && payload.isUrgentSubmission) {
            body = {
              UrgentAssetDirectorStatus: payload.urgentAssetDirStatus,
              UrgentAssetDirectorApprovalDate: payload.urgentAssetDirDate || nowIso,
              UrgentAssetDirectorRejectionReas: payload.urgentAssetDirStatus === 'Rejected' ? payload.urgentAssetDirRejectionReas : '',
            }

            const res = await ops._updateItem(String(wfItemId), body);
            if (!res.ok) throw new Error('Failed to update workflow status.');

            const updated = await _updatePTWForm(formId, 'approve');
            if (!updated) throw new Error('Failed to update PTW form.');

            showBanner(`Approved Successfully.`, { autoHideMs: 3000, fade: true, kind: 'success' });
            goBackToHost();
            return true;
          }

          if (isHSEDirector) {
            body = {
              HSEDirectorStatus: payload.hseDirStatus,
              HSEDirectorApprovalDate: payload.hseDirStatus !== 'Pending' ? nowIso : '',
              HSEDirectorRejectionReason: payload.hseDirStatus === 'Rejected' ? payload.hseDirRejectionReason : '',
            }

            const res1 = await _updatePTWForm(formId, 'approve');
            if (!res1) throw new Error('Failed to update PTW form.');

            const res = await ops._updateItem(String(wfItemId), body);
            if (!res.ok) throw new Error('Failed to update workflow status.');

            showBanner(`Approved Successfully.`, { autoHideMs: 3000, fade: true, kind: 'success' });
            goBackToHost();
            return true;
          }
        }
      }
      catch (e) {
        showBanner('Failed to approve. Please try again.', { autoHideMs: 5000, fade: true, kind: 'error' });
      } finally {
        setIsBusy(false);
      }
      return true;
    }
  }, [validateBeforeApprove, buildPayload, isPerformingAuthority, isPermitIssuer, _paStatus, _piStatus, props.formId, webUrl, props.context.spHttpClient]);

  const _getPTWRejectionReason = React.useCallback(async (): Promise<string> => {
    const payload = buildPayload();
    payloadRef.current = payload;

    // Normalize possible field names used in different branches
    const candidates = [
      payload.paRejectionReason,
      payload.piRejectionReason,
      payload.assetDirRejectionReason,
      payload.hseDirRejectionReason,
      payload.closurePORejectionReason,
      payload.assetManagerRejectionReason,
      payload.urgentAssetDirRejectionReas
    ].filter((v): v is string => typeof v === "string" && v.trim().length > 0);

    return candidates[0] ?? "";
  }, [buildPayload]);

  const approveForm = React.useCallback(async (mode: 'approve'): Promise<boolean> => {
    payloadRef.current = buildPayload();
    const payload = payloadRef.current;
    const validationError = validateBeforeApprove(mode);
    if (validationError) {
      showBanner(validationError);
      return false;
    }
    else {
      setIsBusy(true);
      try {
        // Find workflow item for this form
        const formId = Number(props.formId);
        const wfQuery = `?$select=Id&$filter=PTWForm/Id eq ${formId}`;
        const ops = new SPCrudOperations((props.context as any).spHttpClient, webUrl, 'PTW_Form_Approval_Workflow', wfQuery);
        const wfList = await ops._getItemsWithQuery();
        const wfItemId = Array.isArray(wfList) && wfList[0]?.Id;
        if (!wfItemId) throw new Error('Workflow item not found.');

        const nowIso = new Date().toISOString();
        let body = {};

        if (isAssetDirector) {
          if (payload.isUrgentSubmission) {
            body = {
              UrgentAssetDirectorStatus: payload.urgentAssetDirStatus,
              UrgentAssetDirectorApprovalDate: payload.urgentAssetDirStatus !== 'Pending' ? payload.urgentAssetDirectorApprovalDate : nowIso,
              UrgentAssetDirectorRejectionReas: payload.urgentAssetDirStatus === 'Rejected' ? payload.urgentAssetDirRejectionReas : null,
            }
          }

          if (payload.overallRiskAssessment.toLowerCase().includes('high')) {
            body = {
              AssetDirectorStatus: payload.assetDirStatus,
              AssetDirectorApprovalDate: payload.assetDirStatus !== 'Pending' ? payload.assetDirectorApprovalDate : nowIso,
              AssetDirectorRejectionReason: payload.assetDirStatus === 'Rejected' ? payload.assetDirRejectionReason : null,
            }
          }
        }

        if (isHSEDirector) {
          body = {
            HSEDirectorStatus: payload.hseDirStatus,
            HSEDirectorApprovalDate: payload.hseDirStatus !== 'Pending' ? payload.hseDirectorApprovalDate : nowIso,
          }
        }

        if (isPermitOriginator && _workflowStage?.toLowerCase() === 'Issued'.toLowerCase()) {
          body = {
            POClosureDate: payload.closurePOStatus !== 'Pending' ? nowIso : '',
            POClosureStatus: payload.closurePOStatus,
            POClosureRejectionReason: payload.closurePOStatus === 'Rejected' ? payload.closurePORejectionReason : null,
          }
        }

        if (isAssetManager && _workflowStage?.toLowerCase() === 'ClosedByPO'.toLowerCase()) {
          body = {
            AssetManagerApprovalDate: payload.closureAssetManagerStatus !== 'Pending' ? nowIso : '',
            AssetManagerStatus: payload.closureAssetManagerStatus,
            AssetManagerRejectionReason: payload.closureAssetManagerStatus === 'Rejected' ? payload.assetManagerRejectionReason : null,
          }
        }

        const res = await ops._updateItem(String(wfItemId), body);
        if (!res.ok) throw new Error('Failed to update workflow status.');
        showBanner(`Approved Successfully.`, { autoHideMs: 3000, fade: true, kind: 'success' });
        goBackToHost();
        return true;
      }
      catch (e) {
        showBanner('Failed to approve. Please try again.', { autoHideMs: 5000, fade: true, kind: 'error' });
      } finally {
        setIsBusy(false);
      }
      return true;
    }
  }, [validateBeforeApprove, buildPayload, isAssetDirector, isPermitOriginator, isAssetManager, isHSEDirector, props.formId, webUrl, props.context.spHttpClient]);

  const cancelPTW = React.useCallback(async (mode: 'cancel'): Promise<boolean> => {
    setIsBusy(true);

    try {
      const formId = Number(props.formId);
      const nowIso = new Date().toISOString();
      if (!formId || isNaN(formId)) throw new Error('Invalid form Id.');
      const ptwFormOps = new SPCrudOperations((props.context as any).spHttpClient, webUrl, 'PTW_Form', '');

      const parentBody: any = {
        WorkflowStatus: PTWWorkflowStatus.PermanentlyClosed,
        ToRenewPermit: false
      };

      const parentRes = await ptwFormOps._updateItem(String(formId), parentBody);
      if (!parentRes.ok) throw new Error('Failed to update PTW_Form.');

      const permitsQuery = `?$select=Id,StatusRecord&$filter=PTWForm/Id eq ${formId}`;
      const permitsOps = new SPCrudOperations((props.context as any).spHttpClient, webUrl, 'PTW_Form_Work_Permits', permitsQuery);
      const permits = await permitsOps._getItemsWithQuery();

      if (Array.isArray(permits) && permits.length) {
        await Promise.all(permits.map(async (p: any) => {
          if (String(p.StatusRecord || '').toLowerCase() !== 'closed') {
            await permitsOps._updateItem(String(p.Id), {
              StatusRecord: 'Closed'
            });
          }
        })
        );
      }

      // 3) Update workflow Stage to ClosedByPO in PTW_Form_Approval_Workflow
      const wfQuery = `?$select=Id&$filter=PTWForm/Id eq ${formId}`;
      const wfOps = new SPCrudOperations((props.context as any).spHttpClient, webUrl, 'PTW_Form_Approval_Workflow', wfQuery);
      const wfItems = await wfOps._getItemsWithQuery();
      const wfItemId = Array.isArray(wfItems) && wfItems[0]?.Id;

      if (wfItemId) {
        const wfRes = await wfOps._updateItem(String(wfItemId), {
          Stage: 'ClosedByPO',
          POClosureDate: nowIso,
          POClosureStatus: 'Cancelled'
        });
        if (!wfRes.ok) throw new Error('Failed to update workflow stage.');
      }

      showBanner('Cancelled Successfully.', { autoHideMs: 3000, fade: true, kind: 'success' });
      goBackToHost();
      return true;
    }
    catch (e) {
      showBanner('Failed to cancel. Please try again.', { autoHideMs: 5000, fade: true, kind: 'error' });
    } finally {
      setIsBusy(false);
    }
    return true;

  }, [buildPayload, isPermitOriginator, props.formId, webUrl, props.context.spHttpClient, _workflowStage]);

  const resubmitAfterRejection = React.useCallback(async (mode: 'submitAfterRejection'): Promise<boolean> => {
    payloadRef.current = buildPayload();
    const payload = payloadRef.current;
    const validationError = validateBeforeApprove(mode === "submitAfterRejection" ? "approve" : mode);
    if (validationError) {
      showBanner(validationError);
      return false;
    }
    setIsBusy(true);
    {
      try {
        const formId = Number(props.formId);
        if (!formId || isNaN(formId)) throw new Error('Invalid form Id.');

        const res1 = await _updatePTWForm(formId, 'submitAfterRejection');
        if (!res1) throw new Error('Failed to update PTW form.');

        const spOps = spCrudRef.current ?? new SPCrudOperations((props.context as any).spHttpClient, webUrl, '', '');
        const originatorId = await spOps.ensureUserId(payload.originatorEMail || '');

        // 3) Update workflow IsSubmittedAfterRejection to true in PTW_Form_Approval_Workflow After Resubmission
        const wfQuery = `?$select=Id&$filter=PTWForm/Id eq ${formId}`;
        const wfOps = new SPCrudOperations((props.context as any).spHttpClient, webUrl, 'PTW_Form_Approval_Workflow', wfQuery);
        const wfItems = await wfOps._getItemsWithQuery();
        const wfItemId = Array.isArray(wfItems) && wfItems[0]?.Id;

        let body = {};
        let _paStatusForPOPA = '';
        let _paDateForPOPA = new Date().toISOString();

        if (originatorId === payload.paPickerId) {
          _paStatusForPOPA = payload.paStatus === 'Approved' ? payload.paStatus : 'Approved';
          _paDateForPOPA = _paDateForPOPA;
        } else {
          _paStatusForPOPA = payload.paStatus || 'Pending';
          _paDateForPOPA = payload.paApprovalDate;
        }

        body = {
          IsSubmittedAfterRejection: true,

          StatusRecord: 'New',

          POApprovalDate: payload.poApprovalDate || null,
          POStatus: payload.poStatus || 'Approved',

          PAStatus: _paStatusForPOPA,
          PAApprovalDate: _paStatusForPOPA !== 'Pending' ? _paDateForPOPA : null,
          PARejectionReason: payload.paRejectionReason || '',

          PIApproverId: payload.piPickerId ? Number(payload.piPickerId) : undefined,
          PIStatus: 'Pending',
          PIApprovalDate: null,
          PIRejectionReason: '',

          AssetDirectorApproverId: payload.assetDirPickerId ? Number(payload.assetDirPickerId) : undefined,
          AssetDirectorStatus: 'Pending',
          AssetDirectorApprovalDate: null,
          AssetDirectorRejectionReason: '',

          UrgentAssetDirectorStatus: 'Pending',
          UrgentAssetDirectorApprovalDate: null,
          UrgentAssetDirectorRejectionReas: '',

          HSEDirectorApproverId: payload.hseDirPickerId ? Number(payload.hseDirPickerId) : undefined,
          HSEDirectorStatus: 'Pending',
          HSEDirectorApprovalDate: null,
          HSEDirectorRejectionReason: '',

          AssetManagerApproverId: payload.closureAssetManagerPickerId ? Number(payload.closureAssetManagerPickerId) : undefined,
          AssetManagerStatus: 'Pending',
          AssetManagerApprovalDate: null,
          AssetManagerRejectionReason: '',

          POClosureApproverId: originatorId,
          POClosureDate: null,
          POClosureStatus: 'Pending',

          AssetDirectorReplacerId: payload.assetDirReplacerPickerId ? Number(payload.assetDirReplacerPickerId) : undefined,
          HSEDirectorReplacerId: payload.hseDirReplacerPickerId ? Number(payload.hseDirReplacerPickerId) : undefined,

          IsAssetDirectorReplacer: payload.isAssetDirectorReplacer,
          IsHSEDirectorReplacer: payload.isHSEDirectorReplacer,
        };

        if (wfItemId) {
          const wfRes = await wfOps._updateItem(String(wfItemId), body);
          if (!wfRes.ok) throw new Error('Failed to update workflow stage.');
        }

        showBanner('Cancelled Successfully.', { autoHideMs: 3000, fade: true, kind: 'success' });
        goBackToHost();
        return true;
      }
      catch (e) {
        showBanner('Failed to cancel. Please try again.', { autoHideMs: 5000, fade: true, kind: 'error' });
      } finally {
        setIsBusy(false);
      }
      return true;
    }

  }, [buildPayload, isPermitOriginator, props.formId, webUrl, props.context.spHttpClient, _workflowStage]);

  const issuePermit = React.useCallback(async (mode: 'issuePermit'): Promise<boolean> => {
    payloadRef.current = buildPayload();
    const payload = payloadRef.current;
    const validationError = validateBeforeApprove(mode);
    if (validationError) {
      showBanner(validationError);
      return false;
    }
    else {
      setIsBusy(true);
      try {
        // Find workflow item for this form
        const formId = Number(props.formId);
        const query = `?$select=Id&$filter=PTWForm/Id eq ${formId} && PermitType eq 'new'`;
        const ops = new SPCrudOperations((props.context as any).spHttpClient, webUrl, 'PTW_Form_Work_Permits', query);
        const list = await ops._getItemsWithQuery();
        const itemId = Array.isArray(list) && list[0]?.Id;
        if (!itemId) throw new Error('Permit item not found.');

        let body = {};
        if (isAssetDirector) {
          body = {
            PIStatus: payload.permitStatus,
            PIApprovalDate: payload.permitStatus !== 'Pending' ? payload.permitApprovalDate : null,
          }
        }

        const res = await ops._updateItem(String(itemId), body);
        if (!res.ok) throw new Error('Failed to update Permit status.');
        showBanner(`Approved Successfully.`, { autoHideMs: 3000, fade: true, kind: 'success' });
        goBackToHost();
        return true;
      }
      catch (e) {
        showBanner('Failed to approve. Please try again.', { autoHideMs: 5000, fade: true, kind: 'error' });
      } finally {
        setIsBusy(false);
      }
      return true;
    }
  }, [validateBeforeApprove, buildPayload, props.formId, webUrl, props.context.spHttpClient]);

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
      const spOps = spCrudRef.current ?? new SPCrudOperations((props.context as any).spHttpClient, webUrl, '', '');
      const originatorId = await spOps.ensureUserId(_PermitOriginator?.[0]?.secondaryText || '');

      const validationError = validateBeforeSubmit(originatorId, mode);
      if (validationError) {
        showBanner(validationError);
        return false;
      } else {
        const editFormId = props.formId ? Number(props.formId) : undefined;
        const formStatusRecord = JSON.parse(localStorage.getItem('FormStatusRecord') || '{}');

        if (editFormId === undefined) {
          const savedId = await _createPTWForm(mode, originatorId);
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
  }, [isPermitOriginator, validateBeforeSubmit, props.formId]);

  // Create parent PTWForm item and return its Id
  const _createPTWForm = React.useCallback(async (mode: 'save' | 'submit' | 'renew', spOriginatorId?: number): Promise<number> => {
    const payload = payloadRef.current;

    if (!payload) throw new Error('Form payload is not available');

    const body: any = {
      PermitOriginatorId: spOriginatorId ?? null,
      Title: 'PTW Form' + (spOriginatorId ? ` - ${payload.originatorId}` : ''),
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
      WorkflowStatus: mode === 'submit' ? PTWWorkflowStatus.New : '',
      AttachmentsProvided: payload.attachmentsProvided.toLowerCase() === "yes" ? true : false,
      AttachmentsProvidedDetails: payload.attachmentsDetails ?? '',
      IsUrgentSubmission: !!payload.isUrgentSubmission,
      PreviousReferenceNumber: payload.previousReferenceNumber ?? null,
      PermitsValidityDays: payload.permitPayloadValidityDays
    };

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
      body['PersonnelInvolvedId'] = payload.personnelIds.map(Number);
    }

    spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form', '');
    const newId = await spCrudRef.current._insertItem(body);
    if (!newId) throw new Error('Failed to create PTW Form');

    try {
      const coralReferenceNumber = await spHelpers.assignCoralReferenceNumber(props.context.spHttpClient,
        webUrl, 'PTW_Form', { Id: Number(newId) }, payload.company?.title, 'PTW');
      if (!coralReferenceNumber) throw new Error('Failed to generate Coral Reference Number. Please try again later.');

      setCoralReferenceNumber(coralReferenceNumber);

      if (payload.permitRows?.length && payload.permitRows.some((r: IPermitScheduleRow) => r.isChecked)) {
        const _createdPermits = await _createPTWWorkPermit(Number(newId), payload.permitRows[0]);

        if (!_createdPermits?.length) {
          throw new Error('Failed to create PTW Work Permits');
        }
      }

      if (mode === 'submit') {
        const _createdWorkflow = await _createPTWFormApprovalWorkflow(Number(newId), spOriginatorId);

        if (!_createdWorkflow) {
          throw new Error('Failed to create PTW Form Approval Workflow');
        }
      }

      if (payload.workTaskLists?.length) {
        const _createdTask = await _createPTWTasksJobsDescriptions(Number(newId), payload.workTaskLists, undefined);

        if (!_createdTask?.length) {
          throw new Error('Failed to create PTW Tasks and Job Descriptions');
        }
      }

    } catch (e) {
      console.warn('Failed to create PTW Form:', e);
    }

    return newId as number;
  }, [props.context.spHttpClient, payloadRef.current]);

  const _createPTWWorkPermit = React.useCallback(async (parentId: number, permitRows: IPermitScheduleRow) => {
    const opsDelete = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form_Work_Permits', '');
    await Promise.all([permitRows].map(async (item) => {
      await opsDelete._deleteLookUPItems(Number(parentId), "PTWForm");
    }));

    const requiredItems = [permitRows].filter((row) => row.isChecked && row.type.toLowerCase() === 'new');
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
        Title: item.type === 'new' ? 'New Permit' : 'Renewal Permit',
        PIApproverId: item.piApprover ? Number(item.piApprover?.id) : null,
        StatusRecord: 'new',
      };

      const data = ops._insertItem(body);

      if (!data) throw new Error('Failed to create PTW Work Permits.');
      return typeof data === 'number' ? data : (data);
    });
    const results = await Promise.all(posts);
    return results;
  }, [props.context.spHttpClient, spHelpers]);

  const _createPTWTasksJobsDescriptions = React.useCallback(async (parentId: number, workTaskLists: IRiskTaskRow[], mode: 'submitAfterRejection' | undefined) => {
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
        OrderRecord: index + 1,
      };

      if (mode === 'submitAfterRejection') {
        body.InitialRisk = null;
        body.ResidualRisk = null;
      }

      if (item.safeguardIds?.length) {
        body['SafeguardsId@odata.type'] = 'Collection(Edm.Int32)';
        body['SafeguardsId'] = item.safeguardIds.map(Number);
      }

      // Add custom safeguards -> SP multi-choice field "OtherSafeguards"
      const other = (item.customSafeguards || [])
        .map(s => String(s).trim())
        .filter(Boolean);
      if (other.length) {
        body['OtherSafeguards@odata.type'] = 'Collection(Edm.String)';
        body['OtherSafeguards'] = Array.from(new Set(other));
      }

      const data = ops._insertItem(body);

      if (!data) throw new Error('Failed to create PTW Tasks Descriptions.');
      return typeof data === 'number' ? data : (data);
    });
    const results = await Promise.all(posts);
    return results;
  }, [props.context.spHttpClient]);

  const _updatePTWForm = React.useCallback(async (id: number, mode: 'save' | 'submit' | 'approve' | 'submitAfterRejection'): Promise<boolean> => {
    const payload = payloadRef.current;
    if (!payload) throw new Error('Form payload is not available');

    const spOps = spCrudRef.current ?? new SPCrudOperations((props.context as any).spHttpClient, webUrl, '', '');
    const originatorId = await spOps.ensureUserId(payload.originatorEMail || '');

    let toolboxTalkConductedById: number | undefined = undefined;
    if (payload.toolboxTalkConductedById) {
      toolboxTalkConductedById = await spOps.ensureUserId(payload.toolboxTalkConductedById || '');
    }

    let body: any = {
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
      AttachmentsProvided: payload.attachmentsProvided.toLowerCase() === "yes" ? true : false,
      AttachmentsProvidedDetails: payload.attachmentsDetails ?? '',
      IsUrgentSubmission: !!payload.isUrgentSubmission,
      PreviousReferenceNumber: payload.previousReferenceNumber ?? null,
      PermitsValidityDays: payload.permitPayloadValidityDays,
    };

    if (mode === 'save') {
      body.FormStatusRecord = 'Saved';
      body.WorkflowStatus = '';
    }
    else if (mode === 'submit') {
      body.FormStatusRecord = 'Submitted';
      body.WorkflowStatus = PTWWorkflowStatus.New;
    }
    else if (mode === 'approve') {
      body.WorkflowStatus = PTWWorkflowStatus.InReview;
    }

    if (mode === 'submitAfterRejection') {
      body.OverallRiskAssessment = null;
      body.IsDetailedRiskAssessmentRequired = false;
      body.RiskAssessmentRefNumber = null;
      body.GasTestRequired = null;
      body.GasTestResult = null;
      body.FireWatchNeeded = null;
      body.FireWatchAssigned = null;
      body.ToolboxTalk = null;
      body.ToolboxConductedById = null;
      body.ToolboxTalkHSEReference = null;
      body.ToolBoxTalkDate = null;
      body.FormStatusRecord = 'Submitted';
      body.WorkflowStatus = PTWWorkflowStatus.InReview;
    }
    else if ((isPermitIssuer || (isAssetDirector && payload.isUrgentSubmission) || (isHSEDirector && payload.isUrgentSubmission)) && mode === 'approve') {
      body.OverallRiskAssessment = payload.overallRiskAssessment ?? null;
      body.IsDetailedRiskAssessmentRequired = payload.detailedRiskAssessment ?? false;
      body.RiskAssessmentRefNumber = payload.detailedRiskAssessmentRef ?? null;
      body.GasTestRequired = payload.gasTestRequired?.toLowerCase() === 'yes' ? true : false;
      body.GasTestResult = payload.gasTestResult ?? null;
      body.FireWatchNeeded = payload.fireWatchNeeded?.toLowerCase() === 'yes' ? true : false;
      body.FireWatchAssigned = payload.fireWatchAssigned ?? '';
      body.ToolboxTalk = payload.toolboxTalk ?? null;
      body.ToolboxConductedById = toolboxTalkConductedById ?? null;
      body.ToolboxTalkHSEReference = payload.toolboxHSEReference ?? null;

      if (payload.toolboxTalkDate) {
        const dt = payload.toolboxTalkDate instanceof Date ? payload.toolboxTalkDate : new Date(payload.toolboxTalkDate);
        if (!isNaN(dt.getTime())) body.ToolBoxTalkDate = dt;
      }
    }

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
      body['PersonnelInvolvedId@odata.type'] = 'Collection(Edm.Int32)';
      body['PersonnelInvolvedId'] = payload.personnelIds.map(Number);
    }

    spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form', '');
    const response = await spCrudRef.current._updateItem(String(id), body);
    if (!response.ok) {
      showBanner('Failed to update PTW Form.', { autoHideMs: 5000, fade: true, kind: 'error' });
      return false;
    }

    if (payload.permitRows?.length && payload.permitRows.some((r: IPermitScheduleRow) => r.isChecked)) {
      const _createdPermits = await _createPTWWorkPermit(Number(id), payload.permitRows[0]);

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
      const _createdTask = await _createPTWTasksJobsDescriptions(Number(id), payload.workTaskLists, mode === "submitAfterRejection" ? "submitAfterRejection" : undefined);

      if (!_createdTask?.length) {
        throw new Error('Failed to create PTW Tasks and Job Descriptions');
      }
    }

    return true;
  }, [props.context.spHttpClient, payloadRef.current, isPermitIssuer]);

  const _renewPermit = React.useCallback(async (mode: 'renew'): Promise<boolean> => {
    const payload = buildPayload();
    payloadRef.current = payload;
    const formId = Number(props.formId);
    if (!payload) throw new Error('Form payload is not available');

    const isNumericId = (id: string) => /^[0-9]+$/.test(String(id || ''));
    let renewedPermit: IPermitScheduleRow | undefined;

    if (payload.permitRows && payload.permitRows.length) {
      // Validate there is at least one non-numeric renewal row fully filled
      renewedPermit = payload.permitRows
        .filter((r: IPermitScheduleRow) =>
          r.type === 'renewal' &&
          String(r.statusRecord || '').toLowerCase() === 'new' &&
          !isNumericId(r.id) &&
          !!String(r.date || '').trim() &&
          !!String(r.startTime || '').trim() &&
          !!String(r.endTime || '').trim() &&
          !!r.piApprover
        )
        .sort((a, b) => a.orderRecord - b.orderRecord)[0];

      if (!renewedPermit) {
        showBanner(
          'Add and fully complete a Renewal Permit row (date, start time, end time, approver) before renewing.',
          { autoHideMs: 5000, fade: true, kind: 'error' }
        );
        return false;
      }

      // VALIDATE: date/start/end must be filled
      const hasDate = !!String(renewedPermit.date || '').trim();
      const hasStart = !!String(renewedPermit.startTime || '').trim();
      const hasEnd = !!String(renewedPermit.endTime || '').trim();

      if (!hasDate || !hasStart || !hasEnd) {
        showBanner('Please fill Date, Start Time, and End Time for the selected permit before renewing.', { autoHideMs: 5000, fade: true, kind: 'error' });
        return false;
      }

      // VALIDATE: start < end
      const startMins = spHelpers.parseTimeToMinutes(renewedPermit.startTime);
      const endMins = spHelpers.parseTimeToMinutes(renewedPermit.endTime);
      if (!isNaN(startMins) && !isNaN(endMins) && startMins >= endMins) {
        showBanner('End time must be after start time.', { autoHideMs: 5000, fade: true, kind: 'error' });
        return false;
      }

      const formBody: any = { ToRenewPermit: true };
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form', '');
      const response = await spCrudRef.current._updateItem(String(formId), formBody);

      const permitBody: any = {
        PTWFormId: formId,
        PermitType: renewedPermit.type ?? null,
        PermitDate: renewedPermit.date,
        PermitStartTime: spHelpers.combineDateAndTime(renewedPermit.date.toString(), renewedPermit.startTime),
        PermitEndTime: spHelpers.combineDateAndTime(renewedPermit.date.toString(), renewedPermit.endTime),
        StatusRecord: 'New',
        RecordOrder: renewedPermit.orderRecord,
        PIApproverId: renewedPermit.piApprover ? Number(renewedPermit.piApprover?.id) : null,
        Title: 'Renewal Permit for form #' + formId
      }

      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form_Work_Permits', '');
      const responsePermit = await spCrudRef.current._insertItem(permitBody);

      if (!response.ok) {
        showBanner('Failed to update PTW Form.', { autoHideMs: 5000, fade: true, kind: 'error' });
        return false;
      }

      if (!responsePermit) {
        showBanner('Failed to update PTW Work Permits.', { autoHideMs: 5000, fade: true, kind: 'error' });
        return false;
      }
    }
    else {
      showBanner('No permit rows found to renew.', { autoHideMs: 4000, fade: true, kind: 'warning' });
      return false;
    }
    goBackToHost();
    return true;
  }, [props.context.spHttpClient, payloadRef.current, buildPayload, props.formId, spHelpers]);

  const _approveRenewalPermit = React.useCallback(async (mode: 'approveRenewalPermit'): Promise<boolean> => {
    const payload = buildPayload();
    payloadRef.current = payload;
    if (!payload) throw new Error('Form payload is not available');

    const isNumericId = (id: string) => /^[0-9]+$/.test(String(id || ''));
    let renewedPermit: IPermitScheduleRow | undefined;

    if (payload.permitRows && payload.permitRows.length) {
      // Validate there is at least one non-numeric renewal row fully filled
      renewedPermit = payload.permitRows
        .filter((r: IPermitScheduleRow) =>
          r.type === 'renewal' &&
          String(r.statusRecord || '').toLowerCase() === 'new' &&
          isNumericId(r.id) &&
          !!String(r.date || '').trim() &&
          !!String(r.startTime || '').trim() &&
          !!String(r.endTime || '').trim() &&
          !!r.piApprover).sort((a, b) => a.orderRecord - b.orderRecord)[0];

      const permitBody: any = {
        PIStatus: renewedPermit.piStatus ?? null,
        PIApprovalDate: renewedPermit.piStatus === 'Approved' || renewedPermit.piStatus === 'Rejected' ? new Date() : null,
      }

      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form_Work_Permits', '');
      const responsePermit = await spCrudRef.current._updateItem(String(renewedPermit.id), permitBody);

      if (!responsePermit.ok) {
        showBanner('Failed to approve the renweal permit.', { autoHideMs: 5000, fade: true, kind: 'error' });
        return false;
      }
    }
    else {
      showBanner('No permit rows found to approve.', { autoHideMs: 5000, fade: true, kind: 'warning' });
      return false;
    }
    goBackToHost();
    return true;
  }, [props.context.spHttpClient, payloadRef.current, buildPayload, spHelpers]);

  const _createPTWFormApprovalWorkflow = React.useCallback(async (parentId: number, originatorId: number | undefined, mode?: 'extend') => {
    const payload = payloadRef.current;
    if (originatorId === undefined) return;
    const ops = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form_Approval_Workflow', '');
    let body: any = {};
    const isoDateNow = new Date().toISOString();
    try {
      if (!payload.isUrgentSubmission) {
        let _paStatusForPOPA = '';
        let _paDateForPOPA = isoDateNow;

        if (originatorId === payload.paPickerId) {
          _paStatusForPOPA = payload.paStatus === 'Approved' ? payload.paStatus : 'Approved';
          _paDateForPOPA = _paDateForPOPA;
        } else {
          _paStatusForPOPA = payload.paStatus || 'Pending';
          _paDateForPOPA = payload.paApprovalDate;
        }

        body = {
          PTWFormId: parentId,
          StatusRecord: 'New',
          IsFinalApprover: false,

          POApproverId: originatorId,
          POApprovalDate: payload.poApprovalDate || null,
          POStatus: payload.poStatus || 'Approved',

          PAApproverId: payload.paPickerId ? Number(payload.paPickerId) : undefined,
          PAStatus: _paStatusForPOPA,
          PAApprovalDate: _paStatusForPOPA !== 'Pending' ? _paDateForPOPA : null,
          PARejectionReason: payload.paRejectionReason || '',

          PIApproverId: payload.piPickerId ? Number(payload.piPickerId) : undefined,
          PIStatus: payload.piStatus || 'Pending',
          PIApprovalDate: payload.piStatus !== 'Pending' ? payload.piApprovalDate : null,
          PIRejectionReason: payload.piRejectionReason || '',

          AssetDirectorApproverId: payload.assetDirPickerId ? Number(payload.assetDirPickerId) : undefined,
          AssetDirectorStatus: payload.assetDirStatus || 'Pending',
          AssetDirectorRejectionReason: payload.assetDirRejectionReason || '',

          HSEDirectorApproverId: payload.hseDirPickerId ? Number(payload.hseDirPickerId) : undefined,
          HSEDirectorStatus: payload.hseDirStatus || 'Pending',
          HSEDirectorRejectionReason: payload.hseDirRejectionReason || '',

          AssetManagerApproverId: payload.closureAssetManagerPickerId ? Number(payload.closureAssetManagerPickerId) : undefined,
          AssetManagerStatus: payload.closureAssetManagerStatus || 'Pending',

          POClosureApproverId: originatorId,
          POClosureStatus: payload.closurePoStatus || 'Pending',

          AssetDirectorReplacerId: payload.assetDirReplacerPickerId ? Number(payload.assetDirReplacerPickerId) : undefined,
          HSEDirectorReplacerId: payload.hseDirReplacerPickerId ? Number(payload.hseDirReplacerPickerId) : undefined,
        };
      }

      if (payload.isUrgentSubmission) {
        body = {
          PTWFormId: parentId,
          StatusRecord: 'New',
          IsFinalApprover: false,

          POApproverId: originatorId,
          POApprovalDate: payload.poApprovalDate || null,
          POStatus: payload.poStatus || 'Approved',
          PAStatus: 'Pending',

          PIApproverId: payload.piPickerId ? Number(payload.piPickerId) : undefined,
          PIStatus: payload.piStatus || 'Pending',
          PIApprovalDate: payload.piStatus !== 'Pending' ? payload.piApprovalDate : null,
          PIRejectionReason: payload.piRejectionReason || '',

          AssetDirectorApproverId: payload.assetDirPickerId ? Number(payload.assetDirPickerId) : undefined,
          AssetDirectorReplacerId: payload.assetDirReplacerPickerId ? Number(payload.assetDirReplacerPickerId) : undefined,
          UrgentAssetDirectorStatus: payload.urgentAssetDirStatus || 'Pending',

          HSEDirectorApproverId: payload.hseDirPickerId ? Number(payload.hseDirPickerId) : undefined,
          HSEDirectorStatus: payload.hseDirStatus || 'Pending',
          HSEDirectorRejectionReason: payload.hseDirRejectionReason || '',
          HSEDirectorReplacerId: payload.hseDirReplacerPickerId ? Number(payload.hseDirReplacerPickerId) : undefined,



          POClosureApproverId: originatorId,
          POClosureStatus: payload.closurePoStatus || 'Pending',

          AssetManagerApproverId: payload.closureAssetManagerPickerId ? Number(payload.closureAssetManagerPickerId) : undefined,
          AssetManagerStatus: payload.closureAssetManagerStatus || 'Pending',

          IsAssetDirectorReplacer: payload.isAssetDirectorReplacer,
          IsHSEDirectorReplacer: payload.isHSEDirectorReplacer,
        }
      }

      if (mode === 'extend') {
        if (!payload.isUrgentSubmission) {
          let _paStatusForPOPA = '';
          let _paDateForPOPA = isoDateNow;

          if (originatorId === payload.paPickerId) {
            _paStatusForPOPA = payload.paStatus === 'Approved' ? payload.paStatus : 'Approved';
            _paDateForPOPA = _paDateForPOPA;
          } else {
            _paStatusForPOPA = payload.paStatus || 'Pending';
            _paDateForPOPA = payload.paApprovalDate;
          }

          body = {
            PTWFormId: parentId,
            StatusRecord: 'New',
            IsFinalApprover: false,

            POApproverId: originatorId,
            POApprovalDate: payload.poApprovalDate || null,
            POStatus: payload.poStatus || 'Approved',

            PAApproverId: payload.paPickerId ? Number(payload.paPickerId) : undefined,
            PAStatus: _paStatusForPOPA,
            PAApprovalDate: _paStatusForPOPA !== 'Pending' ? _paDateForPOPA : null,
            PARejectionReason: null,

            PIApproverId: payload.piPickerId ? Number(payload.piPickerId) : undefined,
            PIStatus: 'Pending',
            PIApprovalDate: null,
            PIRejectionReason: null,

            AssetDirectorApproverId: payload.assetDirPickerId ? Number(payload.assetDirPickerId) : undefined,
            AssetDirectorStatus: 'Pending',
            AssetDirectorRejectionReason: null,

            HSEDirectorApproverId: payload.hseDirPickerId ? Number(payload.hseDirPickerId) : undefined,
            HSEDirectorStatus: 'Pending',
            HSEDirectorRejectionReason: null,

            AssetManagerApproverId: payload.closureAssetManagerPickerId ? Number(payload.closureAssetManagerPickerId) : undefined,
            AssetManagerStatus: 'Pending',
            AssetManagerRejectionReason: null,

            POClosureApproverId: originatorId,
            POClosureStatus: 'Pending',

            AssetDirectorReplacerId: payload.assetDirReplacerPickerId ? Number(payload.assetDirReplacerPickerId) : undefined,
            HSEDirectorReplacerId: payload.hseDirReplacerPickerId ? Number(payload.hseDirReplacerPickerId) : undefined,
          };
        }

        if (payload.isUrgentSubmission) {
          body = {
            PTWFormId: parentId,
            StatusRecord: 'New',
            IsFinalApprover: false,

            POApproverId: originatorId,
            POApprovalDate: payload.poApprovalDate || null,
            POStatus: payload.poStatus || 'Approved',
            PAStatus: 'Pending',

            PIApproverId: payload.piPickerId ? Number(payload.piPickerId) : undefined,
            PIStatus: 'Pending',
            PIApprovalDate: null,
            PIRejectionReason: null,

            AssetDirectorApproverId: payload.assetDirPickerId ? Number(payload.assetDirPickerId) : undefined,
            UrgentAssetDirectorStatus: 'Pending',

            HSEDirectorApproverId: payload.hseDirPickerId ? Number(payload.hseDirPickerId) : undefined,
            HSEDirectorStatus: 'Pending',
            HSEDirectorRejectionReason: null,
            HSEDirectorReplacerId: payload.hseDirReplacerPickerId ? Number(payload.hseDirReplacerPickerId) : undefined,

            POClosureApproverId: originatorId,
            POClosureStatus: 'Pending',

            AssetManagerApproverId: payload.closureAssetManagerPickerId ? Number(payload.closureAssetManagerPickerId) : undefined,
            AssetManagerStatus: 'Pending',

            IsAssetDirectorReplacer: payload.isAssetDirectorReplacer,
            IsHSEDirectorReplacer: payload.isHSEDirectorReplacer,
          }
        }
      }
      const data = ops._insertItem(body);
      if (!data) throw new Error('Failed to create PTW Form Approval Workflow.');
      return typeof data === 'number' ? data : (data);
    }
    catch (e) {
      console.warn('Failed to create PTW Form Approval Workflow', e);
    }
  }, [props.context.spHttpClient, payloadRef.current]);

  // ---------------------------
  // Render - Prefill when editing an existing form
  // ---------------------------
  React.useEffect(() => {
    const formId = props.formId;
    if (!formId || prefilledFormId === formId) return;

    // Wait until base items are loaded and itemRows initialized
    if (loading) return;

    let cancelled = false;

    const toPersona = (obj?: { Id?: any; EMail?: string; displayName?: string }): IPersonaProps | undefined => {
      if (!obj) return undefined;
      const text = obj.displayName || '';
      const email = obj.EMail || '';
      const id = obj.Id != null ? String(obj.Id) : text;
      return { text, secondaryText: email, id } as IPersonaProps;
    };

    const load = async () => {
      try {

        const ptwFirstSelect = `?$select=Id,CoralReferenceNumber,AssetID,ProjectTitle,Created,FormStatusRecord,IsDetailedRiskAssessmentRequired,RiskAssessmentRefNumber,WorkHazardsOthers,` +
          `OverallRiskAssessment,GasTestRequired,GasTestResult,WorkflowStatus,IsUrgentSubmission,PreviousReferenceNumber,PermitsValidityDays,ToRenewPermit,` +
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
          `ToolboxTalkHSEReference,ToolBoxTalkDate,ProtectiveSafetyEquipmentsOthers,PrecautionsOthers,FireWatchAssigned,` +
          `Precuations/Id,Precuations/Title,` +
          `ProtectiveSafetyEquipments/Id,ProtectiveSafetyEquipments/Title,` +
          `MachineryInvolved/Id,MachineryInvolved/Title,` +
          `PersonnelInvolved/Id,PersonnelInvolved/Title,` +
          `ToolboxConductedBy/Id,ToolboxConductedBy/Title,ToolboxConductedBy/EMail` +
          `&$expand=Precuations,ProtectiveSafetyEquipments,MachineryInvolved,` +
          `PersonnelInvolved,ToolboxConductedBy` +
          `&$filter=Id eq ${formId}`;

        const ptwWorkPermits = `?$select=Id,PermitType,PermitDate,PermitStartTime,PermitEndTime,RecordOrder,StatusRecord,PIApprovalDate,PIStatus,` +
          `PTWForm/Id,PTWForm/CoralReferenceNumber,` +
          `PIApprover/Id,PIApprover/Title,PIApprover/EMail` +
          `&$expand=PTWForm,PIApprover` +
          `&$filter=PTWForm/Id eq ${formId}`;

        const ptwTaskDescription = `?$select=Id,JobDescription,InitialRisk,ResidualRisk,OrderRecord,OtherSafeguards,` +
          `PTWForm/Id,PTWForm/CoralReferenceNumber,` +
          `Safeguards/Id,Safeguards/Title` +
          `&$expand=PTWForm,Safeguards` +
          `&$filter=PTWForm/Id eq '${formId}'`;

        const workflow: string = `?$select=Id,PTWForm/Id,PTWForm/CoralReferenceNumber,POStatus,PAStatus,PIStatus,UrgentAssetDirectorStatus,AssetDirectorStatus,HSEDirectorStatus,POClosureStatus,AssetManagerStatus,` +
          `POApprovalDate,PAApprovalDate,PIApprovalDate,UrgentAssetDirectorApprovalDate,AssetDirectorApprovalDate,HSEDirectorApprovalDate,POClosureDate,AssetManagerApprovalDate,Stage,IsAssetDirectorReplacer,IsHSEDirectorReplacer,` +
          `POApprover/Id,POApprover/EMail,POApprover/Title,` +
          `PAApprover/Id,PAApprover/EMail,PAApprover/Title,PARejectionReason,` +
          `PIApprover/Id,PIApprover/EMail,PIApprover/Title,PIRejectionReason,` +
          `AssetDirectorApprover/Id,AssetDirectorApprover/EMail,AssetDirectorApprover/Title,AssetDirectorRejectionReason,UrgentAssetDirectorRejectionReas,` +
          `AssetDirectorReplacer/Id,AssetDirectorReplacer/EMail,AssetDirectorReplacer/Title,` +
          `HSEDirectorApprover/Id,HSEDirectorApprover/EMail,HSEDirectorApprover/Title,HSEDirectorRejectionReason,` +
          `HSEDirectorReplacer/Id,HSEDirectorReplacer/EMail,HSEDirectorReplacer/Title,` +
          `POClosureApprover/Id,POClosureApprover/EMail,POClosureApprover/Title,POClosureRejectionReason,` +
          `AssetManagerApprover/Id,AssetManagerApprover/EMail,AssetManagerApprover/Title` +
          `&$expand=PTWForm,POApprover,PAApprover,PIApprover,AssetDirectorApprover,AssetDirectorReplacer,HSEDirectorApprover,HSEDirectorReplacer,POClosureApprover,AssetManagerApprover` +
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

        const formWorkflow = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form_Approval_Workflow', workflow);
        const headerItemsWorkflow = await formWorkflow._getItemsWithQuery();
        const headerWorkflow = Array.isArray(headerItemsWorkflow) ? headerItemsWorkflow[0] : undefined;

        if (headerFirstSelect && !cancelled && headerSecondSelect) {
          // Top-level fields prefill
          if (headerFirstSelect?.FormStatusRecord) {
            setMode(headerFirstSelect?.FormStatusRecord.toLowerCase());
          }

          const permitOriginator = toPersona({ Id: headerFirstSelect?.PermitOriginator?.Id, displayName: headerFirstSelect?.PermitOriginator?.Title, EMail: headerFirstSelect?.PermitOriginator?.EMail });
          setPermitOriginator(permitOriginator ? [permitOriginator] : []);
          setCoralReferenceNumber(headerFirstSelect?.CoralReferenceNumber || '');
          setAssetId(headerFirstSelect?.AssetID);
          const selectedCompany = ptwFormStructure?.companies?.find(c => c.id === headerFirstSelect.CompanyRecord.Id);
          setSelectedCompany(headerFirstSelect?.CompanyRecord ? { id: headerFirstSelect.CompanyRecord.Id, title: headerFirstSelect.CompanyRecord.Title || '', orderRecord: selectedCompany?.orderRecord || 0, fullName: selectedCompany?.fullName || '' } : undefined);
          setProjectTitle(headerFirstSelect?.ProjectTitle || '');
          setSelectedAssetCategory(headerFirstSelect?.AssetCategory ? Number(headerFirstSelect.AssetCategory.Id) : undefined);
          setSelectedAssetDetails(headerFirstSelect?.AssetDetails ? Number(headerFirstSelect.AssetDetails.Id) : undefined);
          setSelectedHacWorkAreaId(headerFirstSelect?.HACClassificationWorkArea?.Id != null ? Number(headerFirstSelect.HACClassificationWorkArea.Id) : undefined);
          setSelectedHacWorkAreaId(headerFirstSelect?.HACClassificationWorkArea?.Id != null ? Number(headerFirstSelect.HACClassificationWorkArea.Id) : undefined);
          setSelectedWorkHazardIds(new Set(Array.isArray(headerFirstSelect.WorkHazards) ? headerFirstSelect.WorkHazards.map((wh: any) => Number(wh.Id)) : []));
          setWorkHazardsOtherText(headerFirstSelect?.WorkHazardsOthers || '');
          setOverAllRiskAssessment(headerFirstSelect?.OverallRiskAssessment);
          setDetailedRiskAssessment(!!headerFirstSelect?.IsDetailedRiskAssessmentRequired);
          setRiskAssessmentReferenceNumber(headerFirstSelect?.RiskAssessmentRefNumber || '');
          setSelectedPrecautionIds(new Set(Array.isArray(headerSecondSelect.Precuations) ? headerSecondSelect.Precuations.map((pc: any) => Number(pc.Id)) : []));
          setPrecautionsOtherText(headerSecondSelect?.PrecautionsOthers || '');
          setGasTestValue(headerFirstSelect?.GasTestRequired ? (headerFirstSelect?.GasTestRequired ? "Yes" : "No") : '');
          setGasTestResult(headerFirstSelect?.GasTestResult || '');
          setFireWatchValue(headerSecondSelect?.FireWatchNeeded ? (headerSecondSelect?.FireWatchNeeded ? "Yes" : "No") : '');
          setFireWatchAssigned(headerSecondSelect?.FireWatchAssigned ? headerSecondSelect.FireWatchAssigned : '');
          setAttachmentsValue(headerSecondSelect?.AttachmentsProvided ? (headerSecondSelect.AttachmentsProvided ? 'Yes' : 'No') : '');
          setAttachmentsResult(headerSecondSelect?.AttachmentsProvidedDetails || '');
          setIsUrgentSubmission(!!headerFirstSelect?.IsUrgentSubmission);
          setPreviousPtwRef(headerFirstSelect?.PreviousReferenceNumber || '');
          setPermitPayloadValidityDays(headerFirstSelect?.PermitsValidityDays || 0);
          setIsIssued(headerFirstSelect?.WorkflowStatus === PTWWorkflowStatus.Issued);

          if (headerFirstSelect?.AssetDetails) {
            const assetDetailId = Number(headerFirstSelect.AssetDetails.Id);
            const cached = (assetCategoriesDetailsList || []).find(d => Number(d.id) === assetDetailId);

            const setFromDetail = (detail: any) => {
              setPiHsePartnerFilteredByCategory(detail?.HSEPartner || []);
              setAssetDirFilteredByCategory(detail?.AssetDirector || []);
              setAssetManagerFilteredByCategory(detail?.AssetManager || []);
            };

            if (cached) {
              setFromDetail({
                HSEPartner: cached.hsePartner,
                AssetDirector: cached.assetDirector,
                AssetManager: cached.assetManager
              });
            }
            else {
              // Fallback: fetch this asset detail directly
              try {
                const query = `?$select=Id,` +
                  `AssetDirector/Id,AssetDirector/Title,AssetDirector/EMail,` +
                  `AssetManager/Id,AssetManager/Title,AssetManager/EMail,` +
                  `HSEPartner/Id,HSEPartner/Title,HSEPartner/EMail` +
                  `&$expand=AssetDirector,AssetManager,HSEPartner` +
                  `&$filter=Id eq ${assetDetailId}`;
                const ops = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'LKP_Asset_Details', query);
                const arr = await ops._getItemsWithQuery();
                const item = Array.isArray(arr) ? arr[0] : undefined;
                if (item) setFromDetail(item);
                else {
                  setPiHsePartnerFilteredByCategory([]);
                  setAssetDirFilteredByCategory([]);
                  setAssetManagerFilteredByCategory([]);
                }
              } catch {
                setPiHsePartnerFilteredByCategory([]);
                setAssetDirFilteredByCategory([]);
                setAssetManagerFilteredByCategory([]);
              }
            }
          }

          if (headerSecondSelect.ProtectiveSafetyEquipments.length > 0) {
            const ids = (headerSecondSelect.ProtectiveSafetyEquipments || [])
              .map((item: any) => {
                if (String(item?.Title || '').toLowerCase().includes('other')) {
                  setProtectiveEquipmentsOtherText(headerSecondSelect?.ProtectiveSafetyEquipmentsOthers || '');
                }
                return Number(item.Id);
              })
              .filter((n: number) => !isNaN(n));
            setSelectedProtectiveEquipmentIds(new Set<number>(ids));
          }

          if (headerSecondSelect?.MachineryInvolved.length > 0) {
            setSelectedMachineryIds(headerSecondSelect?.MachineryInvolved.map((item: any) => Number(item.Id)) || []);
          }
          if (headerSecondSelect?.PersonnelInvolved.length > 0) {
            setSelectedPersonnelIds(headerSecondSelect?.PersonnelInvolved.map((item: any) => Number(item.Id)) || []);
          }

          setToolboxTalk(headerSecondSelect?.ToolboxTalk || '');
          setToolboxHSEReference(headerSecondSelect?.ToolboxTalkHSEReference || '');
          setToolboxTalkDate(headerSecondSelect?.ToolBoxTalkDate ? new Date(headerSecondSelect.ToolBoxTalkDate) : undefined);
          const toolboxConductedBy = toPersona({ Id: headerSecondSelect?.ToolboxConductedBy?.Id, displayName: headerSecondSelect?.ToolboxConductedBy?.Title, EMail: headerSecondSelect?.ToolboxConductedBy?.EMail });
          setToolboxConductedBy(toolboxConductedBy && toolboxConductedBy?.id ? [toolboxConductedBy] : undefined);

          if (headerTaskDescription && headerTaskDescription.length > 0) {
            const tasksList: IRiskTaskRow[] = [];
            headerTaskDescription.forEach((item: any, index) => {
              if (item) {
                tasksList.push({
                  id: item.Id,
                  task: item.JobDescription || '',
                  initialRisk: item.InitialRisk || '',
                  residualRisk: item.ResidualRisk || '',
                  disabledFields: false,
                  orderRecord: item.OrderRecord || 0,
                  customSafeguards: Array.isArray(item.OtherSafeguards) ? item.OtherSafeguards : [],
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
          setWorkPermitRequired(_workCategories?.length > 0);
          if (headerWorkPermits && headerWorkPermits.length > 0) {
            const permitsList: IPermitScheduleRow[] = [];
            headerWorkPermits.sort((a: any, b: any) => a.OrderRecord - b.OrderRecord).forEach((item: any, index) => {
              if (item) {
                const startTime = item?.PermitStartTime ? spHelpers.toHHmm(item.PermitStartTime) : '';
                const endTime = item?.PermitEndTime ? spHelpers.toHHmm(item.PermitEndTime) : '';
                const persona = item.PIApprover ? spHelpers.toPersona(item.PIApprover) : undefined;
                permitsList.push({
                  id: String(item.Id),
                  type: item.PermitType,
                  date: item.PermitDate ? new Date(item.PermitDate).toISOString() : '',
                  startTime: startTime,
                  endTime: endTime,
                  orderRecord: item.RecordOrder,
                  isChecked: item.StatusRecord === 'new' ? true : false,
                  statusRecord: item.StatusRecord || undefined,
                  piApprover: persona ? persona : undefined,
                  piApprovalDate: item.PIApprovalDate ? new Date(item.PIApprovalDate) : undefined,
                  piStatus: item.PIStatus || undefined,
                  piApproverList: _piHsePartnerFilteredByCategory
                });
              }
            });
            setPermitPayload(permitsList.sort((a, b) => a.orderRecord - b.orderRecord));
          } else {
            setPermitPayload([]);
            setPermitPayloadValidityDays(0);
          }

          if (headerWorkflow) {

            const result: IPTWWorkflow = {
              id: headerWorkflow.Id !== undefined && headerWorkflow.Id !== null ? headerWorkflow.Id : undefined,
              PTWFormId: headerWorkflow.PTWForm?.Id !== undefined && headerWorkflow.PTWForm?.Id !== null ? headerWorkflow.PTWForm.Id : undefined,
              CoralReferenceNumber: headerWorkflow.PTWForm?.CoralReferenceNumber !== undefined && headerWorkflow.PTWForm?.CoralReferenceNumber !== null ? headerWorkflow.PTWForm.CoralReferenceNumber : undefined,
              POApprover: headerWorkflow.POApprover !== undefined && headerWorkflow.POApprover !== null ? toPersona(headerWorkflow.POApprover) : undefined,
              POApprovalDate: headerWorkflow.POApprovalDate !== undefined && headerWorkflow.POApprovalDate !== null ? headerWorkflow.POApprovalDate : undefined,
              POStatus: headerWorkflow.POStatus !== undefined && headerWorkflow.POStatus !== null ? headerWorkflow.POStatus : undefined,

              PAApprover: headerWorkflow.PAApprover !== undefined && headerWorkflow.PAApprover !== null ? toPersona(headerWorkflow.PAApprover) : undefined,
              PAApprovalDate: headerWorkflow.PAApprovalDate !== undefined && headerWorkflow.PAApprovalDate !== null ? headerWorkflow.PAApprovalDate : undefined,
              PAStatus: headerWorkflow.PAStatus !== undefined && headerWorkflow.PAStatus !== null ? headerWorkflow.PAStatus : undefined,

              PIApprover: headerWorkflow.PIApprover !== undefined && headerWorkflow.PIApprover !== null ? spHelpers.toPersona(headerWorkflow.PIApprover) : undefined,
              PIApprovalDate: headerWorkflow.PIApprovalDate !== undefined && headerWorkflow.PIApprovalDate !== null ? headerWorkflow.PIApprovalDate : undefined,
              PIStatus: headerWorkflow.PIStatus !== undefined && headerWorkflow.PIStatus !== null ? headerWorkflow.PIStatus : undefined,

              AssetDirectorApprover: headerWorkflow.AssetDirectorApprover !== undefined && headerWorkflow.AssetDirectorApprover !== null ? spHelpers.toPersona(headerWorkflow.AssetDirectorApprover) : undefined,
              AssetDirectorApprovalDate: headerWorkflow.AssetDirectorApprovalDate !== undefined && headerWorkflow.AssetDirectorApprovalDate !== null ? headerWorkflow.AssetDirectorApprovalDate : undefined,
              AssetDirectorStatus: headerWorkflow.AssetDirectorStatus !== undefined && headerWorkflow.AssetDirectorStatus !== null ? headerWorkflow.AssetDirectorStatus : undefined,

              UrgentAssetDirectorRejectionReas: headerWorkflow.UrgentAssetDirectorRejectionReas || '',
              UrgentAssetDirectorApprovalDate: headerWorkflow.UrgentAssetDirectorApprovalDate !== undefined && headerWorkflow.UrgentAssetDirectorApprovalDate !== null ? headerWorkflow.UrgentAssetDirectorApprovalDate : undefined,
              UrgentAssetDirectorStatus: headerWorkflow.UrgentAssetDirectorStatus !== undefined && headerWorkflow.UrgentAssetDirectorStatus !== null ? headerWorkflow.UrgentAssetDirectorStatus : undefined,

              HSEDirectorApprover: headerWorkflow.HSEDirectorApprover !== undefined && headerWorkflow.HSEDirectorApprover !== null ? spHelpers.toPersona(headerWorkflow.HSEDirectorApprover) : undefined,
              HSEDirectorApprovalDate: headerWorkflow.HSEDirectorApprovalDate !== undefined && headerWorkflow.HSEDirectorApprovalDate !== null ? headerWorkflow.HSEDirectorApprovalDate : undefined,
              HSEDirectorStatus: headerWorkflow.HSEDirectorStatus !== undefined && headerWorkflow.HSEDirectorStatus !== null ? headerWorkflow.HSEDirectorStatus : undefined,

              POClosureApprover: headerWorkflow.POClosureApprover !== undefined && headerWorkflow.POClosureApprover !== null ? toPersona(headerWorkflow.POClosureApprover) : undefined,
              POClosureDate: headerWorkflow.POClosureDate !== undefined && headerWorkflow.POClosureDate !== null ? headerWorkflow.POClosureDate : undefined,
              POClosureStatus: headerWorkflow.POClosureStatus !== undefined && headerWorkflow.POClosureStatus !== null ? headerWorkflow.POClosureStatus : undefined,
              POClosureRejectionReason: headerWorkflow.POClosureRejectionReason || '',

              AssetManagerApprover: headerWorkflow.AssetManagerApprover !== undefined && headerWorkflow.AssetManagerApprover !== null ? spHelpers.toPersona(headerWorkflow.AssetManagerApprover) : undefined,
              AssetManagerApprovalDate: headerWorkflow.AssetManagerApprovalDate !== undefined && headerWorkflow.AssetManagerApprovalDate !== null ? headerWorkflow.AssetManagerApprovalDate : undefined,
              AssetManagerStatus: headerWorkflow.AssetManagerStatus !== undefined && headerWorkflow.AssetManagerStatus !== null ? headerWorkflow.AssetManagerStatus : undefined,
              Stage: headerWorkflow.Stage !== undefined && headerWorkflow.Stage !== null ? headerWorkflow.Stage : undefined,

              IsAssetDirectorReplacer: headerWorkflow.IsAssetDirectorReplacer,
              IsHSEDirectorReplacer: headerWorkflow.IsHSEDirectorReplacer,

              AssetDirectorReplacer: headerWorkflow.AssetDirectorReplacer ? spHelpers.toPersona(headerWorkflow.AssetDirectorReplacer) : undefined,
              HSEDirectorReplacer: headerWorkflow.HSEDirectorReplacer ? spHelpers.toPersona(headerWorkflow.HSEDirectorReplacer) : undefined,

              PARejectionReason: headerWorkflow.PARejectionReason || '',
              PIRejectionReason: headerWorkflow.PIRejectionReason || '',
              AssetDirectorRejectionReason: headerWorkflow.AssetDirectorRejectionReason || '',
              HSEDirectorRejectionReason: headerWorkflow.HSEDirectorRejectionReason || '',


            };
            setWorkflowStage(result.Stage || undefined);

            if (!suppressAutoPrefill) {
              setPoDate(result.POApprovalDate ? new Date(result.POApprovalDate) : undefined);
              setPoStatus((result.POStatus as SignOffStatus) ?? undefined);

              setPaPicker(result.PAApprover ? [{ text: result.PAApprover.text || '', secondaryText: result.PAApprover.secondaryText || '', id: result.PAApprover.id || '' }] : []);
              setPaDate(result.PAApprovalDate ? new Date(result.PAApprovalDate) : undefined);
              setPaStatus((result.PAStatus as SignOffStatus) ?? undefined);

              setPiPicker(result.PIApprover ? [{ text: result.PIApprover.text || '', secondaryText: result.PIApprover.secondaryText || '', id: result.PIApprover.id || '' }] : []);
              setPiDate(result.PIApprovalDate ? new Date(result.PIApprovalDate) : undefined);
              setPiStatus((result.PIStatus as SignOffStatus) ?? undefined);

              setAssetDirPicker(result.AssetDirectorApprover ? [{ text: result.AssetDirectorApprover.text || '', secondaryText: result.AssetDirectorApprover.secondaryText || '', id: result.AssetDirectorApprover.id || '' }] : []);
              setAssetDirReplacerPicker(result.AssetDirectorReplacer ? [{ text: result.AssetDirectorReplacer.text || '', secondaryText: result.AssetDirectorReplacer.secondaryText || '', id: result.AssetDirectorReplacer.id || '' }] : []);
              setIsAssetDirectorReplacer(result.IsAssetDirectorReplacer);
              setAssetDirDate(result.AssetDirectorApprovalDate ? new Date(result.AssetDirectorApprovalDate) : undefined);
              setAssetDirStatus((result.AssetDirectorStatus as SignOffStatus) ?? undefined);

              setUrgentAssetDirDate(result.UrgentAssetDirectorApprovalDate ? new Date(result.UrgentAssetDirectorApprovalDate) : undefined);
              setUrgentAssetDirStatus((result.UrgentAssetDirectorStatus as SignOffStatus) ?? undefined);
              setUrgentAssetDirRejectionReas(result.UrgentAssetDirectorRejectionReas || '');

              setHseDirPicker(result.HSEDirectorApprover ? [{ text: result.HSEDirectorApprover.text || '', secondaryText: result.HSEDirectorApprover.secondaryText || '', id: result.HSEDirectorApprover.id || '' }] : []);
              setHseDirReplacerPicker(result.HSEDirectorReplacer ? [{ text: result.HSEDirectorReplacer.text || '', secondaryText: result.HSEDirectorReplacer.secondaryText || '', id: result.HSEDirectorReplacer.id || '' }] : []);
              setIsHseDirectorReplacer(result.IsHSEDirectorReplacer);
              setHseDirDate(result.HSEDirectorApprovalDate ? new Date(result.HSEDirectorApprovalDate) : undefined);
              setHseDirStatus((result.HSEDirectorStatus as SignOffStatus) ?? undefined);

              setClosurePoStatus((result.POClosureStatus as SignOffStatus) ?? undefined);
              setClosurePoDate(result.POClosureDate ? new Date(result.POClosureDate) : undefined);

              setClosureAssetManagerPicker(result.AssetManagerApprover ? [{ text: result.AssetManagerApprover.text || '', secondaryText: result.AssetManagerApprover.secondaryText || '', id: result.AssetManagerApprover.id || '' }] : []);
              setClosureAssetManagerDate(result.AssetManagerApprovalDate ? new Date(result.AssetManagerApprovalDate) : undefined);
              setClosureAssetManagerStatus((result.AssetManagerStatus as SignOffStatus) ?? undefined);
              setPORejectionReason(result.POClosureRejectionReason || '');

              setPaRejectionReason(result.PARejectionReason || '');
              setPiRejectionReason(result.PIRejectionReason || '');
              setAssetDirRejectionReason(result.AssetDirectorRejectionReason || '');
              setHseDirRejectionReason(result.HSEDirectorRejectionReason || '');
            }
          }
        }

        if (!cancelled) setPrefilledFormId(formId);
      } catch (e) {
        showBanner('An error occurred while loading the form for editing.', { autoHideMs: 5000, fade: true, kind: 'error' });
      }
    };

    load();

    return () => { cancelled = true; };
  }, [props.formId, prefilledFormId, loading, props.context, spHelpers]);

  const stageEnabled = React.useMemo(() => {
    const poEnabled = isPermitOriginator; // Originator signs first
    const paEnabled = (isPerformingAuthority && _poStatus.toString() !== 'Approved') || !isSubmitted;
    const piEnabled = (isPermitIssuer && (_paStatus || '').toString() !== 'Approved') || _piUnlockedByPA;
    // High risk requires AD then HSE; otherwise skip to closure after PI
    const assetDirEnabled = isHighRisk && isAssetDirector && _piStatus.toString() !== 'Approved';
    const hseDirEnabled = isHighRisk && isHSEDirector && _assetDirStatus.toString() !== 'Approved';
    const closureEnabled = isAssetManager && (
      (isHighRisk ? _hseDirStatus.toString() != 'Approved' : _piStatus.toString() != 'Approved')
    );
    return { poEnabled, paEnabled, piEnabled, assetDirEnabled, hseDirEnabled, closureEnabled };
  }, [
    isPermitOriginator, isPerformingAuthority, isPermitIssuer, isAssetDirector, isHSEDirector, isAssetManager,
    _poStatus, _paStatus, _piStatus, _assetDirStatus, _hseDirStatus, isHighRisk, _piUnlockedByPA, isSubmitted
  ]);

  const isPIPickerEnabled = React.useCallback((): boolean => {
    const stage = String(_workflowStage || '').toLowerCase();
    const selectedEmail = String(_piPicker?.[0]?.secondaryText || '').toLowerCase();
    const paApproved = String(_paStatus || '').toLowerCase() === 'approved';

    return (
      (isSubmitted && isUniquePermitIssuer && (stage === 'ApprovedFromPAToPI'.toLowerCase() || stage === 'ApprovedFromPOToPI'.toLowerCase()))
      || (isSubmitted && paApproved && stage === 'ApprovedFromPOToPA'.toLowerCase() && selectedEmail == '')
      || (!isSubmitted && _piUnlockedByPA)
      || (isSubmitted && !isIssued && isUniquePermitIssuer)
    );
  }, [_workflowStage, _piPicker, isUniquePermitIssuer, isSubmitted, _piUnlockedByPA, _paStatus]);

  const isPaStatusEnabled = React.useMemo(() => {
    const stage = String(_workflowStage || '').toLowerCase();
    const selectedEmail = String(_paPicker?.[0]?.secondaryText || '').toLowerCase();
    // const piIsSet = _workflowStage
    return (
      isSubmitted &&
      isPerformingAuthority &&
      (stage === 'ApprovedFromPOToPA'.toLowerCase() || stage === 'ApprovedFromPIToPA'.toLowerCase()) &&
      selectedEmail === currentUserEmail
    );
  }, [_workflowStage, _paPicker, currentUserEmail, isSubmitted, isPerformingAuthority]);

  const isPIStatusEnabled = React.useMemo(() => {
    const stage = String(_workflowStage || '').toLowerCase();
    const selectedEmail = String(_piPicker?.[0]?.secondaryText || '').toLowerCase();
    return (
      (isSubmitted &&
        isPermitIssuer &&
        selectedEmail === currentUserEmail &&
        (stage === 'ApprovedFromPAToPI'.toLowerCase() || stage === 'ApprovedFromPOToPI'.toLowerCase()))
      || (!isSubmitted && isPermitIssuer)
    );
  }, [_workflowStage, _piPicker, currentUserEmail, isSubmitted, isPermitIssuer]);

  // const isAssetDirectorStatusEnabled = React.useCallback((): boolean => {
  //   const stage = String(_workflowStage || '').toLowerCase();
  //   const selectedEmail = String(_assetDirPicker?.[0]?.secondaryText || '').toLowerCase();

  //   return (
  //     (isSubmitted && isAssetDirector &&
  //       selectedEmail === currentUserEmail &&
  //       stage === 'ApprovedFromAssetToHSE'.toLowerCase()) ||
  //     (_isUrgentSubmission && isAssetDirector) ||
  //     (isHighRisk && isAssetDirector && stage === 'ApprovedFromAssetToHSE'.toLowerCase())
  //   );

  // }, [_workflowStage, _assetDirPicker, currentUserEmail, isSubmitted, isAssetDirector, _isUrgentSubmission]);

  // const isHSEDirectorStatusEnabled = React.useCallback((): boolean => {
  //   const stage = String(_workflowStage || '').toLowerCase();

  //   return (
  //     (isSubmitted && isHSEDirector && stage === 'ApprovedFromAssetToHSE'.toLowerCase()) ||
  //     (_isUrgentSubmission && isHSEDirector) ||
  //     (isHighRisk && isHSEDirector && stage === 'ApprovedFromAssetToHSE'.toLowerCase())
  //   );

  // }, [_workflowStage, currentUserEmail, isSubmitted, isHSEDirector, _isUrgentSubmission]);

  // const showHighRiskApprovalSection = React.useMemo(() => {
  //   // isSubmitted && isHighRisk) || _isUrgentSubmission

  //   return (
  //     _isUrgentSubmission || (isSubmitted && isHighRisk ) || ()


  //   );
  // }, [_isUrgentSubmission, _piPicker, isSubmitted, isPermitIssuer]);

  React.useEffect(() => {
    const selectedEmail = String((_isAssetDirReplacer ? _assetDirReplacerPicker?.[0]?.secondaryText : _assetDirPicker?.[0]?.secondaryText) || '').toLowerCase();
    setAssetDirStatusEnabled(!!selectedEmail && selectedEmail === currentUserEmail);
  }, [_assetDirPicker, _assetDirReplacerPicker, _isAssetDirReplacer, currentUserEmail]);

  // HSE Director status enabled: main vs replacer
  React.useEffect(() => {
    const selectedEmail = String(
      (_isHseDirReplacer
        ? _hseDirReplacerPicker?.[0]?.secondaryText
        : _hseDirPicker?.[0]?.secondaryText) || ''
    ).toLowerCase();
    setHseDirStatusEnabled(!!selectedEmail && selectedEmail === currentUserEmail);
  }, [_hseDirPicker, _hseDirReplacerPicker, _isHseDirReplacer, currentUserEmail]);

  React.useEffect(() => {
    const selected = _selectedCompany?.id != null
      ? (ptwFormStructure?.companies || []).find(c => Number(c.id) === Number(_selectedCompany.id))
      : undefined;

    setCompanyLogoUrl(selected?.logoUrl ? String(selected.logoUrl) : initialLogoUrl);

    // Replace the first segment of docCode (e.g., COR-...) with first 3 letters of company title
    const prefix = (_selectedCompany?.title || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 3).toUpperCase();
    setDocCode(prev => {
      const base = prev && prev.includes('-') ? prev : 'COR-HSE-21-FOR-005';
      if (!prefix) return base;
      const parts = base.split('-');
      parts[0] = prefix;
      return parts.join('-');
    });
  }, [_selectedCompany, ptwFormStructure?.companies, initialLogoUrl]);

  const permitScheduleRows = React.useMemo(() => {
    if (isSubmitted) return _permitPayload; // show all returned permits
    const firstNew = _permitPayload.find(r => r.type === 'new');
    return firstNew ? [firstNew] : _permitPayload.slice(0, 1);
  }, [isSubmitted, _permitPayload]);

  React.useEffect(() => {
    if (!_piHsePartnerFilteredByCategory?.length) return;

    setPermitPayload(prev =>
      prev.map(r => {
        const isNew = String(r.type || '').toLowerCase() === 'new';
        const targetListRaw = isNew
          ? filterOutCurrentUser(_piHsePartnerFilteredByCategory)
          : _piHsePartnerFilteredByCategory;
        const targetList = Array.isArray(targetListRaw) ? targetListRaw : [];

        const current = Array.isArray(r.piApproverList) ? r.piApproverList : [];

        // shallow compare by id to avoid unnecessary updates
        const sameLength = current.length === targetList.length;
        const sameIds = sameLength && current.every((p, i) => String(p.id) === String(targetList[i].id));

        // Update if empty, undefined, or different from current
        if (!current.length || !sameIds) {
          return { ...r, piApproverList: targetList };
        }
        return r;
      })
    );
  }, [_piHsePartnerFilteredByCategory, filterOutCurrentUser]);

  const canRenewPermit = React.useMemo((): boolean => {
    const pOEMail = (_PermitOriginator?.[0]?.secondaryText || '').toLowerCase();
    const pIEMail = (_piPicker?.[0]?.secondaryText || '').toLowerCase();
    const assetDirectorEMail = (_assetDirPicker?.[0]?.secondaryText || '').toLowerCase();
    const hSEDirectorEMail = (_hseDirPicker?.[0]?.secondaryText || '').toLowerCase();
    const assetManagerEMail = (_closureAssetManagerPicker?.[0]?.secondaryText || '').toLowerCase();
    const loggedInUserIsPOEMail = currentUserEmail.toLowerCase() === pOEMail;
    const isUniquePO = loggedInUserIsPOEMail && (pOEMail !== (pIEMail || assetDirectorEMail || hSEDirectorEMail || assetManagerEMail));

    const rows = Array.isArray(_permitPayload) ? _permitPayload : [];
    const totalAllowed = Number(_permitPayloadValidityDays || 0);
    if (!rows.length || totalAllowed <= 0) return false;

    // Only rows that have a full schedule
    const filled = rows.filter(r =>
      (String(r.date || '').trim() &&
        String(r.startTime || '').trim() &&
        String(r.endTime || '').trim() && r.statusRecord === 'Closed')
      || (String(r.date || '').trim() &&
        String(r.startTime || '').trim() &&
        String(r.endTime || '').trim() && r.statusRecord?.toLowerCase() === 'new' && r.piApprovalDate === undefined || '')

    ).sort((a, b) => a.orderRecord - b.orderRecord);

    if (!filled.length) return false;

    // 1) Capacity: still have remaining permits to use
    const usedCount = filled.length; // or rows.length if placeholders count as used
    const hasRemainingCapacity = totalAllowed >= usedCount;

    // 2) Time: latest permit end time already passed
    const latest = filled.reduce((a, b) => (b.orderRecord > a.orderRecord ? b : a), filled[0]);
    const endDt = spHelpers.combineDateAndTime(latest.date.toString(), latest.endTime);
    const latestExpired = endDt instanceof Date && !isNaN(endDt.getTime()) && endDt.getTime() <= Date.now();

    return (hasRemainingCapacity && latestExpired && isPermitOriginator && isIssued && isUniquePO);

  }, [_permitPayload, _permitPayloadValidityDays, spHelpers, isPermitOriginator,
    isIssued, currentUserEmail, _PermitOriginator, _piPicker, _assetDirPicker, _hseDirPicker, _closureAssetManagerPicker]);

  // const canPOResubmitAfterRejection = React.useMemo((): boolean => {
  //   if (mode !== 'submitted' && !isPermitOriginator) return false;
  //   const stage = String(_workflowStage || '').toLowerCase();
  //   if (stage === 'rejected' && isPermitOriginator) {
  //     setCanPOResubmitAfterRejection(true);
  //     return true;
  //   } else {
  //     setCanPOResubmitAfterRejection(false);
  //     return false;
  //   }

  // }, [mode, isPermitOriginator, _workflowStage]);

  const poCanResubmit = React.useMemo(() => {
    if (mode !== 'submitted' && !isPermitOriginator) return false;
    const stage = String(_workflowStage || '').toLowerCase();
    return stage === 'rejected' && isPermitOriginator && isUniquePermitOriginator;
  }, [mode, isPermitOriginator, _workflowStage, isUniquePermitOriginator]);

  React.useEffect(() => {
    setCanPOResubmitAfterRejection(prev => (prev !== poCanResubmit ? poCanResubmit : prev));
  }, [poCanResubmit]);

  const _addNewPermit = React.useCallback(() => {
    // Allow only one renewal row in "New" status at a time
    const hasNewRenewal = _permitPayload.some(r =>
      r.type === 'renewal' && String(r.statusRecord || '').toLowerCase() === 'new'
    );
    if (hasNewRenewal) {
      showBanner('A renewal permit in "New" status already exists. Complete it before adding another.', {
        kind: 'warning', autoHideMs: 5000, fade: true
      });
      return;
    }

    // Guard: no validity days configured
    if (_permitPayloadValidityDays <= 0) {
      showBanner('Permit validity days not defined. Cannot add renewal.', { kind: 'warning', autoHideMs: 5000, fade: true });
      return;
    }

    // Guard: reached max allowed permits
    if (_permitPayload.length >= _permitPayloadValidityDays) {
      // showBanner(
      //   `Maximum number of permits (${_permitPayloadValidityDays}) reached. Create a new PTW form and reference ${_coralReferenceNumber} for continuation.`,
      //   { kind: 'warning', autoHideMs: 6000, fade: true }
      // );
      setShowExtendDialog(true);
      return;
    }

    // Require last existing permit to be fully populated before adding renewal
    const lastFilled = [..._permitPayload]
      .filter(r => r.date && r.startTime && r.endTime)
      .sort((a, b) => a.orderRecord - b.orderRecord)
      .pop();

    if (!lastFilled) {
      showBanner('Fill the current permit (date, start time, end time) before adding a renewal.', { kind: 'error', autoHideMs: 5000, fade: true });
      return;
    }

    // Check end time passed (optional business rule)
    const lastEnd = spHelpers.combineDateAndTime(lastFilled.date, lastFilled.endTime);
    if (lastEnd && lastEnd.getTime() > Date.now()) {
      showBanner('Current permit has not expired yet. Cannot add renewal.', { kind: 'warning', autoHideMs: 5000, fade: true });
      return;
    }
    // const hsePartners = assetCategoriesDetailsList?.filter(itm => itm.id == _selectedAssetDetails)[0].hsePartner || [];
    // Create new renewal row
    const nextOrder = _permitPayload.reduce((m, r) => Math.max(m, r.orderRecord), 0) + 1;
    const newRow: IPermitScheduleRow = {
      id: `permit-row-${nextOrder - 1}`,
      type: 'renewal',
      date: '',
      startTime: '',
      endTime: '',
      isChecked: false,
      orderRecord: nextOrder,
      statusRecord: 'New',
      piApprover: undefined,
      piApproverList: _piHsePartnerFilteredByCategory,
      piApprovalDate: undefined,
      piStatus: undefined
    };

    setPermitPayload(prev => [...prev, newRow]);
    showBanner('Renewal permit added. Please fill date and times.', { kind: 'success', autoHideMs: 4000, fade: true });
  }, [
    _permitPayload,
    _permitPayloadValidityDays,
    _coralReferenceNumber,
    spHelpers,
    showBanner,
    _piHsePartnerFilteredByCategory
  ]);

  const showRenewalButton = React.useMemo((): boolean => {
    if (mode !== 'submitted' || !isPermitOriginator) return false;
    // true when id has no letters (i.e., purely numeric -> no "text" in id)
    const isNumericId = (id: string) => /^[0-9]+$/.test(String(id || ''));
    const pOEMail = (_PermitOriginator?.[0]?.secondaryText || '').toLowerCase();
    const pIEMail = (_piPicker?.[0]?.secondaryText || '').toLowerCase();
    const assetDirectorEMail = (_assetDirPicker?.[0]?.secondaryText || '').toLowerCase();
    const hSEDirectorEMail = (_hseDirPicker?.[0]?.secondaryText || '').toLowerCase();
    const assetManagerEMail = (_closureAssetManagerPicker?.[0]?.secondaryText || '').toLowerCase();
    const loggedInUserIsPOEMail = currentUserEmail.toLowerCase() === pOEMail.toLowerCase();

    const isUniquePO = loggedInUserIsPOEMail && (pOEMail !== (pIEMail || assetDirectorEMail || hSEDirectorEMail || assetManagerEMail));

    const hasCandidate = (permitScheduleRows || []).some(r =>
      r.type === 'renewal' &&
      (r.statusRecord?.toLowerCase() === 'new' || r.statusRecord?.toLowerCase() === 'open') &&
      !r.piApprovalDate &&                  // no approval date
      !isNumericId(r.id)                   // id doesn't contain text
    );

    if (hasCandidate && permitScheduleRows.length <= _permitPayloadValidityDays && isUniquePO) {
      return true;
    }

    return false;
  }, [mode, isPermitOriginator, permitScheduleRows, _permitPayloadValidityDays, _PermitOriginator, _paPicker, _piPicker, _assetDirPicker, _hseDirPicker, _closureAssetManagerPicker]);

  const permitNeedsApproval = React.useMemo((): boolean => {
    if (mode !== 'submitted' || !isUniquePermitIssuer) return false;

    // true when id has no letters (i.e., purely numeric -> no "text" in id)
    const isNumericId = (id: string) => /^[0-9]+$/.test(String(id || ''));
    const permitNeedApproval = (permitScheduleRows || []).some(r =>
      r.type === 'renewal' && r.statusRecord?.toLowerCase() === 'new' && isNumericId(r.id)
    );

    return permitNeedApproval && isUniquePermitIssuer;
  }, [mode, permitScheduleRows, isUniquePermitIssuer]);

  const showPermitIssuerApprovalButton = React.useMemo((): boolean => {
    if (mode !== 'submitted' || !isPermitIssuer) return false;

    const pOEMail = (_PermitOriginator?.[0]?.secondaryText || '').toLowerCase();
    const pAEMail = (_paPicker?.[0]?.secondaryText || '').toLowerCase();
    const pIEMail = (_piPicker?.[0]?.secondaryText || '').toLowerCase();
    const assetDirectorEMail = (_assetDirPicker?.[0]?.secondaryText || '').toLowerCase();
    const hSEDirectorEMail = (_hseDirPicker?.[0]?.secondaryText || '').toLowerCase();
    const assetManagerEMail = (_closureAssetManagerPicker?.[0]?.secondaryText || '').toLowerCase();

    const loggedInUserIsPIEMail = currentUserEmail.toLowerCase() === pIEMail.toLowerCase();
    const loggedInUserIsHSEDirecotor = currentUserEmail.toLowerCase() === hSEDirectorEMail.toLowerCase();

    const isUniquePIOrHse = (loggedInUserIsPIEMail && (pIEMail !== (pOEMail || pAEMail || assetDirectorEMail || hSEDirectorEMail || assetManagerEMail))) ||
      (loggedInUserIsHSEDirecotor && (hSEDirectorEMail !== (pOEMail || pAEMail || pIEMail || assetDirectorEMail || assetManagerEMail)));

    // true when id has no letters (i.e., purely numeric -> no "text" in id)
    const isNumericId = (id: string) => /^[0-9]+$/.test(String(id || ''));
    const permitNeedApproval = (permitScheduleRows || []).some(r =>
      r.type === 'renewal' && r.statusRecord?.toLowerCase() === 'new' && isNumericId(r.id)
    );
    const approvedPermitsCount = permitScheduleRows.filter(r => r.piApprovalDate !== undefined &&
      r.piStatus !== undefined && isNumericId(r.id) && r.statusRecord?.toLowerCase() === 'closed').length;
    const completedApprovals = approvedPermitsCount === _permitPayloadValidityDays;

    return permitNeedApproval && isPermitIssuer && !completedApprovals && isUniquePIOrHse;
  }, [mode, permitScheduleRows, isPermitIssuer, _PermitOriginator, _paPicker, _piPicker, _assetDirPicker, _hseDirPicker, _closureAssetManagerPicker]);

  const showPOClosureSection = React.useMemo((): boolean => {
    if (mode !== 'submitted' || (!isUniquePermitOriginator && !isAssetManager)) return false;

    // true when id has no letters (i.e., purely numeric -> no "text" in id)
    const isNumericId = (id: string) => /^[0-9]+$/.test(String(id || ''));
    const approvedPermitsCount = permitScheduleRows.filter(r => r.piApprovalDate !== undefined &&
      r.piStatus !== undefined && r.piApprovalDate !== undefined && isNumericId(r.id) &&
      (r.statusRecord?.toLowerCase() === 'closed' || r.piStatus?.toLowerCase() === 'rejected')).length;

    const completedApprovals = approvedPermitsCount === _permitPayloadValidityDays;

    return (isUniquePermitOriginator && completedApprovals) || (isAssetManager && completedApprovals);
  }, [mode, permitScheduleRows, isUniquePermitOriginator, _permitPayloadValidityDays, isAssetManager]);

  const showConfirmButtonForPermitOriginator = React.useMemo((): boolean => {
    if (mode !== 'submitted' || (!isPermitOriginator)) return false;

    // true when id has no letters (i.e., purely numeric -> no "text" in id)
    const isNumericId = (id: string) => /^[0-9]+$/.test(String(id || ''));
    const closedAndRejectedPermits = permitScheduleRows.filter(r =>
      r.piStatus !== undefined && isNumericId(r.id) &&
      (r.statusRecord?.toLowerCase() === 'closed' || r.piStatus?.toLowerCase() === 'rejected')).length;

    const completedApprovals = closedAndRejectedPermits === _permitPayloadValidityDays;

    return (isPermitOriginator && completedApprovals);
  }, [mode, permitScheduleRows, isPermitOriginator, _permitPayloadValidityDays]);

  const showCancelPTWForm = React.useMemo((): boolean => {
    if (mode !== 'submitted' || !isPermitOriginator) return false;
    const stage = String(_workflowStage || '').toLowerCase();
    if (isPermitOriginator && isUniquePermitOriginator && !(stage === 'rejected' || stage === 'ClosedByPO'.toLowerCase() || stage === 'ClosedByAssetManager'.toLowerCase()
      || stage === 'Permanently Closed'.toLowerCase()) && !isIssued) {
      return true;
    }

    return false;
  }, [mode, isPermitOriginator, _workflowStage, isIssued, _PermitOriginator, _paPicker, _piPicker, _assetDirPicker, _hseDirPicker, _closureAssetManagerPicker]);

  const permitIssuerIsRejecting = React.useMemo((): boolean => {
    if (mode !== 'submitted' && !isPermitIssuer) return false;
    const piStatus = String(_piStatus || '').toLowerCase();
    if (piStatus === 'rejected') {
      return true;
    } else {
      return false;
    }

  }, [mode, isPermitIssuer, _piStatus]);

  const disableRiskControlsIssuedForm = React.useMemo((): boolean => {
    if ((isUniqueHSEDirector || isUniquePermitIssuer) && !isIssued) {
      if (_isUrgentSubmission && isHighRisk) {
        return true;
      } else {
        return true;
      }
    }
    return false;
  }, [isIssued, isUniqueHSEDirector, isUniquePermitIssuer, _isUrgentSubmission, isHighRisk]);

  // Small helpers: set only if not suppressed OR current selection is empty
  const safeSetPicker = React.useCallback(
    (current: IPersonaProps[] | undefined, setFn: (items: IPersonaProps[]) => void, items: IPersonaProps[] | undefined) => {
      if (suppressAutoPrefill) {
        if (!current || current.length === 0) setFn(items || []);
      } else {
        setFn(items || []);
      }
    }, [suppressAutoPrefill]);

  React.useEffect(() => {

    if (!_canPOResubmitAfterRejection || !isUniquePermitOriginator) {
      rejectionResetDoneRef.current = false; // allow future run when rejection happens again
      return;
    }
    if (rejectionResetDoneRef.current) return; // already executed once for current rejection
    rejectionResetDoneRef.current = true;

    // Lock any later auto-prefill from other effects
    setSuppressAutoPrefill(true);

    // Clear PA/PI so PO must reassign
    setPaPicker([]);
    setPaStatus('Pending');
    setPaDate(undefined);
    setPaRejectionReason('');

    setPiPicker([]);
    setPiStatus('Pending');
    setPiDate(undefined);
    setPiRejectionReason('');
    setPiUnlockedByPA(false);

    // If an Asset Detail is selected, prefill pickers/options from it
    const detail = (_selectedAssetDetails != null)
      ? (assetCategoriesDetailsList || []).find(d => Number(d.id) === Number(_selectedAssetDetails))
      : undefined;

    if (detail) {
      // Populate available PI options from HSE Partner; keep PI unselected so PO must choose
      setPiHsePartnerFilteredByCategory(detail.hsePartner || []);
      setAssetDirFilteredByCategory(detail.assetDirector || []);
      setAssetManagerFilteredByCategory(detail.assetManager || []);

      // Selections will be preserved empty due to suppressAutoPrefill; only fill if empty
      setAssetDirPicker([]);
      setAssetDirReplacerPicker([]);
      setHseDirPicker([]);
      setHseDirReplacerPicker([]);
      setClosureAssetManagerPicker([]);
    }

    if (isHighRisk || _isUrgentSubmission) {
      // Asset Director (main + replacer)
      setIsAssetDirectorReplacer(false);
      setAssetDirPicker([]);
      setAssetDirReplacerPicker([]);
      setAssetDirStatus('Pending');
      setAssetDirDate(undefined);
      setAssetDirRejectionReason('');

      // HSE Director (main + replacer)
      setIsHseDirectorReplacer(false);
      setHseDirPicker([]);
      setHseDirReplacerPicker([]);
      setHseDirStatus('Pending');
      setHseDirDate(undefined);
      setHseDirRejectionReason('');
    }

    showBanner(
      'This PTW was rejected. Please reassign Performing Authority and Permit Issuer. Director fields were refreshed from Asset Details.',
      { autoHideMs: 6000, fade: true, kind: 'warning' }
    );

  }, [_canPOResubmitAfterRejection, _selectedAssetDetails, assetCategoriesDetailsList, isHighRisk,
    _isUrgentSubmission, showBanner]);

  const today = React.useMemo(() => {
    const d = new Date();
    d.setHours(0, 0, 0, 0);
    return d;
  }, []);

  const validateNewPermitExtension = (item: any): string[] => {
    const errors: string[] = [];
    if (!item.date) errors.push("Date is required.");
    if (!item.startTime) errors.push("Start Time is required.");
    if (!item.endTime) errors.push("Expiry Time is required.");
    if (!item.piApproverId) errors.push("Permit Approver is required.");

    if (errors.length > 0) return errors;
    // Convert date and times to Date objects
    const startDateTime = new Date(`${item.date.split('T')[0]}T${item.startTime}`);
    const endDateTime = new Date(`${item.date.split('T')[0]}T${item.endTime}`);

    // Validate start < end
    if (startDateTime >= endDateTime) {
      errors.push("Start Time must be earlier than end Time.");
    }
    return errors;
  };

  const handleExtendClick = async () => {
    const item: any = {
      date: selectedDate?.toISOString(),
      startTime,
      endTime,
      piApproverId: selectedApprover
    };

    const errors = validateNewPermitExtension(item);
    if (errors.length > 0) {
      setErrors(errors);
      setIsTeachingBubbleVisible(true);
      return;
    }
    setIsTeachingBubbleVisible(false);
    setShowExtendDialog(false);
    setIsBusy(true);
    setBusyLabel('Creating PTW renewal...');
    try {
      const extendedFormId = await _extendCurrentPTW(Number(_PermitOriginator?.[0]?.id), item);
      if (extendedFormId) {
        // setSuccessMessage("PTW renewal has been successfully created.");
        goBackToHost();
        alert('PTW renewal has been successfully created.');
      }
    }
    catch (error) {
      setIsBusy(false);
      setShowExtendDialog(false);
      showBanner(error instanceof Error ? error.message : 'An unexpected error occurred. Reload and try again.', { kind: 'error', autoHideMs: 5000, fade: true });
    }
    finally {
      setIsBusy(false);
      setBusyLabel(''); // Hide loader
    }
  };

  const _extendCurrentPTW = React.useCallback(async (spOriginatorId?: number, item?: any): Promise<number> => {
    payloadRef.current = buildPayload();
    const payload = payloadRef.current;

    if (!payload) throw new Error('Form payload is not available');

    const spOps = spCrudRef.current ?? new SPCrudOperations((props.context as any).spHttpClient, webUrl, '', '');
    const originatorId = await spOps.ensureUserId(payload.originatorEMail || '');

    let toolboxTalkConductedById: number | undefined = undefined;
    if (payload.toolboxTalkConductedById) {
      toolboxTalkConductedById = await spOps.ensureUserId(payload.toolboxTalkConductedById || '');
    }
    const body: any = {
      PermitOriginatorId: originatorId ?? null,
      Title: 'PTW Form' + (originatorId ? ` - ${originatorId}` : ''),
      AssetID: payload.assetId ?? null,
      AssetCategoryId: payload.assetCategoryId ? Number(payload.assetCategoryId) : null,
      AssetDetailsId: payload.assetDetailsId ? Number(payload.assetDetailsId) : null,
      CompanyRecordId: payload.company?.id ? Number(payload.company.id) : null,
      ProjectTitle: payload.projectTitle ?? null,
      HACClassificationWorkAreaId: payload.hacWorkAreaId ? Number(payload.hacWorkAreaId) : null,
      WorkHazardsOthers: payload.workHazardsOtherText ?? null,
      ProtectiveSafetyEquipmentsOthers: payload.protectiveEquipmentsOtherText ?? null,
      PrecautionsOthers: payload.precautionsOtherText ?? null,
      FormStatusRecord: 'Submitted',
      WorkflowStatus: PTWWorkflowStatus.New,
      IsUrgentSubmission: payload.isUrgentSubmission === "" ? null : payload.isUrgentSubmission,
      PreviousReferenceNumber: payload.reference ?? null,
      RejectionReason: null,
      OverallRiskAssessment: payload.overallRiskAssessment ?? null,
      IsDetailedRiskAssessmentRequired: payload.detailedRiskAssessment === "" ? null : payload.detailedRiskAssessment,
      RiskAssessmentRefNumber: payload.detailedRiskAssessmentRef ?? null,
      GasTestRequired: payload.gasTestRequired === "" ? null : payload.gasTestRequired,
      GasTestResult: payload.gasTestResult ?? null,
      FireWatchNeeded: payload.fireWatchNeeded === "" ? null : payload.fireWatchNeeded,
      FireWatchAssigned: payload.fireWatchAssigned,
      AttachmentsProvided: payload.attachmentsProvided === "" ? null : payload.attachmentsProvided,
      AttachmentsProvidedDetails: payload.attachmentsDetails ?? '',
      ToolboxTalk: payload.toolboxTalk === "" ? null : payload.toolboxTalk,
      ToolBoxTalkDate: payload.toolboxTalkDate,
      ToolboxConductedById: toolboxTalkConductedById ?? null,
      ToolboxTalkHSEReference: payload.toolboxHSEReference ?? null,
      PermitsValidityDays: payload.permitPayloadValidityDays,
    };

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
      body['PersonnelInvolvedId'] = payload.personnelIds.map(Number);
    }

    spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form', '');
    const newId = await spCrudRef.current._insertItem(body);
    if (!newId) throw new Error('Failed to create PTW Form');

    try {
      const coralReferenceNumber = await spHelpers.assignCoralReferenceNumber(props.context.spHttpClient,
        webUrl, 'PTW_Form', { Id: Number(newId) }, payload.company?.title, 'PTW');
      if (!coralReferenceNumber) throw new Error('Failed to generate Coral Reference Number. Please try again later.');

      setCoralReferenceNumber(coralReferenceNumber);

      if (item) {
        const _renewalPermit: any = {
          PTWFormId: Number(newId),
          PermitType: 'new',
          PermitDate: item.date || null,
          PermitStartTime: spHelpers.combineDateAndTime(item.date.toString(), item.startTime) || null,
          PermitEndTime: spHelpers.combineDateAndTime(item.date.toString(), item.endTime) || null,
          RecordOrder: 1,
          StatusRecord: 'New',
          PIApproverId: item.piApproverId ? Number(item.piApproverId) : null,
          PIApprovalDate: null,
          PIStatus: null
        }

        const ops = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form_Work_Permits', '');
        const data = ops._insertItem(_renewalPermit);
        if (!data) throw new Error('Failed to create PTW Work Permits.');
      }

      const _createdWorkflow = await _createPTWFormApprovalWorkflow(Number(newId), spOriginatorId, 'extend');
      if (!_createdWorkflow) {
        throw new Error('Failed to create PTW Form Approval Workflow');
      }

      if (payload.workTaskLists?.length) {
        const _createdTask = await _createPTWTasksJobsDescriptions(Number(newId), payload.workTaskLists, undefined);

        if (!_createdTask?.length) {
          throw new Error('Failed to create PTW Tasks and Job Descriptions');
        }
      }

    } catch (e) {
      console.warn('Failed to extend PTW Form:', e);
    }

    return newId as number;
  }, [props.context.spHttpClient, buildPayload]);

  // const _deleteSavedForm = React.useCallback(async (mode: 'delete'): Promise<boolean> => {
  //   const editFormId = props.formId ? Number(props.formId) : undefined;

  //   if (!editFormId) {
  //     showBanner('Cannot delete the PTW form. Form ID is missing.', { autoHideMs: 5000, fade: true, kind: 'error' });
  //     return false;
  //   }
  //   try {

  //     setBusyLabel('Deleting related items…');
  //     setIsBusy(true);
  //     await new Promise((resolve) => setTimeout(resolve, 500)); 

  //     const spCrud1 = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form_Work_Permits', '');
  //     await spCrud1._deleteLookUPItems(editFormId, 'PTWForm');

  //     const spCrud2 = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form_Job_Descriptions', '');
  //     await spCrud2._deleteLookUPItems(editFormId, 'PTWForm');

  //     const spCrud3 = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form_Approval_Workflow', '');
  //     await spCrud3._deleteLookUPItems(editFormId, 'PTWForm');

  //     spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'PTW_Form', '');
  //     await spCrudRef.current._deleteItem(editFormId!);

  //     goBackToHost();
  //     return true;
  //   } catch (error) {
  //     console.error('Error deleting related items:', error);
  //     showBanner('Failed to delete PTW form and related items.', { autoHideMs: 5000, fade: true, kind: 'error' });
  //     return false;
  //   }
  //   finally {
  //     setIsBusy(false);
  //     setBusyLabel('');
  //   }
  // }, [props.context.spHttpClient, props.formId, showBanner]);

  // When we start submitting/updating, scroll to where the loader overlay is rendered
  React.useEffect(() => {
    if (!isExportingPdf) return;
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
  }, [isExportingPdf]);

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



  const lowestValidityPermit = React.useMemo(() => {

    const lowest = (_selectedPermitTypeList || [])
      .filter((pt): pt is IWorkCategory => pt != null && (pt as any).renewalValidity != null)
      .reduce<IWorkCategory | undefined>((min, pt) => {
        if (!min) return pt;
        const minVal = typeof min.renewalValidity === 'number' ? min.renewalValidity : Number.POSITIVE_INFINITY;
        const ptVal = typeof pt.renewalValidity === 'number' ? pt.renewalValidity : Number.POSITIVE_INFINITY;
        return minVal < ptVal ? min : pt;
      }, undefined);

    return lowest;
  }, [_selectedPermitTypeList]);


  if (loading) {
    return (
      <div className={styles.loadingContainer}>
        <Spinner label={"Preparing PTW form.. "} size={SpinnerSize.large} />
      </div>
    );
  }

  function onInputChange(input: string): string { const outlookRegEx = /<.*>/g; const emailAddress = outlookRegEx.exec(input); if (emailAddress && emailAddress[0]) return emailAddress[0].substring(1, emailAddress[0].length - 1); return input; }

  const options = _piHsePartnerFilteredByCategory.map((m: IPersonaProps) => ({
    key: String(m.id),
    text: m.title || m.text || ''
  }));

  return (
    <div style={{ position: 'relative' }} ref={containerRef} data-export-mode={exportMode ? 'true' : 'false'}>
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
          ref={overlayRef}
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
          <Spinner label="Preparing PDF…" size={SpinnerSize.large} />
        </div>
      )}

      <Dialog
        hidden={!showExtendDialog}
        onDismiss={() => setShowExtendDialog(false)}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: 'Extend Work Permit',
          subText: `Maximum number of permit(s) ${_permitPayloadValidityDays} are reached. ` +
            `Do you want to extend your current work permit? This will create a new PTW form with the same details and reference number ${_coralReferenceNumber}.`
        }}
        modalProps={{
          isBlocking: true,
          styles: { main: { maxWidth: 1000, minWidth: 800, width: 900 } },
        }}
      >
        <DialogContent>
          {/* Date */}
          <DatePicker
            label="Date"
            value={selectedDate}
            strings={defaultDatePickerStrings}
            onSelectDate={(date) => setSelectedDate(date || undefined)}
            minDate={today}
            allowTextInput={false}
            styles={datePickerBlackStyles}
          />

          {/* Start Time */}
          <TextField label="Starting Time" type="time" value={startTime} onChange={(_, val) => setStartTime(val || '')} step={60} />

          {/* Expiry Time */}
          <TextField label="Expiry Time" type="time" value={endTime} onChange={(_, val) => setEndTime(val || '')} step={60} />

          {/* Permit Approver */}
          <ComboBox
            label="Permit Approver"
            placeholder="Select Approver"
            options={options}
            selectedKey={selectedApprover}
            onChange={(_, opt) => setSelectedApprover(opt?.key as string)}
            useComboBoxAsMenuWidth
          />
        </DialogContent>
        <DialogFooter>
          <>
            <div ref={buttonRef}>
              <PrimaryButton
                text="Yes, Extend"
                onClick={handleExtendClick}
              />
            </div>

            {isTeachingBubbleVisible && (
              <TeachingBubble
                target={buttonRef.current}
                headline="Validation Errors"
                hasCloseButton
                onDismiss={() => setIsTeachingBubbleVisible(false)}>
                <ul style={{ margin: 0, paddingLeft: '20px' }}>
                  {errors.map((err, idx) => (
                    <li key={idx}>{err}</li>
                  ))}
                </ul>
              </TeachingBubble>
            )}
          </>

          <DefaultButton text="Cancel" onClick={() => setShowExtendDialog(false)} />
        </DialogFooter>
      </Dialog>


      <form id="ptwFormMain">
        <div id="formTitleSection" style={{ pageBreakAfter: exportMode ? 'always' : 'auto' }}>
          <div className={styles.ptwformHeader} >
            <div>
              <img src={companyLogoUrl} alt="Logo" className={styles.formLogo} />
            </div>
            <div className={styles.ptwFormTitleLogo}>

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

        <div id="formHeaderInfo" className={styles.formBody} style={{ pageBreakAfter: exportMode ? 'always' : 'auto' }}>
          {/* Administrative Note */}
          <div className={`row mb-1`} id="administrativeNoteDiv">
            <div className={`form-group col-md-6`}>
              <NormalPeoplePicker label={"Permit Originator"} onResolveSuggestions={_onFilterChanged} itemLimit={1}
                className={'ms-PeoplePicker'}
                key={'permitOriginator'}
                removeButtonAriaLabel={'Remove'}
                onInputChange={onInputChange}
                resolveDelay={150}
                styles={peoplePickerBlackStyles}
                selectedItems={_PermitOriginator}
                inputProps={{ placeholder: 'Enter name or email' }}
                pickerSuggestionsProps={suggestionProps}
                disabled={true}
              />
            </div>

          </div>

          <div className={`row mb-1`} >
            <div className={`form-group col-md-6`}>
              <TextField label="PTW Ref #" readOnly value={_coralReferenceNumber} />
            </div>

            {/* NEW: Previous PTW reference */}
            <div className={`form-group col-md-6`} id="previousPtwRefDiv">
              <TextField
                label="Previous PTW Ref #"
                value={_previousPtwRef}
                onChange={(_, v) => setPreviousPtwRef(v || '')}
                readOnly={true}
              />
            </div>
          </div>

          <div className='row' id="permitOriginatorDiv">
            <div className={`form-group col-md-6`}>
              <ComboBox
                label="Company"
                placeholder="Select a company"
                options={ptwFormStructure?.companies?.sort((a, b) => (a.orderRecord || 0) - (b.orderRecord || 0))
                  .map(c => ({ key: c.id, text: c.title || '', fullName: c.fullName || '' })) || []}
                selectedKey={_selectedCompany?.id}
                onChange={(_e, item) => setSelectedCompany(item ? {
                  id: Number(item.key), title: item.text, orderRecord: 0,
                  fullName: ptwFormStructure?.companies?.find(c => c.id === Number(item.key))?.fullName || ''
                } : undefined)}
                useComboBoxAsMenuWidth={true}
              />
            </div>
            <div className={`form-group col-md-6`}>
              <TextField
                label="Asset ID"
                value={_assetId}
                onChange={(_, newValue) => setAssetId(newValue || '')}
              />

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
                styles={!_selectedAssetCategory ? comboBoxBlackStyles : undefined}
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

          {
            exportMode ? null : (
              <div className="row" id="urgentToggleDiv">
                <div className="form-group col-md-12">
                  <Toggle
                    inlineLabel
                    label={`Urgent submission (bypass Submission Range Interval${_coralFormList?.SubmissionRangeInterval ? `: ${_coralFormList.SubmissionRangeInterval}h` : ''})`}
                    checked={_isUrgentSubmission}
                    onChange={(_, chk) => {
                      setIsUrgentSubmission(!!chk)
                    }}
                    disabled={uiDisabled(isSubmitted)}
                    styles={isSubmitted ? customToggleStyles : undefined}
                  />
                  <small className="text-muted">
                    Use only for urgent PTW forms that must be submitted earlier than the norm interval.
                  </small>
                </div>
              </div>
            )
          }
        </div>

        <div id="permitScheduleSectionContainer" style={{ pageBreakAfter: exportMode ? 'always' : 'auto' }}>
          <div className='row pb-3' id="permitScheduleSection">
            <PermitSchedule
              workCategories={ptwFormStructure?.workCategories?.sort((a, b) => a.orderRecord - b.orderRecord) || []}
              selectedPermitTypeList={_selectedPermitTypeList.sort((a, b) => a.orderRecord - b.orderRecord)}
              permitRows={permitScheduleRows}
              onPermitTypeChange={handlePermitTypeChange}
              onPermitRowUpdate={updatePermitRow}
              styles={styles}
              permitsValidityDays={_permitPayloadValidityDays}
              isPermitIssuer={permitNeedsApproval}
              piApproverList={_piHsePartnerFilteredByCategory}
              isIssued={isIssued}
              isSubmitted={isSubmitted}
              exportMode={exportMode}
            />
            {/* Action buttons under PermitSchedule */}
            {(() => {
              const showRenewActions = mode === 'submitted' && canRenewPermit && isPermitOriginator && permitScheduleRows.length > 0;
              if (!showRenewActions && !showRenewalButton && !isPermitOriginator) return null; // render nothing if no action applies
              return (
                <div className="col-md-12" id="permitScheduleActions"
                  style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 4 }}>
                  {showRenewalButton && (
                    <PrimaryButton
                      text="Renew Permit"
                      onClick={() => _renewPermit('renew')}
                      disabled={!showRenewalButton || isBusy}
                    />
                  )}

                  {canRenewPermit && (
                    <DefaultButton
                      iconProps={{ iconName: 'Add' }}
                      text="Add Renewal Permit"
                      onClick={() => _addNewPermit()}
                      styles={{ label: { fontWeight: 600 } as any }}
                      disabled={isBusy}
                    />
                  )}
                </div>
              );
            })()}
          </div>
        </div>

        <div id="formContentSection" style={{ pageBreakAfter: exportMode ? 'always' : 'auto' }}>
          {workPermitRequired && (
            <>
              <div className="row pb-3" id="hacClassificationWorkAreaSection" style={{ pageBreakAfter: exportMode ? 'always' : 'auto' }}>
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

              <div className="row pb-3" id="workHazardSection" style={{ pageBreakAfter: exportMode ? 'always' : 'auto' }}>
                <div>
                  <Label className={styles.ptwLabel}>Work Hazards</Label>
                  <div className="text-center pb-3">
                    <small className="text-muted" style={{ fontStyle: 'italic', fontSize: '0.8rem' }}>
                      if 3 or more working hazards, detailed job description/tasks shall be provided below.
                    </small>
                  </div>
                </div>

                <>
                  {exportMode ? (
                    <CheckBoxDistributerComponent id="workHazardsComponent"
                      optionList={ptwFormStructure?.workHazardosList?.filter(h => _selectedWorkHazardIds.has(Number(h.id)))
                        .sort((a, b) => (a.orderRecord || 0) - (b.orderRecord || 0)) || []}
                      selectedIds={Array.from(_selectedWorkHazardIds)}
                      onChange={(ids) => setSelectedWorkHazardIds(new Set(ids))}
                      othersTextValue={_workHazardsOtherText}
                      onOthersChange={(checked, othersText) => setWorkHazardsOtherText(othersText)}
                    />) :
                    (
                      <CheckBoxDistributerComponent id="workHazardsComponent"
                        optionList={ptwFormStructure?.workHazardosList || []}
                        selectedIds={Array.from(_selectedWorkHazardIds)}
                        onChange={(ids) => setSelectedWorkHazardIds(new Set(ids))}
                        othersTextValue={_workHazardsOtherText}
                        onOthersChange={(checked, othersText) => setWorkHazardsOtherText(othersText)}
                      />)
                  }
                </>

              </div>

              {_selectedWorkHazardIds.size >= 3 && (
                <div className="row pb-2" id="riskAssessmentListSection" style={{ pageBreakAfter: exportMode ? 'always' : 'auto' }}>
                  <div className="form-group col-md-12">
                    <RiskAssessmentList
                      initialRiskOptions={ptwFormStructure?.initialRisk || []}
                      residualRiskOptions={ptwFormStructure?.residualRisk || []}
                      safeguards={filteredSafeguards || []}
                      overallRiskOptions={ptwFormStructure?.overallRiskAssessment || []}
                      selectedOverallRisk={_overAllRiskAssessment || ''}
                      disableRiskControls={!disableRiskControlsIssuedForm}
                      defaultRows={_riskAssessmentsTasks?.sort((a, b) => a.orderRecord - b.orderRecord) || []}
                      onChange={handleRiskTasksChange}
                      onOverallRiskChange={handleOverallRiskChange}
                      onDetailedRiskChange={handleDetailedRiskChange}
                      onDetailedRiskRefChange={handleDetailedRiskRefChange}
                      l2Required={_detailedRiskAssessment}
                      l2Ref={_riskAssessmentReferenceNumber}
                      exportMode={exportMode}
                    />
                  </div>
                </div>
              )}

              <div className="row pb-3" id="precautionsSection" style={{ pageBreakAfter: exportMode ? 'always' : 'auto' }}>
                <div>
                  <Label className={styles.ptwLabel}>Precautions Required</Label>
                </div>

                <div className="form-group col-md-12">
                  <div className={styles.checkboxContainer}>
                    <>
                      {
                        exportMode ?
                          <CheckBoxDistributerComponent id="precautionsComponent"
                            optionList={ptwFormStructure?.precuationsItems?.filter(h => _selectedPrecautionIds.has(Number(h.id)))
                              .sort((a, b) => (a.orderRecord || 0) - (b.orderRecord || 0)) || []}
                            selectedIds={Array.from(_selectedPrecautionIds)}
                            onChange={(ids) => setSelectedPrecautionIds(new Set(ids))}
                            othersTextValue={_precautionsOtherText}
                            onOthersChange={(checked, othersText) => setPrecautionsOtherText(othersText)}
                          />
                          : (
                            <CheckBoxDistributerComponent id="precautionsComponent"
                              optionList={ptwFormStructure?.precuationsItems || []}
                              selectedIds={Array.from(_selectedPrecautionIds)}
                              onChange={(ids) => setSelectedPrecautionIds(new Set(ids))}
                              othersTextValue={_precautionsOtherText}
                              onOthersChange={(checked, othersText) => setPrecautionsOtherText(othersText)}
                            />
                          )
                      }
                    </>

                  </div>
                </div>
              </div>

              <Separator />
              <div className='row pb-3' id="gasTestFireWatchAttachmentsSection">
                {(isIssued || isUniquePermitIssuer || isUniqueHSEDirector) && (
                  <>
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
                                onChange={() => {
                                  setGasTestValue(prev => (prev === gas ? '' : gas));
                                  setGasTestResult('');
                                }
                                }
                                disabled={uiDisabled(!(isUniquePermitIssuer || isUniqueHSEDirector))}
                              />
                            </div>
                          ))}

                          <Label style={{ paddingRight: '10px' }}>Gas Test Result:</Label>
                        </div>
                        <div style={{ flex: '1' }}>
                          {(
                            () => {
                              const disabled = _gasTestValue !== 'Yes' && !(isUniquePermitIssuer || isUniqueHSEDirector);
                              return (
                                <TextField
                                  type="text" style={{ padding: '4px 6px', border: '1px solid #ccc', borderRadius: '4px' }}
                                  placeholder={!disabled ? "Enter result" : ''}
                                  disabled={uiDisabled(disabled)}
                                  value={_gasTestResult}
                                  onChange={(e, newValue) => setGasTestResult(newValue || '')}
                                  styles={disabled ? textFieldBlackStyles : undefined}
                                />
                              );
                            }
                          )()}
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
                                disabled={uiDisabled(!(isUniquePermitIssuer || isUniqueHSEDirector))}
                              />
                            </div>
                          ))}
                          <Label style={{ paddingRight: '10px' }}>Firewatch Assigned:</Label>
                        </div>
                        <div style={{ flex: '1' }}>

                          {(() => {
                            const disabled = _fireWatchValue !== 'Yes' && !(isUniquePermitIssuer || isUniqueHSEDirector)
                            return (
                              <TextField type="text" style={{ padding: '4px 6px', border: '1px solid #ccc', borderRadius: '4px' }}
                                placeholder={!disabled ? "Enter name" : ''}
                                disabled={uiDisabled(disabled)}
                                value={_fireWatchAssigned}
                                onChange={(e, newValue) => setFireWatchAssigned(newValue || '')}
                                styles={disabled ? textFieldBlackStyles : undefined}
                              />);
                          })()}

                        </div>
                      </div>
                    </div>
                  </>
                )}
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
                    {(
                      () => {
                        const disabled = _attachmentsValue !== 'Yes';
                        return (
                          <TextField type="text" style={{ padding: '4px 6px', border: '1px solid #ccc', borderRadius: '4px' }}
                            placeholder={!disabled ? "Enter detail" : ''}
                            disabled={uiDisabled(disabled)}
                            value={_attachmentsResult}
                            onChange={(e, newValue) => setAttachmentsResult(newValue || '')}
                            styles={disabled ? textFieldBlackStyles : undefined}
                          />
                        );
                      }
                    )()}
                  </div>
                </div>

              </div>
              <Separator />

              <div className="row pb-3" id="protectiveSafetyEquipmentSection" style={{ pageBreakAfter: exportMode ? 'always' : 'auto' }}>
                <div>
                  <Label className={styles.ptwLabel}>Protective & Safety Equipment</Label>
                </div>

                <div className="form-group col-md-12">
                  <div className={styles.checkboxContainer}>
                    {
                      exportMode ?
                        <CheckBoxDistributerComponent id="protectiveSafetyEquipmentComponent"
                          optionList={ptwFormStructure?.protectiveSafetyEquipments?.filter(p => _selectedProtectiveEquipmentIds.has(Number(p.id)))
                            .sort((a, b) => (a.orderRecord || 0) - (b.orderRecord || 0)) || []}
                          selectedIds={Array.from(_selectedProtectiveEquipmentIds)}
                          onChange={(ids) => setSelectedProtectiveEquipmentIds(new Set(ids))}
                          othersTextValue={_protectiveEquipmentsOtherText}
                          onOthersChange={(checked, othersText) => setProtectiveEquipmentsOtherText(othersText)}
                        />
                        : (
                          <CheckBoxDistributerComponent id="protectiveSafetyEquipmentComponent"
                            optionList={ptwFormStructure?.protectiveSafetyEquipments || []}
                            selectedIds={Array.from(_selectedProtectiveEquipmentIds)}
                            onChange={(ids) => setSelectedProtectiveEquipmentIds(new Set(ids))}
                            othersTextValue={_protectiveEquipmentsOtherText}
                            onOthersChange={(checked, othersText) => setProtectiveEquipmentsOtherText(othersText)}
                          />
                        )
                    }
                  </div>
                </div>
              </div>

              <div className='row pb-3' id="machineryToolsSection" style={{ pageBreakAfter: exportMode ? 'always' : 'auto' }}>
                <div>
                  <Label className={styles.ptwLabel}>Machinery Involved / Tools</Label>
                </div>
                <div className="form-group col-md-12">
                  {
                    exportMode ? null :
                      (<div className='col-md-12'>
                        <ComboBox
                          key={`machinery-${_selectedMachineryIds?.slice().sort((a, b) => a - b).join('_')}`}
                          placeholder="Select machinery/tools"
                          options={machineryOptions as any}
                          // selectedKey={_selectedMachineryIds}
                          onChange={onMachineryChange}
                          multiSelect
                          useComboBoxAsMenuWidth
                          styles={comboBoxBlackStyles}
                        />
                      </div>)
                  }

                  <div className='col-md-12'>
                    <div style={{ borderRadius: 4, padding: 8, marginTop: 8, width: '100%' }}>
                      {selectedMachinery?.length === 0 || selectedMachinery === undefined ? (
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
              <div className='row pb-3' id="personnelInvolvedSection" style={{ breakAfter: 'page' }}>
                <div>
                  <Label className={styles.ptwLabel}>Personnel Involved</Label>
                </div>
                <div className="form-group col-md-12">
                  {exportMode ? null :
                    (<ComboBox
                      key={`personnel-${_selectedPersonnelIds?.slice().sort((a, b) => a - b).join('_')}`}
                      placeholder="Select personnel"
                      options={personnelOptions as any}
                      onChange={onPersonnelChange}
                      // selectedKey={_selectedPersonnelIds}
                      multiSelect
                      useComboBoxAsMenuWidth
                      styles={comboBoxBlackStyles}
                    />)}
                  <div style={{ borderRadius: 4, padding: 8, marginTop: 8, width: '100%' }}>
                    {selectedPersonnel?.length === 0 || selectedPersonnel === undefined ? (
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

              <div className="row pb-3" id="InstructionsSection" style={{ breakAfter: 'page' }}>
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
              {(isIssued || isUniquePermitIssuer || isUniqueHSEDirector) && (
                <div className="row pb-3" id="toolboxTalkSection" style={{ breakAfter: exportMode ? 'page' : 'auto' }} >
                  <div className="col-md-3" style={{ display: 'flex', alignItems: 'center' }}>
                    {(() => {
                      const disabled = !(isUniquePermitIssuer || isUniqueHSEDirector)
                      return (<Checkbox
                        label="Toolbox Talk (TBT); complete details if applicable"
                        checked={!!_selectedToolboxTalk}
                        onChange={(_, chk) => {
                          const isChecked = !!chk;
                          setToolboxTalk(isChecked);
                          if (!isChecked) {
                            setToolboxConductedBy(undefined);
                            setTimeout(() => setToolboxConductedBy([]), 0);
                            setToolboxHSEReference('');
                            setToolboxTalkDate(undefined);
                          }
                        }}
                        disabled={uiDisabled(disabled)}
                        styles={disabled ? comboBoxBlackStyles : undefined}
                      />);
                    })()}

                  </div>

                  <div className="col-md-4">
                    <Label>Conducted By (Title)</Label>
                    <NormalPeoplePicker
                      onResolveSuggestions={_onFilterChanged}
                      itemLimit={1}
                      className={'ms-PeoplePicker'}
                      key={`toolboxConductedBy-${_selectedToolboxTalk ? 'on' : 'off'}-${_selectedToolboxConductedBy?.[0]?.id || 'none'}`}
                      removeButtonAriaLabel={'Remove'}
                      onInputChange={onInputChange}
                      resolveDelay={150}
                      styles={peoplePickerBlackStyles}
                      selectedItems={_selectedToolboxConductedBy || []}
                      onChange={(items) => setToolboxConductedBy(items && items.length ? items : [])}
                      inputProps={{ placeholder: 'Enter name or email' }}
                      pickerSuggestionsProps={suggestionProps}
                      disabled={!(_selectedToolboxTalk)}
                    />
                  </div>

                  <div className="col-md-3">
                    <Label>HSE TBT Reference</Label>
                    <TextField
                      placeholder="Enter reference"
                      value={String(_toolboxHSEReference || '')}
                      onChange={(_, v) => setToolboxHSEReference(v || '')}
                      disabled={!_selectedToolboxTalk}
                    />
                  </div>

                  <div className="col-md-2">
                    <Label>Date</Label>
                    <DatePicker
                      placeholder="Select date"
                      value={_selectedToolboxTalkDate}
                      onSelectDate={(date) => setToolboxTalkDate(date ?? undefined)}
                      disabled={!_selectedToolboxTalk}
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
                    readOnly={true}
                  />
                  <DatePicker
                    disabled={true}
                    placeholder="Select date"
                    value={_poDate ? _poDate : new Date()}
                    styles={datePickerBlackStyles}
                  />
                </div>

                {!_isUrgentSubmission && (
                  <>
                    {/* Performing Authority (PA) */}
                    <div className="col-md-4" style={{ padding: 8 }}>
                      <Label style={{ fontWeight: 600 }}>Performing Authority (PA)</Label>
                      <ComboBox
                        placeholder="Select Performing Authority"
                        disabled={!(stageEnabled.paEnabled || suppressAutoPrefill)}
                        options={getOptionsForGroup('PerformingAuthorityGroup')}
                        selectedKey={_paPicker?.[0]?.id || undefined}
                        onChange={onSingleApproverChange('PerformingAuthorityGroup', (items) => {
                          const selectedPersona = { id: items[0].id, displayName: items[0].text, secondaryText: items[0].secondaryText } as IPersonaProps;
                          setPaPicker(selectedPersona ? [selectedPersona] : []);
                        }, setPaStatusEnabled)}
                        useComboBoxAsMenuWidth
                        styles={comboBoxBlackStyles}
                        className={'pb-1'}
                      />
                      <DatePicker
                        disabled={true}
                        placeholder="Select date"
                        value={_paDate ? new Date(_paDate) : new Date()}
                        styles={datePickerBlackStyles}
                      />
                      <ComboBox
                        disabled={!(isPaStatusEnabled || suppressAutoPrefill)}
                        placeholder="Status"
                        options={statusOptions.filter(opt => opt.text.toLowerCase() !== 'closed')}
                        selectedKey={_paStatus}
                        useComboBoxAsMenuWidth
                        styles={comboBoxBlackStyles}
                        onChange={(_, opt) => {
                          setPaRejectionReason('');
                          setPaStatus((opt?.key as SignOffStatus) ?? 'Pending');
                          if (isSubmitted && ((opt?.key as SignOffStatus) === 'Pending' || (opt?.key as SignOffStatus) === 'Rejected')) {
                            setPiPicker([]);
                            setPiStatus('Pending');
                          }
                        }}
                      />
                      {/* Show reason only when Rejected */}
                      {_paStatus === 'Rejected' && (
                        <TextField
                          label="Rejection Reason"
                          placeholder="Enter reason for rejection"
                          value={_paRejectionReason}
                          onChange={(_, v) => setPaRejectionReason(v || '')}
                          required
                          autoAdjustHeight
                          rows={2}
                        // errorMessage={isPaStatusEnabled && !_paRejectionReason.trim() ? 'Rejection reason is required.' : undefined}
                        />
                      )}
                    </div>

                    {/* Permit Issuer (PI) */}
                    {!(_paStatus === 'Pending' || _paStatus === 'Rejected') && (

                      <div className="col-md-4" style={{ padding: 8 }}>
                        <Label style={{ fontWeight: 600 }}>Permit Issuer (PI)</Label>
                        <ComboBox
                          placeholder="Select Permit Issuer"
                          // disabled={!stageEnabled.piEnabled}
                          disabled={!(isPIPickerEnabled() || suppressAutoPrefill)}
                          options={_piHsePartnerFilteredByCategory?.map(m => ({
                            key: String(m.id),
                            text: m.title || m.text || ''
                          }))}
                          selectedKey={_piPicker?.[0]?.id || undefined}
                          onChange={onPermitIssuerChange((items) => setPiPicker(items), setPiStatusEnabled)}
                          useComboBoxAsMenuWidth
                          styles={!(isPIPickerEnabled() || suppressAutoPrefill) ? comboBoxBlackStyles : undefined}
                          className={'pb-1'}
                        />
                        <DatePicker
                          disabled={true}
                          placeholder="Select date"
                          value={_piDate ? new Date(_piDate) : new Date()}
                          styles={datePickerBlackStyles}
                        />
                        <ComboBox
                          disabled={!isPIStatusEnabled}
                          placeholder="Status"
                          options={statusOptions.filter(opt => opt.text.toLowerCase() !== 'closed')}
                          selectedKey={_piStatus}
                          onChange={(_, opt) => {
                            setPiRejectionReason('');
                            setPiStatus((opt?.key as SignOffStatus) ?? 'Pending');
                            if (opt && (opt?.key as SignOffStatus) !== 'Pending') {
                              setPiDate(new Date());
                            }
                          }
                          }
                          useComboBoxAsMenuWidth
                          styles={!isPIStatusEnabled ? comboBoxBlackStyles : undefined}
                        />
                        {/* Show reason only when Rejected */}
                        {_piStatus === 'Rejected' && (
                          <TextField
                            label="Rejection Reason"
                            placeholder="Enter reason for rejection"
                            value={_piRejectionReason}
                            onChange={(_, v) => setPiRejectionReason(v || '')}
                            required
                            autoAdjustHeight
                            rows={2}
                          />
                        )}
                      </div>
                    )}
                  </>
                )
                }
              </div>
              {/* )} */}

              {/* URGENT PTW Approval (if applicable) - visible when submitted and is urgent */}
              {(_isUrgentSubmission) && (
                <div className="row pb-3" id="urgentApprovalSection" style={{ border: '1px solid #c8c6c4', borderRadius: 4, background: '#e9edf7', pageBreakAfter: exportMode ? 'always' : 'auto' }}>

                  <div className="col-md-12" style={{ paddingTop: 8 }}>
                    <Label style={{ fontWeight: 600 }}>
                      Urgent PTW Approval
                    </Label>
                  </div>

                  <div className="col-md-6" style={{ padding: 8 }}>
                    {(() => {
                      const enabled = (isUniquePermitOriginator && !isIssued && _workflowStage === undefined);
                      return (
                        <Toggle
                          id='UrgentAssetDirector'
                          inlineLabel
                          label={_isAssetDirReplacer ? 'Asset Director' : 'Delegate Asset Director'}
                          checked={!!_isAssetDirReplacer}
                          onChange={(_, chk) => setIsAssetDirectorReplacer(!!chk)}
                          disabled={!enabled}
                        />
                      );
                    })()}

                    <Label style={{ fontWeight: 600 }}>{_isAssetDirReplacer ? 'Delegate Asset Director' : 'Asset Director'}</Label>
                    <NormalPeoplePicker
                      onResolveSuggestions={_onFilterChanged} itemLimit={1}
                      className={'ms-PeoplePicker pb-1'}
                      // key={_isAssetDirReplacer ? 'assetDirectorReplacer' : 'assetDirector'}
                      removeButtonAriaLabel={'Remove'}
                      onInputChange={onInputChange}
                      resolveDelay={150}
                      styles={peoplePickerBlackStyles}
                      selectedItems={
                        _isAssetDirReplacer
                          ? (_assetDirReplacerPicker?.[0]?.id ? _assetDirReplacerPicker : [])
                          : (_assetDirPicker?.[0]?.id ? _assetDirPicker : [])
                      }
                      pickerSuggestionsProps={suggestionProps}
                      disabled={true}
                    />
                    <DatePicker
                      disabled={true}
                      styles={datePickerBlackStyles}
                      placeholder="Select date"
                      value={_urgentAssetDirDate ? new Date(_urgentAssetDirDate) : new Date()}
                    />
                    {(() => {
                      const enabled = isAssetDirector && String(_workflowStage || '').toLowerCase() === 'ApprovedFromPOtoAssetUrgent'.toLowerCase();
                      // const isApprovedWhenHighRisk = isHighRisk && String(_workflowStage || '').toLowerCase() === 'ApprovedFromPOToPI'.toLowerCase() && _piStatus === 'Approved';
                      return (
                        <ComboBox
                          disabled={!enabled}
                          placeholder="Status"
                          options={statusOptions.filter(opt => opt.text.toLowerCase() !== 'Cancelled'.toLowerCase() && opt.text.toLowerCase() !== 'Closed'.toLowerCase())}
                          selectedKey={_urgentAssetDirStatus}
                          styles={comboBoxBlackStyles}
                          onChange={(_, opt) => {
                            setUrgentAssetDirRejectionReas('');
                            setUrgentAssetDirStatus((opt?.key as SignOffStatus) ?? 'Pending')
                          }
                          }
                          useComboBoxAsMenuWidth
                        />);
                    })()}

                    {_assetDirStatus === 'Rejected' && (
                      <TextField
                        label="Rejection Reason"
                        placeholder="Enter reason for rejection"
                        value={_urgentAssetDirRejectionReas}
                        onChange={(_, v) => setUrgentAssetDirRejectionReas(v || '')}
                        required
                        autoAdjustHeight
                        rows={2}
                      />
                    )}

                  </div>

                  {/* Permit Issuer (PI) */}
                  {isSubmitted && (
                    <div className="col-md-6" style={{ padding: 8 }}>
                      <div style={{ height: '38px' }}>
                      </div>
                      <Label style={{ fontWeight: 600 }}>Permit Issuer (PI)</Label>
                      <ComboBox
                        placeholder="Select Permit Issuer"
                        // disabled={!stageEnabled.piEnabled}
                        disabled={!(isPIPickerEnabled() || suppressAutoPrefill)}
                        options={_piHsePartnerFilteredByCategory?.map(m => ({
                          key: String(m.id),
                          text: m.title || m.text || ''
                        }))}
                        selectedKey={_piPicker?.[0]?.id || undefined}
                        onChange={onPermitIssuerChange((items) => setPiPicker(items), setPiStatusEnabled)}
                        useComboBoxAsMenuWidth
                        styles={comboBoxBlackStyles}
                        className={'pb-1'}
                      />
                      <DatePicker
                        disabled={true}
                        styles={datePickerBlackStyles}
                        placeholder="Select date"
                        value={_piDate ? new Date(_piDate) : new Date()}
                      />
                      <ComboBox
                        disabled={!isPIStatusEnabled}
                        placeholder="Status"
                        options={statusOptions.filter(opt => opt.text.toLowerCase() !== 'closed')}
                        selectedKey={_piStatus}
                        styles={comboBoxBlackStyles}
                        onChange={(_, opt) => {
                          setPiRejectionReason('');
                          setPiStatus((opt?.key as SignOffStatus) ?? 'Pending');
                          if (opt && (opt?.key as SignOffStatus) !== 'Pending') {
                            setPiDate(new Date());
                          }
                        }
                        }
                        useComboBoxAsMenuWidth
                      />


                      {/* Show reason only when Rejected */}
                      {_piStatus === 'Rejected' && (
                        <TextField
                          label="Rejection Reason"
                          placeholder="Enter reason for rejection"
                          value={_piRejectionReason}
                          onChange={(_, v) => setPiRejectionReason(v || '')}
                          required
                          autoAdjustHeight
                          rows={2}
                        />
                      )}
                    </div>
                  )}

                </div>
              )
              }

              {/* HIGH RISK PTW Approval (if applicable) - visible when submitted and overall risk is High */}
              {(isSubmitted && isHighRisk && !permitIssuerIsRejecting) && (
                <div className="row pb-3" id="highRiskApprovalSection" style={{ border: '1px solid #c8c6c4', borderRadius: 4, background: '#e9edf7', pageBreakAfter: exportMode ? 'always' : 'auto' }}>

                  <div className="col-md-12" style={{ paddingTop: 8 }}>
                    <Label style={{ fontWeight: 600 }}>
                      HIGH RISK PTW Approval
                    </Label>
                  </div>

                  <div className="col-md-6" style={{ padding: 8 }}>

                    {(() => {
                      const toogleAssetDirectorStatus = (isUniquePermitIssuer && !isIssued && !_isUrgentSubmission);
                      return (
                        <Toggle
                          id='HighRiskAssetDirector'
                          inlineLabel
                          label={_isAssetDirReplacer ? 'Asset Director' : 'Delegate Asset Director'}
                          checked={!!_isAssetDirReplacer}
                          onChange={(_, chk) => setIsAssetDirectorReplacer(!!chk)}
                          disabled={!toogleAssetDirectorStatus}
                        />
                      );
                    })()}

                    <Label style={{ fontWeight: 600 }}>{_isAssetDirReplacer ? 'Delegate Asset Director' : 'Asset Director'}</Label>
                    <NormalPeoplePicker
                      onResolveSuggestions={_onFilterChanged} itemLimit={1}
                      className={'ms-PeoplePicker pb-1'}
                      key={_isAssetDirReplacer ? 'assetDirectorReplacer' : 'assetDirector'}
                      removeButtonAriaLabel={'Remove'}
                      onInputChange={onInputChange}
                      resolveDelay={150}
                      styles={peoplePickerBlackStyles}
                      selectedItems={
                        _isAssetDirReplacer
                          ? (_assetDirReplacerPicker?.[0]?.id ? _assetDirReplacerPicker : [])
                          : (_assetDirPicker?.[0]?.id ? _assetDirPicker : [])
                      }
                      pickerSuggestionsProps={suggestionProps}
                      disabled={true}
                    />
                    <DatePicker
                      disabled={true}
                      styles={datePickerBlackStyles}
                      placeholder="Select date"
                      value={_assetDirDate ? new Date(_assetDirDate) : new Date()}
                    />
                    {(() => {
                      const enabled = isAssetDirector && String(_workflowStage || '').toLowerCase() === 'ApprovedFromPIToAsset'.toLowerCase();
                      return (
                        <ComboBox
                          disabled={!enabled}
                          placeholder="Status"
                          options={statusOptions.filter(opt => opt.text.toLowerCase() !== 'Cancelled'.toLowerCase() && opt.text.toLowerCase() !== 'Closed'.toLowerCase())}
                          selectedKey={_assetDirStatus}
                          styles={comboBoxBlackStyles}
                          onChange={(_, opt) => {
                            setAssetDirRejectionReason('');
                            setAssetDirStatus((opt?.key as SignOffStatus) ?? 'Pending')
                          }
                          }
                          useComboBoxAsMenuWidth
                        />);
                    })()}

                    {/* Show reason only when Rejected */}
                    {_assetDirStatus === 'Rejected' && (
                      <TextField
                        label="Rejection Reason"
                        placeholder="Enter reason for rejection"
                        value={_assetDirRejectionReason}
                        onChange={(_, v) => setAssetDirRejectionReason(v || '')}
                        required
                        autoAdjustHeight
                        rows={2}
                      />
                    )}

                  </div>

                  <div className="col-md-6" style={{ padding: 8 }}>
                    {(() => {
                      const toogleHSEDirectorStatus = (isUniqueHSEDirector && !isIssued);
                      return (
                        <Toggle
                          inlineLabel
                          label={_isHseDirReplacer ? 'HSE Director' : 'Delegate HSE Director'}
                          checked={!!_isHseDirReplacer}
                          onChange={(_, chk) => setIsHseDirectorReplacer(!!chk)}
                          disabled={!toogleHSEDirectorStatus}
                        />
                      );
                    })()}

                    <Label style={{ fontWeight: 600 }}>{_isHseDirReplacer ? 'Delegate HSE Director' : 'HSE Director'}</Label>
                    <NormalPeoplePicker
                      onResolveSuggestions={_onFilterChanged}
                      itemLimit={1}
                      className={'ms-PeoplePicker'}
                      key={_isHseDirReplacer ? 'hseDirectorReplacer' : 'hseDirector'}
                      removeButtonAriaLabel={'Remove'}
                      onInputChange={onInputChange}
                      resolveDelay={150}
                      styles={peoplePickerBlackStyles}
                      selectedItems={
                        _isHseDirReplacer
                          ? (_hseDirReplacerPicker?.[0]?.id ? _hseDirReplacerPicker : [])
                          : (_hseDirPicker?.[0]?.id ? _hseDirPicker : [])
                      }
                      pickerSuggestionsProps={suggestionProps}
                      disabled={true}
                    />
                    <DatePicker
                      disabled={true}
                      styles={datePickerBlackStyles}
                      placeholder="Select date"
                      value={_hseDirDate ? new Date(_hseDirDate) : new Date()}
                    />

                    {(() => {
                      const enabled = isHSEDirector && String(_workflowStage || '').toLowerCase() === 'approvedfromassettohse';
                      return (
                        <ComboBox
                          disabled={!enabled}
                          placeholder="Status"
                          options={statusOptions.filter(opt => opt.text.toLowerCase() !== 'Cancelled'.toLowerCase() && opt.text.toLowerCase() !== 'Closed'.toLowerCase())}
                          selectedKey={_hseDirStatus}
                          styles={!enabled ? comboBoxBlackStyles : undefined}
                          onChange={(_, opt) => {
                            setHseDirRejectionReason('');
                            setHseDirStatus((opt?.key as SignOffStatus) ?? 'Pending');
                          }}
                          useComboBoxAsMenuWidth
                        />
                      );
                    }
                    )()}

                    {/* Show reason only when Rejected */}
                    {_hseDirStatus === 'Rejected' && (
                      <TextField
                        label="Rejection Reason"
                        placeholder="Enter reason for rejection"
                        value={_hseDirRejectionReason}
                        onChange={(_, v) => setHseDirRejectionReason(v || '')}
                        required
                        autoAdjustHeight
                        rows={2}
                      />
                    )}
                  </div>
                </div>
              )}

              {/* PTW Closure */}
              {showPOClosureSection && (
                <div className="row pb-3" id="ptwClosureSection"
                  style={{ border: '1px solid #c8c6c4', borderRadius: 4, background: '#e9edf7', pageBreakAfter: exportMode ? 'always' : 'auto' }}>
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
                      readOnly={true}
                    />
                    <DatePicker
                      placeholder="Select date"
                      styles={datePickerBlackStyles}
                      value={_closurePoDate ? new Date(_closurePoDate) : new Date()}
                      disabled={true}
                    />

                    {(() => {
                      const enabled = isPermitOriginator && String(_workflowStage || '').toLowerCase() === 'Issued'.toLowerCase();
                      return (
                        <ComboBox
                          disabled={!enabled}
                          placeholder='Status'
                          options={statusOptions.filter(opt => opt.text.toLowerCase() === 'approved' || opt.text.toLowerCase() === 'pending' || opt.text.toLowerCase() === 'rejected')}
                          selectedKey={_closurePoStatus}
                          styles={!enabled ? comboBoxBlackStyles : undefined}
                          onChange={(_, opt) => {
                            setClosurePoStatus((opt?.key as SignOffStatus) ?? 'Pending');
                          }
                          }
                          useComboBoxAsMenuWidth
                        />
                      );
                    }
                    )()}

                    {_closurePoStatus === 'Rejected' && (
                      <TextField
                        label="Rejection Reason"
                        placeholder="Enter reason for rejection"
                        value={_poRejectionReason}
                        onChange={(_, v) => setPORejectionReason(v || '')}
                        required
                        autoAdjustHeight
                        rows={2}
                      />
                    )}
                  </div>

                  <div className="col-md-6" style={{ padding: 8 }}>
                    <Label style={{ fontWeight: 600 }}>Asset Manager</Label>

                    <ComboBox
                      placeholder="Select Asset Manager"
                      disabled={!stageEnabled.closureEnabled}
                      options={_assetManagerFilteredByCategory?.map(m => ({
                        key: String(m.id),
                        text: m.title || m.text || ''
                      }))}
                      selectedKey={_closureAssetManagerPicker?.[0]?.id || undefined}
                      // onChange={onAssetManagerChange}
                      useComboBoxAsMenuWidth
                      styles={comboBoxBlackStyles}
                      className={'pb-1'}
                    />

                    <DatePicker
                      disabled={true}
                      styles={datePickerBlackStyles}
                      placeholder="Select date"
                      value={_closureAssetManagerDate ? new Date(_closureAssetManagerDate) : new Date()}
                    />

                    {(() => {
                      const enabled = isAssetManager && String(_workflowStage || '').toLowerCase() === 'ClosedByPO'.toLowerCase();
                      return (
                        <ComboBox
                          disabled={!enabled}
                          placeholder='Status'
                          options={statusOptions.filter(opt => opt.text.toLowerCase() === 'approved' || opt.text.toLowerCase() === 'pending' || opt.text.toLowerCase() === 'rejected')}
                          selectedKey={_closureAssetManagerStatus}
                          onChange={(_, opt) => setClosureAssetManagerStatus((opt?.key as SignOffStatus) ?? 'Pending')}
                          useComboBoxAsMenuWidth
                          styles={!enabled ? comboBoxBlackStyles : undefined}
                        />
                      );
                    }
                    )()}

                    {/* Show reason only when Rejected */}
                    {_closureAssetManagerStatus === 'Rejected' && (
                      <TextField
                        label="Rejection Reason"
                        placeholder="Enter reason for rejection"
                        value={_asssetManagerRejectionReason}
                        onChange={(_, v) => setAssetManagerRejectionReason(v || '')}
                        required
                        autoAdjustHeight
                        rows={2}
                      />
                    )}

                  </div>
                </div>
              )}
            </>
          )}
        </div>

      </form >

      <Separator />

      {/* Bottom action buttons */}
      <div id="formButtonsSection" className="no-pdf" style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 8, marginBottom: 8 }}>
        <DefaultButton text="Close" onClick={handleCancel} />

        {showCancelPTWForm && isUniquePermitOriginator && (
          <DefaultButton
            text="Cancel PTW"
            onClick={() => cancelPTW('cancel')}
          />
        )}

        {(mode === "submitted" || mode === "rejected") && (
          <ExportPdfControls
            targetRef={containerRef}
            coralReferenceNumber={_coralReferenceNumber}
            originator={_PermitOriginator?.[0]?.text}
            exportMode={exportMode}
            onExportModeChange={setExportMode}
            onBusyChange={setIsExportingPdf}
            onError={(m) => showBanner(m)}
            docCode={docCode}
            docVersion='V04'
            companyName={_selectedCompany?.fullName}
            selectedWorkCategory={lowestValidityPermit ? lowestValidityPermit.title : ''}
          />
        )
        }

        {(mode === "new" || mode === "saved") &&
          <>
            {/* <DefaultButton text="Save"
              onClick={() => submitForm('save')}
              disabled={!isPermitOriginator || isBusy}
            /> */}

            <DefaultButton text="Submit"
              onClick={() => submitForm('submit')}
              disabled={!isPermitOriginator || isBusy}
            />
          </>
        }
        {(mode === "submitted") && _workflowStage?.toLowerCase() !== "ClosedByAssetManager".toLowerCase() &&
          _workflowStage?.toLowerCase() !== "Rejected".toLowerCase() && (
            (
              (isPerformingAuthority && (_workflowStage?.toLowerCase() == "ApprovedFromPOToPA".toLowerCase() || _workflowStage?.toLowerCase() == "ApprovedFromPIToPA".toLowerCase())) ||
              (isPermitIssuer && isUniquePermitIssuer && _workflowStage?.toLowerCase() == "ApprovedFromPAToPI".toLowerCase()) ||
              (isPermitIssuer && isUniquePermitIssuer && _workflowStage?.toLowerCase() == "ApprovedFromPOToPI".toLowerCase()) ||
              (isAssetDirector && _isUrgentSubmission && _workflowStage?.toLowerCase() == "ApprovedFromPOtoAssetUrgent".toLowerCase()) ||
              (isHSEDirector && isUniqueHSEDirector && _isUrgentSubmission && _workflowStage?.toLowerCase() == "ApprovedFromAssetToHSE".toLowerCase()) ||
              (isAssetDirector && _workflowStage?.toLowerCase() == "ApprovedFromPIToAsset".toLowerCase() && isHighRisk) ||
              (isHSEDirector && isUniqueHSEDirector && _workflowStage?.toLowerCase() == "ApprovedFromAssetToHSE".toLowerCase() && isHighRisk)
            ) && (
              <PrimaryButton id="approveFormWWithUpdate" text="Confirm" onClick={() => approveFormWWithUpdate('approve')} disabled={isBusy} />
            )
          )}

        {(mode === "submitted") && _workflowStage?.toLowerCase() !== "ClosedByAssetManager".toLowerCase()
          && _workflowStage?.toLowerCase() !== "Rejected".toLowerCase() &&
          ((isPermitOriginator && isUniquePermitOriginator && _workflowStage?.toLowerCase() == "ApprovedFromHSEToPO".toLowerCase()) ||
            (isPermitOriginator && isUniquePermitOriginator && showConfirmButtonForPermitOriginator) ||
            (isAssetManager && _workflowStage?.toLowerCase() == "ClosedByPO".toLowerCase()) ||
            (isAssetManager && _workflowStage?.toLowerCase() == "ApprovedFromPOtoAssetmanager".toLowerCase())
          ) &&
          (<PrimaryButton id="approveForm" text="Confirm" onClick={() => approveForm('approve')} disabled={isBusy} />)
        }

        {showPermitIssuerApprovalButton && !isUniquePermitOriginator && (
          <PrimaryButton
            text="Approve Renewal Permit"
            onClick={() => _approveRenewalPermit('approveRenewalPermit')}
            disabled={!isPermitIssuer || isBusy}
          />
        )}

        {poCanResubmit && (
          <PrimaryButton
            text="Resubmit PTW"
            onClick={() => resubmitAfterRejection('submitAfterRejection')}
          />
        )}

      </div>

      <div id="formFooterSection" className='row'>
        <div className='col-md-12 col-lg-12 col-xl-12 col-sm-12'>
          <DocumentMetaBanner docCode={docCode} version='V04' effectiveDate='06-AUG-2024' page={1} companyName={_selectedCompany?.fullName} />
        </div>
      </div>

    </div >
  );


}