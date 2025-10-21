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
  IconButton
} from '@fluentui/react';
import { NormalPeoplePicker, IBasePickerSuggestionsProps, IBasePickerStyles } from '@fluentui/react/lib/Pickers';
import { IGraphResponse, IGraphUserResponse, ILKPItemInstructionsForUse } from '../../../Interfaces/Common/ICommon';
import { MSGraphClientV3 } from '@microsoft/sp-http-msgraph';
import { IUser } from '../../../Interfaces/Common/IUser';
import { SPCrudOperations } from "../../../Classes/SPCrudOperations";
import { SPHelpers } from "../../../Classes/SPHelpers";
import { IAssetCategoryDetails, IAssetsDetails, ICoralForm, IEmployeePeronellePassport, ILookupItem, IPTWForm, ISagefaurdsItem, IWorkCategory } from '../../../Interfaces/PtwForm/IPTWForm';
import { CheckBoxDistributerComponent } from './CheckBoxDistributerComponent';
import RiskAssessmentList from './RiskAssessmentList';
import { CheckBoxDistributerOnlyComponent } from './CheckBoxDistributerOnlyComponent';
import { DocumentMetaBanner } from '../../../Components/DocumentMetaBanner';
import { ICoralFormsList } from '../../../Interfaces/Common/ICoralFormsList';

export default function PTWForm(props: IPTWFormProps) {

  // Helpers and refs
  const formName = "Permit To Work";
  const spCrudRef = React.useRef<SPCrudOperations | undefined>(undefined);
  const spHelpers = React.useMemo(() => new SPHelpers(), []);
  const [_users, setUsers] = React.useState<IUser[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [, setCoralFormsList] = React.useState<ICoralFormsList>({ Id: "" });

  const [ptwFormStructure, setPTWFormStructure] = React.useState<IPTWForm>({ issuanceInstrunctions: [], personnalInvolved: [] });
  const [itemInstructionsForUse, setItemInstructionsForUse] = React.useState<ILKPItemInstructionsForUse[]>([]);
  const [personnelInvolved, setPersonnelInvolved] = React.useState<IEmployeePeronellePassport[]>([]);
  const [, setAssetDetails] = React.useState<IAssetCategoryDetails[]>([]);
  const [safeguards, setSafeguards] = React.useState<ISagefaurdsItem[]>([]);
  const [filteredSafeguards, setFilteredSafeguards] = React.useState<ISagefaurdsItem[]>([]);
  // const webUrl = props.context.pageContext.web.absoluteUrl;

  // Form State to used on update or submit
  const [_coralReferenceNumber, setCoralReferenceNumber] = React.useState<string>('');
  const [_PermitOriginator, setPermitOriginator] = React.useState<IPersonaProps[]>([]);
  const [_projectTitle, setProjectTitle] = React.useState<string>('');
  const [_assetId, setAssetId] = React.useState<string>('');
  const [_selectedAssetCategory, setSelectedAssetCategory] = React.useState<string | number | undefined>(undefined);
  const [_selectedAssetDetails, setSelectedAssetDetails] = React.useState<string | number | undefined>(undefined);
  const [_gasTestValue, setGasTestValue] = React.useState('');
  const [_gasTestResult, setGasTestResult] = React.useState('');
  const [_fireWatchValue, setFireWatchValue] = React.useState('');
  const [_fireWatchAssigned, setFireWatchAssigned] = React.useState('');
  const [_attachmentsValue, setAttachmentsValue] = React.useState('');
  const [_attachmentsResult, setAttachmentsResult] = React.useState('');
  const [_selectedWorkHazardIds, setSelectedWorkHazardIds] = React.useState<Set<number>>(new Set());
  const [_selectedPermitTypeList, setSelectedPermitTypeList] = React.useState<IWorkCategory[]>([]);
  const [_permitPayload, setPermitPayload] = React.useState<IPermitScheduleRow[]>([]);
  const [_selectedHacWorkAreaId, setSelectedHacWorkAreaId] = React.useState<number | undefined>(undefined);
  const [_selectedMachineryIds, setSelectedMachineryIds] = React.useState<number[]>([]);
  const [_selectedPersonnelIds, setSelectedPersonnelIds] = React.useState<number[]>([]);
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

  // ---------------------------
  // Data-loading functions (ported)
  // ---------------------------
  const _getUsers = React.useCallback(async (): Promise<IUser[]> => {
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
    `&$expand=CoralFormId,CompanyRecord,` +
    `WorkCategory,HACWorkArea,WorkHazards,Machinery,PrecuationItems,` +
    `ProtectiveSafetyEquiment`
  ), []);

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
      return result;
    } catch (error) {
      return [];
    }
  }, [props.context]);

  // Modified _getAssetDetails function
  const _getAssetDetails = React.useCallback(async () => {
    try {
      const query: string = `?$select=Id,Title,OrderRecord,` +
        `Manager/Id,Manager/EMail,` +
        `HSEPartner/Id,HSEPartner/EMail,` +
        `AssetCategoryRecord/Id,AssetCategoryRecord/Title` +
        `&$expand=AssetCategoryRecord,Manager,HSEPartner`;

      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'LKP_Asset_Details', query);
      const data = await spCrudRef.current._getItemsWithQuery();

      // Get asset categories first
      const categories = await _getAssetCategories();

      // Group asset details by category
      const categoriesWithDetails: IAssetCategoryDetails[] = [];
      // Create a map to group details by category ID
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
          };

          if (!detailsByCategory.has(categoryId)) {
            detailsByCategory.set(categoryId, []);
          }
          detailsByCategory.get(categoryId)!.push(assetDetail);
        }
      });

      // Create the final structure with categories and their associated details
      categories.forEach((category) => {
        if (category.id) {
          const categoryDetails = detailsByCategory.get(category.id as number) || [];

          const categoryWithDetails: IAssetCategoryDetails = {
            id: category.id,
            title: category.title,
            orderRecord: category.orderRecord,
            assetsDetails: categoryDetails,
          };

          categoriesWithDetails.push(categoryWithDetails);
        }
      });

      // Update the PTW form structure with both categories and all details
      setPTWFormStructure(prev => ({
        ...prev,
        assetsCategories: categories,
        assetsDetails: data.map((obj: any) => ({
          id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
          title: obj.Title !== undefined && obj.Title !== null ? obj.Title : undefined,
          orderRecord: obj.OrderRecord !== undefined && obj.OrderRecord !== null ? obj.OrderRecord : undefined,
          assetCategoryId: obj.AssetCategoryRecord?.Id !== undefined && obj.AssetCategoryRecord?.Id !== null ? obj.AssetCategoryRecord.Id : undefined,
        }))
      }));

      // Set the categorized asset details
      // setAssetDetails(categoriesWithDetails);

    } catch (error) {
      setAssetDetails([]);
      setPTWFormStructure(prev => ({
        ...prev,
        assetsCategories: [],
        assetsDetails: []
      }));
    }
  }, [props.context, _getAssetCategories]);

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
      const fetchedUsers = await _getUsers();
      const coralListResult = await _getCoralFormsList();
      await Promise.all([
        _getPTWFormStructure(),
        _getAssetCategories(),
        _getAssetDetails(),
        _getPersonnelInvolved(),
        _getWorkSafeguards(),
      ]);

      if (coralListResult && coralListResult?.hasInstructionForUse) {
        await _getLKPItemInstructionsForUse(formName);
      }

      console.log('Fetched users:', _users?.length > 0 ? _users[0].displayName : 'none');
      if (!cancelled) {
        try {
          const currentUserEmail = props.context.pageContext.user.email;
          const current = fetchedUsers.find(u => u.email === currentUserEmail);
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
  const _onFilterChanged = (filterText: string, currentPersonas: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> => {
    if (filterText) {
      let filteredPersonas: IPersonaProps[] = [];
      if (_users && _users.length > 0) {
        filteredPersonas = _users
          .filter(user =>
            user.displayName?.toLowerCase().includes(filterText.toLowerCase()) ||
            user.email?.toLowerCase().includes(filterText.toLowerCase())
          )
          .map(user => ({
            text: user.displayName || '',
            secondaryText: user.email || '',
            id: user.id
          }));
      }
      return filteredPersonas.filter(persona =>
        !currentPersonas.some(currentPersona => currentPersona.id === persona.id)
      );
    } else {
      return [];
    }
  };
  // Asset category options
  const assetCategoryOptions: IDropdownOption[] = React.useMemo(() => {
    if (!ptwFormStructure?.assetsCategories) return [];
    return ptwFormStructure.assetsCategories.map(item => ({
      key: item.id,
      text: item.title || ''
    }));
  }, [ptwFormStructure?.assetsCategories]);

  // Asset details options (filtered by selected category)
  const assetDetailsOptions: IDropdownOption[] = React.useMemo(() => {
    if (!ptwFormStructure?.assetsDetails) return [];

    // If no category is selected, return all asset details
    if (!_selectedAssetCategory) {
      return ptwFormStructure.assetsDetails.map(item => ({
        key: item.id,
        text: item.title || ''
      }));
    }

    // / Filter asset details based on selected category
    // Note: This assumes there's a relationship between asset category and details
    // You may need to adjust this logic based on your data structure
    return ptwFormStructure.assetsDetails
      .filter(item => item.assetCategoryId === _selectedAssetCategory) // Adjust this condition based on your data structure
      .map(item => ({
        key: item.id,
        text: item.title || ''
      }));
  }, [ptwFormStructure?.assetsDetails, _selectedAssetCategory]);

  // Handle asset category change
  const onAssetCategoryChange = (event: React.FormEvent<IComboBox>, item: IDropdownOption | undefined): void => {
    setSelectedAssetCategory(item ? item.key : undefined);
    setSelectedAssetDetails(undefined);
  };

  // Handle asset details change
  const onAssetDetailsChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption | undefined): void => {
    setSelectedAssetDetails(item ? item.key : undefined);
  };

  // Machinery/Tools - multi-select ComboBox wiring
  const machineryOptions = React.useMemo(() => {
    const items = ptwFormStructure?.machinaries || [];
    return items.map(m => ({ key: m.id, text: m.title, selected: _selectedMachineryIds.includes(Number(m.id)) }));
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
    return _selectedMachineryIds
      .map(id => byId.get(Number(id)))
      .filter((m): m is ILookupItem => !!m);
  }, [ptwFormStructure?.machinaries, _selectedMachineryIds]);

  const removeMachinery = React.useCallback((id: number) => {
    setSelectedMachineryIds(prev => prev.filter(x => x !== id));
  }, []);

  // Personnel Involved - multi-select ComboBox wiring
  const personnelOptions = React.useMemo(() => {
    return (personnelInvolved || []).map(p => ({
      key: p.Id,
      text: p.fullName || '',
      selected: _selectedPersonnelIds.includes(Number(p.Id))
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
    return _selectedPersonnelIds
      .map(id => byId.get(Number(id)))
      .filter((p): p is IEmployeePeronellePassport => !!p);
  }, [personnelInvolved, _selectedPersonnelIds]);

  const removePersonnel = React.useCallback((id: number) => {
    setSelectedPersonnelIds(prev => prev.filter(x => x !== id));
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
      } else {
        // Minimum number of renewals among selected categories
        const minRenewals = Math.min(...selectedItems.map(cat => (cat.renewalValidity ?? 0)));

        // Preserve any existing row values when possible
        const existingById = new Map(_permitPayload.map(r => [r.id, r] as const));

        const rows: IPermitScheduleRow[] = [];
        // Always include the New Permit row
        rows.push(
          existingById.get('permit-row-0') ?? {
            id: 'permit-row-0', type: 'new', date: '', startTime: '', endTime: '', isChecked: false
          }
        );

        // If renewalValidity indicates "renewable N times", render N renewal rows (1..N)
        for (let i = 1; i < minRenewals; i++) {
          const id = `permit-row-${i}`;
          rows.push(
            existingById.get(id) ?? {
              id, type: 'renewal', date: '', startTime: '', endTime: '', isChecked: false
            }
          );
        }

        setPermitPayload(rows);
      }

      return { ...prev, workCategories: nextWorkCategories } as IPTWForm;
    });
  }, [_permitPayload, setPTWFormStructure]);

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
    setPermitPayload(prevItems =>
      prevItems.map(item => {
        if (item.id !== rowId) return item;
        // Base update for the edited field and selection state
        const next = { ...item, [field]: value, isChecked: !!checked } as IPermitScheduleRow;
        // If the row was just deselected via the checkbox, clear the other inputs
        if (field === 'type' && !checked) {
          return { ...next, date: '', startTime: '', endTime: '' };
        }
        return next;
      })
    );
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

  // const showBanner = useCallback((text: string, opts?: { autoHideMs?: number; fade?: boolean, kind?: BannerKind }) => {
  //   setBannerText(text);
  //   setBannerTick(t => t + 1);
  //   setBannerOpts(opts);
  // }, []);

  // const hideBanner = useCallback(() => {
  //   showBanner(``);
  //   setBannerText(undefined);
  //   setBannerOpts(undefined);
  // }, []);

  // Navigate back to host list view (via callback or URL params)
  // const goBackToHost = useCallback(() => {
  //   if (typeof props.onClose === 'function') {
  //     props.onClose();
  //     return;
  //   }
  //   const url = new URL(window.location.href);
  //   url.searchParams.delete('mode');
  //   url.searchParams.delete('formId');
  //   window.location.href = url.toString();
  // }, [props.onClose]);

  // const handleCancel = useCallback(() => {
  //   goBackToHost();
  // }, [goBackToHost]);

  // When we start submitting/updating, scroll to where the loader overlay is rendered
  // useEffect(() => {
  //   if (!isSubmitting) return;
  //   // Wait for overlay to render, then scroll it into view
  //   requestAnimationFrame(() => {
  //     if (overlayRef.current && overlayRef.current.scrollIntoView) {
  //       try { overlayRef.current.scrollIntoView({ behavior: 'smooth', block: 'center' }); } catch { /* ignore */ }
  //     } else if (containerRef.current) {
  //       try { containerRef.current.scrollTo({ top: 0, behavior: 'smooth' }); } catch { /* ignore */ }
  //     } else {
  //       try { window.scrollTo({ top: 0, behavior: 'smooth' }); } catch { /* ignore */ }
  //     }
  //   });
  // }, [isSubmitting]);

  // Handlers


  // ---------------------------
  // Render
  // ---------------------------

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

    <div style={{ position: 'relative' }}>

      {/* {isSubmitting && !exportMode && (
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
      )} */}

      <form >
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

        <div id="formHeaderInfo" className={styles.formBody}>
          {/* Administrative Note */}
          <div>
            <div className={`form-group col-md-12 ${styles.adminNote}`}
              style={{ display: 'flex', justifyContent: 'space-between' }}>
              <Label>Grey areas are for administrative use only</Label>

              <div className={`form-group col-md-4`}>

                <TextField label="PTW Ref #" underlined disabled defaultValue={_coralReferenceNumber}
                  styles={{ root: { color: '#000', fontWeight: 500, backgroundColor: '#f4f4f4' } }}
                  onChange={(_, newValue) => setCoralReferenceNumber(newValue || '')} />
              </div>
            </div>
          </div>

          <div className='row'>
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
              <TextField
                label="Asset ID"
                value={_assetId}
                onChange={(_, newValue) => setAssetId(newValue || '')} />
            </div>
          </div>

          <div className={`row`}>
            <div className={`form-group col-md-6`}>
              <ComboBox
                label="Asset Category"
                placeholder="Select an asset category"
                options={assetCategoryOptions}
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
                onChange={() => onAssetDetailsChange}
                disabled={!_selectedAssetCategory}
                styles={comboBoxBlackStyles}
                useComboBoxAsMenuWidth={true}
              />
            </div>
          </div>

          <div className={`row`}>
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
              workCategories={ptwFormStructure?.workCategories || []}
              selectedPermitTypeList={_selectedPermitTypeList}
              permitRows={_permitPayload}
              onPermitTypeChange={handlePermitTypeChange}
              onPermitRowUpdate={updatePermitRow}
              styles={styles}
            />

          </div>

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
                <CheckBoxDistributerOnlyComponent id="precautionsComponent" optionList={ptwFormStructure?.precuationsItems || []} />
              </div>
            </div>
          </div>

          <div className='pb-3' id="gasTestAndFireWatch">
            <Label className={`${styles.ptwLabel} me-3`}>
              Gas Test Required
            </Label>
            {/* Gas Test Required Section */}
            <div className="form-group col-md-12 d-flex align-items-center mb-2" style={{ paddingLeft: '30px' }}>
              <div className={`col-md-3 ${styles.checkboxContainer}`}>
                {ptwFormStructure?.gasTestRequired?.map((gas, i) => (
                  <div key={i} className={styles.checkboxItem}>
                    <Checkbox
                      label={gas}
                      checked={_gasTestValue === gas}
                      onChange={() => setGasTestValue(gas)}
                    />
                  </div>
                ))}
              </div>

              <div className={`ms-4`} style={{ display: 'flex', alignItems: 'center', flex: '1', justifyContent: 'flex-end', paddingRight: '20px' }}>
                <Label style={{ paddingRight: ' 10px' }}>Gas Test Result:</Label>
                <TextField
                  type="text"
                  className={styles.resultInput}
                  placeholder="Enter result"
                  disabled={_gasTestValue !== 'Yes'}
                  value={_gasTestResult}
                  onChange={(e, newValue) => setGasTestResult(newValue || '')}
                />
              </div>
            </div>

            {/* Fire Watch Needed Section */}
            <div className=''>
              <Label className={`${styles.ptwLabel} me-3`}>Fire Watch Needed</Label>
              <div className="form-group col-md-12 d-flex align-items-center mb-2" style={{ paddingLeft: '30px' }}>
                <div className={`col-md-3 ${styles.checkboxContainer}`}>
                  {ptwFormStructure?.fireWatchNeeded?.map((item, i) => (
                    <div key={i} className={styles.checkboxItem}>
                      <Checkbox
                        label={item}
                        checked={_fireWatchValue === item}
                        onChange={() => setFireWatchValue(item)}
                      />
                    </div>
                  ))}
                </div>

                <div className={`ms-4`} style={{ display: 'flex', alignItems: 'center', flex: '1', justifyContent: 'flex-end', paddingRight: '20px' }}>
                  <Label style={{ paddingRight: ' 10px' }}>Firewatch Assigned:</Label>
                  <TextField className={styles.resultInput}
                    placeholder="Enter name"
                    disabled={_fireWatchValue !== 'Yes'}
                    value={_fireWatchAssigned}
                    onChange={(e, newValue) => setFireWatchAssigned(newValue || '')}
                  />
                </div>
              </div>
            </div>
          </div>

          <div className="row pb-3" id="protectiveSafetyEquipmentSection" >
            <div>
              <Label className={styles.ptwLabel}>Protective & Safety Equipment</Label>
            </div>

            <div className="form-group col-md-12">
              <div className={styles.checkboxContainer}>
                <CheckBoxDistributerComponent id="protectiveSafetyEquipmentComponent" optionList={ptwFormStructure?.protectiveSafetyEquipments || []} />
              </div>
            </div>
          </div>

          <div className='row pb-3' id="machineryToolsSection">
            <div>
              <Label className={styles.ptwLabel}>Machinery Involved / Tools</Label>
            </div>
            <div className="form-group col-md-12">
              <ComboBox
                key={`machinery-${_selectedMachineryIds.slice().sort((a, b) => a - b).join('_')}`}
                placeholder="Select machinery/tools"
                options={machineryOptions as any}
                onChange={onMachineryChange}
                multiSelect
                useComboBoxAsMenuWidth
                styles={comboBoxBlackStyles}
              />
            </div>

            <div style={{ border: '1px solid #e1e1e1', borderRadius: 4, padding: 8, marginTop: 8, width: '100%' }}>
              {selectedMachinery.length === 0 ? (
                <span style={{ color: '#605e5c', fontStyle: 'italic' }}>No machines selected</span>
              ) : (
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                  {selectedMachinery.map(m => (
                    <span key={m.id}
                      style={{
                        background: '#f3f2f1',
                        border: '1px solid #c8c6c4',
                        borderRadius: 12,
                        padding: '2px 6px',
                        display: 'inline-flex', alignItems: 'center', gap: 6
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

          <div className="row pb-3" id="attachmentsProvidedSection">
            <Label className={`${styles.ptwLabel} me-3`}>
              Attachment(s) provided
            </Label>
            {/* Gas Test Required Section */}
            <div className="form-group col-md-12 d-flex align-items-center mb-2" style={{ paddingLeft: '30px' }}>
              <div className={`col-md-3 ${styles.checkboxContainer}`}>
                {ptwFormStructure?.attachmentsProvided?.map((attachment, i) => (
                  <div key={i} className={styles.checkboxItem}>
                    <Checkbox
                      label={attachment}
                      checked={_attachmentsValue === attachment}
                      onChange={() => setAttachmentsValue(attachment)}
                    />
                  </div>
                ))}
              </div>

              <div className={`ms-4`} style={{ display: 'flex', alignItems: 'center', flex: '1', justifyContent: 'flex-end', paddingRight: '20px' }}>
                <Label style={{ paddingRight: ' 10px' }}>Details:</Label>
                <TextField
                  type="text"
                  className={styles.resultInput}
                  placeholder="Enter result"
                  disabled={_attachmentsValue !== 'Yes'}
                  value={_attachmentsResult}
                  onChange={(e, newValue) => setAttachmentsResult(newValue || '')}
                />
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

          {/* Personnel Involved - placed under Attachments section */}
          <div className='row pb-3' id="personnelInvolvedSection">
            <div>
              <Label className={styles.ptwLabel}>Personnel Involved</Label>
            </div>
            <div className="form-group col-md-12">
              <ComboBox
                key={`personnel-${_selectedPersonnelIds.slice().sort((a, b) => a - b).join('_')}`}
                placeholder="Select personnel"
                options={personnelOptions as any}
                onChange={onPersonnelChange}
                multiSelect
                useComboBoxAsMenuWidth
                styles={comboBoxBlackStyles}
              />
              <div style={{ border: '1px solid #e1e1e1', borderRadius: 4, padding: 8, marginTop: 8 }}>
                {selectedPersonnel.length === 0 ? (
                  <span style={{ color: '#605e5c', fontStyle: 'italic' }}>No personnel selected</span>
                ) : (
                  <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                    {selectedPersonnel.map(p => (
                      <span key={p.Id}
                        style={{ background: '#f3f2f1', border: '1px solid #c8c6c4', borderRadius: 12, padding: '2px 6px', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
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

        </div>
        <div id="formFooterSection" className='row'>
          <div className='col-md-12 col-lg-12 col-xl-12 col-sm-12'>
            <DocumentMetaBanner docCode='COR-HSE-21-FOR-005' version='V04' effectiveDate='06-AUG-2024' />
          </div>
        </div>
      </form>
    </div>
  );
}