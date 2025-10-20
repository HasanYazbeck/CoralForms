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
  MessageBar
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

export default function PTWForm(props: IPTWFormProps) {

  // Helpers and refs
  const spCrudRef = React.useRef<SPCrudOperations | undefined>(undefined);
  const spHelpers = React.useMemo(() => new SPHelpers(), []);
  const [_users, setUsers] = React.useState<IUser[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [_ptwFormStructure, setPTWFormStructure] = React.useState<IPTWForm>({ issuanceInstrunctions: [], personnalInvolved: [] });
  const [_itemInstructionsForUse, setItemInstructionsForUse] = React.useState<ILKPItemInstructionsForUse[]>([]);
  const [, setPersonnelInvolved] = React.useState<IEmployeePeronellePassport[]>([]);
  const [_assetDetails, setAssetDetails] = React.useState<IAssetCategoryDetails[]>([]);
  const [_safeguards, setSafeguards] = React.useState<ISagefaurdsItem[]>([]);
  const [_filteredSafeguards, setFilteredSafeguards] = React.useState<ISagefaurdsItem[]>([]);
  const [_selectedPermitTypeList, setSelectedPermitTypeList] = React.useState<IWorkCategory[]>([]);
  const [_permitPayload, setPermitPayload] = React.useState<IPermitScheduleRow[]>([]);
  // const webUrl = props.context.pageContext.web.absoluteUrl;

  // Form State
  const [_coralReferenceNumber, setCoralReferenceNumber] = React.useState<string>('');
  const [_PermitOriginator, setPermitOriginator] = React.useState<IPersonaProps[]>([]);
  const [projectTitle, setProjectTitle] = React.useState<string>('');
  const [assetId, setAssetId] = React.useState<string>('');
  const [selectedAssetCategory, setSelectedAssetCategory] = React.useState<string | number | undefined>(undefined);
  const [selectedAssetDetails, setSelectedAssetDetails] = React.useState<string | number | undefined>(undefined);
  const [gasTestValue, setGasTestValue] = React.useState('');
  const [gasTestResult, setGasTestResult] = React.useState('');
  const [fireWatchValue, setFireWatchValue] = React.useState('');
  const [fireWatchAssigned, setFireWatchAssigned] = React.useState('');
  const [selectedWorkHazardIds, setSelectedWorkHazardIds] = React.useState<Set<number>>(new Set());

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
        `&$expand=EmployeeRecord`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'Employee_Personelle_Passport', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: IEmployeePeronellePassport[] = [];
      data.forEach((obj: any) => {
        if (obj) {
          const temp: IEmployeePeronellePassport = {
            Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            fullName: obj.EmployeeId?.FullName !== undefined && obj.EmployeeId?.FullName !== null ? obj.EmployeeId.FullName : undefined,
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
      await _getPTWFormStructure();
      await _getAssetCategories();
      await _getAssetDetails();
      await _getPersonnelInvolved();
      await _getWorkSafeguards();

      if (_ptwFormStructure && _ptwFormStructure.id && _ptwFormStructure.coralForm?.hasInstructionsForUse) {
        await _getLKPItemInstructionsForUse('PTW');
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
    if (!_ptwFormStructure?.assetsCategories) return [];
    return _ptwFormStructure.assetsCategories.map(item => ({
      key: item.id,
      text: item.title || ''
    }));
  }, [_ptwFormStructure?.assetsCategories]);

  // Asset details options (filtered by selected category)
  const assetDetailsOptions: IDropdownOption[] = React.useMemo(() => {
    if (!_ptwFormStructure?.assetsDetails) return [];

    // If no category is selected, return all asset details
    if (!selectedAssetCategory) {
      return _ptwFormStructure.assetsDetails.map(item => ({
        key: item.id,
        text: item.title || ''
      }));
    }

    // / Filter asset details based on selected category
    // Note: This assumes there's a relationship between asset category and details
    // You may need to adjust this logic based on your data structure
    return _ptwFormStructure.assetsDetails
      .filter(item => item.assetCategoryId === selectedAssetCategory) // Adjust this condition based on your data structure
      .map(item => ({
        key: item.id,
        text: item.title || ''
      }));
  }, [_ptwFormStructure?.assetsDetails, selectedAssetCategory]);

  // Handle asset category change
  const onAssetCategoryChange = (event: React.FormEvent<IComboBox>, item: IDropdownOption | undefined): void => {
    setSelectedAssetCategory(item ? item.key : undefined);
    setSelectedAssetDetails(undefined);
  };

  // Handle asset details change
  const onAssetDetailsChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption | undefined): void => {
    setSelectedAssetDetails(item ? item.key : undefined);
  };

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
        setFilteredSafeguards((_safeguards || []).filter(s => s.workCategoryId !== undefined && selectedIds.has(s.workCategoryId)));
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
      setFilteredSafeguards((_safeguards || []).filter(s => s.workCategoryId !== undefined && ids.has(s.workCategoryId)));
    } else {
      setFilteredSafeguards(_safeguards || []);
    }
  }, [_safeguards, _selectedPermitTypeList]);

  const updatePermitRow = React.useCallback((rowId: string, field: string, value: string, checked: boolean) => {
    setPermitPayload((prevItems) =>
      prevItems.map((item) =>
        item.id === rowId ? { ...item, [field]: value, isChecked: checked } : item
      )
    );
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
                <span className={styles.formArTitle}>{_ptwFormStructure?.coralForm?.arTitle}</span>
                <span className={styles.formTitle}>{_ptwFormStructure?.coralForm?.title}</span>
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
                {/* 
                <TextField label="Ref#" value={_coralReferenceNumber}
                  styles={{ root: { color: '#000', fontWeight: 500 } }}
                  onChange={(_, newValue) => setCoralReferenceNumber(newValue || '')} /> */}
              </div>
            </div>
          </div>

          <div className='row'>
            <div className={`form-group col-md-6`}>
              <NormalPeoplePicker label={"Permit Originator"} onResolveSuggestions={_onFilterChanged} itemLimit={1}
                className={'ms-PeoplePicker'} key={'permitOriginator'} removeButtonAriaLabel={'Remove'}
                onInputChange={onInputChange} resolveDelay={150}
                styles={peoplePickerBlackStyles}
                // onChange={handlePermitOriginatorChange}
                selectedItems={_PermitOriginator}
                inputProps={{ placeholder: 'Enter name or email' }}
                pickerSuggestionsProps={suggestionProps}
                disabled={true}
              />
            </div>

            <div className={`form-group col-md-6`}>
              <TextField
                label="Asset ID"
                value={assetId}
                onChange={(_, newValue) => setAssetId(newValue || '')} />
            </div>
          </div>

          <div className={`row`}>
            <div className={`form-group col-md-6`}>
              <ComboBox
                label="Asset Category"
                placeholder="Select an asset category"
                options={assetCategoryOptions}
                selectedKey={selectedAssetCategory}
                onChange={(_e, ch) => onAssetCategoryChange(_e, ch)}
                // onChange={() => onAssetCategoryChange}
                styles={comboBoxBlackStyles}
                useComboBoxAsMenuWidth={true}
              />
            </div>
            <div className={`form-group col-md-6`}>
              <ComboBox
                label="Asset Details"
                placeholder="Select asset details"
                options={assetDetailsOptions}
                selectedKey={selectedAssetDetails}
                onChange={() => onAssetDetailsChange}
                disabled={!selectedAssetCategory}
                styles={comboBoxBlackStyles}
                useComboBoxAsMenuWidth={true}
              />
            </div>
          </div>

          <div className={`row`}>
            <div className={`form-group col-md-12`}>
              <TextField
                label="Project Title / Description"
                value={projectTitle}
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
              workCategories={_ptwFormStructure?.workCategories || []}
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
              optionList={_ptwFormStructure?.hacWorkAreas || []}
              colSpacing='col-2' />
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
              optionList={_ptwFormStructure?.workHazardosList || []}
              selectedIds={Array.from(selectedWorkHazardIds)}
              onChange={(ids) => setSelectedWorkHazardIds(new Set(ids))}
            />
          </div>

          {selectedWorkHazardIds.size >= 3 && (
            <div className="row pb-2" id="riskAssessmentListSection">
              <div className="form-group col-md-12">
                <RiskAssessmentList
                  initialRiskOptions={_ptwFormStructure?.initialRisk || []}
                  residualRiskOptions={_ptwFormStructure?.residualRisk || []}
                  safeguards={_filteredSafeguards || []}
                  overallRiskOptions={_ptwFormStructure?.overallRiskAssessment || []}
                // onChange={(state) => {
                //   setRiskAssessmentState(state);  
                //   // TODO: store this in your form state for submit
                //   // Example: setRiskAssessmentState(state);
                //   // state.rows, state.overallRisk, state.l2Required, state.l2Ref
                //   // console.log('RiskAssessmentList change', state);
                // }}
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
                <CheckBoxDistributerOnlyComponent id="precautionsComponent" optionList={_ptwFormStructure?.precuationsItems || []} />
              </div>
            </div>
          </div>

          <div className='row pb-3' id="gasTestAndFireWatch">

            {/* Gas Test Required Section */}
            <div className="form-group col-md-12 d-flex align-items-center mb-2">
              <Label className={`${styles.ptwLabel} me-3`} style={{ minWidth: '180px' }}>
                Gas Test Required
              </Label>

              <div className={styles.checkboxContainer}>
                {_ptwFormStructure?.gasTestRequired?.map((gas, i) => (
                  <div key={i} className={styles.checkboxItem}>
                    <Checkbox
                      label={gas}
                      checked={gasTestValue === gas}
                      onChange={() => setGasTestValue(gas)}
                    />
                  </div>
                ))}
              </div>

              <div className={`${styles.resultContainer} ms-4`}>
                <Label className="me-2">Gas Test Result:</Label>
                <input
                  type="text"
                  className={styles.resultInput}
                  placeholder="Enter result"
                  disabled={gasTestValue !== 'Yes'}
                  value={gasTestResult}
                  onChange={e => setGasTestResult(e.target.value)}
                />
              </div>
            </div>

            {/* Fire Watch Needed Section */}
            <div className='row'>

            </div>
            <div className="form-group col-md-12 d-flex align-items-center">
              <Label className={`me-3`} style={{ minWidth: '180px' }}>
                Fire Watch Needed
              </Label>

              <div className={styles.checkboxContainer}>
                {_ptwFormStructure?.fireWatchNeeded?.map((item, i) => (
                  <div key={i} className={styles.checkboxItem}>
                    <Checkbox
                      label={item}
                      checked={fireWatchValue === item}
                      onChange={() => setFireWatchValue(item)}
                    />
                  </div>
                ))}
              </div>

              <div className={`${styles.resultContainer} ms-4`}>
                <Label className="me-2">Firewatch Assigned:</Label>
                <input
                  type="text"
                  className={styles.resultInput}
                  placeholder="Enter name"
                  disabled={fireWatchValue !== 'Yes'}
                  value={fireWatchAssigned}
                  onChange={e => setFireWatchAssigned(e.target.value)}
                />
              </div>
            </div>
          </div>

          <div className="row pb-3" id="protectiveSafetyEquipmentSection" >
            <div>
              <Label className={styles.ptwLabel}>Protective & Safety Equipment</Label>
            </div>

            <div className="form-group col-md-12">
              <div className={styles.checkboxContainer}>
                <CheckBoxDistributerComponent id="protectiveSafetyEquipmentComponent" optionList={_ptwFormStructure?.protectiveSafetyEquipments || []} />
              </div>
            </div>
          </div>

          <div className="row pb-3" id="attachmentsProvidedSection">
            <div id="PdfInstructionsSegment">
              {/* Instructions For Use */}
              <Stack horizontal id="InstructionsStack">
                {_itemInstructionsForUse && _itemInstructionsForUse.length > 0 && (
                  <div style={{ marginTop: 12 }}>
                    <Label>Instructions for Use:</Label>
                    <div style={{ backgroundColor: "#f3f2f1", padding: 10, borderRadius: 4 }}>
                      {_itemInstructionsForUse.map((instr: ILKPItemInstructionsForUse, idx: number) => (
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
          </div>
        </div>
      </form >
    </div >
  );
}