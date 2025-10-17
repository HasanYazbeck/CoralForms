import * as React from 'react';
import type { IPTWFormProps } from './IPTWFormProps';

import styles from './PtwForm.module.scss';
import { IPersonaProps, Spinner, SpinnerSize } from '@fluentui/react';
import { IGraphResponse, IGraphUserResponse, ILKPItemInstructionsForUse } from '../../../Interfaces/Common/ICommon';
import { MSGraphClientV3 } from '@microsoft/sp-http-msgraph';
import { IUser } from '../../../Interfaces/Common/IUser';
import { SPCrudOperations } from "../../../Classes/SPCrudOperations";
import { SPHelpers } from "../../../Classes/SPHelpers";
import { ICoralForm, IEmployeePeronellePassport, ILookupItem, IPTWForm } from '../../../Interfaces/PtwForm/IPTWForm';

export default function PTWForm(props: IPTWFormProps) {

  // Helpers and refs
  const spCrudRef = React.useRef<SPCrudOperations | undefined>(undefined);
  const spHelpers = React.useMemo(() => new SPHelpers(), []);
  const [_users, setUsers] = React.useState<IUser[]>([]);
  const [_PermitOriginator, setPermitOriginator] = React.useState<IPersonaProps[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [_ptwFormStructure, setPTWFormStructure] = React.useState<IPTWForm>({ issuanceInstrunctions: [], personnalInvolved: [] });
  const [, setItemInstructionsForUse] = React.useState<ILKPItemInstructionsForUse[]>([]);
  const [, setPersonnelInvolved] = React.useState<IEmployeePeronellePassport[]>([]);
  // const webUrl = props.context.pageContext.web.absoluteUrl;

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
    `?$select=Id,AttachmentsProvided,InitialRisk,ResidualRisk,OverallRiskAssessment,FireWatchNeeded,GasTestRequired,PersonnelInvolvedId,` +
    `CoralFormId/Title,CoralFormId/ArabicTitle,` +
    `AssetCategory/Id,AssetCategory/Title,AssetCategory/OrderRecord,` +
    `AssetDetails/Id,AssetDetails/Title,AssetDetails/OrderRecord,` +
    `CompanyRecord/Id,CompanyRecord/Title,CompanyRecord/RecordOrder,` +
    `WorkCategory/Id,WorkCategory/Title,WorkCategory/OrderRecord,WorkCategory/RenewalValidity,` +
    `HACWorkArea/Id,HACWorkArea/Title,HACWorkArea/OrderRecord,` +
    `WorkHazards/Id,WorkHazards/Title,WorkHazards/OrderRecord,` +
    `Machinery/Id,Machinery/Title,Machinery/OrderRecord,` +
    `PrecuationItems/Id,PrecuationItems/Title,PrecuationItems/OrderRecord,` +
    `ProtectiveSafetyEquiment/Id,ProtectiveSafetyEquiment/Title,ProtectiveSafetyEquiment/OrderRecord` +
    `&$expand=CoralFormId,AssetCategory,AssetDetails,CompanyRecord,` +
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
        if (obj.CompanyRecord !== undefined && obj.CompanyRecord !== null) {
          _companies.push({ id: obj.CompanyRecord.Id, title: obj.CompanyRecord.Title, orderRecord: obj.CompanyRecord.OrderRecord || 0 });
        }

        const _assetsCategories: ILookupItem[] = [];
        if (obj.AssetsCategories !== undefined && obj.AssetsCategories !== null) {
          _assetsCategories.push({ id: obj.AssetsCategories.Id, title: obj.AssetsCategories.Title, orderRecord: obj.AssetsCategories.OrderRecord || 0 });
        }

        const _assetsDetails: ILookupItem[] = [];
        if (obj.AssetsDetails !== undefined && obj.AssetsDetails !== null) {
          _assetsDetails.push({ id: obj.AssetsDetails.Id, title: obj.AssetsDetails.Title, orderRecord: obj.AssetsDetails.OrderRecord || 0 });
        }

        const _workCategories: ILookupItem[] = [];
        if (obj.WorkCategory !== undefined && obj.WorkCategory !== null) {
          _workCategories.push({ id: obj.WorkCategory.Id, title: obj.WorkCategory.Title, orderRecord: obj.WorkCategory.OrderRecord || 0 });
        }

        const _hacWorkAreas: ILookupItem[] = [];
        if (obj.HACWorkArea !== undefined && obj.HACWorkArea !== null) {
          _hacWorkAreas.push({ id: obj.HACWorkArea.Id, title: obj.HACWorkArea.Title, orderRecord: obj.HACWorkArea.OrderRecord || 0 });
        }

        const _workHazardosList: ILookupItem[] = [];
        if (obj.WorkHazards !== undefined && obj.WorkHazards !== null) {
          _workHazardosList.push({ id: obj.WorkHazards.Id, title: obj.WorkHazards.Title, orderRecord: obj.WorkHazards.OrderRecord || 0 });
        }

        const _machineryList: ILookupItem[] = [];
        if (obj.Machinery !== undefined && obj.Machinery !== null) {
          _machineryList.push({ id: obj.Machinery.Id, title: obj.Machinery.Title, orderRecord: obj.Machinery.OrderRecord || 0 });
        }

        const _precuationsItemsList: ILookupItem[] = [];
        if (obj.PrecuationItems !== undefined && obj.PrecuationItems !== null) {
          _precuationsItemsList.push({ id: obj.PrecuationItems.Id, title: obj.PrecuationItems.Title, orderRecord: obj.PrecuationItems.OrderRecord || 0 });
        }

        const _protectiveSafetyEquipmentsList: ILookupItem[] = [];
        if (obj.ProtectiveSafetyEquiment !== undefined && obj.ProtectiveSafetyEquiment !== null) {
          _protectiveSafetyEquipmentsList.push({ id: obj.ProtectiveSafetyEquiment.Id, title: obj.ProtectiveSafetyEquiment.Title, orderRecord: obj.ProtectiveSafetyEquiment.OrderRecord || 0 });
        }

        result = {
          id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
          coralForm: coralForm, companies: _companies, assetsCategories: _assetsCategories,
          assetsDetails: _assetsDetails, workCategories: _workCategories, hacWorkAreas: _hacWorkAreas,
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
        };
        setPTWFormStructure(result);
      }

    } catch (error) {
      setPTWFormStructure({ issuanceInstrunctions: [], personnalInvolved: [] });
    }
  }, [props.context, spHelpers, spCrudRef, ptwStructureSelect]);

  const _getLKPItemInstructionsForUse = React.useCallback(async (formName?: string) => {
    try {
      const query: string = `?$select=Id,FormName,RecordOrder,Description&$expand=Author&$filter=substringof('${formName}', FormName)&$orderby=RecordOrder asc`;
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
      const query: string = `?$select=Id,EmployeeId/Id,EmployeeId/FullName,EmployeeId/EmailAddressRecord,IsHSEInductionCompleted,IsFireFightingTrained` +
        `&$expand=EmployeeId`;
      spCrudRef.current = new SPCrudOperations((props.context as any).spHttpClient, props.context.pageContext.web.absoluteUrl, 'LKP_Item_Instructions_For_Use', query);
      const data = await spCrudRef.current._getItemsWithQuery();
      const result: IEmployeePeronellePassport[] = [];
      data.forEach((obj: any) => {
        if (obj) {
          const temp: IEmployeePeronellePassport = {
            Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
            fullName: obj.EmployeeId?.FullName !== undefined && obj.EmployeeId?.FullName !== null ? obj.EmployeeId.FullName : undefined,
            EMailAddress: obj.EmployeeId?.EmailAddressRecord !== undefined && obj.EmployeeId?.EmailAddressRecord !== null ? obj.EmployeeId.EmailAddressRecord : undefined,
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

  // Initial load of users
  React.useEffect(() => {
    let cancelled = false;
    const load = async () => {
      setLoading(true);
      const fetchedUsers = await _getUsers();
      await _getPTWFormStructure();
      await _getPersonnelInvolved();

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

  return (
    <div style={{ position: 'relative' }}>
      <div />
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
        <div id="formHeader">
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
        <div>
          {/* <BannerComponent text={bannerText} kind={bannerOpts?.kind || 'error'}
            autoHideMs={bannerOpts?.autoHideMs} fade={bannerOpts?.fade} onDismiss={() => { setBannerText(undefined); setBannerOpts(undefined); }} /> */}
          {/* Form body will go here */}

        </div>
      </form>
    </div>
  );
}

