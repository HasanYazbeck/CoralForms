import { IEmployeeProps } from "../PpeForm/IEmployeeProps";
import { ILKPItemInstructionsForUse } from "../Common/ICommon";

export interface ICoralForm {
    id?: number;
    title?: string;
    arTitle?: string;
    hasInstructionsForUse?: boolean;
    hasWorkflow?: boolean;
    hostingWebPartPath?: string;
    submissionRangeInterval?: number;
}

export interface ILookupItem {
    id: number;
    title: string;
    orderRecord: number;
}

interface IAssetsDetails extends ILookupItem {
    assetCategoryId?: number;
}

interface IWorkCategory extends ILookupItem {
    priority?: string;
    renewalValidity?: number;
}

export interface IEmployeePeronellePassport extends IEmployeeProps {
    isHSEInductionCompleted?: boolean;
    hseInductionDate?: Date;
    hseInductionValidity?: number;
    isPTWCertified?: boolean;
    ptwCertificationDate?: Date;
    ptwCertificationValidity?: number;
    isHSETrained?: boolean;
    hseTrainingDate?: Date;
    hseTrainingValidity?: number;
    isFireFightingTrained?: boolean;
    fireFightingTrainingDate?: Date;
    fireFightingTrainingValidity?: number;
}

export interface IPTWForm {
    id?: string | number;
    coralForm?: ICoralForm;
    companies?: ILookupItem[];
    assetsCategories?: ILookupItem[];
    assetsDetails?: IAssetsDetails[];
    workCategories?: IWorkCategory[];
    hacWorkAreas?: ILookupItem[];
    workHazardosList?: ILookupItem[];
    machinaries?: ILookupItem[];
    precuationsItems?: ILookupItem[];
    protectiveSafetyEquipments?: ILookupItem[];
    attachmentsProvided?: string[];
    gasTestRequired?: string;
    fireWatchNeeded?: string;
    overallRiskAssessment?: string;
    initialRisk?: string[];
    residualRisk?: string[];
    personnalInvolved: IEmployeePeronellePassport[];
    issuanceInstrunctions: ILKPItemInstructionsForUse[];
}