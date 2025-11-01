import { IEmployeeProps } from "../PpeForm/IEmployeeProps";
import { ILKPItemInstructionsForUse } from "../Common/ICommon";
import { IPersonaProps } from "@fluentui/react";

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

export interface IAssetsDetails extends ILookupItem {
    assetCategoryId: number;
    assetManager: IPersonaProps | undefined;
    assetDirector: IPersonaProps | undefined;
    hsePartner: IPersonaProps | undefined;
}

export interface IAssetCategoryDetails extends ILookupItem {
    assetsDetails?: IAssetsDetails[];
}

export interface IWorkCategory extends ILookupItem {
    priority?: string;
    renewalValidity?: number;
    isChecked: boolean;
}

export interface ISagefaurdsItem extends ILookupItem {
    workCategoryId: number;
    workCategoryTitle: string;
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
    gasTestRequired?: string[];
    fireWatchNeeded?: string[];
    overallRiskAssessment?: string[];
    initialRisk?: string[];
    residualRisk?: string[];
    personnalInvolved: IEmployeePeronellePassport[];
    issuanceInstrunctions: ILKPItemInstructionsForUse[];
}

export interface IPTWWorkflow {
    id: number | undefined;
    PTWFormId: number | undefined;
    CoralReferenceNumber: string | undefined;
    POApprover: IPersonaProps | undefined;
    POApprovalDate: Date | undefined;
    POStatus: "Approved" | "Rejected" | "Pending" | "Closed" | undefined;
    PAApprover: IPersonaProps | undefined;
    PAApprovalDate: Date | undefined;
    PAStatus: "Approved" | "Rejected" | "Pending" | "Closed" | undefined;
    PIApprover: IPersonaProps | undefined;
    PIApprovalDate: Date | undefined;
    PIStatus: "Approved" | "Rejected" | "Pending" | "Closed" | undefined;
    AssetDirectorApprover: IPersonaProps | undefined;
    AssetDirectorApprovalDate: Date | undefined;
    AssetDirectorStatus: "Approved" | "Rejected" | "Pending" | "Closed" | undefined;
    HSEDirectorApprover: IPersonaProps | undefined;
    HSEDirectorApprovalDate: Date | undefined;
    HSEDirectorStatus: "Approved" | "Rejected" | "Pending" | "Closed" | undefined;
    POClosureApprover: IPersonaProps | undefined;
    POClosureDate: Date | undefined;
    POClosureStatus: "Approved" | "Rejected" | "Pending" | "Closed" | undefined;
    AssetManagerApprover: IPersonaProps | undefined;
    AssetManagerApprovalDate: Date | undefined;
    AssetManagerStatus: "Approved" | "Rejected" | "Pending" | "Closed" | undefined;
}