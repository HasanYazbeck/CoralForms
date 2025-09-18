import { SPListItem } from "./ICommon";

export interface ICoralFormsList extends SPListItem {
    hasInstructionForUse?: boolean | undefined;
    hasWorkflow?: boolean | undefined;
}