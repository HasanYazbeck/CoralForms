import { ISPListItem } from "./ICommon";

export interface ICoralFormsList extends ISPListItem {
    hasInstructionForUse?: boolean | undefined;
    hasWorkflow?: boolean | undefined;
}