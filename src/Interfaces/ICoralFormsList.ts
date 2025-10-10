import { ISPListItem } from "./ICommon";

export interface ICoralFormsList extends ISPListItem {
    hasInstructionForUse?: boolean | undefined;
    hasWorkflow?: boolean | undefined;
    SubmissionRangeInterval?: number | undefined; // Number of days before a new form can be submitted

}