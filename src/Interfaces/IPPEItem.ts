import { SPListItem } from './ICommon';
import { IPPEItemDetails } from './IPPEItemDetails';

export interface IPPEItem extends SPListItem {
    Required?: boolean | undefined;
    // hasInstructionForUse?: boolean | undefined;
    // hasWorkflow?: boolean | undefined;
    PPEDetails?: IPPEItemDetails[] | undefined;
    Brands?: string[] | undefined;
}

