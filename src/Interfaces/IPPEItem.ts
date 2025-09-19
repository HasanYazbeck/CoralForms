import { SPListItem } from './ICommon';
import { IPPEItemDetails } from './IPPEItemDetails';

export interface IPPEItem extends SPListItem {
    IsRequired?: boolean | undefined;
    Brands?: string[] | undefined;
    PPEItemsDetails?: IPPEItemDetails[] | undefined;
}

