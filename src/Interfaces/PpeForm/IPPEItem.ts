import { ISPListItem } from './../Common/ICommon';
import { IPPEItemDetails } from './IPPEItemDetails';

export interface IPPEItem extends ISPListItem {
    Brands?: string[] | undefined;
    PPEItemsDetails?: IPPEItemDetails[] | undefined;
}

