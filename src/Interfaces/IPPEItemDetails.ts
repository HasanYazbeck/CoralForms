import { ISPListItem } from "./ICommon";
import { IPPEItem } from "./IPPEItem";

 export interface IPPEItemDetails extends ISPListItem {
    PPEItem: IPPEItem | undefined;
    Sizes: string [] | undefined;
}