import { SPListItem } from "./ICommon";
import { IPPEItem } from "./IPPEItem";

 export interface IPPEItemDetails extends SPListItem {
    PPEItem: IPPEItem | undefined;
    Types: string | undefined;
    Sizes: string [] | undefined;
}