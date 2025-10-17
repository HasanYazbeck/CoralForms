import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IPersonaProps } from "@fluentui/react";
import { IPPEItem } from "../../../Interfaces/PpeForm/IPPEItem";
import { ICoralFormsList } from "../../../Interfaces/Common/ICoralFormsList";
import { IPPEItemDetails } from "../../../Interfaces/PpeForm/IPPEItemDetails";
import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls";

export interface IPpeFormWebPartProps {
  context: WebPartContext;
  ThemeColor: string | undefined;
  IsDarkTheme: boolean;
  HasTeamsContext: boolean;
  onClose?: () => void;
  onSubmitted?: (newFormId?: number) => void;
  formId?: number;
  useTargetAudience: boolean;
  targetAudience: IPropertyFieldGroupOrPerson[];
}

export interface IPpeFormWebPartState {
  JobTitle: string;
  Department: string;
  Division: string;
  Company: string;
  Employee: IPersonaProps[];
  EmployeeId: string | undefined;
  Submitter: IPersonaProps[],
  Requester: IPersonaProps[],
  isReplacementChecked: boolean,
  PPEItems: IPPEItem[];
  CoralFormsList: ICoralFormsList;
  PPEItemsDetails: IPPEItemDetails[];
  PPEItemsRows?: { Item: string; Brands?: string; Required: boolean; Details: string; Qty: string; Size: string; Selected?: boolean }[];
}
