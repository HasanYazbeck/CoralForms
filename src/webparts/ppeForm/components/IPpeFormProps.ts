import { WebPartContext } from "@microsoft/sp-webpart-base";
// import { IUser } from "../../../Interfaces/IUser";
import { IPersonaProps } from "@fluentui/react";
import { IPPEItem } from "../../../Interfaces/IPPEItem";
import { ICoralFormsList } from "../../../Interfaces/ICoralFormsList";
import { IPPEItemDetails } from "../../../Interfaces/IPPEItemDetails";

export interface IPpeFormWebPartProps {
  context: WebPartContext;
  ThemeColor: string | undefined;
  IsDarkTheme: boolean;
  HasTeamsContext: boolean;
  // Optional callbacks from host to control navigation
  onClose?: () => void;
  onSubmitted?: (newFormId?: number) => void;
  // Optional: when provided, the form opens in edit mode and loads this form
  formId?: number;
}

export interface IPpeFormWebPartState {
  // SelectedEmployeeId?: string;
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
