import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IUser } from "../../../Interfaces/IUser";
import { IPersonaProps } from "@fluentui/react";
import { IPPEItem } from "../../../Interfaces/IPPEItem";
import { ICoralFormsList } from "../../../Interfaces/ICoralFormsList";
import { IPPEItemDetails } from "../../../Interfaces/IPPEItemDetails";

export interface IPpeFormWebPartProps {
  context: WebPartContext;
  Users: IUser[];
  IsLoading?: Boolean;
  ThemeColor: string | undefined;
  IsDarkTheme: boolean;
  HasTeamsContext: boolean;
  PPEItems: IPPEItem[];
  CoralFormsList: ICoralFormsList;
  PPEItemDetails: IPPEItemDetails[];
  // EnvironmentMessage: string;
  // userDisplayName: string;
  // titleBackgroundColor?: string;
  // buttonColor?: string;
  // buttonBorder?:string;
  // buttonBack?:string;

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
