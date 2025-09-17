import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IUser } from "../../../Interfaces/IUser";
import { IPersonaProps } from "@fluentui/react";

export interface IPpeFormWebPartProps {
  Context: WebPartContext;
  Users: IUser[];
  IsLoading?: Boolean;
  ThemeColor: string | undefined;
  IsDarkTheme: boolean;
  HasTeamsContext: boolean;

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
}
