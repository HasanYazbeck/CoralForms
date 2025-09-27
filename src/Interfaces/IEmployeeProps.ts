import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ICommon, ISPListItem } from "./ICommon";

export interface IEmployeeProps extends ISPListItem{
  employeeID?: number | undefined;
  fullName?: string;
  jobTitle?: ICommon | undefined;
  company?: ICommon | undefined;
  division?: ICommon | undefined;
  department?: ICommon | undefined;
  employmentStatus?: string | undefined;
  manager?: IEmployeeProps | undefined;
}

enum PPERequestReason {
  New = 1,
  Replacement = 2
}

export interface IPPEForm extends IEmployeeProps {
  context: WebPartContext
  requestorName: IEmployeeProps;
  dateRequested: Date;
  reasonOfRequest: PPERequestReason;
  replacementReason?: string;
}

export interface IEmployeesPPEItemsCriteria extends ISPListItem {
  employeeID?: number | undefined;
  fullName?: string | undefined;
  safetyHelmet?: string | undefined;
  reflectiveVest?: string  | undefined;
  safetyShoes?: string  | undefined;
  rainSuit?: string | undefined;
  uniformCoveralls?: string | undefined;
  uniformTop?: string | undefined;
  uniformPants?: string | undefined;
  winterJacket?: string | undefined;
}