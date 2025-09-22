import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ICommon, ISPListItem } from "./ICommon";

export interface IEmployeeProps extends ISPListItem{
  employeeID: number | undefined;
  fullName: string;
  jobTitle: ICommon | undefined;
  company: ICommon | undefined;
  division: ICommon | undefined;
  department: ICommon | undefined;
  employmentStatus: string | undefined;
  manager: IEmployeeProps | undefined;
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

export interface IEmployeesPPEItemsCriteria extends IEmployeeProps {
  safetyHelmet: string | undefined;
  reflectiveVest: string  | undefined;
  safetyShoes: string  | undefined;
  rainSuit?: ICommon | undefined;
  uniformCoveralls?: ICommon | undefined;
  uniformTop?: ICommon | undefined;
  uniformPants?: ICommon | undefined;
  winterJacket?: ICommon | undefined;
}