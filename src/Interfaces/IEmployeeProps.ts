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
  EMailAddress?: string | undefined;
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
  additionalPPEItems?: string | undefined;
}