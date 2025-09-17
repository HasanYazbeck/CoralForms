import { WebPartContext } from "@microsoft/sp-webpart-base";

interface IEmployeeProps {
  employeeName: string;
  jobTitle: string;
  company: string;
  division: string;
  department: string;
}

enum PPERequestReason {
  New = 1,
  Replacement = 2
}

export interface IPPEForm extends IEmployeeProps{
  context: WebPartContext
  requestorName: IEmployeeProps;
  dateRequested: Date;
  reasonOfRequest: PPERequestReason;
  replacementReason?: string;
}