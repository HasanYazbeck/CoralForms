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
  requestorName: IEmployeeProps;
  dateRequested: Date;
  reasonOfRequest: PPERequestReason;
  replacementReason?: string;
}