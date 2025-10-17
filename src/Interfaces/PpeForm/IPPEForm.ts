import { IPersonaProps } from '@fluentui/react';
import { IEmployeeProps } from './../PpeForm/IEmployeeProps';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IUser } from './../Common/IUser';
import { IPPEItemDetails } from './../PpeForm/IPPEItemDetails';

// export interface IPPEForm extends ISPListItem {
//     EmployeeID?: number;
//     EmployeeRecord?: IPersonaProps | undefined;
//     JobTitle: ICommon | undefined;
//     Company?: ICommon | undefined;
//     Division?: ICommon | undefined;
//     Department?: ICommon | undefined;
//     RequesterNameID?: number | undefined;
//     SubmitterNameID?: number | undefined;
//     ReasonForRequest?: string | undefined;
//     ReplacementReason?: string | undefined;
// }

enum PPERequestReason {
  New = 'New',
  Replacement = 'Replacement'
}

export interface IPPEForm extends IEmployeeProps {
  context?: WebPartContext
  requesterName: IUser | undefined;
  submitterName: IUser | undefined;
  dateRequested: Date | undefined;
  reasonForRequest: PPERequestReason;
  replacementReason: string;
  employeeRecord: IPersonaProps | undefined;
  ppeItems? : IPPEItemDetails[];
}

