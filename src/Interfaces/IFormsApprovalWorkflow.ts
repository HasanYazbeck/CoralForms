
import { IPersonaProps } from '@fluentui/react';
import { ICommon, ISPListItem } from './ICommon';

export interface IFormsApprovalWorkflow extends ISPListItem {
  FormName?: ICommon | undefined;
  EmployeeId?: number | undefined;
  SignOffName: string | undefined;
  DepartmentManager?: IPersonaProps | undefined;
  Status: ICommon | undefined;
  Reason: string | undefined;
  Date: Date | undefined;
  Order?: number | undefined;
  IsFinalFormApprover: boolean | undefined;
};