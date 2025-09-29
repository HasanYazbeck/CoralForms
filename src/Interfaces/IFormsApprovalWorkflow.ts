
import { IPersonaProps } from '@fluentui/react';
import { ISPListItem } from './ICommon';

export interface IFormsApprovalWorkflow extends ISPListItem {
  FormName?: string | undefined;
  EmployeeId?: number | undefined;
  SignOffName: string | undefined;
  DepartmentManager?: IPersonaProps | undefined;
  Status: string | undefined;
  Reason: string | undefined;
  Date: Date | undefined;
  Order?: number | undefined;
  IsFinalFormApprover: boolean | undefined;
};