
import { IPersonaProps } from '@fluentui/react';
import { ICommon, ISPListItem } from './../Common/ICommon';

export interface IFormsApprovalWorkflow extends ISPListItem {
  FormName?: ICommon | undefined;
  EmployeeId?: number | undefined;
  SignOffName: string | undefined;
  FinalLevel: number | undefined;
  ApproverGroup?: IPersonaProps | undefined;
  DepartmentManagerApprover?: IPersonaProps | undefined;
  Status: ICommon | undefined;
  Reason: string | undefined;
  Date: Date | undefined;
  Order?: number | undefined;
  IsFinalFormApprover: boolean | undefined;
  ModifiedByPersona?: IPersonaProps | undefined;
  ApproversNamesList: Record<string, IPersonaProps[]>
};