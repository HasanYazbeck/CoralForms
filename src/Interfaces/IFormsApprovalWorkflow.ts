 
import { IPersonaProps } from '@fluentui/react';
import {ISPListItem} from './ICommon';

 export interface IFormsApprovalWorkflow extends ISPListItem {
    FormName: string | undefined;
    SignOffName: string | undefined;
    EmployeeId?: number | undefined;
    DepartmentManager?: IPersonaProps | undefined;
  };