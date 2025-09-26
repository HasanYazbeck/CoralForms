 
 import {ISPListItem} from './ICommon';
import { IEmployeeProps } from './IEmployeeProps';

 export interface IFormsApprovalWorkflow extends ISPListItem {
    FormName: string | undefined;
    DepartmentName: string | undefined;
    Manager: IEmployeeProps| undefined; // User ID
  };