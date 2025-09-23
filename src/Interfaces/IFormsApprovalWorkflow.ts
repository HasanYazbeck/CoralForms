 
 import {ISPListItem} from './ICommon';
import { IEmployeeProps } from './IEmployeeProps';

 export interface IFormsApprovalWorkflow extends ISPListItem {
    FormName: string | undefined;
    Order:  number | 0;
    DepartmentName: string | undefined;
    Manager: IEmployeeProps| undefined; // User ID
  };