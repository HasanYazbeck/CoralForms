import {IUser} from './IUser'; 
  
  
  export interface ICommon {
    id: string | undefined;
    label?: string;
    title?: string;
  }

  export interface DateRange {
    StartDate: Date | undefined;
    EndDate: Date | undefined;
  }

  export interface IGraphResponse {
  value: IGraphUserResponse[];
  "@odata.nextLink"?: string;
}

  export interface IGraphUserResponse {
  id: string;
  displayName: string;
  mail: string;
  department?: string;
  jobTitle?: string;
  mobilePhone?: string;
  officeLocation?: string;
  manager?: {
    displayName: string;
    id: string;
  };
}

 export type FileWithPreview = {
    File: File;
    Preview?: string;
    Id: string;
  };

  export interface ISPListItem {
    Id: string;
    Title?: string;
    Created?: Date | undefined;
    CreatedBy?: IUser | undefined;
    Modified?: Date | undefined;
    ModifiedBy?: IUser | undefined;
    Attachments?: FileWithPreview[] | undefined;
    Order?: number | undefined;
    // [key: string]: any; // for other dynamic fields
  };

  export interface ILKPItemInstructionsForUse  extends ISPListItem {
    FormName: string;
    Description: string;
  }
