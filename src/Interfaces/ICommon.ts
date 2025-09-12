
  export interface ICommon {
    id: string | undefined;
    label?: string;
  }

  export type IDeviceType = {
    Id?: string | undefined;
    Title?: string | undefined;
  }

  export type IDeviceCategory = {
    Id?: string | undefined;
    Title?: string | undefined;
    DeviceTypes?: IDeviceType;
  }

  export type IDevice = {
    Id?: string | undefined;
    Title?: string | undefined;
    DeviceType?: IDeviceType | undefined;
    DeviceCategory?: IDeviceCategory | undefined;
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