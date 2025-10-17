import { IWorkCategory } from "./IPTWForm";

export interface IPermitScheduleRow {
  id: string;
  type: 'new' | 'renewal';
  date: string;
  startTime: string;
  startAmPm: 'AM' | 'PM';
  endTime: string;
  endAmPm: 'AM' | 'PM';
}

export interface IPermitScheduleProps {
  workCategories: IWorkCategory[];
  selectedPermitType?: IWorkCategory;
  permitRows: IPermitScheduleRow[];
  onPermitTypeChange: (workCategory: IWorkCategory | undefined) => void;
  onPermitRowUpdate: (rowId: string, field: string, value: string) => void;
  styles?: any;
}