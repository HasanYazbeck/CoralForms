import { IWorkCategory } from "./IPTWForm";

export interface IPermitScheduleRow {
  id: string;
  type: 'new' | 'renewal';
  date: string;
  startTime: string;
  endTime: string;
  isChecked: boolean;
}

export interface IPermitScheduleProps {
  workCategories: IWorkCategory[];
  selectedPermitType?: IWorkCategory;
  permitRows: IPermitScheduleRow[];
  onPermitTypeChange: (workCategory: IWorkCategory | undefined) => void;
  onPermitRowUpdate: (rowId: string, field: string, value: string , checked :boolean) => void;
  styles?: any;
}