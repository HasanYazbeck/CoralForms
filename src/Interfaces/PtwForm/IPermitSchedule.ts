import { IWorkCategory } from "./IPTWForm";

export interface IPermitScheduleRow {
  id: string;
  type: 'new' | 'renewal';
  date: string;
  startTime: string;
  endTime: string;
  isChecked: boolean;
  orderRecord: number;
}

export interface IPermitScheduleProps {
  workCategories: IWorkCategory[];
  selectedPermitTypeList: IWorkCategory[];
  permitRows: IPermitScheduleRow[];
  onPermitTypeChange: (checked: boolean | undefined, workCategory: IWorkCategory | undefined) => void;
  onPermitRowUpdate: (rowId: string, field: string, value: string , checked :boolean | undefined) => void;
  styles?: any;
  isEndTimeOptionDisabled?: (row: IPermitScheduleRow, optionTime: string) => boolean;
}