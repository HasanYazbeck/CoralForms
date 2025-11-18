import { IPersonaProps } from "@fluentui/react";
import { IWorkCategory } from "./IPTWForm";

export type WorkPermitStatus = 'New' | 'Open' | 'Closed' | 'Cancelled';

export interface IPermitScheduleRow {
  id: string;
  type: 'new' | 'renewal';
  date: string;
  startTime: string;
  endTime: string;
  isChecked: boolean;
  orderRecord: number;
  statusRecord?: WorkPermitStatus | undefined;
  piApprover: IPersonaProps | undefined;
  piApproverList?: IPersonaProps[] | undefined;
  piApprovalDate?: Date | undefined;
  piStatus?: string | undefined;
}

export interface IPermitScheduleProps {
  workCategories: IWorkCategory[];
  selectedPermitTypeList: IWorkCategory[];
  permitRows: IPermitScheduleRow[];
  onPermitTypeChange: (checked: boolean | undefined, workCategory: IWorkCategory | undefined) => void;
  onPermitRowUpdate: (rowId: string, field: string, value: string, checked: boolean | undefined) => void;
  styles?: any;
  isEndTimeOptionDisabled?: (row: IPermitScheduleRow, optionTime: string) => boolean;
  permitsValidityDays: number;
  permitStatus?: boolean;
  isPermitIssuer?: boolean;
  piApproverList: IPersonaProps[] | undefined;
  isIssued: boolean;
  isSubmitted: boolean;
}