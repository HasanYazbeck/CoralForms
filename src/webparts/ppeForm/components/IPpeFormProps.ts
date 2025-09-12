import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IUser } from "../../../Interfaces/IUser";

export interface IPpeFormWebPartProps {
  Context: WebPartContext;
  Users: IUser[];
  Departments: string[];
  JobTitles: string[];
  IsLoading?: Boolean;
}

export interface IPpeFormWebPartState {
  selectedEmployeeId?: string;
  jobTitle: string;
  department: string;
  division: string;
  company: string;
}
