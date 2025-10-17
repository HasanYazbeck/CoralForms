import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls";
import { IPTWForm } from "../../../Interfaces/PtwForm/IPTWForm";

export interface IPTWFormProps {
  context: WebPartContext;
  ThemeColor: string | undefined;
  IsDarkTheme: boolean;
  useTargetAudience: boolean;
  targetAudience: IPropertyFieldGroupOrPerson[];
  onClose?: () => void;
  onSubmitted?: (newFormId?: number) => void;
  formId?: number;
  formStructure?: IPTWForm;
}

