export interface IUser {
  id: string;
  displayName: string;
  email?: string | undefined;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  mobilePhone?: string;
  profileImageUrl?: string; // for lazy-loaded photo
  isSelected?: boolean;
  manager?: { displayName: string; id: string };
  company?: string;

  title?: string | undefined;
  loginName?: string | undefined;
  isSiteAdmin?: boolean | undefined;
  isEmailAuthenticationGuestUser?: boolean | undefined;
  isHiddenInUI?: boolean | undefined;
  isShareByEmailGuestUser?: boolean | undefined;
  principalType?: number | undefined;
}