// export interface IUser {
//   Department?: string;
//   DisplayName: string;
//   Name?: string;
//   Id: string;
//   JobTitle?: string;
//   OfficeLocation?: string;
//   Email?: string;
//   MobilePhone?: string | undefined;
//   ProfileImageUrl?: string;
//   IsSelected?: boolean;
//   Manager?: { DisplayName: string, Id: string}
// }

export interface IUser {
  id: string;
  displayName: string;
  email?: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  mobilePhone?: string;
  profileImageUrl?: string; // for lazy-loaded photo
  isSelected?: boolean;
  manager?: { displayName: string; id: string };
  company?: string; 
}