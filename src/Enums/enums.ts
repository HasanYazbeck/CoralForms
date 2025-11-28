
// This enum is used when adding a column to SharePoint List, then a fieldType should be passed.
export enum FieldTypeKind {
  SingleLineOfText = 2,
  MultipleLinesOfText = 3,
  Number = 4,
  DateTime = 6,
  Choice = 7,
  Lookup = 8,
  YesNo = 9,
  PersonOrGroup = 10,
  HyperlinkOrPicture = 11
}

export enum FormsApprovalLevels {
  FLA = 1,
  SLA = 2,
  TLA = 3,
  FOLA = 4,
  FILA = 5
}

export enum PTWWorkflowStatus {
  New = 'New',
  InReview = 'In Review',
  Issued = 'Issued',
  Open = 'Open',
  Renewed = 'Renewed',
  Closed = 'Closed',
  PermanentlyClosed = 'Permanently Closed',
  Rejected = 'Rejected',
  Cancelled = 'Cancelled'
}

export enum PTWWorflowStage {
  ApprovedFromPOToPA = 'ApprovedFromPOToPA',
  ApprovedFromPAToPI = 'ApprovedFromPAToPI',
  ApprovedFromPOToPI = 'ApprovedFromPOToPI',
  ApprovedFromPIToAsset = 'ApprovedFromPIToAsset',
  ApprovedFromAssetToHSE = 'ApprovedFromAssetToHSE',
  ApprovedFromPOtoAssetUrgent = 'ApprovedFromPOtoAssetUrgent',
  Issued = 'Issued',
  Rejected = 'Rejected',
  ClosedByPO = 'ClosedByPO',
  ClosedByAssetManager = 'ClosedByAssetManager',
  PermanentlyClosed = 'Permanently Closed',
}

export enum FormStatusRecord {
  Saved = 'Saved',
  Submitted = 'Submitted',
  Rejected = 'Rejected'
}