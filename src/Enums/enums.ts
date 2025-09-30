
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