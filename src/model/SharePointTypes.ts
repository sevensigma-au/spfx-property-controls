export interface ISpList {
  Title?: string;
}

export const SpListProperty = {
  title: 'Title'
};

export interface ISpField {
  Title?: string;
  InternalName?: string;
  TypeAsString?: string;
  FieldTypeKind?: number;
  Hidden?: boolean;
  CanBeDeleted?: boolean;
}

export const SpFieldProperty = {
  title: 'Title',
  internalName: 'InternalName',
  typeAsString: 'TypeAsString',
  fieldTypeKind: 'FieldTypeKind',
  hidden: 'Hidden',
  canBeDeleted: 'CanBeDeleted'
};

export const enum SpFieldType {
  invalid,
  integer,
  text,
  note,
  dateTime,
  counter,
  choice,
  lookup,
  boolean,
  number,
  currency,
  url,
  computed,
  threading,
  guid,
  multichoice,
  gridchoice,
  calculated,
  file,
  attachments,
  user,
  recurrence,
  crossProjectLink,
  modStat,
  error,
  contentTypeId,
  pageSeparator,
  threadIndex,
  workflowStatus,
  allDayEvent,
  workflowEventType,
  maxItems
}