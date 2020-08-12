export interface FieldDef {
  "@odata.id": string;
}

export default function getParentColumnDefinition(fielddef: FieldDef) {
  const listid = fielddef["@odata.id"].substr(fielddef["@odata.id"].indexOf("Lists(guid'") + 11, 36);
  return { 'parameters': { 'FieldTypeKind': 7, 'Title': 'Parent', 'LookupListId': listid, 'LookupFieldName': 'ID' } };
}