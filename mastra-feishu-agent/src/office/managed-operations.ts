// Existing business operations own these endpoints. Both discovery and raw API
// validation consult this table, so raw writes cannot bypass their preconditions.
export const managedOperations: Record<string, string> = {
  'calendar.v4.calendar.create': 'office_calendar_container_create',
  'calendar.v4.calendar.patch': 'office_calendar_container_update',
  'calendar.v4.calendar.delete': 'office_calendar_container_delete',
  'calendar.v4.calendar.event.create': 'office_calendar_create',
  'calendar.v4.calendar.event.patch': 'office_calendar_update',
  'calendar.v4.calendar.event.delete': 'office_calendar_delete',
  'task.v2.task.create': 'office_task_create',
  'task.v2.task.patch': 'office_task_update',
  'task.v2.task.delete': 'office_task_delete',
  'docx.v1.document.create': 'office_doc_create',
  'docx.v1.document.block.patch': 'office_doc_update',
  'bitable.v1.app.table.record.create': 'office_table_create_record',
  'bitable.v1.app.table.record.update': 'office_table_update_record',
  'drive.v1.file.create_folder': 'office_drive_create_folder',
  'drive.v1.file.copy': 'office_drive_copy_file',
  'im.v1.message.create': 'office_message_send',
};

export function managedOperation(id: string, method: string, path: string): string | undefined {
  if (managedOperations[id]) return managedOperations[id];
  // Legacy/private Sheets v2 endpoints may be added to a future SDK catalog.
  if (method === 'PUT' && /\/sheets\/v2\/spreadsheets\/:[\w]+\/values$/.test(path)) return 'office_table_update_cell';
  if (method === 'POST' && /\/sheets\/v2\/spreadsheets\/:[\w]+\/values_append$/.test(path)) return 'office_table_append_row';
}
