export function joinMergeValues(values, connector) {
  return values.map(value => String(value ?? '').trim()).filter(Boolean).join(connector);
}

export function shiftedColumn(originalColumn, removedColumns) {
  return originalColumn - removedColumns.filter(column => column < originalColumn).length;
}
