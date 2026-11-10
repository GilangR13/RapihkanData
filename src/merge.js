export function joinMergeValues(values, connector) {
  return values.map(value => String(value ?? '').trim()).filter(Boolean).join(connector);
}

export function shiftedColumn(originalColumn, removedColumns) {
  return originalColumn - removedColumns.filter(column => column < originalColumn).length;
}

export function groupColumns(group, groups) {
  const columns = [...group.columns];
  for (const id of group.groupIds || []) {
    const previous = groups.find(item => item.id === id);
    if (previous) columns.push(...groupColumns(previous, groups));
  }
  return [...new Set(columns)].sort((a, b) => a - b);
}

export function finalMergeGroups(groups) {
  const usedAsInput = new Set(groups.flatMap(group => group.groupIds || []));
  return groups.filter(group => !usedAsInput.has(group.id));
}

export function mergeGroupValue(group, groups, readColumn) {
  const inputs = [
    ...group.columns.map(column => ({ column, read: () => readColumn(column) })),
    ...(group.groupIds || []).map(id => {
      const previous = groups.find(item => item.id === id);
      return previous ? { column: groupColumns(previous, groups)[0], read: () => mergeGroupValue(previous, groups, readColumn) } : null;
    }).filter(Boolean),
  ];
  inputs.sort((a, b) => a.column - b.column);
  return joinMergeValues(inputs.map(input => input.read()), group.separator);
}
