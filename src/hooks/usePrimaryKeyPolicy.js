import { useState } from 'react';

const ROW_SEPARATOR = '\u241f';

const normalizeRowLength = (row, columnCount) =>
  Array.from({ length: columnCount }, (_, index) => String(row[index] ?? '').trim());

const isMeaningfulRow = (row) => row.some((cell) => cell !== '');

export default function usePrimaryKeyPolicy() {
  const [blockedRows, setBlockedRows] = useState([]);

  const clearBlockedRows = () => {
    setBlockedRows([]);
  };

  const resetBlockedRows = () => {
    setBlockedRows([]);
  };

  const validateRows = ({ headers, rows, primaryKeyIndex, source }) => {
    const safeHeaders = headers.length ? headers : ['PermissionId'];
    const columnCount = safeHeaders.length;
    const resolvedPrimaryKeyIndex = Math.min(primaryKeyIndex, columnCount - 1);

    const normalizedRows = rows.map((row) => normalizeRowLength(row, columnCount));
    const seenPrimaryKeys = new Set();
    const nextBlockedRows = [];

    normalizedRows.forEach((row, rowIndex) => {
      if (!isMeaningfulRow(row)) return;

      const rawPrimaryKey = row[resolvedPrimaryKeyIndex] ?? '';
      const normalizedPrimaryKey = rawPrimaryKey.toLowerCase();

      if (!normalizedPrimaryKey) {
        nextBlockedRows.push({
          id: `${source}-missing-key-${rowIndex}`,
          row,
          source,
          keyValue: '(empty)',
          reason: 'Primary key cannot be empty.',
        });
        return;
      }

      if (seenPrimaryKeys.has(normalizedPrimaryKey)) {
        nextBlockedRows.push({
          id: `${source}-duplicate-key-${rowIndex}-${normalizedPrimaryKey}${ROW_SEPARATOR}${row.join(ROW_SEPARATOR)}`,
          row,
          source,
          keyValue: rawPrimaryKey,
          reason: 'Duplicate primary key value.',
        });
        return;
      }

      seenPrimaryKeys.add(normalizedPrimaryKey);
    });

    if (nextBlockedRows.length > 0) {
      setBlockedRows((prev) => [...nextBlockedRows, ...prev]);

      return {
        isValid: false,
        headers: safeHeaders,
        rows: normalizedRows,
        resolvedPrimaryKeyIndex,
        error: `${nextBlockedRows.length} row(s) blocked. Primary key "${safeHeaders[resolvedPrimaryKeyIndex]}" must be unique and non-empty.`,
      };
    }

    return {
      isValid: true,
      headers: safeHeaders,
      rows: normalizedRows,
      resolvedPrimaryKeyIndex,
      error: '',
    };
  };

  return {
    blockedRows,
    clearBlockedRows,
    resetBlockedRows,
    validateRows,
  };
}
