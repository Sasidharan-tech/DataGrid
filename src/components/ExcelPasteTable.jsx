"use client";

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import './ExcelPasteTable.css';
import usePrimaryKeyPolicy from '../hooks/usePrimaryKeyPolicy';

const STATIC_PERMISSION_DATA = {
  headers: ['PermissionId', 'PermissionName', 'Module', 'AccessLevel', 'Status'],
  rows: [
    ['PERM-001', 'View Dashboard', 'Dashboard', 'Read', 'Active'],
    ['PERM-002', 'Manage Users', 'Users', 'Write', 'Active'],
    ['PERM-003', 'View Reports', 'Reports', 'Read', 'Active'],
    ['PERM-004', 'Edit Reports', 'Reports', 'Write', 'Active'],
    ['PERM-005', 'Delete Reports', 'Reports', 'Delete', 'Inactive'],
  ],
};

const FILTER_OPERATORS = [
  { value: 'contains', label: 'contains' },
  { value: 'equals', label: 'equals' },
  { value: 'startsWith', label: 'starts with' },
  { value: 'endsWith', label: 'ends with' },
  { value: 'notEquals', label: 'not equals' },
  { value: 'greaterThan', label: '> greater than' },
  { value: 'lessThan', label: '< less than' },
];

const PAGE_SIZE_OPTIONS = [5, 10, 25, 50, 100];
const DEFAULT_ROWS_PER_PAGE = 10;

const createEmptyFilterConfig = () => ({
  condition1: { operator: 'contains', value: '' },
  logic: 'AND',
  condition2: { operator: 'contains', value: '' },
});

const normalizeCell = (value) => String(value ?? '').toLowerCase();

const matchesOperator = (cellValue, operator, query) => {
  if (!query) return true;

  const normalizedCell = normalizeCell(cellValue);
  const normalizedQuery = String(query).toLowerCase();

  switch (operator) {
    case 'equals':
      return normalizedCell === normalizedQuery;
    case 'startsWith':
      return normalizedCell.startsWith(normalizedQuery);
    case 'endsWith':
      return normalizedCell.endsWith(normalizedQuery);
    case 'notEquals':
      return normalizedCell !== normalizedQuery;
    case 'greaterThan': {
      const numCell = parseFloat(cellValue);
      const numQuery = parseFloat(query);
      if (!isNaN(numCell) && !isNaN(numQuery)) return numCell > numQuery;
      return normalizedCell > normalizedQuery;
    }
    case 'lessThan': {
      const numCell = parseFloat(cellValue);
      const numQuery = parseFloat(query);
      if (!isNaN(numCell) && !isNaN(numQuery)) return numCell < numQuery;
      return normalizedCell < normalizedQuery;
    }
    case 'contains':
    default:
      return normalizedCell.includes(normalizedQuery);
  }
};

function detectDelimiter(text) {
  const firstContentLine = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .find((line) => line.length > 0);

  if (!firstContentLine) return '\t';

  const tabCount = (firstContentLine.match(/\t/g) || []).length;
  const commaCount = (firstContentLine.match(/,/g) || []).length;
  const semicolonCount = (firstContentLine.match(/;/g) || []).length;

  const scored = [
    { delimiter: '\t', count: tabCount },
    { delimiter: ',', count: commaCount },
    { delimiter: ';', count: semicolonCount },
  ].sort((a, b) => b.count - a.count);

  return scored[0].count > 0 ? scored[0].delimiter : '\t';
}

function parseTabularText(text) {
  const input = text.replace(/\r\n/g, '\n').trim();

  if (!input) {
    return { error: 'No data found. Paste Excel-style table data to continue.' };
  }

  const delimiter = detectDelimiter(input);

  const rawRows = input
    .split('\n')
    .map((row) => row.split(delimiter).map((cell) => cell.trim()));

  if (!rawRows.length || rawRows.every((row) => row.every((cell) => cell === ''))) {
    return { error: 'Unable to parse data. Ensure rows and columns are properly delimited.' };
  }

  const normalizedRows = rawRows.filter(
    (row) => row.length > 1 || row[0] !== ''
  );

  if (!normalizedRows.length) {
    return { error: 'No usable rows found after parsing.' };
  }

  const maxColumns = Math.max(...normalizedRows.map((row) => row.length));

  const paddedRows = normalizedRows.map((row) => {
    if (row.length === maxColumns) return row;
    return [...row, ...Array(maxColumns - row.length).fill('')];
  });

  const headerRow = paddedRows[0].map((header, index) =>
    header || `Column ${index + 1}`
  );

  const bodyRows = paddedRows.slice(1);

  return {
    headers: headerRow,
    rows: bodyRows,
    delimiter,
    error: '',
  };
}

const ExcelPasteTable = () => {
  const [headers, setHeaders] = useState([...STATIC_PERMISSION_DATA.headers]);
  const [rows, setRows] = useState(
    STATIC_PERMISSION_DATA.rows.map((row) => [...row])
  );
  const [error, setError] = useState('');
  const [primaryKeyIndex, setPrimaryKeyIndex] = useState(0);
  const [enforceNotNull, setEnforceNotNull] = useState(true);
  const [detectedDelimiter, setDetectedDelimiter] = useState('\t');
  const [searchTerm, setSearchTerm] = useState('');
  const [debouncedSearchTerm, setDebouncedSearchTerm] = useState('');
  const [filterColumn, setFilterColumn] = useState('all');
  const [groupByColumns, setGroupByColumns] = useState([]);
  const [orderByColumn, setOrderByColumn] = useState('none');
  const [orderDirection, setOrderDirection] = useState('asc');
  const [columnFilters, setColumnFilters] = useState({});
  const [menuColumnIndex, setMenuColumnIndex] = useState(null);
  const [menuPosition, setMenuPosition] = useState({ top: 0, left: 0 });
  const [menuDraft, setMenuDraft] = useState(createEmptyFilterConfig());
  const [draggedGroupColumn, setDraggedGroupColumn] = useState(null);
  const [isGroupDropActive, setIsGroupDropActive] = useState(false);
  const [visibleRowIndexes, setVisibleRowIndexes] = useState([]);
  const [currentPage, setCurrentPage] = useState(1);
  const [rowsPerPage, setRowsPerPage] = useState(DEFAULT_ROWS_PER_PAGE);
  const [selection, setSelection] = useState(null);
  const [isDraggingSelection, setIsDraggingSelection] = useState(false);
  const [editingCell, setEditingCell] = useState(null);
  const [editingValue, setEditingValue] = useState('');
  const [selectedRows, setSelectedRows] = useState({});
  const [hiddenColumns, setHiddenColumns] = useState(new Set());
  const [showColumnPanel, setShowColumnPanel] = useState(false);
  const fileInputRef = useRef(null);
  const { blockedRows, clearBlockedRows, resetBlockedRows, validateRows } = usePrimaryKeyPolicy();

  useEffect(() => {
    const timer = setTimeout(() => {
      setDebouncedSearchTerm(searchTerm);
    }, 200);
    return () => clearTimeout(timer);
  }, [searchTerm]);

  useEffect(() => {
    if (!rows.length) {
      setVisibleRowIndexes([]);
      return;
    }

    const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });

    const query = debouncedSearchTerm.trim().toLowerCase();
    let indexes = rows.map((_, index) => index);

    if (enforceNotNull && headers.length > 0) {
      const resolvedPrimaryKeyIndex = Math.min(primaryKeyIndex, headers.length - 1);
      indexes = indexes.filter((rowIndex) => {
        const primaryKeyValue = String(rows[rowIndex]?.[resolvedPrimaryKeyIndex] ?? '').trim();
        return primaryKeyValue !== '';
      });
    }

    if (query) {
      indexes = indexes.filter((rowIndex) => {
        const row = rows[rowIndex] || [];
        if (filterColumn === 'all') {
          return row.some((cell) => String(cell).toLowerCase().includes(query));
        }

        const selectedColumnIndex = Number(filterColumn);
        const cellValue = row[selectedColumnIndex] ?? '';
        return String(cellValue).toLowerCase().includes(query);
      });
    }

    const hasColumnFilters = Object.keys(columnFilters).length > 0;
    if (hasColumnFilters) {
      indexes = indexes.filter((rowIndex) => {
        const row = rows[rowIndex] || [];

        return Object.entries(columnFilters).every(([columnKey, filterConfig]) => {
          const columnIndex = Number(columnKey);
          const cellValue = row[columnIndex] ?? '';

          const firstMatch = matchesOperator(
            cellValue,
            filterConfig.condition1.operator,
            filterConfig.condition1.value
          );

          const hasSecondValue = String(filterConfig.condition2.value || '').trim() !== '';
          if (!hasSecondValue) {
            return firstMatch;
          }

          const secondMatch = matchesOperator(
            cellValue,
            filterConfig.condition2.operator,
            filterConfig.condition2.value
          );

          return filterConfig.logic === 'OR'
            ? firstMatch || secondMatch
            : firstMatch && secondMatch;
        });
      });
    }

    const getCellValue = (rowIndex, columnIndex) =>
      String(rows[rowIndex]?.[columnIndex] ?? '').trim();

    const compareByColumn = (leftIndex, rightIndex, columnIndex, direction = 'asc') => {
      const leftValue = getCellValue(leftIndex, columnIndex);
      const rightValue = getCellValue(rightIndex, columnIndex);
      const baseCompare = collator.compare(leftValue, rightValue);

      if (baseCompare !== 0) {
        return direction === 'asc' ? baseCompare : -baseCompare;
      }

      return leftIndex - rightIndex;
    };

    if (groupByColumns.length > 0 || orderByColumn !== 'none') {
      const orderColumnIndex = Number(orderByColumn);
      indexes = [...indexes].sort((leftIndex, rightIndex) => {
        for (let i = 0; i < groupByColumns.length; i += 1) {
          const groupColumnIndex = Number(groupByColumns[i]);
          const groupCompare = compareByColumn(leftIndex, rightIndex, groupColumnIndex, 'asc');
          if (groupCompare !== 0) {
            return groupCompare;
          }
        }

        if (orderByColumn === 'none') {
          return leftIndex - rightIndex;
        }

        return compareByColumn(leftIndex, rightIndex, orderColumnIndex, orderDirection);
      });
    }

    setVisibleRowIndexes(indexes);
  }, [
    rows,
    headers,
    primaryKeyIndex,
    enforceNotNull,
    debouncedSearchTerm,
    filterColumn,
    columnFilters,
    groupByColumns,
    orderByColumn,
    orderDirection,
  ]);

  useEffect(() => {
    const totalPages = Math.max(1, Math.ceil(visibleRowIndexes.length / rowsPerPage));
    if (currentPage > totalPages) {
      setCurrentPage(totalPages);
    }
  }, [visibleRowIndexes, currentPage, rowsPerPage]);

  useEffect(() => {
    setCurrentPage(1);
  }, [debouncedSearchTerm, filterColumn, columnFilters, groupByColumns, orderByColumn, orderDirection, rowsPerPage]);

  const normalizeSelection = (currentSelection) => {
    if (!currentSelection) return null;

    const top = Math.min(currentSelection.startRow, currentSelection.endRow);
    const bottom = Math.max(currentSelection.startRow, currentSelection.endRow);
    const left = Math.min(currentSelection.startCol, currentSelection.endCol);
    const right = Math.max(currentSelection.startCol, currentSelection.endCol);

    return { top, bottom, left, right };
  };

  const selectedRange = useMemo(() => normalizeSelection(selection), [selection]);

  const isCellSelected = (rowIndex, columnIndex) => {
    if (!selectedRange) return false;
    return (
      rowIndex >= selectedRange.top &&
      rowIndex <= selectedRange.bottom &&
      columnIndex >= selectedRange.left &&
      columnIndex <= selectedRange.right
    );
  };

  const parseClipboardMatrix = (text) => {
    const normalized = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').trimEnd();
    if (!normalized) return [];

    const delimiter = detectDelimiter(normalized);

    return normalized
      .split('\n')
      .map((line) => line.split(delimiter).map((cell) => cell.trim()))
      .filter((row) => row.length > 1 || row[0] !== '');
  };

  const applyRowsWithPrimaryKeyPolicy = (
    nextHeaders,
    candidateRows,
    source,
    overridePrimaryKeyIndex
  ) => {
    const basePrimaryKeyIndex =
      typeof overridePrimaryKeyIndex === 'number' ? overridePrimaryKeyIndex : primaryKeyIndex;
    const validation = validateRows({
      headers: nextHeaders,
      rows: candidateRows,
      primaryKeyIndex: basePrimaryKeyIndex,
      source,
    });

    if (!validation.isValid) {
      setError(validation.error);
      return false;
    }

    setHeaders(validation.headers);
    setRows(validation.rows);
    setPrimaryKeyIndex(validation.resolvedPrimaryKeyIndex);
    setSelectedRows({});

    if (filterColumn !== 'all' && Number(filterColumn) >= validation.headers.length) {
      setFilterColumn('all');
    }

    setGroupByColumns((prev) =>
      prev.filter((columnKey) => Number(columnKey) < validation.headers.length)
    );

    if (orderByColumn !== 'none' && Number(orderByColumn) >= validation.headers.length) {
      setOrderByColumn('none');
      setOrderDirection('asc');
    }

    setError('');
    return true;
  };

  const getSelectedMatrix = () => {
    if (!selectedRange) return [];

    const matrix = [];
    for (let rowIndex = selectedRange.top; rowIndex <= selectedRange.bottom; rowIndex += 1) {
      const selectedRow = [];
      for (
        let columnIndex = selectedRange.left;
        columnIndex <= selectedRange.right;
        columnIndex += 1
      ) {
        selectedRow.push(rows[rowIndex]?.[columnIndex] ?? '');
      }
      matrix.push(selectedRow);
    }

    return matrix;
  };

  const copySelectedCellsToClipboard = async () => {
    if (!selectedRange) return;

    const content = getSelectedMatrix()
      .map((row) => row.join('\t'))
      .join('\n');

    try {
      await navigator.clipboard.writeText(content);
      setError('');
    } catch {
      setError('Unable to copy automatically. Use Ctrl+C while table is focused.');
    }
  };

  const ensureGridCapacity = (anchorRow, anchorCol, matrix) => {
    const matrixWidth = Math.max(...matrix.map((row) => row.length));
    const requiredColumns = anchorCol + matrixWidth;
    const requiredRows = anchorRow + matrix.length;

    const nextHeaders = [...headers];
    while (nextHeaders.length < requiredColumns) {
      nextHeaders.push(`Column ${nextHeaders.length + 1}`);
    }

    const nextRows = rows.map((row) => {
      const cloned = [...row];
      while (cloned.length < nextHeaders.length) {
        cloned.push('');
      }
      return cloned;
    });

    while (nextRows.length < requiredRows) {
      nextRows.push(Array(nextHeaders.length).fill(''));
    }

    return { nextHeaders, nextRows };
  };

  const pasteMatrixIntoGrid = (anchorRow, anchorCol, matrix) => {
    if (!matrix.length) return;

    const { nextHeaders, nextRows } = ensureGridCapacity(anchorRow, anchorCol, matrix);

    matrix.forEach((matrixRow, rowOffset) => {
      matrixRow.forEach((value, colOffset) => {
        nextRows[anchorRow + rowOffset][anchorCol + colOffset] = value;
      });
    });

    const applied = applyRowsWithPrimaryKeyPolicy(nextHeaders, nextRows, 'paste');
    if (!applied) return;

    const lastRow = anchorRow + matrix.length - 1;
    const widestRow = Math.max(...matrix.map((row) => row.length));
    const lastCol = anchorCol + widestRow - 1;
    setSelection({
      startRow: anchorRow,
      startCol: anchorCol,
      endRow: lastRow,
      endCol: lastCol,
    });
  };

  const applyParsedData = (text) => {
    const result = parseTabularText(text);

    if (result.error) {
      setHeaders([]);
      setRows([]);
      setError(result.error);
      return;
    }

    applyRowsWithPrimaryKeyPolicy(result.headers, result.rows, 'paste');
    setDetectedDelimiter(result.delimiter);
    setSelection(null);
    setSelectedRows({});
    setEditingCell(null);
  };

  const handlePaste = (event) => {
    const targetTag = event.target.tagName.toLowerCase();
    if (targetTag === 'input') {
      return;
    }

    event.preventDefault();
    const pastedText = event.clipboardData.getData('text');

    if (headers.length > 0 && selectedRange) {
      const matrix = parseClipboardMatrix(pastedText);
      if (!matrix.length) {
        setError('Pasted content is empty.');
        return;
      }

      pasteMatrixIntoGrid(selectedRange.top, selectedRange.left, matrix);
      setError('');
      return;
    }

    applyParsedData(pastedText);
  };

  const handleCopy = (event) => {
    if (!selectedRange) return;

    event.preventDefault();
    const content = getSelectedMatrix()
      .map((row) => row.join('\t'))
      .join('\n');
    event.clipboardData.setData('text/plain', content);
    setError('');
  };

  const clearSelectedCells = () => {
    if (!selectedRange) return;

    const nextRows = rows.map((row, rowIndex) =>
      row.map((cell, columnIndex) =>
        isCellSelected(rowIndex, columnIndex) ? '' : cell
      )
    );

    applyRowsWithPrimaryKeyPolicy(headers, nextRows, 'delete-cells');
  };

  const handleKeyDown = (event) => {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'c') {
      if (!selectedRange) return;
      event.preventDefault();
      copySelectedCellsToClipboard();
      return;
    }

    if (event.key === 'Delete') {
      if (!selectedRange) return;
      event.preventDefault();
      clearSelectedCells();
      return;
    }

    if (event.key === 'Escape') {
      setSelection(null);
      setEditingCell(null);
    }

    if (event.key === 'Enter' && selectedRange) {
      if (
        selectedRange.top === selectedRange.bottom &&
        selectedRange.left === selectedRange.right
      ) {
        const currentValue = rows[selectedRange.top]?.[selectedRange.left] ?? '';
        setEditingCell({ row: selectedRange.top, col: selectedRange.left });
        setEditingValue(currentValue);
      }
    }
  };

  const updateHeaderCell = (columnIndex, value) => {
    setHeaders((prevHeaders) =>
      prevHeaders.map((header, index) => (index === columnIndex ? value : header))
    );
  };

  const updateBodyCell = (rowIndex, columnIndex, value) => {
    const nextRows = rows.map((row, currentRowIndex) =>
        currentRowIndex === rowIndex
          ? row.map((cell, currentColumnIndex) =>
              currentColumnIndex === columnIndex ? value : cell
            )
          : row
      );

    applyRowsWithPrimaryKeyPolicy(headers, nextRows, 'edit');
  };

  const startSelection = (rowIndex, columnIndex) => {
    setSelection({
      startRow: rowIndex,
      startCol: columnIndex,
      endRow: rowIndex,
      endCol: columnIndex,
    });
    setIsDraggingSelection(true);
    setEditingCell(null);
    setError('');
  };

  const extendSelection = (rowIndex, columnIndex) => {
    if (!isDraggingSelection) return;

    setSelection((prevSelection) => {
      if (!prevSelection) return prevSelection;
      return {
        ...prevSelection,
        endRow: rowIndex,
        endCol: columnIndex,
      };
    });
  };

  const commitEditCell = () => {
    if (!editingCell) return;

    updateBodyCell(editingCell.row, editingCell.col, editingValue);
    setEditingCell(null);
    setEditingValue('');
  };

  const cancelEditCell = () => {
    setEditingCell(null);
    setEditingValue('');
  };

  const handleAddColumn = () => {
    const nextColumnIndex = headers.length + 1;
    const nextHeaders = [...headers, `Column ${nextColumnIndex}`];
    const nextRows = rows.map((row) => [...row, '']);
    applyRowsWithPrimaryKeyPolicy(nextHeaders, nextRows, 'add-column');
  };

  const handleRemoveColumn = (columnIndex) => {
    if (headers.length <= 1) {
      setHeaders(['PermissionId']);
      setRows([]);
      resetBlockedRows();
      return;
    }

    const nextHeaders = headers.filter((_, index) => index !== columnIndex);
    const nextRows = rows.map((row) => row.filter((_, index) => index !== columnIndex));
    applyRowsWithPrimaryKeyPolicy(nextHeaders, nextRows, 'remove-column');
  };

  const handleAddRow = () => {
    const nextHeaders = headers.length === 0 ? ['Column 1'] : headers;
    const columnCount = nextHeaders.length;
    const nextRows = [...rows, Array(columnCount).fill('')];
    applyRowsWithPrimaryKeyPolicy(nextHeaders, nextRows, 'add-row');
  };

  const toggleRowSelection = (rowIndex) => {
    setSelectedRows((prevSelectedRows) => ({
      ...prevSelectedRows,
      [rowIndex]: !prevSelectedRows[rowIndex],
    }));
  };

  const toggleAllRowsSelection = () => {
    const selectedCount = Object.values(selectedRows).filter(Boolean).length;

    if (selectedCount === rows.length) {
      setSelectedRows({});
      return;
    }

    const all = {};
    rows.forEach((_, rowIndex) => {
      all[rowIndex] = true;
    });
    setSelectedRows(all);
  };

  const deleteSelectedRows = () => {
    const selectedIndexes = new Set(
      Object.entries(selectedRows)
        .filter(([, isSelected]) => isSelected)
        .map(([index]) => Number(index))
    );

    if (!selectedIndexes.size) return;

    const nextRows = rows.filter((_, rowIndex) => !selectedIndexes.has(rowIndex));
    applyRowsWithPrimaryKeyPolicy(headers, nextRows, 'delete-row');
    setSelectedRows({});
    setSelection(null);
    setError('');
  };

  const handlePrimaryKeyChange = (event) => {
    const nextPrimaryKeyIndex = Number(event.target.value);
    setPrimaryKeyIndex(nextPrimaryKeyIndex);
    applyRowsWithPrimaryKeyPolicy(headers, rows, 'primary-key-change', nextPrimaryKeyIndex);
  };

  const handleHeaderDragStart = (columnIndex) => {
    setDraggedGroupColumn(columnIndex);
  };

  const handleHeaderDragEnd = () => {
    setDraggedGroupColumn(null);
    setIsGroupDropActive(false);
  };

  const handleGroupDrop = () => {
    if (draggedGroupColumn === null) return;
    setGroupByColumns((prev) => {
      const key = String(draggedGroupColumn);
      if (prev.includes(key)) return prev;
      return [...prev, key];
    });
    setDraggedGroupColumn(null);
    setIsGroupDropActive(false);
  };

  const handleRemoveGroupColumn = (columnKey) => {
    setGroupByColumns((prev) => prev.filter((currentKey) => currentKey !== columnKey));
  };

  const handleClearGroupColumns = () => {
    setGroupByColumns([]);
  };

  const openColumnMenu = (columnIndex, buttonElement) => {
    setMenuColumnIndex((current) => {
      if (current === columnIndex) {
        return null;
      }

      if (buttonElement) {
        const rect = buttonElement.getBoundingClientRect();
        setMenuPosition({
          top: rect.bottom + 6,
          left: Math.max(12, rect.right - 230),
        });
      }

      const existingFilter = columnFilters[columnIndex] || createEmptyFilterConfig();
      setMenuDraft({
        condition1: { ...existingFilter.condition1 },
        logic: existingFilter.logic,
        condition2: { ...existingFilter.condition2 },
      });
      return columnIndex;
    });
  };

  const closeColumnMenu = () => {
    setMenuColumnIndex(null);
  };

  const updateMenuDraft = (path, value) => {
    setMenuDraft((prev) => {
      if (path === 'logic') {
        return { ...prev, logic: value };
      }

      const [conditionKey, fieldKey] = path.split('.');
      return {
        ...prev,
        [conditionKey]: {
          ...prev[conditionKey],
          [fieldKey]: value,
        },
      };
    });
  };

  const applyColumnFilter = () => {
    if (menuColumnIndex === null) return;

    const hasFirst = String(menuDraft.condition1.value || '').trim() !== '';
    const hasSecond = String(menuDraft.condition2.value || '').trim() !== '';

    if (!hasFirst && !hasSecond) {
      setColumnFilters((prev) => {
        const next = { ...prev };
        delete next[menuColumnIndex];
        return next;
      });
      closeColumnMenu();
      return;
    }

    setColumnFilters((prev) => ({
      ...prev,
      [menuColumnIndex]: {
        condition1: { ...menuDraft.condition1 },
        logic: menuDraft.logic,
        condition2: { ...menuDraft.condition2 },
      },
    }));
    closeColumnMenu();
  };

  const clearColumnFilter = () => {
    if (menuColumnIndex === null) return;

    setColumnFilters((prev) => {
      const next = { ...prev };
      delete next[menuColumnIndex];
      return next;
    });
    setMenuDraft(createEmptyFilterConfig());
    closeColumnMenu();
  };

  const setSortAscending = () => {
    if (menuColumnIndex === null) return;
    setOrderByColumn(String(menuColumnIndex));
    setOrderDirection('asc');
    closeColumnMenu();
  };

  const setSortDescending = () => {
    if (menuColumnIndex === null) return;
    setOrderByColumn(String(menuColumnIndex));
    setOrderDirection('desc');
    closeColumnMenu();
  };

  const clearSort = () => {
    setOrderByColumn('none');
    setOrderDirection('asc');
    closeColumnMenu();
  };

  const handleClear = () => {
    setHeaders([...STATIC_PERMISSION_DATA.headers]);
    setRows(STATIC_PERMISSION_DATA.rows.map((row) => [...row]));
    setError('');
    setPrimaryKeyIndex(0);
    setDetectedDelimiter('\t');
    setSelection(null);
    setSelectedRows({});
    setEditingCell(null);
    resetBlockedRows();
    setSearchTerm('');
    setDebouncedSearchTerm('');
    setFilterColumn('all');
    setGroupByColumns([]);
    setOrderByColumn('none');
    setOrderDirection('asc');
    setEnforceNotNull(true);
    setColumnFilters({});
    setMenuColumnIndex(null);
    setMenuDraft(createEmptyFilterConfig());
    setHiddenColumns(new Set());
    setRowsPerPage(DEFAULT_ROWS_PER_PAGE);
  };

  const exportToCSV = useCallback(() => {
    const visibleHeaderIndexes = headers
      .map((_, index) => index)
      .filter((index) => !hiddenColumns.has(index));

    const headerLine = visibleHeaderIndexes
      .map((index) => {
        const cell = String(headers[index] ?? '');
        return cell.includes(',') || cell.includes('"') || cell.includes('\n')
          ? `"${cell.replace(/"/g, '""')}"`
          : cell;
      })
      .join(',');

    const dataLines = visibleRowIndexes.map((rowIndex) => {
      const row = rows[rowIndex] || [];
      return visibleHeaderIndexes
        .map((colIndex) => {
          const cell = String(row[colIndex] ?? '');
          return cell.includes(',') || cell.includes('"') || cell.includes('\n')
            ? `"${cell.replace(/"/g, '""')}"`
            : cell;
        })
        .join(',');
    });

    const csvContent = [headerLine, ...dataLines].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'datagrid-export.csv';
    link.click();
    URL.revokeObjectURL(url);
  }, [headers, rows, visibleRowIndexes, hiddenColumns]);

  const handleFileImport = (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      const text = e.target?.result;
      if (typeof text === 'string') {
        applyParsedData(text);
      }
    };
    reader.readAsText(file);
    event.target.value = '';
  };

  const toggleColumnVisibility = (columnIndex) => {
    setHiddenColumns((prev) => {
      const next = new Set(prev);
      if (next.has(columnIndex)) {
        next.delete(columnIndex);
      } else {
        next.add(columnIndex);
      }
      return next;
    });
  };

  const showAllColumns = () => setHiddenColumns(new Set());
  const hideAllColumns = () => setHiddenColumns(new Set(headers.map((_, i) => i)));

  const selectedRowCount = Object.values(selectedRows).filter(Boolean).length;
  const areAllRowsSelected = rows.length > 0 && selectedRowCount === rows.length;
  const totalPages = Math.max(1, Math.ceil(visibleRowIndexes.length / rowsPerPage));
  const pageStart = (currentPage - 1) * rowsPerPage;
  const pageEnd = pageStart + rowsPerPage;
  const paginatedRowIndexes = visibleRowIndexes.slice(pageStart, pageEnd);
  const groupedPageItems = useMemo(() => {
    if (groupByColumns.length === 0) {
      return paginatedRowIndexes.map((rowIndex) => ({ type: 'row', rowIndex }));
    }

    const groupCounts = new Map();

    paginatedRowIndexes.forEach((rowIndex) => {
      let path = '';
      groupByColumns.forEach((columnKey, levelIndex) => {
        const groupColumnIndex = Number(columnKey);
        const rawValue = String(rows[rowIndex]?.[groupColumnIndex] ?? '').trim();
        const groupValue = rawValue || '(empty)';
        path = path ? `${path}__${groupValue}` : groupValue;
        const countKey = `${levelIndex}:${path}`;
        groupCounts.set(countKey, (groupCounts.get(countKey) || 0) + 1);
      });
    });

    const items = [];
    const previousGroupValues = Array(groupByColumns.length).fill(null);

    paginatedRowIndexes.forEach((rowIndex) => {
      let path = '';

      groupByColumns.forEach((columnKey, levelIndex) => {
        const groupColumnIndex = Number(columnKey);
        const rawValue = String(rows[rowIndex]?.[groupColumnIndex] ?? '').trim();
        const groupValue = rawValue || '(empty)';

        path = path ? `${path}__${groupValue}` : groupValue;

        if (previousGroupValues[levelIndex] !== groupValue) {
          for (let resetIndex = levelIndex; resetIndex < previousGroupValues.length; resetIndex += 1) {
            previousGroupValues[resetIndex] = null;
          }

          previousGroupValues[levelIndex] = groupValue;
          const countKey = `${levelIndex}:${path}`;

          items.push({
            type: 'group',
            level: levelIndex,
            columnLabel: headers[groupColumnIndex] || `Column ${groupColumnIndex + 1}`,
            value: groupValue,
            count: groupCounts.get(countKey) || 0,
            key: `${countKey}-${rowIndex}`,
          });
        }
      });

      items.push({ type: 'row', rowIndex });
    });

    return items;
  }, [groupByColumns, headers, paginatedRowIndexes, rows]);

  return (
    <section
      className="excel-table-tool"
      onPaste={handlePaste}
      onCopy={handleCopy}
      onKeyDown={handleKeyDown}
      onMouseUp={() => setIsDraggingSelection(false)}
      tabIndex={0}
      aria-label="Paste-enabled table"
    >
      {/* ── Toolbar ── */}
      <div className="excel-table-tool__toolbar">
        <div className="excel-table-tool__summary">
          <span>{rows.length} rows</span>
          <span>·</span>
          <span>{headers.length} columns</span>
          {selectedRowCount > 0 && (
            <>
              <span>·</span>
              <span className="excel-table-tool__summary-selected">{selectedRowCount} selected</span>
            </>
          )}
          {Object.keys(columnFilters).length > 0 && (
            <span className="excel-table-tool__summary-badge">
              {Object.keys(columnFilters).length} filter{Object.keys(columnFilters).length !== 1 ? 's' : ''} active
            </span>
          )}
        </div>
        <div className="excel-table-tool__toolbar-actions">
          <button
            type="button"
            className="btn btn--add"
            onClick={handleAddRow}
            title="Add a new empty row"
          >
            + Row
          </button>
          <button
            type="button"
            className="btn btn--add"
            onClick={handleAddColumn}
            title="Add a new empty column"
          >
            + Col
          </button>
          <button
            type="button"
            className="btn btn--remove"
            onClick={deleteSelectedRows}
            disabled={selectedRowCount === 0}
            title="Delete selected rows"
          >
            Delete Selected
          </button>
          <button
            type="button"
            className="btn btn--copy"
            onClick={copySelectedCellsToClipboard}
            disabled={!selectedRange}
            title="Copy selected cells to clipboard"
          >
            Copy
          </button>
          <button
            type="button"
            className="btn btn--export"
            onClick={exportToCSV}
            disabled={rows.length === 0}
            title="Export visible data to CSV"
          >
            Export CSV
          </button>
          <label className="btn btn--import" title="Import CSV or TXT file">
            Import File
            <input
              ref={fileInputRef}
              type="file"
              accept=".csv,.txt,.tsv"
              className="excel-table-tool__sr-only"
              onChange={handleFileImport}
            />
          </label>
          <button
            type="button"
            className="btn btn--columns"
            onClick={() => setShowColumnPanel((v) => !v)}
            title="Show/hide columns"
          >
            Columns ▾
          </button>
          <button
            type="button"
            className="btn btn--ghost"
            onClick={handleClear}
            title="Reset to demo data"
          >
            Reset
          </button>
        </div>
      </div>

      {/* ── Column Visibility Panel ── */}
      {showColumnPanel && headers.length > 0 && (
        <div className="excel-table-tool__column-panel">
          <div className="excel-table-tool__column-panel-header">
            <span>Column Visibility</span>
            <div className="excel-table-tool__column-panel-actions">
              <button type="button" onClick={showAllColumns}>Show all</button>
              <button type="button" onClick={hideAllColumns}>Hide all</button>
              <button type="button" onClick={() => setShowColumnPanel(false)}>✕</button>
            </div>
          </div>
          <div className="excel-table-tool__column-panel-list">
            {headers.map((header, index) => (
              <label key={`vis-${index}`} className="excel-table-tool__column-panel-item">
                <input
                  type="checkbox"
                  checked={!hiddenColumns.has(index)}
                  onChange={() => toggleColumnVisibility(index)}
                />
                {header || `Column ${index + 1}`}
              </label>
            ))}
          </div>
        </div>
      )}

      {/* ── Filters Bar ── */}
      <div className="excel-table-tool__filters">
        <div
          className={`excel-table-tool__group-dropzone ${isGroupDropActive ? 'is-active' : ''}`}
          onDragOver={(event) => {
            event.preventDefault();
            setIsGroupDropActive(true);
          }}
          onDragLeave={() => setIsGroupDropActive(false)}
          onDrop={(event) => {
            event.preventDefault();
            handleGroupDrop();
          }}
        >
          {groupByColumns.length === 0 ? (
            'Drag column headers here to group rows'
          ) : (
            <>
              <span>Grouped by:</span>
              <div className="excel-table-tool__group-chips">
                {groupByColumns.map((columnKey) => {
                  const groupColumnIndex = Number(columnKey);
                  const label = headers[groupColumnIndex] || `Column ${groupColumnIndex + 1}`;

                  return (
                    <span key={`group-chip-${columnKey}`} className="excel-table-tool__group-chip">
                      {label}
                      <button
                        type="button"
                        onClick={() => handleRemoveGroupColumn(columnKey)}
                        aria-label={`Remove group ${label}`}
                      >
                        ×
                      </button>
                    </span>
                  );
                })}
              </div>
              <button
                type="button"
                className="excel-table-tool__group-clear"
                onClick={handleClearGroupColumns}
              >
                Clear Group
              </button>
            </>
          )}
        </div>

        <label className="excel-table-tool__pk-selector">
          Primary Key
          <select
            className="excel-table-tool__filter-select"
            value={primaryKeyIndex}
            onChange={handlePrimaryKeyChange}
            aria-label="Primary key column"
          >
            {headers.map((header, index) => (
              <option key={`pk-${index}`} value={index}>
                {header || `Column ${index + 1}`}
              </option>
            ))}
          </select>
        </label>

        <label className="excel-table-tool__pk-selector">
          NOT NULL
          <input
            type="checkbox"
            checked={enforceNotNull}
            onChange={(event) => setEnforceNotNull(event.target.checked)}
            aria-label="Enforce primary key NOT NULL"
          />
        </label>

        <input
          type="text"
          className="excel-table-tool__search-input"
          value={searchTerm}
          onChange={(event) => setSearchTerm(event.target.value)}
          placeholder="Search data..."
          aria-label="Search data"
        />
        <select
          className="excel-table-tool__filter-select"
          value={filterColumn}
          onChange={(event) => setFilterColumn(event.target.value)}
          aria-label="Filter by column"
        >
          <option value="all">All Columns</option>
          {headers.map((header, index) => (
            <option key={`filter-${index}`} value={String(index)}>
              {header || `Column ${index + 1}`}
            </option>
          ))}
        </select>

        <label className="excel-table-tool__pk-selector">
          ORDER BY
          <select
            className="excel-table-tool__filter-select"
            value={orderByColumn}
            onChange={(event) => setOrderByColumn(event.target.value)}
            aria-label="Order by column"
          >
            <option value="none">None</option>
            {headers.map((header, index) => (
              <option key={`order-${index}`} value={String(index)}>
                {header || `Column ${index + 1}`}
              </option>
            ))}
          </select>
        </label>

        <label className="excel-table-tool__pk-selector">
          Direction
          <select
            className="excel-table-tool__filter-select"
            value={orderDirection}
            onChange={(event) => setOrderDirection(event.target.value)}
            disabled={orderByColumn === 'none'}
            aria-label="Order direction"
          >
            <option value="asc">ASC</option>
            <option value="desc">DESC</option>
          </select>
        </label>

        <span className="excel-table-tool__filter-count">
          Showing {visibleRowIndexes.length} of {rows.length}
        </span>
      </div>

      {/* ── Table ── */}
      <div className="excel-table-tool__table-wrapper">
        <table className="excel-table-tool__table">
          {headers.length > 0 && (
            <thead>
              <tr>
                <th className="excel-table-tool__row-select-col">
                  <input
                    type="checkbox"
                    checked={areAllRowsSelected}
                    onChange={toggleAllRowsSelection}
                    aria-label="Select all rows"
                  />
                </th>
                {headers.map((header, index) => {
                  if (hiddenColumns.has(index)) return null;
                  return (
                    <th
                      key={`${index}-${header}`}
                      draggable
                      onDragStart={() => handleHeaderDragStart(index)}
                      onDragEnd={handleHeaderDragEnd}
                    >
                      <div className="excel-table-tool__header-cell">
                        <span className="excel-table-tool__drag-handle" title="Drag to Group By">
                          ⋮⋮
                        </span>
                        <input
                          type="text"
                          className="excel-table-tool__header-input"
                          value={header}
                          onChange={(event) => updateHeaderCell(index, event.target.value)}
                          aria-label={`Header ${index + 1}`}
                        />
                        <button
                          type="button"
                          className="excel-table-tool__icon-btn"
                          onClick={() => handleRemoveColumn(index)}
                          aria-label={`Remove column ${index + 1}`}
                          title="Remove column"
                        >
                          ×
                        </button>
                        <button
                          type="button"
                          className="excel-table-tool__menu-btn"
                          onClick={(event) => openColumnMenu(index, event.currentTarget)}
                          aria-label={`Open menu for ${header || `Column ${index + 1}`}`}
                          title="Column menu"
                        >
                          ☰
                        </button>
                      </div>
                    </th>
                  );
                })}
              </tr>
            </thead>
          )}
          <tbody>
            {headers.length === 0 ? (
              <tr>
                <td className="excel-table-tool__empty-cell">
                  Click this table and paste Excel data (Ctrl+V), or use Import File above
                </td>
              </tr>
            ) : rows.length === 0 ? (
              <tr>
                <td className="excel-table-tool__single-row-note" colSpan={headers.length + 1}>
                  Header row parsed successfully. No data rows found.
                </td>
              </tr>
            ) : visibleRowIndexes.length === 0 ? (
              <tr>
                <td className="excel-table-tool__single-row-note" colSpan={headers.length + 1}>
                  No rows match the current search/filter.
                </td>
              </tr>
            ) : (
              groupedPageItems.map((item) => {
                if (item.type === 'group') {
                  const colSpan = headers.length - hiddenColumns.size + 1;
                  return (
                    <tr
                      key={item.key}
                      className={`excel-table-tool__group-row excel-table-tool__group-row-level-${Math.min(item.level + 1, 3)}`}
                    >
                      <td colSpan={colSpan}>
                        <span className="excel-table-tool__group-row-icon" aria-hidden="true">▶</span>
                        <strong>{item.columnLabel}</strong>: {item.value}
                        <span className="excel-table-tool__group-row-count"> ({item.count})</span>
                      </td>
                    </tr>
                  );
                }

                const rowIndex = item.rowIndex;
                const row = rows[rowIndex] || [];

                return (
                  <tr
                    key={`row-${rowIndex}`}
                    className={selectedRows[rowIndex] ? 'excel-table-tool__row-selected' : ''}
                  >
                    <td className="excel-table-tool__row-select-col">
                      <input
                        type="checkbox"
                        checked={!!selectedRows[rowIndex]}
                        onChange={() => toggleRowSelection(rowIndex)}
                        aria-label={`Select row ${rowIndex + 1}`}
                      />
                    </td>
                    {row.map((cell, cellIndex) => {
                      if (hiddenColumns.has(cellIndex)) return null;
                      return (
                        <td
                          key={`cell-${rowIndex}-${cellIndex}`}
                          className={isCellSelected(rowIndex, cellIndex) ? 'excel-table-tool__cell-selected' : ''}
                          onMouseDown={() => startSelection(rowIndex, cellIndex)}
                          onMouseEnter={() => extendSelection(rowIndex, cellIndex)}
                          onDoubleClick={() => {
                            setEditingCell({ row: rowIndex, col: cellIndex });
                            setEditingValue(cell);
                          }}
                        >
                          {editingCell?.row === rowIndex && editingCell?.col === cellIndex ? (
                            <input
                              type="text"
                              className="excel-table-tool__cell-input"
                              value={editingValue}
                              onChange={(event) => setEditingValue(event.target.value)}
                              onBlur={commitEditCell}
                              onKeyDown={(event) => {
                                if (event.key === 'Enter') {
                                  commitEditCell();
                                }

                                if (event.key === 'Escape') {
                                  cancelEditCell();
                                }
                              }}
                              autoFocus
                              aria-label={`Row ${rowIndex + 1} Column ${cellIndex + 1}`}
                            />
                          ) : (
                            <span className="excel-table-tool__cell-value">{cell}</span>
                          )}
                        </td>
                      );
                    })}
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>

      {/* ── Column Filter Menu ── */}
      {menuColumnIndex !== null && (
        <div
          className="excel-table-tool__column-menu"
          style={{ top: `${menuPosition.top}px`, left: `${menuPosition.left}px` }}
        >
          <button
            type="button"
            className="excel-table-tool__menu-item"
            onClick={setSortAscending}
          >
            Sort A → Z
          </button>
          <button
            type="button"
            className="excel-table-tool__menu-item"
            onClick={setSortDescending}
          >
            Sort Z → A
          </button>
          <button
            type="button"
            className="excel-table-tool__menu-item"
            onClick={clearSort}
          >
            Remove Sort
          </button>

          <div className="excel-table-tool__menu-divider" />

          <div className="excel-table-tool__menu-label">Show rows where:</div>
          <div className="excel-table-tool__menu-row">
            <select
              className="excel-table-tool__menu-select"
              value={menuDraft.condition1.operator}
              onChange={(event) =>
                updateMenuDraft('condition1.operator', event.target.value)
              }
            >
              {FILTER_OPERATORS.map((operator) => (
                <option key={operator.value} value={operator.value}>
                  {operator.label}
                </option>
              ))}
            </select>
            <input
              type="text"
              className="excel-table-tool__menu-input"
              value={menuDraft.condition1.value}
              onChange={(event) =>
                updateMenuDraft('condition1.value', event.target.value)
              }
              placeholder="Enter value"
            />
          </div>

          <div className="excel-table-tool__menu-row">
            <select
              className="excel-table-tool__menu-select excel-table-tool__menu-logic"
              value={menuDraft.logic}
              onChange={(event) => updateMenuDraft('logic', event.target.value)}
            >
              <option value="AND">And</option>
              <option value="OR">Or</option>
            </select>
          </div>

          <div className="excel-table-tool__menu-row">
            <select
              className="excel-table-tool__menu-select"
              value={menuDraft.condition2.operator}
              onChange={(event) =>
                updateMenuDraft('condition2.operator', event.target.value)
              }
            >
              {FILTER_OPERATORS.map((operator) => (
                <option key={`${operator.value}-2`} value={operator.value}>
                  {operator.label}
                </option>
              ))}
            </select>
            <input
              type="text"
              className="excel-table-tool__menu-input"
              value={menuDraft.condition2.value}
              onChange={(event) =>
                updateMenuDraft('condition2.value', event.target.value)
              }
              placeholder="Enter value"
            />
          </div>

          <div className="excel-table-tool__menu-actions">
            <button
              type="button"
              className="excel-table-tool__menu-apply"
              onClick={applyColumnFilter}
            >
              FILTER
            </button>
            <button
              type="button"
              className="excel-table-tool__menu-clear"
              onClick={clearColumnFilter}
            >
              CLEAR
            </button>
            <button
              type="button"
              className="excel-table-tool__menu-close"
              onClick={closeColumnMenu}
            >
              ✕
            </button>
          </div>
        </div>
      )}

      {/* ── Pagination ── */}
      {visibleRowIndexes.length > 0 && (
        <div className="excel-table-tool__pagination">
          <label className="excel-table-tool__page-size-label">
            Rows per page:
            <select
              className="excel-table-tool__page-size-select"
              value={rowsPerPage}
              onChange={(event) => setRowsPerPage(Number(event.target.value))}
              aria-label="Rows per page"
            >
              {PAGE_SIZE_OPTIONS.map((size) => (
                <option key={size} value={size}>{size}</option>
              ))}
            </select>
          </label>

          <button
            type="button"
            className="excel-table-tool__page-btn"
            onClick={() => setCurrentPage((prev) => Math.max(1, prev - 1))}
            disabled={currentPage === 1}
          >
            Prev
          </button>

          <span className="excel-table-tool__page-info">
            Page {currentPage} of {totalPages}
          </span>

          {Array.from({ length: totalPages }, (_, index) => index + 1)
            .filter((pageNumber) => {
              if (totalPages <= 7) return true;
              if (pageNumber === 1 || pageNumber === totalPages) return true;
              return Math.abs(pageNumber - currentPage) <= 2;
            })
            .reduce((acc, pageNumber, idx, arr) => {
              if (idx > 0 && pageNumber - arr[idx - 1] > 1) {
                acc.push({ ellipsis: true, key: `ellipsis-${pageNumber}` });
              }
              acc.push({ pageNumber, key: `page-${pageNumber}` });
              return acc;
            }, [])
            .map((item) =>
              item.ellipsis ? (
                <span key={item.key} className="excel-table-tool__page-ellipsis">…</span>
              ) : (
                <button
                  key={item.key}
                  type="button"
                  className={`excel-table-tool__page-btn ${
                    item.pageNumber === currentPage ? 'is-active' : ''
                  }`}
                  onClick={() => setCurrentPage(item.pageNumber)}
                >
                  {item.pageNumber}
                </button>
              )
            )}

          <button
            type="button"
            className="excel-table-tool__page-btn"
            onClick={() => setCurrentPage((prev) => Math.min(totalPages, prev + 1))}
            disabled={currentPage === totalPages}
          >
            Next
          </button>
        </div>
      )}

      {/* ── Primary Key Violations Log ── */}
      {blockedRows.length > 0 && (
        <section className="excel-table-tool__duplicates">
          <div className="excel-table-tool__duplicates-header">
            <h3>Primary Key Violations ({blockedRows.length})</h3>
            <div className="excel-table-tool__duplicates-actions">
              <button type="button" className="btn btn--ghost" onClick={clearBlockedRows}>
                Clear Log
              </button>
            </div>
          </div>
          <div className="excel-table-tool__duplicates-list">
            {blockedRows.map((item) => (
              <div key={item.id} className="excel-table-tool__duplicate-item">
                <span className="excel-table-tool__duplicate-source">
                  {item.source} • {item.keyValue}
                </span>
                <span className="excel-table-tool__duplicate-values">
                  {item.reason} — {item.row.join(' | ')}
                </span>
              </div>
            ))}
          </div>
        </section>
      )}

      {error && <p className="excel-table-tool__error">{error}</p>}
    </section>
  );
};

export default ExcelPasteTable;
