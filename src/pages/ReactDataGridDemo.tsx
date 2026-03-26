import React, { useState, useCallback, useMemo, useRef, useEffect } from 'react';
import { DataGrid, Column, SortColumn, RenderEditCellProps, SelectColumn } from 'react-data-grid';
import ExcelJS from 'exceljs';
import FileUpload from '../components/FileUpload';
import './DemoPage.css';
import './ReactDataGridDemo.css';
import 'react-data-grid/lib/styles.css';

interface Row {
  [key: string]: any;
  _rowIndex?: number; // For row identification
}

interface DropdownOptions {
  [columnKey: string]: string[];
}

// Helper function for extracting sheet names from formulas
const extractAllSheetNamesFromFormula = (formula: string): string[] => {
  const names: string[] = [];
  const regex = /'([^']+)'!|(?<![A-Za-z_])([A-Za-z_]\w*)!/g;
  let m: RegExpExecArray | null;
  while ((m = regex.exec(formula)) !== null) {
    names.push(m[1] || m[2]);
  }
  return names;
};

const resolveRangeReference = (
  formula: string,
  workbook: ExcelJS.Workbook
): string[] => {
  const match = formula.match(
    /^'?([^'!]+)'?!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/i
  );
  if (!match) return [];
  const [, sheetName, startCol, startRowStr, , endRowStr] = match;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return [];
  const start = parseInt(startRowStr, 10);
  const end = parseInt(endRowStr, 10);
  const colIndex = startCol.toUpperCase().split('').reduce((acc, ch, i, arr) =>
    acc + (ch.charCodeAt(0) - 64) * Math.pow(26, arr.length - i - 1), 0);
  const values: string[] = [];
  for (let r = start; r <= end; r++) {
    const cell = ws.getRow(r).getCell(colIndex);
    let v: any = cell.value;
    if (v == null) continue;
    if (typeof v === 'object' && 'result' in v) v = v.result;
    if (typeof v === 'object' && 'richText' in v)
      v = (v as any).richText.map((rt: any) => rt.text).join('');
    const str = String(v).trim();
    if (str) values.push(str);
  }
  return values;
};

// Simple text editor for regular cells
function TextEditor<TRow>({ row, column, onRowChange, onClose }: RenderEditCellProps<TRow>) {
  const inputRef = useRef<HTMLInputElement>(null);
  const value = (row as any)[column.key] ?? '';

  useEffect(() => {
    inputRef.current?.focus();
    inputRef.current?.select();
  }, []);

  return (
    <input
      ref={inputRef}
      className="rdg-text-editor"
      value={value}
      onChange={(e) => onRowChange({ ...row, [column.key]: e.target.value })}
      onBlur={() => onClose(true)}
      onKeyDown={(e) => {
        if (e.key === 'Enter') onClose(true);
        if (e.key === 'Escape') onClose(false);
      }}
    />
  );
}

// Dropdown Editor Component
function DropdownEditor({
  row,
  column,
  onRowChange,
  onClose,
  options,
}: RenderEditCellProps<Row> & { options: string[] }) {
  const [search, setSearch] = useState('');
  const inputRef = useRef<HTMLInputElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    setTimeout(() => inputRef.current?.focus(), 0);
  }, []);

  const filtered = options.filter((opt) =>
    opt.toLowerCase().includes(search.toLowerCase())
  );

  const handleSelect = (value: string) => {
    onRowChange({ ...row, [column.key]: value }, true);
  };

  return (
    <div ref={containerRef} className="rdg-dropdown-editor">
      <input
        ref={inputRef}
        type="text"
        className="rdg-dropdown-search"
        placeholder="Search..."
        value={search}
        onChange={(e) => setSearch(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Escape') onClose();
          if (e.key === 'Enter' && filtered.length === 1) {
            handleSelect(filtered[0]);
          }
        }}
      />
      <div className="rdg-dropdown-list">
        {filtered.map((opt) => (
          <div
            key={opt}
            className={`rdg-dropdown-item${row[column.key] === opt ? ' selected' : ''}`}
            onMouseDown={(e) => {
              e.preventDefault();
              handleSelect(opt);
            }}
          >
            {opt}
          </div>
        ))}
        {filtered.length === 0 && (
          <div className="rdg-dropdown-empty">No matches</div>
        )}
      </div>
    </div>
  );
}

const ReactDataGridDemo: React.FC = () => {
  const [rows, setRows] = useState<Row[]>([]);
  const [columns, setColumns] = useState<Column<Row>[]>([]);
  const [fileName, setFileName] = useState('');
  const [sortColumns, setSortColumns] = useState<readonly SortColumn[]>([]);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [activeSheet, setActiveSheet] = useState(0);
  const [allSheetsData, setAllSheetsData] = useState<{ rows: Row[]; columns: Column<Row>[]; dropdownOptions: DropdownOptions }[]>([]);
  const [dropdownOptions, setDropdownOptions] = useState<DropdownOptions>({});
  const [selectedRows, setSelectedRows] = useState<ReadonlySet<number>>(new Set());

  // Add row index for selection tracking
  const rowsWithIndex = useMemo(() => 
    rows.map((row, idx) => ({ ...row, _rowIndex: idx })), 
    [rows]
  );

  const rowKeyGetter = (row: Row) => row._rowIndex ?? 0;

  // Sort rows while preserving _rowIndex
  const sortedRowsWithIndex = useMemo(() => {
    if (sortColumns.length === 0) return rowsWithIndex;

    return [...rowsWithIndex].sort((a, b) => {
      for (const sort of sortColumns) {
        const { columnKey, direction } = sort;
        if (columnKey === '_rowIndex') continue;
        const aVal = (a as Row)[columnKey];
        const bVal = (b as Row)[columnKey];

        if (aVal == null && bVal == null) continue;
        if (aVal == null) return direction === 'ASC' ? -1 : 1;
        if (bVal == null) return direction === 'ASC' ? 1 : -1;

        let comparison = 0;
        if (typeof aVal === 'number' && typeof bVal === 'number') {
          comparison = aVal - bVal;
        } else {
          comparison = String(aVal).localeCompare(String(bVal));
        }

        if (comparison !== 0) {
          return direction === 'ASC' ? comparison : -comparison;
        }
      }
      return 0;
    });
  }, [rowsWithIndex, sortColumns]);

  const deleteSelectedRows = () => {
    if (selectedRows.size === 0) return;
    const newRows = rows.filter((_, idx) => !selectedRows.has(idx));
    setRows(newRows);
    setSelectedRows(new Set());
    // Update allSheetsData
    const updatedSheets = [...allSheetsData];
    if (updatedSheets[activeSheet]) {
      updatedSheets[activeSheet] = { ...updatedSheets[activeSheet], rows: newRows };
      setAllSheetsData(updatedSheets);
    }
  };

  const addRow = () => {
    const newRow: Row = {};
    if (columns.length > 0) {
      columns.forEach(col => {
        if (col.key !== '_rowIndex') {
          newRow[col.key] = '';
        }
      });
    }
    const newRows = [...rows, newRow];
    setRows(newRows);
    // Update allSheetsData
    const updatedSheets = [...allSheetsData];
    if (updatedSheets[activeSheet]) {
      updatedSheets[activeSheet] = { ...updatedSheets[activeSheet], rows: newRows };
      setAllSheetsData(updatedSheets);
    }
  };

  const handleFileUpload = useCallback(async (file: File) => {
    setFileName(file.name);
    
    const arrayBuffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);

    // Step 1: Build defined names map
    const definedNameMap = new Map<string, string[]>();
    try {
      const names = (workbook as any).definedNames;
      if (names?.model) {
        Object.entries(names.model).forEach(([name, def]: [string, any]) => {
          const ranges = Array.isArray(def.ranges) ? def.ranges : [def.ranges];
          definedNameMap.set(name, ranges.filter(Boolean).map(String));
        });
      }
    } catch (_) {}

    // Step 2: Identify sheets to exclude (hidden + lookup sheets)
    const sheetsToExclude = new Set<string>();
    
    workbook.eachSheet((ws) => {
      if (ws.state === 'hidden' || ws.state === 'veryHidden') {
        sheetsToExclude.add(ws.name);
      }
      // Exclude sheets that look like lookup/options sheets
      if (/^(Options|Lookup)_/i.test(ws.name)) {
        sheetsToExclude.add(ws.name);
      }
    });

    // Find sheets referenced in data validation formulas
    workbook.eachSheet((ws) => {
      const dvModel = (ws as any).dataValidations?.model;
      if (!dvModel) return;
      Object.keys(dvModel).forEach((addr) => {
        const dv = dvModel[addr];
        if (dv?.type !== 'list' || !dv.formulae) return;
        const raw: string = dv.formulae[0];
        if (!raw) return;
        const clean = raw.replace(/^=/, '');
        for (const name of extractAllSheetNamesFromFormula(clean)) {
          sheetsToExclude.add(name);
        }
        if (!clean.includes('!') && definedNameMap.has(clean)) {
          for (const range of definedNameMap.get(clean)!) {
            for (const name of extractAllSheetNamesFromFormula(range)) {
              sheetsToExclude.add(name);
            }
          }
        }
      });
    });

    // Helper: resolve dropdown options from formula
    const resolveDropdownOptions = (formula: string, currentSheet: ExcelJS.Worksheet): string[] => {
      const clean = formula.replace(/^=/, '');
      // Sheet reference like 'SheetName'!A1:A10
      if (/^'?[^(]+!\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+$/i.test(clean)) {
        return resolveRangeReference(clean, workbook);
      }
      // Defined name
      if (definedNameMap.has(clean)) {
        const values: string[] = [];
        for (const range of definedNameMap.get(clean)!) {
          values.push(...resolveRangeReference(range, workbook));
        }
        if (values.length > 0) return values;
      }
      // Local range reference
      if (/^\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+$/i.test(clean)) {
        return resolveRangeReference(`'${currentSheet.name}'!${clean}`, workbook);
      }
      // Extract from referenced sheet
      const refSheetNames = extractAllSheetNamesFromFormula(clean);
      if (refSheetNames.length > 0) {
        const colMatch = clean.match(/!\$?([A-Z]+)/i);
        if (colMatch) {
          const refSheet = workbook.getWorksheet(refSheetNames[0]);
          if (refSheet) {
            const colIndex = colMatch[1].toUpperCase().split('').reduce((acc, ch, i, arr) =>
              acc + (ch.charCodeAt(0) - 64) * Math.pow(26, arr.length - i - 1), 0);
            const values: string[] = [];
            refSheet.eachRow({ includeEmpty: false }, (row) => {
              const cell = row.getCell(colIndex);
              let v: any = cell.value;
              if (v == null) return;
              if (typeof v === 'object' && 'result' in v) v = v.result;
              const str = String(v).trim();
              if (str) values.push(str);
            });
            return values;
          }
        }
      }
      // Comma-separated values
      return clean.replace(/^"/, '').replace(/"$/, '').split(',').map((s) => s.trim()).filter(Boolean);
    };

    // Step 3: Build data for visible sheets only
    const sheets: { rows: Row[]; columns: Column<Row>[]; dropdownOptions: DropdownOptions }[] = [];
    const names: string[] = [];

    workbook.eachSheet((worksheet) => {
      if (sheetsToExclude.has(worksheet.name)) return;
      
      names.push(worksheet.name);
      
      const sheetDropdownOptions: DropdownOptions = {};
      const columnIndexToKey: { [colIndex: number]: string } = {};
      
      // Get headers from first row
      const headerRow = worksheet.getRow(1);
      const headers: string[] = [];
      headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const header = cell.text || `Column${colNumber}`;
        headers.push(header);
        columnIndexToKey[colNumber] = header;
      });

      // Parse data validations for dropdowns
      const dataValidations = (worksheet as any).dataValidations;
      if (dataValidations?.model) {
        Object.entries(dataValidations.model).forEach(([address, validation]: [string, any]) => {
          if (validation?.type === 'list' && validation.formulae?.[0]) {
            const formula = validation.formulae[0];
            const options = resolveDropdownOptions(formula, worksheet);
            
            if (options.length > 0) {
              // Parse address to get column
              const match = address.match(/([A-Z]+)/);
              if (match) {
                const colLetter = match[1];
                const colIndex = colLetter.split('').reduce((acc, char, i, arr) =>
                  acc + (char.charCodeAt(0) - 64) * Math.pow(26, arr.length - i - 1), 0);
                const colKey = columnIndexToKey[colIndex];
                if (colKey && !sheetDropdownOptions[colKey]) {
                  sheetDropdownOptions[colKey] = options;
                }
              }
            }
          }
        });
      }

      // Parse rows data
      const jsonData: Row[] = [];
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const rowData: Row = {};
        headers.forEach((header, idx) => {
          const cell = row.getCell(idx + 1);
          let value: any = cell.value;
          if (value && typeof value === 'object' && 'result' in value) {
            value = value.result;
          }
          rowData[header] = cell.text || value || '';
        });
        jsonData.push(rowData);
      });

      // Create columns with dropdown editors where applicable
      const cols: Column<Row>[] = headers.map((header) => {
        const hasDropdown = sheetDropdownOptions[header] && sheetDropdownOptions[header].length > 0;
        const options = sheetDropdownOptions[header] || [];
        
        return {
          key: header,
          name: hasDropdown ? `${header} ▼` : header,
          resizable: true,
          sortable: true,
          editable: true,
          width: 150,
          renderCell: ({ row }: { row: Row }) => {
            const value = row[header];
            if (hasDropdown) {
              const isValid = value === '' || value == null || options.includes(String(value));
              return (
                <span className={`rdg-dropdown-cell${isValid ? '' : ' rdg-invalid-value'}`}>
                  {!isValid && <span className="rdg-invalid-icon" title="Value not in dropdown list">⚠</span>}
                  <span className="rdg-cell-value">{value}</span>
                  <span className="rdg-dropdown-indicator">▼</span>
                </span>
              );
            }
            if (typeof value === 'number') {
              return <span style={{ color: '#1d4ed8' }}>{value.toLocaleString()}</span>;
            }
            return <span>{value}</span>;
          },
          renderEditCell: hasDropdown
            ? (props: RenderEditCellProps<Row>) => (
                <DropdownEditor {...props} options={options} />
              )
            : TextEditor,
        };
      });

      sheets.push({ rows: jsonData, columns: cols, dropdownOptions: sheetDropdownOptions });
    });

    setSheetNames(names);
    setAllSheetsData(sheets);
    if (sheets.length > 0) {
      setRows(sheets[0].rows);
      setColumns(sheets[0].columns);
      setDropdownOptions(sheets[0].dropdownOptions);
    }
    setActiveSheet(0);
    setSortColumns([]);
    setSelectedRows(new Set());
  }, []);

  const handleSheetChange = (index: number) => {
    setActiveSheet(index);
    setRows(allSheetsData[index]?.rows || []);
    setColumns(allSheetsData[index]?.columns || []);
    setDropdownOptions(allSheetsData[index]?.dropdownOptions || {});
    setSortColumns([]);
    setSelectedRows(new Set());
  };

  const handleRowsChange = (newRows: Row[]) => {
    setRows(newRows);
  };

  const createSampleData = () => {
    const statusOptions = ['Active', 'On Leave', 'Remote', 'Terminated'];
    const departmentOptions = ['Engineering', 'Marketing', 'Sales', 'HR', 'Finance', 'Operations'];
    
    const sampleDropdownOptions: DropdownOptions = {
      department: departmentOptions,
      status: statusOptions,
    };

    const sampleRows: Row[] = [
      { id: 1, name: 'John Doe', department: 'Engineering', salary: 75000, startDate: '2020-01-15', status: 'Active' },
      { id: 2, name: 'Jane Smith', department: 'Marketing', salary: 65000, startDate: '2021-03-22', status: 'Active' },
      { id: 3, name: 'Bob Johnson', department: 'Sales', salary: 85000, startDate: '2018-07-01', status: 'Active' },
      { id: 4, name: 'Alice Brown', department: 'HR', salary: 70000, startDate: '2019-11-10', status: 'On Leave' },
      { id: 5, name: 'Charlie Wilson', department: 'Engineering', salary: 72000, startDate: '2022-02-28', status: 'Active' },
      { id: 6, name: 'Diana Miller', department: 'Finance', salary: 78000, startDate: '2020-06-15', status: 'Active' },
      { id: 7, name: 'Edward Davis', department: 'Engineering', salary: 82000, startDate: '2019-04-20', status: 'Active' },
      { id: 8, name: 'Fiona Garcia', department: 'Marketing', salary: 68000, startDate: '2021-09-01', status: 'Remote' },
      { id: 9, name: 'Invalid User', department: 'InvalidDept', salary: 50000, startDate: '2023-01-01', status: 'Unknown' },
    ];

    const sampleColumns: Column<Row>[] = [
      { key: 'id', name: 'ID', width: 60, sortable: true, editable: true, renderEditCell: TextEditor },
      { 
        key: 'name', 
        name: 'Name', 
        width: 150, 
        sortable: true,
        resizable: true,
        editable: true,
        renderEditCell: TextEditor,
      },
      { 
        key: 'department', 
        name: 'Department ▼', 
        width: 140, 
        sortable: true,
        resizable: true,
        editable: true,
        renderCell: ({ row }) => {
          const isValid = departmentOptions.includes(row.department);
          return (
            <span className={`rdg-dropdown-cell${isValid ? '' : ' rdg-invalid-value'}`}>
              {!isValid && row.department && <span className="rdg-invalid-icon" title="Value not in dropdown list">⚠</span>}
              <span className="rdg-cell-value">{row.department}</span>
              <span className="rdg-dropdown-indicator">▼</span>
            </span>
          );
        },
        renderEditCell: (props: RenderEditCellProps<Row>) => (
          <DropdownEditor {...props} options={departmentOptions} />
        ),
      },
      { 
        key: 'salary', 
        name: 'Salary', 
        width: 100, 
        sortable: true,
        resizable: true,
        editable: true,
        renderEditCell: TextEditor,
        renderCell: ({ row }) => (
          <span style={{ color: '#1d4ed8' }}>${row.salary.toLocaleString()}</span>
        ),
      },
      { 
        key: 'startDate', 
        name: 'Start Date', 
        width: 120, 
        sortable: true,
        resizable: true,
        editable: true,
        renderEditCell: TextEditor,
        renderCell: ({ row }) => (
          <span style={{ color: '#7c3aed' }}>{row.startDate}</span>
        ),
      },
      { 
        key: 'status', 
        name: 'Status ▼', 
        width: 120, 
        sortable: true,
        resizable: true,
        editable: true,
        renderCell: ({ row }) => {
          const isValid = statusOptions.includes(row.status);
          const colors: { [key: string]: string } = {
            'Active': '#10b981',
            'On Leave': '#f59e0b',
            'Remote': '#3b82f6',
            'Terminated': '#ef4444',
          };
          return (
            <span className={`rdg-dropdown-cell${isValid ? '' : ' rdg-invalid-value'}`}>
              {!isValid && row.status && <span className="rdg-invalid-icon" title="Value not in dropdown list">⚠</span>}
              <span 
                style={{ 
                  background: isValid ? (colors[row.status] || '#6b7280') : '#fecaca',
                  color: isValid ? 'white' : '#991b1b',
                  padding: '2px 8px',
                  borderRadius: '4px',
                  fontSize: '0.75rem',
                }}
              >
                {row.status}
              </span>
              <span className="rdg-dropdown-indicator">▼</span>
            </span>
          );
        },
        renderEditCell: (props: RenderEditCellProps<Row>) => (
          <DropdownEditor {...props} options={statusOptions} />
        ),
      },
    ];

    setRows(sampleRows);
    setColumns(sampleColumns);
    setDropdownOptions(sampleDropdownOptions);
    setFileName('sample-data.xlsx');
    setSheetNames(['Employees']);
    setAllSheetsData([{ rows: sampleRows, columns: sampleColumns, dropdownOptions: sampleDropdownOptions }]);
    setActiveSheet(0);
    setSelectedRows(new Set());
  };

  return (
    <div className="demo-page">
      <header className="demo-header">
        <h1>React Data Grid</h1>
        <span className="license-tag">MIT</span>
      </header>

      <div className="demo-info">
        <p>
          React Data Grid is a feature-rich grid component for React. It provides excellent performance
          with virtual scrolling and supports sorting, filtering, cell editing, and column resizing.
        </p>
        <div className="features-list">
          <span className="feature">✓ Virtual Scrolling</span>
          <span className="feature">✓ Column Sorting</span>
          <span className="feature">✓ Column Resizing</span>
          <span className="feature">✓ Dropdown Editors</span>
          <span className="feature">✓ Excel Validation Support</span>
        </div>
      </div>

      <div className="upload-section">
        <FileUpload onFileUpload={handleFileUpload} />
        <div className="sample-data-option">
          <p>or</p>
          <button className="sample-data-btn" onClick={createSampleData}>
            Load Sample Data
          </button>
        </div>
      </div>

      {rows.length > 0 && (
        <div className="result-section">
          <div className="result-header">
            <h2>
              <span className="file-icon">📊</span>
              {fileName || 'Data Grid'}
              <span className="editable-badge">Editable</span>
            </h2>
            {sheetNames.length > 1 && (
              <div className="sheet-tabs">
                {sheetNames.map((name, idx) => (
                  <button
                    key={name}
                    className={`sheet-tab ${idx === activeSheet ? 'active' : ''}`}
                    onClick={() => handleSheetChange(idx)}
                  >
                    {name}
                  </button>
                ))}
              </div>
            )}
          </div>

          <div className="rdg-toolbar">
            <button className="rdg-add-btn" onClick={addRow}>
              + Add Row
            </button>
            <button 
              className="rdg-delete-btn" 
              onClick={deleteSelectedRows}
              disabled={selectedRows.size === 0}
            >
              Delete Selected ({selectedRows.size})
            </button>
          </div>

          <div className="rdg-wrapper">
            <DataGrid
              columns={[SelectColumn, ...columns]}
              rows={sortedRowsWithIndex}
              onRowsChange={(newRows) => {
                const cleanRows = newRows.map(({ _rowIndex, ...rest }) => rest);
                handleRowsChange(cleanRows);
              }}
              sortColumns={sortColumns}
              onSortColumnsChange={setSortColumns}
              rowKeyGetter={rowKeyGetter}
              selectedRows={selectedRows}
              onSelectedRowsChange={setSelectedRows}
              className="rdg-light"
              style={{ height: '100%' }}
            />
          </div>

          <div className="data-info">
            <h4>Interactive Features</h4>
            <p>• Select rows using checkboxes, then click "Delete Selected"</p>
            <p>• Double-click cells to edit (dropdown columns show searchable list)</p>
            <p>• Click column headers to sort (click again to reverse)</p>
            <p>• Drag column edges to resize</p>
            <p>• {rows.length} rows × {columns.length} columns | {selectedRows.size} selected</p>
            {Object.keys(dropdownOptions).length > 0 && (
              <p>• Dropdown columns: {Object.keys(dropdownOptions).join(', ')}</p>
            )}
          </div>
        </div>
      )}

      <div className="code-example">
        <div className="code-header">Usage Example</div>
        <pre>
          <code>{`import DataGrid from 'react-data-grid';
import 'react-data-grid/lib/styles.css';

const columns = [
  { key: 'id', name: 'ID' },
  { key: 'name', name: 'Name', sortable: true },
  { key: 'salary', name: 'Salary', sortable: true },
];

const rows = [
  { id: 1, name: 'John', salary: 75000 },
  { id: 2, name: 'Jane', salary: 65000 },
];

function MyGrid() {
  return (
    <DataGrid
      columns={columns}
      rows={rows}
      sortColumns={sortColumns}
      onSortColumnsChange={setSortColumns}
    />
  );
}`}</code>
        </pre>
      </div>
    </div>
  );
};

export default ReactDataGridDemo;
