import React, { useState, useCallback, useMemo } from 'react';
import { RevoGrid } from '@revolist/react-datagrid';
import type { ColumnRegular } from '@revolist/revogrid';
import SelectColumnType from '@revolist/revogrid-column-select';
import ExcelJS from 'exceljs';
import FileUpload from '../components/FileUpload';
import './DemoPage.css';
import './RevoGridDemo.css';

// ——————————————————————————————————————————————————————
// Helper functions for Excel parsing
// ——————————————————————————————————————————————————————

const parseCellRef = (ref: string): { col: number; row: number } | null => {
  const match = ref.match(/^\$?([A-Z]+)\$?(\d+)$/i);
  if (!match) return null;
  const colStr = match[1].toUpperCase();
  const row = parseInt(match[2], 10);
  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }
  return { col, row };
};

const resolveRangeReference = (
  formula: string,
  workbook: ExcelJS.Workbook
): string[] => {
  const clean = formula.replace(/^=/, '');
  const m = clean.match(
    /^'?([^'!]+)'?!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/i
  );
  if (!m) return [];
  const [, sheetName, startCol, startRow, endCol, endRow] = m;
  const refSheet = workbook.getWorksheet(sheetName);
  if (!refSheet) return [];
  const startRowNum = parseInt(startRow, 10);
  const endRowNum = parseInt(endRow, 10);
  const startColRef = parseCellRef(`${startCol}1`);
  const endColRef = parseCellRef(`${endCol}1`);
  if (!startColRef || !endColRef) return [];
  const values: string[] = [];
  for (let row = startRowNum; row <= endRowNum; row++) {
    for (let col = startColRef.col; col <= endColRef.col; col++) {
      const cell = refSheet.getCell(row, col);
      let cellValue: any = cell.value;
      if (cellValue === null || cellValue === undefined) continue;
      if (typeof cellValue === 'object' && 'result' in cellValue)
        cellValue = cellValue.result;
      else if (typeof cellValue === 'object' && 'richText' in cellValue)
        cellValue = (cellValue as ExcelJS.CellRichTextValue).richText
          .map((rt) => rt.text)
          .join('');
      const str = String(cellValue).trim();
      if (str) values.push(str);
    }
  }
  return values;
};

const parseAddressRange = (
  address: string
): {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
} | null => {
  const rangeMatch = address.match(
    /^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/i
  );
  if (rangeMatch) {
    const s = parseCellRef(`${rangeMatch[1]}${rangeMatch[2]}`);
    const e = parseCellRef(`${rangeMatch[3]}${rangeMatch[4]}`);
    if (s && e) {
      return {
        startRow: s.row - 1,
        startCol: s.col - 1,
        endRow: e.row - 1,
        endCol: e.col - 1,
      };
    }
  }
  const cellMatch = address.match(/^\$?([A-Z]+)\$?(\d+)$/i);
  if (cellMatch) {
    const ref = parseCellRef(`${cellMatch[1]}${cellMatch[2]}`);
    if (ref) {
      return {
        startRow: ref.row - 1,
        startCol: ref.col - 1,
        endRow: ref.row - 1,
        endCol: ref.col - 1,
      };
    }
  }
  return null;
};

// ——————————————————————————————————————————————————————
// Types
// ——————————————————————————————————————————————————————

interface DropdownConfig {
  [key: string]: string[]; // columnProp -> options array
}

interface SheetData {
  name: string;
  columns: ColumnRegular[];
  source: Record<string, any>[];
  dropdowns: DropdownConfig;
}

// ——————————————————————————————————————————————————————
// Component
// ——————————————————————————————————————————————————————

const RevoGridDemo: React.FC = () => {
  const [sheets, setSheets] = useState<SheetData[]>([]);
  const [activeSheetIndex, setActiveSheetIndex] = useState(0);
  const [showGrid, setShowGrid] = useState(false);
  const [fileName, setFileName] = useState<string>('');
  const [error, setError] = useState<string | null>(null);
  const [selectedRowIndex, setSelectedRowIndex] = useState<number | null>(null);

  // Download selected row as Excel file
  const downloadSelectedRow = useCallback(async () => {
    if (selectedRowIndex === null || !sheets[activeSheetIndex]) return;

    const activeSheet = sheets[activeSheetIndex];
    const rowData = activeSheet.source[selectedRowIndex];
    if (!rowData) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Row Data');

    // Add headers
    const headers = activeSheet.columns.map((col) => col.name);
    worksheet.addRow(headers);

    // Style header row
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE2E8F0' },
    };

    // Add data row
    const values = activeSheet.columns.map((col) => rowData[col.prop as string] ?? '');
    worksheet.addRow(values);

    // Auto-fit columns
    worksheet.columns.forEach((col, idx) => {
      const header = headers[idx] || '';
      const value = String(values[idx] || '');
      col.width = Math.max(header.length, value.length, 10) + 2;
    });

    // Generate and download
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `row_${selectedRowIndex + 1}_${activeSheet.name || 'data'}.xlsx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }, [selectedRowIndex, sheets, activeSheetIndex]);

  const handleFileSelect = useCallback(async (file: File) => {
    setError(null);
    setFileName(file.name);

    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(arrayBuffer);

      const parsedSheets: SheetData[] = [];

      workbook.eachSheet((worksheet, sheetId) => {
        // Skip hidden sheets
        if (worksheet.state === 'hidden' || worksheet.state === 'veryHidden') {
          return;
        }

        const rowCount = worksheet.rowCount;
        const colCount = worksheet.columnCount;
        if (rowCount === 0 || colCount === 0) return;

        // Build columns from header row (row 1)
        const columns: ColumnRegular[] = [];
        const headerRow = worksheet.getRow(1);
        for (let c = 1; c <= colCount; c++) {
          const cell = headerRow.getCell(c);
          let headerValue = cell.value;
          if (typeof headerValue === 'object' && headerValue !== null) {
            if ('richText' in headerValue) {
              headerValue = (headerValue as ExcelJS.CellRichTextValue).richText
                .map((rt) => rt.text)
                .join('');
            } else if ('result' in headerValue) {
              headerValue = (headerValue as any).result;
            }
          }
          const headerText = headerValue ? String(headerValue) : `Column ${c}`;
          columns.push({
            prop: `col_${c}`,
            name: headerText,
            size: Math.max(headerText.length * 10 + 20, 100),
            sortable: true,
          });
        }

        // Build source data from rows 2+
        const source: Record<string, any>[] = [];
        for (let r = 2; r <= rowCount; r++) {
          const row = worksheet.getRow(r);
          const rowData: Record<string, any> = {};
          for (let c = 1; c <= colCount; c++) {
            const cell = row.getCell(c);
            let cellValue = cell.value;
            if (typeof cellValue === 'object' && cellValue !== null) {
              if ('richText' in cellValue) {
                cellValue = (cellValue as ExcelJS.CellRichTextValue).richText
                  .map((rt) => rt.text)
                  .join('');
              } else if ('result' in cellValue) {
                cellValue = (cellValue as any).result;
              }
            }
            rowData[`col_${c}`] = cellValue ?? '';
          }
          source.push(rowData);
        }

        // Parse data validations for dropdowns
        const dropdowns: DropdownConfig = {};
        const validations = (worksheet as any).dataValidations?.model;
        if (validations) {
          for (const address of Object.keys(validations)) {
            const rule = validations[address];
            if (rule.type !== 'list') continue;

            let options: string[] = [];
            if (rule.formulae && rule.formulae.length > 0) {
              const formula = rule.formulae[0];
              if (typeof formula === 'string') {
                if (formula.includes('!')) {
                  options = resolveRangeReference(formula, workbook);
                } else {
                  options = formula
                    .replace(/^"/, '')
                    .replace(/"$/, '')
                    .split(',')
                    .map((s) => s.trim());
                }
              }
            }

            if (options.length > 0) {
              const range = parseAddressRange(address);
              if (range) {
                for (let col = range.startCol; col <= range.endCol; col++) {
                  const colProp = `col_${col + 1}`;
                  dropdowns[colProp] = options;
                }
              }
            }
          }
        }

        parsedSheets.push({
          name: worksheet.name,
          columns,
          source,
          dropdowns,
        });
      });

      if (parsedSheets.length === 0) {
        setError('No visible sheets found in the workbook');
        return;
      }

      setSheets(parsedSheets);
      setActiveSheetIndex(0);
      setShowGrid(true);
    } catch (err) {
      setError(`Error parsing Excel file: ${err}`);
      console.error(err);
    }
  }, []);

  const activeSheet = sheets[activeSheetIndex];

  // Build columns with dropdown editors applied
  const columnsWithDropdowns = useMemo(() => {
    if (!activeSheet) return [];
    return activeSheet.columns.map((col) => {
      const options = activeSheet.dropdowns[col.prop as string];
      if (options && options.length > 0) {
        return {
          ...col,
          columnType: 'select',
          source: options,
        };
      }
      return col;
    });
  }, [activeSheet]);

  // Build source with row class for selected row
  const sourceWithRowClass = useMemo(() => {
    if (!activeSheet) return [];
    return activeSheet.source.map((row, idx) => ({
      ...row,
      _rowClass: idx === selectedRowIndex ? 'selected-row' : '',
    }));
  }, [activeSheet, selectedRowIndex]);

  // Column types including select
  const columnTypes = useMemo(() => ({
    select: new SelectColumnType(),
  }), []);

  return (
    <div className="demo-page">
      <header className="demo-header">
        <h2>RevoGrid Demo</h2>
        <p>High-performance virtual data grid with Excel-like features</p>
        <div className="package-info">
          <span className="package-name">@revolist/revogrid</span>
          <span className="package-license">MIT License</span>
        </div>
      </header>

      <section className="demo-section">
        <h3>Upload Excel File</h3>
        <FileUpload onFileUpload={handleFileSelect} accept=".xlsx,.xls" />
        {fileName && <p className="file-name">Loaded: {fileName}</p>}
        {error && <p className="error-message">{error}</p>}
      </section>

      {showGrid && activeSheet && (
        <section className="demo-section spreadsheet-section">
          <h3>RevoGrid View</h3>

          {/* Sheet tabs */}
          {sheets.length > 1 && (
            <div className="sheet-tabs">
              {sheets.map((sheet, idx) => (
                <button
                  key={sheet.name}
                  className={`sheet-tab ${idx === activeSheetIndex ? 'active' : ''}`}
                  onClick={() => setActiveSheetIndex(idx)}
                >
                  {sheet.name}
                </button>
              ))}
            </div>
          )}

          {/* Row actions toolbar */}
          <div className="revogrid-toolbar">
            <span className="toolbar-info">
              {selectedRowIndex !== null 
                ? `Row ${selectedRowIndex + 1} selected` 
                : 'Click a row to select it'}
            </span>
            <button
              className="download-row-btn"
              onClick={downloadSelectedRow}
              disabled={selectedRowIndex === null}
            >
              ⬇ Download Selected Row
            </button>
          </div>

          <div className="revogrid-container">
            <RevoGrid
              columns={columnsWithDropdowns}
              source={sourceWithRowClass}
              columnTypes={columnTypes}
              theme="compact"
              resize={true}
              autoSizeColumn={true}
              filter={true}
              range={true}
              readonly={false}
              rowClass="_rowClass"
              onBeforecellfocus={(e) => {
                const detail = e.detail;
                if (detail && typeof detail.rowIndex === 'number') {
                  setSelectedRowIndex(detail.rowIndex);
                }
              }}
            />
          </div>

          <div className="data-info">
            <h4>Sheet Info</h4>
            <p>Rows: {activeSheet.source.length}</p>
            <p>Columns: {activeSheet.columns.length}</p>
            {Object.keys(activeSheet.dropdowns).length > 0 && (
              <p>
                Dropdown columns:{' '}
                {Object.keys(activeSheet.dropdowns)
                  .map((prop) => {
                    const col = activeSheet.columns.find((c) => c.prop === prop);
                    return col?.name || prop;
                  })
                  .join(', ')}
              </p>
            )}
          </div>
        </section>
      )}

      <section className="demo-section">
        <h3>Features</h3>
        <div className="feature-grid">
          <div className="feature-card">
            <h4>Virtual Scrolling</h4>
            <p>Handles 100k+ rows efficiently with virtualized rendering</p>
          </div>
          <div className="feature-card">
            <h4>Column Resize</h4>
            <p>Drag column borders to resize, with auto-size support</p>
          </div>
          <div className="feature-card">
            <h4>Filtering</h4>
            <p>Built-in column filtering with multiple filter types</p>
          </div>
          <div className="feature-card">
            <h4>Range Selection</h4>
            <p>Excel-like range selection and copy/paste</p>
          </div>
          <div className="feature-card">
            <h4>Sorting</h4>
            <p>Click column headers to sort ascending/descending</p>
          </div>
          <div className="feature-card">
            <h4>Cell Editing</h4>
            <p>Double-click cells to edit, with custom editors support</p>
          </div>
        </div>
      </section>

      <section className="demo-section">
        <h3>Code Example</h3>
        <pre className="code-block">
{`import { RevoGrid } from '@revolist/react-datagrid';
import type { ColumnRegular } from '@revolist/revogrid';

const columns: ColumnRegular[] = [
  { prop: 'name', name: 'Name', size: 150, sortable: true },
  { prop: 'value', name: 'Value', size: 100, sortable: true }
];

const source = [
  { name: 'Item 1', value: 100 },
  { name: 'Item 2', value: 200 }
];

<RevoGrid
  columns={columns}
  source={source}
  theme="compact"
  resize={true}
  filter={true}
  range={true}
/>`}
        </pre>
      </section>
    </div>
  );
};

export default RevoGridDemo;
