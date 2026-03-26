import React, { useState, useCallback, useRef, useEffect } from 'react';
import jspreadsheet from 'jspreadsheet-ce';
import 'jspreadsheet-ce/dist/jspreadsheet.css';
import 'jsuites/dist/jsuites.css';
import ExcelJS from 'exceljs';
import FileUpload from '../components/FileUpload';
import './DemoPage.css';
import './JSpreadsheetDemo.css';

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

interface ColumnConfig {
  type?: string;
  title: string;
  width?: number;
  source?: string[];
}

interface SheetData {
  name: string;
  data: any[][];
  columns: ColumnConfig[];
}

// ——————————————————————————————————————————————————————
// Component
// ——————————————————————————————————————————————————————

const JSpreadsheetDemo: React.FC = () => {
  const [sheets, setSheets] = useState<SheetData[]>([]);
  const [activeSheetIndex, setActiveSheetIndex] = useState(0);
  const [showSpreadsheet, setShowSpreadsheet] = useState(false);
  const [fileName, setFileName] = useState<string>('');
  const [error, setError] = useState<string | null>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const jspreadsheetRef = useRef<any>(null);

  const handleFileUpload = useCallback(async (file: File) => {
    setError(null);
    setFileName(file.name);

    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(arrayBuffer);

      const parsedSheets: SheetData[] = [];

      workbook.eachSheet((worksheet) => {
        // Skip hidden sheets
        if (worksheet.state === 'hidden' || worksheet.state === 'veryHidden') {
          return;
        }

        const rowCount = worksheet.rowCount;
        const colCount = worksheet.columnCount;
        if (rowCount === 0 || colCount === 0) return;

        // Build columns from header row (row 1)
        const columns: ColumnConfig[] = [];
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
            title: headerText,
            width: Math.max(headerText.length * 10 + 20, 100),
          });
        }

        // Build data from rows 2+
        const data: any[][] = [];
        for (let r = 2; r <= rowCount; r++) {
          const row = worksheet.getRow(r);
          const rowData: any[] = [];
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
            rowData.push(cellValue ?? '');
          }
          data.push(rowData);
        }

        // Parse data validations for dropdowns
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
                  if (columns[col]) {
                    columns[col].type = 'dropdown';
                    columns[col].source = options;
                  }
                }
              }
            }
          }
        }

        parsedSheets.push({
          name: worksheet.name,
          data,
          columns,
        });
      });

      if (parsedSheets.length === 0) {
        setError('No visible sheets found in the workbook');
        return;
      }

      setSheets(parsedSheets);
      setActiveSheetIndex(0);
      setShowSpreadsheet(true);
    } catch (err) {
      setError(`Error parsing Excel file: ${err}`);
      console.error(err);
    }
  }, []);

  // Initialize/update jspreadsheet when sheet changes
  useEffect(() => {
    if (!showSpreadsheet || sheets.length === 0 || !containerRef.current) return;

    const activeSheet = sheets[activeSheetIndex];
    if (!activeSheet) return;

    // Destroy existing instance
    if (jspreadsheetRef.current) {
      jspreadsheetRef.current.destroy();
      jspreadsheetRef.current = null;
    }

    // Clear container
    containerRef.current.innerHTML = '';

    // Convert columns to jspreadsheet format
    const jssColumns = activeSheet.columns.map((col) => ({
      title: col.title,
      width: col.width || 120,
      type: col.type || 'text',
      source: col.source || [],
      autocomplete: col.type === 'dropdown',
    }));

    // Create new jspreadsheet instance
    jspreadsheetRef.current = jspreadsheet(containerRef.current, {
      worksheets: [{
        data: activeSheet.data.length > 0 ? activeSheet.data : [[]],
        columns: jssColumns,
        minDimensions: [activeSheet.columns.length, Math.max(activeSheet.data.length, 10)],
        tableOverflow: true,
        tableWidth: '100%',
        tableHeight: '450px',
        search: true,
      }],
    } as any);

    return () => {
      if (jspreadsheetRef.current) {
        jspreadsheetRef.current.destroy();
        jspreadsheetRef.current = null;
      }
    };
  }, [showSpreadsheet, sheets, activeSheetIndex]);

  const activeSheet = sheets[activeSheetIndex];

  return (
    <div className="demo-page">
      <header className="demo-header">
        <h2>jspreadsheet CE Demo</h2>
        <p>Lightweight vanilla JavaScript spreadsheet with Excel-like features</p>
        <div className="package-info">
          <span className="package-name">jspreadsheet-ce</span>
          <span className="package-license">MIT License</span>
        </div>
      </header>

      <section className="demo-section">
        <h3>Upload Excel File</h3>
        <FileUpload onFileUpload={handleFileUpload} accept=".xlsx,.xls" />
        {fileName && <p className="file-name">Loaded: {fileName}</p>}
        {error && <p className="error-message">{error}</p>}
      </section>

      {showSpreadsheet && activeSheet && (
        <section className="demo-section spreadsheet-section">
          <h3>jspreadsheet View</h3>

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

          <div className="jspreadsheet-container" ref={containerRef}></div>

          <div className="data-info">
            <h4>Sheet Info</h4>
            <p>Rows: {activeSheet.data.length}</p>
            <p>Columns: {activeSheet.columns.length}</p>
            {activeSheet.columns.some((c) => c.type === 'dropdown') && (
              <p>
                Dropdown columns:{' '}
                {activeSheet.columns
                  .filter((c) => c.type === 'dropdown')
                  .map((c) => c.title)
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
            <h4>Lightweight</h4>
            <p>Small bundle size, pure JavaScript with no dependencies</p>
          </div>
          <div className="feature-card">
            <h4>Column Types</h4>
            <p>Text, dropdown, checkbox, calendar, color picker, and more</p>
          </div>
          <div className="feature-card">
            <h4>Pagination</h4>
            <p>Built-in pagination for large datasets</p>
          </div>
          <div className="feature-card">
            <h4>Search</h4>
            <p>Built-in search functionality across all cells</p>
          </div>
          <div className="feature-card">
            <h4>Sorting</h4>
            <p>Click column headers to sort data</p>
          </div>
          <div className="feature-card">
            <h4>Copy/Paste</h4>
            <p>Excel-compatible copy and paste support</p>
          </div>
        </div>
      </section>

      <section className="demo-section">
        <h3>Code Example</h3>
        <pre className="code-block">
{`import jspreadsheet from 'jspreadsheet-ce';
import 'jspreadsheet-ce/dist/jspreadsheet.css';

const container = document.getElementById('spreadsheet');

jspreadsheet(container, {
  data: [
    ['Item 1', 100, true],
    ['Item 2', 200, false],
  ],
  columns: [
    { title: 'Name', width: 150 },
    { title: 'Value', width: 100 },
    { title: 'Active', type: 'checkbox', width: 80 }
  ],
  search: true,
  pagination: 50,
  columnSorting: true,
});`}
        </pre>
      </section>
    </div>
  );
};

export default JSpreadsheetDemo;
