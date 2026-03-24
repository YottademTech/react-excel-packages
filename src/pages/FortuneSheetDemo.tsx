import React, { useState, useCallback, useRef } from 'react';
import { Workbook, WorkbookInstance } from '@fortune-sheet/react';
import '@fortune-sheet/react/dist/index.css';
import ExcelJS from 'exceljs';
import FileUpload from '../components/FileUpload';
import './DemoPage.css';
import './FortuneSheetDemo.css';

interface SheetData {
  name: string;
  celldata: any[];
  row?: number;
  column?: number;
  dataVerification?: { [key: string]: any };
}

const FortuneSheetDemo: React.FC = () => {
  const [sheets, setSheets] = useState<SheetData[]>([]);
  const [fileName, setFileName] = useState('');
  const [showSpreadsheet, setShowSpreadsheet] = useState(false);
  const [showToolbar, setShowToolbar] = useState(false);
  const [showFormulaBar, setShowFormulaBar] = useState(false);
  const workbookRef = useRef<WorkbookInstance>(null);

  const handleFileUpload = useCallback(async (file: File) => {
    setFileName(file.name);
    
    const arrayBuffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);

    const fortuneSheets: SheetData[] = [];

    workbook.eachSheet((worksheet) => {
      const celldata: any[] = [];
      const dataVerification: { [key: string]: any } = {};
      
      let maxRow = 0;
      let maxCol = 0;

      // Process cells
      worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        const r = rowNumber - 1; // Fortune Sheet uses 0-based indexing
        maxRow = Math.max(maxRow, r);
        
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
          const c = colNumber - 1; // Fortune Sheet uses 0-based indexing
          maxCol = Math.max(maxCol, c);
          
          let value = cell.value;
          let displayValue = '';
          
          // Handle different cell value types
          if (value === null || value === undefined) {
            return;
          } else if (typeof value === 'object' && 'result' in value) {
            // Formula result
            value = value.result;
            displayValue = String(value || '');
          } else if (typeof value === 'object' && 'richText' in value) {
            // Rich text
            displayValue = (value as ExcelJS.CellRichTextValue).richText
              .map((rt) => rt.text)
              .join('');
            value = displayValue;
          } else if (value instanceof Date) {
            displayValue = value.toLocaleDateString();
          } else {
            displayValue = String(value);
          }

          const cellValue: any = {
            r,
            c,
            v: {
              v: value,
              m: displayValue,
            },
          };

          // Handle cell types
          if (typeof value === 'number') {
            cellValue.v.ct = { fa: 'General', t: 'n' };
          }

          // Handle data validation for this cell
          if (cell.dataValidation) {
            const dv = cell.dataValidation;
            const key = `${r}_${c}`;
            
            if (dv.type === 'list') {
              let options: string[] = [];
              if (dv.formulae && dv.formulae.length > 0) {
                // Parse the formula - could be like '"Option1,Option2,Option3"'
                const formula = dv.formulae[0];
                options = formula
                  .replace(/^"/, '')
                  .replace(/"$/, '')
                  .split(',')
                  .map((s: string) => s.trim());
              }
              
              dataVerification[key] = {
                type: 'dropdown',
                type2: null,
                value1: options.join(','),
                value2: '',
                checked: false,
                remote: false,
                prohibitInput: false,
                hintShow: false,
                hintText: '',
              };
            } else if (dv.type === 'whole' || dv.type === 'decimal') {
              dataVerification[key] = {
                type: 'number',
                type2: dv.operator || 'between',
                value1: dv.formulae?.[0] || '',
                value2: dv.formulae?.[1] || '',
                checked: false,
                remote: false,
                prohibitInput: !!dv.showErrorMessage,
                hintShow: !!dv.showErrorMessage,
                hintText: dv.error || '',
              };
            }
          }

          celldata.push(cellValue);
        });
      });

      fortuneSheets.push({
        name: worksheet.name,
        celldata,
        row: Math.max(maxRow + 1, 50),
        column: Math.max(maxCol + 1, 26),
        dataVerification: Object.keys(dataVerification).length > 0 ? dataVerification : undefined,
      });
    });

    setSheets(fortuneSheets);
    setShowSpreadsheet(true);
  }, []);

  const createSampleData = () => {
    // Create dataVerification for dropdown cells
    const dataVerification: { [key: string]: any } = {};
    
    // Department dropdown (column 1, rows 1-4)
    for (let r = 1; r <= 4; r++) {
      dataVerification[`${r}_1`] = {
        type: 'dropdown',
        type2: null,
        value1: 'Engineering,Marketing,Sales,HR,Finance,Operations',
        value2: '',
        checked: false,
        remote: false,
        prohibitInput: false,
        hintShow: false,
        hintText: '',
      };
    }
    
    // Status dropdown (column 4, rows 1-4)
    for (let r = 1; r <= 4; r++) {
      dataVerification[`${r}_4`] = {
        type: 'dropdown',
        type2: null,
        value1: 'Active,On Leave,Terminated',
        value2: '',
        checked: false,
        remote: false,
        prohibitInput: false,
        hintShow: false,
        hintText: '',
      };
    }

    const sampleSheets: SheetData[] = [
      {
        name: 'Employee Data',
        celldata: [
          // Header row
          { r: 0, c: 0, v: { v: 'Name', m: 'Name', bg: '#f8fafc', bl: 1 } },
          { r: 0, c: 1, v: { v: 'Department', m: 'Department', bg: '#f8fafc', bl: 1 } },
          { r: 0, c: 2, v: { v: 'Salary', m: 'Salary', bg: '#f8fafc', bl: 1 } },
          { r: 0, c: 3, v: { v: 'Start Date', m: 'Start Date', bg: '#f8fafc', bl: 1 } },
          { r: 0, c: 4, v: { v: 'Status', m: 'Status', bg: '#f8fafc', bl: 1 } },
          // Data rows
          { r: 1, c: 0, v: { v: 'John Doe', m: 'John Doe' } },
          { r: 1, c: 1, v: { v: 'Engineering', m: 'Engineering' } },
          { r: 1, c: 2, v: { v: 75000, m: '$75,000', ct: { fa: '$#,##0', t: 'n' } } },
          { r: 1, c: 3, v: { v: '2020-01-15', m: '2020-01-15' } },
          { r: 1, c: 4, v: { v: 'Active', m: 'Active', fc: '#10b981' } },
          
          { r: 2, c: 0, v: { v: 'Jane Smith', m: 'Jane Smith' } },
          { r: 2, c: 1, v: { v: 'Marketing', m: 'Marketing' } },
          { r: 2, c: 2, v: { v: 65000, m: '$65,000', ct: { fa: '$#,##0', t: 'n' } } },
          { r: 2, c: 3, v: { v: '2021-03-22', m: '2021-03-22' } },
          { r: 2, c: 4, v: { v: 'Active', m: 'Active', fc: '#10b981' } },
          
          { r: 3, c: 0, v: { v: 'Bob Johnson', m: 'Bob Johnson' } },
          { r: 3, c: 1, v: { v: 'Sales', m: 'Sales' } },
          { r: 3, c: 2, v: { v: 85000, m: '$85,000', ct: { fa: '$#,##0', t: 'n' } } },
          { r: 3, c: 3, v: { v: '2018-07-01', m: '2018-07-01' } },
          { r: 3, c: 4, v: { v: 'Active', m: 'Active', fc: '#10b981' } },
          
          { r: 4, c: 0, v: { v: 'Alice Brown', m: 'Alice Brown' } },
          { r: 4, c: 1, v: { v: 'HR', m: 'HR' } },
          { r: 4, c: 2, v: { v: 70000, m: '$70,000', ct: { fa: '$#,##0', t: 'n' } } },
          { r: 4, c: 3, v: { v: '2019-11-10', m: '2019-11-10' } },
          { r: 4, c: 4, v: { v: 'On Leave', m: 'On Leave', fc: '#f59e0b' } },
          
          // Formula example
          { r: 6, c: 0, v: { v: 'Total Salary:', m: 'Total Salary:', bl: 1 } },
          { r: 6, c: 2, v: { v: 295000, m: '$295,000', f: '=SUM(C2:C5)', ct: { fa: '$#,##0', t: 'n' } } },
        ],
        row: 50,
        column: 26,
        dataVerification,
      },
    ];

    setSheets(sampleSheets);
    setFileName('sample-data.xlsx');
    setShowSpreadsheet(true);
  };

  return (
    <div className="demo-page">
      <header className="demo-header">
        <h1>Fortune Sheet</h1>
        <span className="license-tag">MIT</span>
      </header>

      <div className="demo-info">
        <p>
          Fortune Sheet is a full-featured Excel-like spreadsheet component for React. 
          It's a TypeScript fork of Luckysheet with modern React support and extensive Excel functionality.
        </p>
        <div className="features-list">
          <span className="feature">✓ Full Excel UI</span>
          <span className="feature">✓ Formula Support</span>
          <span className="feature">✓ Cell Formatting</span>
          <span className="feature">✓ Merge Cells</span>
          <span className="feature">✓ Freeze Panes</span>
        </div>
      </div>

      {!showSpreadsheet && (
        <div className="upload-section">
          <FileUpload onFileUpload={handleFileUpload} />
          <div className="sample-data-option">
            <p>or</p>
            <button className="sample-data-btn" onClick={createSampleData}>
              Load Sample Data
            </button>
          </div>
        </div>
      )}

      {showSpreadsheet && sheets.length > 0 && (
        <div className="result-section fortune-result">
          <div className="result-header">
            <h2>
              <span className="file-icon">📊</span>
              {fileName || 'Fortune Sheet'}
              <span className="editable-badge">Full Excel UI</span>
            </h2>
            <button 
              className="reset-btn"
              onClick={() => {
                setShowSpreadsheet(false);
                setSheets([]);
                setFileName('');
              }}
            >
              Upload New File
            </button>
          </div>

          <div className="toolbar-toggles">
            <label className="toggle-label">
              <input
                type="checkbox"
                checked={showToolbar}
                onChange={(e) => setShowToolbar(e.target.checked)}
              />
              Show Toolbar
            </label>
            <label className="toggle-label">
              <input
                type="checkbox"
                checked={showFormulaBar}
                onChange={(e) => setShowFormulaBar(e.target.checked)}
              />
              Show Formula Bar
            </label>
          </div>

          <div className="fortune-sheet-wrapper">
            <Workbook
              ref={workbookRef}
              data={sheets}
              showToolbar={showToolbar}
              showFormulaBar={showFormulaBar}
              onChange={(data) => {
                console.log('Sheet data changed:', data);
              }}
            />
          </div>

          <div className="data-info">
            <h4>Fortune Sheet Configuration</h4>
            <p>• <code>showToolbar={'{false}'}</code> - Hide the formatting toolbar</p>
            <p>• <code>showFormulaBar={'{false}'}</code> - Hide the formula bar</p>
            <p>• Right-click context menus still available</p>
            <p>• Sheet tabs at bottom</p>
            <p>• Supports 400+ Excel formulas</p>
          </div>
        </div>
      )}

      <div className="code-example">
        <div className="code-header">Usage Example (Clean UI)</div>
        <pre>
          <code>{`import { Workbook } from '@fortune-sheet/react';
import '@fortune-sheet/react/dist/index.css';

const data = [
  {
    name: 'Sheet1',
    celldata: [
      { r: 0, c: 0, v: { v: 'Hello', m: 'Hello' } },
      { r: 0, c: 1, v: { v: 'World', m: 'World' } },
    ],
  },
];

function MyWorkbook() {
  return (
    <div style={{ height: 600 }}>
      <Workbook 
        data={data}
        showToolbar={false}      // Hide toolbar
        showFormulaBar={false}   // Hide formula bar
      />
    </div>
  );
}`}</code>
        </pre>
      </div>
    </div>
  );
};

export default FortuneSheetDemo;
