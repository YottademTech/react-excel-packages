import React, { useState, useCallback, useMemo } from 'react';
import Spreadsheet, { CellBase, Matrix } from 'react-spreadsheet';
import * as XLSX from 'xlsx';
import FileUpload from '../components/FileUpload';
import './DemoPage.css';
import './ReactSpreadsheetDemo.css';

interface CustomCell extends CellBase {
  value: string | number | boolean | null;
  className?: string;
}

const ReactSpreadsheetDemo: React.FC = () => {
  const [data, setData] = useState<Matrix<CustomCell>>([]);
  const [fileName, setFileName] = useState('');
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [activeSheet, setActiveSheet] = useState(0);
  const [allSheetsData, setAllSheetsData] = useState<Matrix<CustomCell>[]>([]);

  const columnLabels = useMemo(() => {
    if (data.length === 0) return [];
    const maxCols = Math.max(...data.map(row => row.length));
    return Array.from({ length: maxCols }, (_, i) => {
      let result = '';
      let n = i;
      while (n >= 0) {
        result = String.fromCharCode((n % 26) + 65) + result;
        n = Math.floor(n / 26) - 1;
      }
      return result;
    });
  }, [data]);

  const rowLabels = useMemo(() => {
    return data.map((_, i) => String(i + 1));
  }, [data]);

  const handleFileUpload = useCallback((file: File) => {
    setFileName(file.name);
    const reader = new FileReader();

    reader.onload = (e) => {
      const arrayBuffer = e.target?.result as ArrayBuffer;
      const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });

      const sheets: Matrix<CustomCell>[] = [];
      const names: string[] = [];

      workbook.SheetNames.forEach((sheetName) => {
        names.push(sheetName);
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, {
          header: 1,
          raw: false,
          dateNF: 'yyyy-mm-dd',
        }) as any[][];

        const spreadsheetData: Matrix<CustomCell> = jsonData.map((row) =>
          row.map((cell) => ({
            value: cell ?? '',
            className: typeof cell === 'number' ? 'cell-number' : undefined,
          }))
        );

        sheets.push(spreadsheetData);
      });

      setSheetNames(names);
      setAllSheetsData(sheets);
      setData(sheets[0] || []);
      setActiveSheet(0);
    };

    reader.readAsArrayBuffer(file);
  }, []);

  const handleSheetChange = (index: number) => {
    setActiveSheet(index);
    setData(allSheetsData[index] || []);
  };

  const handleDataChange = (newData: Matrix<CustomCell>) => {
    setData(newData);
    // Update the stored data for this sheet
    const updatedAllSheets = [...allSheetsData];
    updatedAllSheets[activeSheet] = newData;
    setAllSheetsData(updatedAllSheets);
  };

  // Create sample data for demo
  const createSampleData = () => {
    const sampleData: Matrix<CustomCell> = [
      [{ value: 'Name' }, { value: 'Age' }, { value: 'Department' }, { value: 'Start Date' }, { value: 'Salary' }],
      [{ value: 'John Doe' }, { value: 32 }, { value: 'Engineering' }, { value: '2020-01-15' }, { value: 75000 }],
      [{ value: 'Jane Smith' }, { value: 28 }, { value: 'Marketing' }, { value: '2021-03-22' }, { value: 65000 }],
      [{ value: 'Bob Johnson' }, { value: 45 }, { value: 'Sales' }, { value: '2018-07-01' }, { value: 85000 }],
      [{ value: 'Alice Brown' }, { value: 35 }, { value: 'HR' }, { value: '2019-11-10' }, { value: 70000 }],
      [{ value: 'Charlie Wilson' }, { value: 29 }, { value: 'Engineering' }, { value: '2022-02-28' }, { value: 72000 }],
    ];
    
    setData(sampleData);
    setFileName('sample-data.xlsx');
    setSheetNames(['Sample Data']);
    setAllSheetsData([sampleData]);
    setActiveSheet(0);
  };

  return (
    <div className="demo-page">
      <header className="demo-header">
        <h1>React Spreadsheet</h1>
        <span className="license-tag">MIT</span>
      </header>

      <div className="demo-info">
        <p>
          React Spreadsheet is a lightweight, customizable spreadsheet component for React.
          It provides basic Excel-like functionality with a clean API and is great for simple data editing.
        </p>
        <div className="features-list">
          <span className="feature">✓ Editable Cells</span>
          <span className="feature">✓ Copy/Paste</span>
          <span className="feature">✓ Keyboard Navigation</span>
          <span className="feature">✓ Custom Cell Renderers</span>
          <span className="feature">✓ Lightweight</span>
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

      {data.length > 0 && (
        <div className="result-section">
          <div className="result-header">
            <h2>
              <span className="file-icon">📊</span>
              {fileName || 'Spreadsheet'}
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

          <div className="react-spreadsheet-wrapper">
            <Spreadsheet
              data={data}
              onChange={handleDataChange}
              columnLabels={columnLabels}
              rowLabels={rowLabels}
            />
          </div>

          <div className="data-info">
            <h4>Interactive Features</h4>
            <p>• Click any cell to edit its value</p>
            <p>• Use arrow keys to navigate</p>
            <p>• Select multiple cells and copy/paste</p>
            <p>• Tab to move to next cell</p>
          </div>
        </div>
      )}

      <div className="code-example">
        <div className="code-header">Usage Example</div>
        <pre>
          <code>{`import Spreadsheet from 'react-spreadsheet';

const data = [
  [{ value: 'Name' }, { value: 'Age' }],
  [{ value: 'John' }, { value: 30 }],
  [{ value: 'Jane' }, { value: 25 }],
];

function MySpreadsheet() {
  const [data, setData] = useState(initialData);
  
  return (
    <Spreadsheet
      data={data}
      onChange={setData}
      columnLabels={['A', 'B']}
    />
  );
}`}</code>
        </pre>
      </div>
    </div>
  );
};

export default ReactSpreadsheetDemo;
