import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import FileUpload from '../components/FileUpload';
import './DemoPage.css';

interface CellInfo {
  value: any;
  type: string;
  format?: string;
  validation?: {
    type: string;
    options?: string[];
  };
}

interface SheetData {
  name: string;
  data: any[][];
  cellInfo: { [key: string]: CellInfo };
  merges?: XLSX.Range[];
}

const XlsxDemo: React.FC = () => {
  const [sheets, setSheets] = useState<SheetData[]>([]);
  const [activeSheet, setActiveSheet] = useState(0);
  const [fileName, setFileName] = useState('');

  const getCellType = (cell: XLSX.CellObject | undefined): string => {
    if (!cell) return 'empty';
    switch (cell.t) {
      case 's': return 'string';
      case 'n': return 'number';
      case 'b': return 'boolean';
      case 'd': return 'date';
      case 'e': return 'error';
      default: return 'unknown';
    }
  };

  const parseDataValidation = (ws: XLSX.WorkSheet): { [key: string]: CellInfo['validation'] } => {
    const validations: { [key: string]: CellInfo['validation'] } = {};
    
    // Check for data validations in the worksheet
    if ((ws as any)['!dataValidation']) {
      const dvs = (ws as any)['!dataValidation'];
      dvs.forEach((dv: any) => {
        if (dv.type === 'list' && dv.sqref) {
          const options = dv.formula1?.split(',').map((s: string) => s.trim().replace(/"/g, '')) || [];
          validations[dv.sqref] = {
            type: 'dropdown',
            options,
          };
        }
      });
    }
    
    return validations;
  };

  const handleFileUpload = useCallback((file: File) => {
    setFileName(file.name);
    const reader = new FileReader();
    
    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      const workbook = XLSX.read(data, { 
        type: 'array',
        cellDates: true,
        cellStyles: true,
        cellNF: true,
      });

      const parsedSheets: SheetData[] = workbook.SheetNames.map((sheetName) => {
        const ws = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, dateNF: 'yyyy-mm-dd' }) as any[][];
        
        // Build cell info map
        const cellInfo: { [key: string]: CellInfo } = {};
        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        
        for (let r = range.s.r; r <= range.e.r; r++) {
          for (let c = range.s.c; c <= range.e.c; c++) {
            const cellAddress = XLSX.utils.encode_cell({ r, c });
            const cell = ws[cellAddress] as XLSX.CellObject | undefined;
            
            if (cell) {
              cellInfo[cellAddress] = {
                value: cell.v,
                type: getCellType(cell),
                format: typeof cell.z === 'string' ? cell.z : undefined,
              };
            }
          }
        }

        // Parse data validations
        const validations = parseDataValidation(ws);
        Object.entries(validations).forEach(([ref, validation]) => {
          if (cellInfo[ref]) {
            cellInfo[ref].validation = validation;
          }
        });

        return {
          name: sheetName,
          data: jsonData,
          cellInfo,
          merges: ws['!merges'],
        };
      });

      setSheets(parsedSheets);
      setActiveSheet(0);
    };

    reader.readAsArrayBuffer(file);
  }, []);

  const formatCellValue = (value: any, rowIdx: number, colIdx: number): string => {
    if (value === null || value === undefined) return '';
    
    // Handle Date objects
    if (value instanceof Date) {
      return value.toLocaleDateString();
    }
    
    return String(value);
  };

  const getCellClassName = (rowIdx: number, colIdx: number): string => {
    const cellAddress = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
    const info = sheets[activeSheet]?.cellInfo[cellAddress];
    
    const classes = ['cell'];
    if (info) {
      classes.push(`cell-type-${info.type}`);
      if (info.validation?.type === 'dropdown') {
        classes.push('cell-dropdown');
      }
    }
    
    return classes.join(' ');
  };

  const renderCell = (value: any, rowIdx: number, colIdx: number) => {
    const cellAddress = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
    const info = sheets[activeSheet]?.cellInfo[cellAddress];
    
    if (info?.validation?.type === 'dropdown') {
      return (
        <div className="dropdown-cell">
          <span>{formatCellValue(value, rowIdx, colIdx)}</span>
          <span className="dropdown-indicator">▼</span>
          {info.validation.options && (
            <div className="dropdown-options">
              {info.validation.options.map((opt, i) => (
                <div key={i} className="dropdown-option">{opt}</div>
              ))}
            </div>
          )}
        </div>
      );
    }

    return formatCellValue(value, rowIdx, colIdx);
  };

  return (
    <div className="demo-page">
      <header className="demo-header">
        <h1>XLSX (SheetJS)</h1>
        <span className="license-tag">Apache 2.0</span>
      </header>

      <div className="demo-info">
        <p>
          SheetJS is the most popular spreadsheet parser and writer for JavaScript. 
          It supports reading and writing various formats including XLSX, XLS, CSV, and more.
        </p>
        <div className="features-list">
          <span className="feature">✓ Read/Write Excel</span>
          <span className="feature">✓ Formulas</span>
          <span className="feature">✓ Data Validation</span>
          <span className="feature">✓ Date Handling</span>
          <span className="feature">✓ Multiple Sheets</span>
        </div>
      </div>

      <FileUpload onFileUpload={handleFileUpload} />

      {sheets.length > 0 && (
        <div className="result-section">
          <div className="result-header">
            <h2>
              <span className="file-icon">📄</span>
              {fileName}
            </h2>
            <div className="sheet-tabs">
              {sheets.map((sheet, idx) => (
                <button
                  key={sheet.name}
                  className={`sheet-tab ${idx === activeSheet ? 'active' : ''}`}
                  onClick={() => setActiveSheet(idx)}
                >
                  {sheet.name}
                </button>
              ))}
            </div>
          </div>

          <div className="spreadsheet-container">
            <table className="spreadsheet">
              <thead>
                <tr>
                  <th className="row-header"></th>
                  {sheets[activeSheet]?.data[0]?.map((_, colIdx) => (
                    <th key={colIdx} className="col-header">
                      {XLSX.utils.encode_col(colIdx)}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {sheets[activeSheet]?.data.map((row, rowIdx) => (
                  <tr key={rowIdx}>
                    <td className="row-header">{rowIdx + 1}</td>
                    {row.map((cell, colIdx) => (
                      <td
                        key={colIdx}
                        className={getCellClassName(rowIdx, colIdx)}
                      >
                        {renderCell(cell, rowIdx, colIdx)}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="cell-legend">
            <h4>Cell Types Detected:</h4>
            <div className="legend-items">
              <span className="legend-item"><span className="legend-color type-string"></span> String</span>
              <span className="legend-item"><span className="legend-color type-number"></span> Number</span>
              <span className="legend-item"><span className="legend-color type-date"></span> Date</span>
              <span className="legend-item"><span className="legend-color type-dropdown"></span> Dropdown</span>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default XlsxDemo;
