import React, { useState, useCallback } from 'react';
import ExcelJS from 'exceljs';
import FileUpload from '../components/FileUpload';
import './DemoPage.css';

interface CellData {
  value: any;
  style: {
    fill?: string;
    fontColor?: string;
    bold?: boolean;
    italic?: boolean;
    alignment?: string;
  };
  type: string;
  validation?: {
    type: string;
    formulae?: string[];
  };
}

interface SheetData {
  name: string;
  data: CellData[][];
  columnWidths: number[];
}

const ExcelJsDemo: React.FC = () => {
  const [sheets, setSheets] = useState<SheetData[]>([]);
  const [activeSheet, setActiveSheet] = useState(0);
  const [fileName, setFileName] = useState('');
  const [metadata, setMetadata] = useState<any>(null);

  const getCellType = (cell: ExcelJS.Cell): string => {
    if (cell.value === null || cell.value === undefined) return 'empty';
    if (cell.type === ExcelJS.ValueType.Number) return 'number';
    if (cell.type === ExcelJS.ValueType.Date) return 'date';
    if (cell.type === ExcelJS.ValueType.Boolean) return 'boolean';
    if (cell.type === ExcelJS.ValueType.Formula) return 'formula';
    if (cell.type === ExcelJS.ValueType.Hyperlink) return 'hyperlink';
    if (cell.type === ExcelJS.ValueType.RichText) return 'richtext';
    return 'string';
  };

  const extractCellValue = (cell: ExcelJS.Cell): any => {
    if (cell.value === null || cell.value === undefined) return '';
    
    // Handle formula results
    if (typeof cell.value === 'object' && 'result' in cell.value) {
      return cell.value.result;
    }
    
    // Handle rich text
    if (typeof cell.value === 'object' && 'richText' in cell.value) {
      return (cell.value as ExcelJS.CellRichTextValue).richText
        .map((rt) => rt.text)
        .join('');
    }
    
    // Handle hyperlinks
    if (typeof cell.value === 'object' && 'hyperlink' in cell.value) {
      return (cell.value as ExcelJS.CellHyperlinkValue).text || cell.value.hyperlink;
    }
    
    // Handle dates
    if (cell.value instanceof Date) {
      return cell.value.toLocaleDateString();
    }
    
    return cell.value;
  };

  const extractCellStyle = (cell: ExcelJS.Cell): CellData['style'] => {
    const style: CellData['style'] = {};
    
    if (cell.fill && cell.fill.type === 'pattern' && cell.fill.fgColor) {
      const color = cell.fill.fgColor;
      if (color.argb) {
        style.fill = `#${color.argb.substring(2)}`;
      }
    }
    
    if (cell.font) {
      if (cell.font.color?.argb) {
        style.fontColor = `#${cell.font.color.argb.substring(2)}`;
      }
      style.bold = cell.font.bold;
      style.italic = cell.font.italic;
    }
    
    if (cell.alignment?.horizontal) {
      style.alignment = cell.alignment.horizontal;
    }
    
    return style;
  };

  const handleFileUpload = useCallback(async (file: File) => {
    setFileName(file.name);
    
    const arrayBuffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);

    // Extract metadata
    setMetadata({
      creator: workbook.creator,
      lastModifiedBy: workbook.lastModifiedBy,
      created: workbook.created?.toLocaleDateString(),
      modified: workbook.modified?.toLocaleDateString(),
    });

    const parsedSheets: SheetData[] = [];

    workbook.eachSheet((worksheet) => {
      const data: CellData[][] = [];
      const columnWidths: number[] = [];
      
      // Get column widths
      worksheet.columns.forEach((col, idx) => {
        columnWidths[idx] = col.width || 10;
      });

      worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        const rowData: CellData[] = [];
        
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const cellData: CellData = {
            value: extractCellValue(cell),
            style: extractCellStyle(cell),
            type: getCellType(cell),
          };
          
          // Check for data validation
          if (cell.dataValidation) {
            cellData.validation = {
              type: cell.dataValidation.type || 'unknown',
              formulae: cell.dataValidation.formulae,
            };
          }
          
          // Ensure array is large enough
          while (rowData.length < colNumber - 1) {
            rowData.push({ value: '', style: {}, type: 'empty' });
          }
          rowData[colNumber - 1] = cellData;
        });
        
        data[rowNumber - 1] = rowData;
      });

      parsedSheets.push({
        name: worksheet.name,
        data,
        columnWidths,
      });
    });

    setSheets(parsedSheets);
    setActiveSheet(0);
  }, []);

  const getCellStyle = (cell: CellData): React.CSSProperties => {
    const style: React.CSSProperties = {};
    
    if (cell.style.fill) {
      style.backgroundColor = cell.style.fill;
    }
    if (cell.style.fontColor) {
      style.color = cell.style.fontColor;
    }
    if (cell.style.bold) {
      style.fontWeight = 'bold';
    }
    if (cell.style.italic) {
      style.fontStyle = 'italic';
    }
    if (cell.style.alignment) {
      style.textAlign = cell.style.alignment as any;
    }
    
    return style;
  };

  const getCellClassName = (cell: CellData): string => {
    const classes = ['cell', `cell-type-${cell.type}`];
    if (cell.validation) {
      classes.push('cell-dropdown');
    }
    return classes.join(' ');
  };

  const getColumnLetter = (index: number): string => {
    let result = '';
    let n = index;
    while (n >= 0) {
      result = String.fromCharCode((n % 26) + 65) + result;
      n = Math.floor(n / 26) - 1;
    }
    return result;
  };

  return (
    <div className="demo-page">
      <header className="demo-header">
        <h1>ExcelJS</h1>
        <span className="license-tag">MIT</span>
      </header>

      <div className="demo-info">
        <p>
          ExcelJS is a powerful library for reading, manipulating, and writing spreadsheet data.
          It excels at preserving cell styles, data validation, and supports streaming for large files.
        </p>
        <div className="features-list">
          <span className="feature">✓ Rich Cell Styling</span>
          <span className="feature">✓ Data Validation</span>
          <span className="feature">✓ Conditional Formatting</span>
          <span className="feature">✓ Images Support</span>
          <span className="feature">✓ Streaming API</span>
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

          {metadata && (
            <div className="data-info">
              <h4>Document Metadata (ExcelJS Feature)</h4>
              <p><strong>Creator:</strong> {metadata.creator || 'N/A'}</p>
              <p><strong>Last Modified By:</strong> {metadata.lastModifiedBy || 'N/A'}</p>
              <p><strong>Created:</strong> {metadata.created || 'N/A'}</p>
              <p><strong>Modified:</strong> {metadata.modified || 'N/A'}</p>
            </div>
          )}

          <div className="spreadsheet-container">
            <table className="spreadsheet">
              <thead>
                <tr>
                  <th className="row-header"></th>
                  {sheets[activeSheet]?.data[0]?.map((_, colIdx) => (
                    <th key={colIdx} className="col-header">
                      {getColumnLetter(colIdx)}
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
                        className={getCellClassName(cell)}
                        style={getCellStyle(cell)}
                      >
                        {cell.validation ? (
                          <div className="dropdown-cell">
                            <span>{cell.value}</span>
                            <span className="dropdown-indicator">▼</span>
                            {cell.validation.formulae && (
                              <div className="dropdown-options">
                                {cell.validation.formulae[0]
                                  ?.replace(/"/g, '')
                                  .split(',')
                                  .map((opt, i) => (
                                    <div key={i} className="dropdown-option">
                                      {opt.trim()}
                                    </div>
                                  ))}
                              </div>
                            )}
                          </div>
                        ) : (
                          cell.value
                        )}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="cell-legend">
            <h4>Cell Types & Styles Preserved:</h4>
            <div className="legend-items">
              <span className="legend-item"><span className="legend-color type-number"></span> Number</span>
              <span className="legend-item"><span className="legend-color type-date"></span> Date</span>
              <span className="legend-item"><span className="legend-color type-dropdown"></span> Dropdown</span>
              <span className="legend-item">🎨 Background Colors</span>
              <span className="legend-item"><strong>B</strong> Bold</span>
              <span className="legend-item"><em>I</em> Italic</span>
            </div>
          </div>
        </div>
      )}

      <div className="code-example">
        <div className="code-header">Usage Example</div>
        <pre>
          <code>{`import ExcelJS from 'exceljs';

const workbook = new ExcelJS.Workbook();
await workbook.xlsx.load(buffer);

workbook.eachSheet((worksheet) => {
  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      console.log(cell.value);
      console.log(cell.style);
      console.log(cell.dataValidation);
    });
  });
});`}</code>
        </pre>
      </div>
    </div>
  );
};

export default ExcelJsDemo;
