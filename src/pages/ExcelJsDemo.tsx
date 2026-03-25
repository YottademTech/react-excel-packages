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
    resolvedOptions?: string[]; // Resolved dropdown options from sheet references
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

  // Helper function to parse cell reference (e.g., "$A$1" or "A1") to column and row
  const parseCellRef = (ref: string): { col: number; row: number } | null => {
    const match = ref.match(/^\$?([A-Z]+)\$?(\d+)$/i);
    if (!match) return null;
    
    const colStr = match[1].toUpperCase();
    const row = parseInt(match[2], 10);
    
    // Convert column letters to number (A=1, B=2, ..., Z=26, AA=27, etc.)
    let col = 0;
    for (let i = 0; i < colStr.length; i++) {
      col = col * 26 + (colStr.charCodeAt(i) - 64);
    }
    
    return { col, row };
  };

  // Helper function to resolve sheet reference formula to actual values
  const resolveSheetReference = (
    formula: string,
    workbook: ExcelJS.Workbook
  ): string[] => {
    // Remove leading = if present
    formula = formula.replace(/^=/, '');
    
    // Check for sheet reference pattern: 'SheetName'!$A$1:$A$10 or SheetName!A1:A10
    const sheetRefMatch = formula.match(/^'?([^'!]+)'?!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/i);
    
    if (!sheetRefMatch) {
      return [];
    }
    
    const [, sheetName, startCol, startRow, endCol, endRow] = sheetRefMatch;
    
    // Find the referenced worksheet
    const refSheet = workbook.getWorksheet(sheetName);
    if (!refSheet) {
      console.warn(`Referenced sheet "${sheetName}" not found`);
      return [];
    }
    
    const values: string[] = [];
    const startRowNum = parseInt(startRow, 10);
    const endRowNum = parseInt(endRow, 10);
    const startColRef = parseCellRef(`${startCol}1`);
    const endColRef = parseCellRef(`${endCol}1`);
    
    if (!startColRef || !endColRef) return [];
    
    // Extract values from the referenced range
    for (let row = startRowNum; row <= endRowNum; row++) {
      for (let col = startColRef.col; col <= endColRef.col; col++) {
        const cell = refSheet.getCell(row, col);
        let cellValue = cell.value;
        
        if (cellValue === null || cellValue === undefined) continue;
        
        // Handle different value types
        if (typeof cellValue === 'object' && 'result' in cellValue) {
          cellValue = cellValue.result;
        } else if (typeof cellValue === 'object' && 'richText' in cellValue) {
          cellValue = (cellValue as ExcelJS.CellRichTextValue).richText
            .map((rt) => rt.text)
            .join('');
        }
        
        const strValue = String(cellValue).trim();
        if (strValue) {
          values.push(strValue);
        }
      }
    }
    
    return values;
  };

  // Helper function to extract sheet name from a formula reference
  const extractSheetNameFromFormula = (formula: string): string | null => {
    formula = formula.replace(/^=/, '');
    const match = formula.match(/^'?([^'!]+)'?!/i);
    return match ? match[1] : null;
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

    // First pass: identify sheets that are used as lookup sources for data validation
    const lookupSheetNames = new Set<string>();
    
    workbook.eachSheet((worksheet) => {
      worksheet.eachRow({ includeEmpty: false }, (row) => {
        row.eachCell({ includeEmpty: false }, (cell) => {
          if (cell.dataValidation?.type === 'list' && cell.dataValidation.formulae) {
            const formula = cell.dataValidation.formulae[0];
            if (formula && formula.includes('!')) {
              const sheetName = extractSheetNameFromFormula(formula);
              if (sheetName) {
                lookupSheetNames.add(sheetName);
              }
            }
          }
        });
      });
    });

    console.log('Lookup sheets to hide:', Array.from(lookupSheetNames));

    const parsedSheets: SheetData[] = [];

    workbook.eachSheet((worksheet) => {
      // Skip sheets that are only used as lookup sources
      if (lookupSheetNames.has(worksheet.name)) {
        console.log(`Skipping lookup sheet: "${worksheet.name}"`);
        return;
      }

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
            const dv = cell.dataValidation;
            cellData.validation = {
              type: dv.type || 'unknown',
              formulae: dv.formulae,
            };
            
            // Resolve dropdown options from sheet references
            if (dv.type === 'list' && dv.formulae && dv.formulae.length > 0) {
              const formula = dv.formulae[0];
              
              if (formula.includes('!')) {
                // This is a sheet reference - resolve it
                const options = resolveSheetReference(formula, workbook);
                if (options.length > 0) {
                  cellData.validation.resolvedOptions = options;
                  console.log(`Resolved dropdown from "${formula}":`, options);
                }
              } else {
                // Inline list like "Option1,Option2,Option3"
                const options = formula
                  .replace(/^"/, '')
                  .replace(/"$/, '')
                  .split(',')
                  .map((s: string) => s.trim())
                  .filter((s: string) => s.length > 0);
                if (options.length > 0) {
                  cellData.validation.resolvedOptions = options;
                }
              }
            }
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
                            {cell.validation.resolvedOptions && cell.validation.resolvedOptions.length > 0 ? (
                              <div className="dropdown-options">
                                {cell.validation.resolvedOptions.map((opt, i) => (
                                  <div key={i} className="dropdown-option">
                                    {opt}
                                  </div>
                                ))}
                              </div>
                            ) : cell.validation.formulae && (
                              <div className="dropdown-options">
                                <div className="dropdown-option dropdown-formula">
                                  Source: {cell.validation.formulae[0]}
                                </div>
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
