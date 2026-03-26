import React, { useState, useCallback, useMemo, useRef, useEffect } from 'react';
import Spreadsheet, {
  CellBase,
  Matrix,
  DataEditorProps,
  DataViewerProps,
  Selection,
  RangeSelection,
  EntireRowsSelection,
} from 'react-spreadsheet';
import ExcelJS from 'exceljs';
import FileUpload from '../components/FileUpload';
import './DemoPage.css';
import './ReactSpreadsheetDemo.css';

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Helper functions (shared logic with FortuneSheetDemo)
// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

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

const extractAllSheetNamesFromFormula = (formula: string): string[] => {
  const names: string[] = [];
  const regex = /'([^']+)'!|(?<![A-Za-z_])([A-Za-z_]\w*)!/g;
  let m;
  while ((m = regex.exec(formula)) !== null) {
    names.push(m[1] || m[2]);
  }
  return Array.from(new Set(names));
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

const extractColumnValues = (
  sheet: ExcelJS.Worksheet,
  colLetter: string
): string[] => {
  const colRef = parseCellRef(`${colLetter}1`);
  if (!colRef) return [];
  const values: string[] = [];
  sheet.eachRow({ includeEmpty: false }, (row) => {
    const cell = row.getCell(colRef.col);
    let v: any = cell.value;
    if (v === null || v === undefined) return;
    if (typeof v === 'object' && 'result' in v) v = v.result;
    else if (typeof v === 'object' && 'richText' in v)
      v = (v as ExcelJS.CellRichTextValue).richText
        .map((rt) => rt.text)
        .join('');
    const str = String(v).trim();
    if (str) values.push(str);
  });
  return values;
};

const extractFieldName = (sheetName: string): string => {
  let name = sheetName.replace(/^(Options|Lookup)_/i, '');
  name = name.replace(/^[a-z]+_/i, '');
  return name.toLowerCase();
};

const columnMatchScore = (fieldName: string, headerNorm: string): number => {
  if (fieldName === headerNorm) return 10000;
  if (headerNorm.includes(fieldName)) return fieldName.length * 100;
  if (fieldName.includes(headerNorm)) return headerNorm.length * 100;
  const fw = fieldName.match(/[a-z]{3,}/g) || [];
  const hw = headerNorm.match(/[a-z]{3,}/g) || [];
  let score = 0;
  for (const f of fw) {
    for (const h of hw) {
      if (f === h) score += f.length * 10;
      else if (h.includes(f) || f.includes(h))
        score += Math.min(f.length, h.length) * 5;
    }
  }
  return score;
};

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Extended cell type with dropdown support
// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

interface DropdownCell extends CellBase {
  dropdownOptions?: string[];
}

// Custom DataViewer - shows dropdown indicator
const DropdownDataViewer: React.FC<DataViewerProps<DropdownCell>> = ({ cell }) => {
  const hasDropdown = cell?.dropdownOptions && cell.dropdownOptions.length > 0;
  return (
    <span className="rs-cell-viewer">
      <span className="rs-cell-value">{cell?.value ?? ''}</span>
      {hasDropdown && <span className="rs-dropdown-arrow">▾</span>}
    </span>
  );
};

// Custom DataEditor - searchable dropdown list

const DropdownDataEditor: React.FC<DataEditorProps<CellBase>> = ({
  cell,
  onChange,
  exitEditMode,
}) => {
  const options = ((cell as DropdownCell)?.dropdownOptions) || [];
  const [search, setSearch] = useState('');
  const inputRef = useRef<HTMLInputElement>(null);
  const dropdownRef = useRef<HTMLDivElement>(null);
  const [position, setPosition] = useState<{ top: number; left: number } | null>(null);

  useEffect(() => {
    setTimeout(() => inputRef.current?.focus(), 0);
  }, []);

  // Position the dropdown below the active cell using fixed positioning
  useEffect(() => {
    const activeCell = document.querySelector('.Spreadsheet__active-cell');
    if (activeCell && dropdownRef.current) {
      const rect = activeCell.getBoundingClientRect();
      setPosition({
        top: rect.bottom + 2,
        left: rect.left,
      });
    }
  }, []);

  const filtered = options.filter((opt) =>
    opt.toLowerCase().includes(search.toLowerCase())
  );

  const handleSelect = (opt: string) => {
    onChange({ ...cell!, value: opt } as CellBase);
    exitEditMode();
  };

  return (
    <div 
      ref={dropdownRef}
      className="rs-dropdown-editor"
      style={position ? { top: position.top, left: position.left } : { visibility: 'hidden' }}
    >
      <input
        ref={inputRef}
        type="text"
        className="rs-dropdown-search"
        placeholder="Search..."
        value={search}
        onChange={(e) => setSearch(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Escape') exitEditMode();
          if (e.key === 'Enter' && filtered.length === 1)
            handleSelect(filtered[0]);
        }}
      />
      <div className="rs-dropdown-list">
        {filtered.map((opt) => (
          <div
            key={opt}
            className={`rs-dropdown-item${
              String(cell?.value) === opt ? ' selected' : ''
            }`}
            onMouseDown={(e) => {
              e.preventDefault();
              handleSelect(opt);
            }}
          >
            {opt}
          </div>
        ))}
        {filtered.length === 0 && (
          <div className="rs-dropdown-empty">No matches</div>
        )}
      </div>
    </div>
  );
};

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Column width utilities
// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

const computeColumnWidths = (data: Matrix<DropdownCell>): number[] => {
  if (data.length === 0) return [];
  const maxCols = Math.max(...data.map((row) => row.length));
  const widths: number[] = new Array(maxCols).fill(60);
  for (let c = 0; c < maxCols; c++) {
    let maxChars = 0;
    for (let r = 0; r < data.length; r++) {
      const cell = data[r]?.[c];
      if (!cell) continue;
      const text = String(cell.value ?? '');
      if (text.length > maxChars) maxChars = text.length;
    }
    widths[c] = Math.max(maxChars * 8 + 16, 60);
  }
  return widths;
};

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Main component
// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

const ReactSpreadsheetDemo: React.FC = () => {
  const [data, setData] = useState<Matrix<DropdownCell>>([]);
  const [fileName, setFileName] = useState('');
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [activeSheet, setActiveSheet] = useState(0);
  const [allSheetsData, setAllSheetsData] = useState<Matrix<DropdownCell>[]>(
    []
  );
  const [columnWidths, setColumnWidths] = useState<number[]>([]);
  const [showSpreadsheet, setShowSpreadsheet] = useState(false);
  const [hideRowIndicators, setHideRowIndicators] = useState(false);
  const [hideColumnIndicators, setHideColumnIndicators] = useState(false);
  const [selectedRow, setSelectedRow] = useState<number | null>(null);

  // Store dropdown options per sheet: Map<colIndex, string[]>
  const dropdownMapRef = useRef<Map<number, string[]>[]>([]);

  const columnLabels = useMemo(() => {
    if (data.length === 0) return [];
    const maxCols = Math.max(...data.map((row) => row.length));
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

  // Generate dynamic column width CSS
  const colWidthCSS = useMemo(() => {
    if (columnWidths.length === 0) return '';
    return columnWidths
      .map(
        (w, i) =>
          `.rs-auto-widths td:nth-of-type(${i + 1}) { min-width: ${w}px; max-width: ${Math.max(w, 200)}px; }`
      )
      .join('\n');
  }, [columnWidths]);

  // â”€â”€ File upload handler using ExcelJS â”€â”€
  const handleFileUpload = useCallback(async (file: File) => {
    setFileName(file.name);

    const arrayBuffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);

    // â”€â”€ Step 1: Defined names map â”€â”€
    const definedNameMap = new Map<string, string[]>();
    try {
      const model = workbook.definedNames?.model;
      if (model && Array.isArray(model)) {
        for (const entry of model) {
          if (entry.name && entry.ranges) {
            definedNameMap.set(entry.name, entry.ranges);
          }
        }
      }
    } catch (_) {}

    // â”€â”€ Step 2: Identify sheets to exclude â”€â”€
    const sheetsToExclude = new Set<string>();

    workbook.eachSheet((ws) => {
      if (ws.state === 'hidden' || ws.state === 'veryHidden') {
        sheetsToExclude.add(ws.name);
      }
    });

    workbook.eachSheet((ws) => {
      const dvModel = (ws as any).dataValidations?.model as
        | Record<string, any>
        | undefined;
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

    const sheetsWithValidation = new Set<string>();
    workbook.eachSheet((ws) => {
      const dvModel = (ws as any).dataValidations?.model as
        | Record<string, any>
        | undefined;
      if (dvModel && Object.keys(dvModel).length > 0) {
        sheetsWithValidation.add(ws.name);
      }
    });
    definedNameMap.forEach((ranges) => {
      for (const range of ranges) {
        const names = extractAllSheetNamesFromFormula(range);
        for (const name of names) {
          if (!sheetsWithValidation.has(name)) {
            sheetsToExclude.add(name);
          }
        }
      }
    });

    // Helper: resolve dropdown options
    const resolveDropdownOptions = (
      formula: string,
      currentSheet: ExcelJS.Worksheet
    ): string[] => {
      const clean = formula.replace(/^=/, '');
      if (/^'?[^(]+!\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+$/i.test(clean)) {
        return resolveRangeReference(clean, workbook);
      }
      if (definedNameMap.has(clean)) {
        const values: string[] = [];
        for (const range of definedNameMap.get(clean)!) {
          values.push(...resolveRangeReference(range, workbook));
        }
        if (values.length > 0) return values;
      }
      if (/^\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+$/i.test(clean)) {
        return resolveRangeReference(
          `'${currentSheet.name}'!${clean}`,
          workbook
        );
      }
      const refSheetNames = extractAllSheetNamesFromFormula(clean);
      if (refSheetNames.length > 0) {
        const colMatch = clean.match(/!\$?([A-Z]+)/i);
        if (colMatch) {
          const refSheet = workbook.getWorksheet(refSheetNames[0]);
          if (refSheet) return extractColumnValues(refSheet, colMatch[1]);
        }
      }
      return clean
        .replace(/^"/, '')
        .replace(/"$/, '')
        .split(',')
        .map((s) => s.trim())
        .filter(Boolean);
    };

    // â”€â”€ Step 3: Build spreadsheet data for visible sheets â”€â”€
    const allSheets: Matrix<DropdownCell>[] = [];
    const names: string[] = [];
    const allDropdownMaps: Map<number, string[]>[] = [];

    workbook.eachSheet((worksheet) => {
      if (sheetsToExclude.has(worksheet.name)) return;

      names.push(worksheet.name);

      let maxRow = 0;
      let maxCol = 0;

      const cellMap = new Map<string, { value: any; display: string }>();
      const dvMap = new Map<string, string[]>();

      worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        const r = rowNumber - 1;
        maxRow = Math.max(maxRow, r);
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
          const c = colNumber - 1;
          maxCol = Math.max(maxCol, c);
          let value = cell.value;
          let displayValue = '';
          if (value === null || value === undefined) return;
          if (typeof value === 'object' && 'result' in value) {
            value = value.result;
            displayValue = String(value || '');
          } else if (typeof value === 'object' && 'richText' in value) {
            displayValue = (value as ExcelJS.CellRichTextValue).richText
              .map((rt) => rt.text)
              .join('');
            value = displayValue;
          } else if (value instanceof Date) {
            displayValue = value.toLocaleDateString();
          } else {
            displayValue = String(value);
          }
          cellMap.set(`${r}_${c}`, { value, display: displayValue });
        });
      });

      // Parse data validations
      const dvModel = (worksheet as any).dataValidations?.model as
        | Record<string, any>
        | undefined;
      if (dvModel) {
        Object.keys(dvModel).forEach((addr) => {
          const dv = dvModel[addr];
          if (!dv || dv.type !== 'list' || !dv.formulae) return;
          const bounds = parseAddressRange(addr);
          if (!bounds) return;
          const options = dv.formulae[0]
            ? resolveDropdownOptions(dv.formulae[0], worksheet)
            : [];
          if (options.length === 0) return;
          for (let r = bounds.startRow; r <= bounds.endRow; r++) {
            for (let c = bounds.startCol; c <= bounds.endCol; c++) {
              dvMap.set(`${r}_${c}`, options);
            }
          }
        });
      }

      // Build column-level dropdown map
      const sheetDropdownMap = new Map<number, string[]>();
      dvMap.forEach((options, key) => {
        const c = Number(key.split('_')[1]);
        if (!sheetDropdownMap.has(c)) {
          sheetDropdownMap.set(c, options);
        }
      });
      allDropdownMaps.push(sheetDropdownMap);

      // Build 2D data matrix
      const sheetData: Matrix<DropdownCell> = [];
      for (let r = 0; r <= maxRow; r++) {
        const row: (DropdownCell | undefined)[] = [];
        for (let c = 0; c <= maxCol; c++) {
          const key = `${r}_${c}`;
          const cellInfo = cellMap.get(key);
          const dropdownOptions = dvMap.get(key);
          const isDropdown = dropdownOptions && dropdownOptions.length > 0;
          const columnHasDropdown = sheetDropdownMap.has(c);
          const isHeader = r === 0;
          
          const cellObj: DropdownCell = {
            value: cellInfo?.display ?? '',
            className: isHeader
              ? columnHasDropdown
                ? 'rs-header-cell rs-header-dropdown'
                : 'rs-header-cell'
              : isDropdown
                ? 'rs-dropdown-cell'
                : typeof cellInfo?.value === 'number'
                  ? 'cell-number'
                  : undefined,
          };
          if (isDropdown) {
            cellObj.dropdownOptions = dropdownOptions;
            cellObj.DataEditor = DropdownDataEditor;
            cellObj.DataViewer = DropdownDataViewer;
          }
          row.push(cellObj);
        }
        sheetData.push(row);
      }

      allSheets.push(sheetData);
    });

    // â”€â”€ Step 4: Auto-map hidden lookup sheets â”€â”€
    const hiddenLookups: {
      name: string;
      fieldName: string;
      values: string[];
    }[] = [];
    workbook.eachSheet((ws) => {
      if (ws.state !== 'hidden' && ws.state !== 'veryHidden') return;
      if (!/^(Options|Lookup)_/i.test(ws.name)) return;
      const values: string[] = [];
      ws.eachRow({ includeEmpty: false }, (row) => {
        const cell = row.getCell(1);
        let v: any = cell.value;
        if (v === null || v === undefined) return;
        if (typeof v === 'object' && 'result' in v) v = v.result;
        if (typeof v === 'object' && 'richText' in v)
          v = (v as ExcelJS.CellRichTextValue).richText
            .map((rt) => rt.text)
            .join('');
        const s = String(v).trim();
        if (s) values.push(s);
      });
      if (values.length > 0) {
        hiddenLookups.push({
          name: ws.name,
          fieldName: extractFieldName(ws.name),
          values,
        });
      }
    });

    if (hiddenLookups.length > 0) {
      allSheets.forEach((sheetData, sheetIdx) => {
        const sheetDropdownMap = allDropdownMaps[sheetIdx];
        if (sheetDropdownMap && sheetDropdownMap.size > 0) return;

        const headers: { col: number; norm: string }[] = [];
        if (sheetData.length > 0) {
          sheetData[0].forEach((cell, c) => {
            if (cell?.value) {
              headers.push({
                col: c,
                norm: String(cell.value)
                  .toLowerCase()
                  .replace(/[^a-z]/g, ''),
              });
            }
          });
        }
        if (headers.length === 0) return;

        const candidates: { idx: number; col: number; score: number }[] = [];
        hiddenLookups.forEach((lookup, idx) => {
          headers.forEach((h) => {
            const score = columnMatchScore(lookup.fieldName, h.norm);
            if (score > 0) candidates.push({ idx, col: h.col, score });
          });
        });

        candidates.sort((a, b) => b.score - a.score);
        const usedCols = new Set<number>();
        const usedSheets = new Set<number>();

        for (const c of candidates) {
          if (usedCols.has(c.col) || usedSheets.has(c.idx)) continue;
          usedCols.add(c.col);
          usedSheets.add(c.idx);

          const opts = hiddenLookups[c.idx].values;
          const maxDataRow = Math.min(sheetData.length, 200);
          for (let r = 1; r < maxDataRow; r++) {
            if (!sheetData[r]) sheetData[r] = [];
            if (!sheetData[r][c.col]) {
              sheetData[r][c.col] = { value: '' };
            }
            const cell = sheetData[r][c.col]!;
            if (!(cell as DropdownCell).dropdownOptions) {
              (cell as DropdownCell).dropdownOptions = opts;

              cell.DataEditor = DropdownDataEditor;
            }
          }

          if (!allDropdownMaps[sheetIdx]) {
            allDropdownMaps[sheetIdx] = new Map();
          }
          allDropdownMaps[sheetIdx].set(c.col, opts);
        }
      });
    }

    dropdownMapRef.current = allDropdownMaps;
    setSheetNames(names);
    setAllSheetsData(allSheets);
    setData(allSheets[0] || []);
    setColumnWidths(computeColumnWidths(allSheets[0] || []));
    setActiveSheet(0);
    setShowSpreadsheet(true);
  }, []);

  const handleSheetChange = (index: number) => {
    setActiveSheet(index);
    setData(allSheetsData[index] || []);
    setColumnWidths(computeColumnWidths(allSheetsData[index] || []));
  };

  const handleDataChange = (newData: Matrix<DropdownCell>) => {
    // Preserve dropdown properties on cells that lost them during edit
    const preserved = newData.map((row, r) =>
      row.map((cell, c) => {
        if (!cell) return cell;
        const original = data[r]?.[c] as DropdownCell | undefined;
        if (
          original?.dropdownOptions &&
          !(cell as DropdownCell).dropdownOptions
        ) {
          return {
            ...cell,
            className: 'rs-dropdown-cell',
            dropdownOptions: original.dropdownOptions,
            DataEditor: DropdownDataEditor, DataViewer: DropdownDataViewer,
          };
        }
        return cell;
      })
    );

    setData(preserved);
    setColumnWidths(computeColumnWidths(preserved));
    const updatedAllSheets = [...allSheetsData];
    updatedAllSheets[activeSheet] = preserved;
    setAllSheetsData(updatedAllSheets);
  };

  const addRow = () => {
    if (data.length === 0) return;
    const colCount = Math.max(...data.map((r) => r.length));
    const newRow: (DropdownCell | undefined)[] = [];
    const sheetDropdownMap = dropdownMapRef.current[activeSheet];
    for (let c = 0; c < colCount; c++) {
      const opts = sheetDropdownMap?.get(c);
      if (opts && opts.length > 0) {
        newRow.push({
          value: '',
          className: 'rs-dropdown-cell',
          dropdownOptions: opts,
          DataEditor: DropdownDataEditor, DataViewer: DropdownDataViewer,
        });
      } else {
        newRow.push({ value: '' });
      }
    }
    const newData = [...data, newRow];
    setData(newData);
    const updatedAllSheets = [...allSheetsData];
    updatedAllSheets[activeSheet] = newData;
    setAllSheetsData(updatedAllSheets);
  };

  const deleteSelectedRow = () => {
    if (selectedRow === null) {
      alert('Please select a row first by clicking on a cell');
      return;
    }
    if (data.length <= 1) {
      alert('Cannot delete the last remaining row');
      return;
    }
    // Don't allow deleting the header row (row 0)
    if (selectedRow === 0) {
      alert('Cannot delete the header row');
      return;
    }
    const newData = data.filter((_, index) => index !== selectedRow);
    setData(newData);
    const updatedAllSheets = [...allSheetsData];
    updatedAllSheets[activeSheet] = newData;
    setAllSheetsData(updatedAllSheets);
    setSelectedRow(null);
  };

  const handleSelect = (selection: Selection) => {
    // Track the selected row for deletion
    if (selection instanceof RangeSelection) {
      const range = selection.range;
      // Use the start row of the selection
      setSelectedRow(range.start.row);
    } else if (selection instanceof EntireRowsSelection) {
      setSelectedRow(selection.start);
    } else {
      setSelectedRow(null);
    }
  };

  // Sample data with dropdowns
  const createSampleData = () => {
    const departments = [
      'Engineering',
      'Marketing',
      'Sales',
      'HR',
      'Finance',
      'Operations',
    ];
    const statuses = ['Active', 'On Leave', 'Terminated'];

    const sampleData: Matrix<DropdownCell> = [
      [
        { value: 'Name', className: 'rs-header-cell' },
        { value: 'Department', className: 'rs-header-cell rs-header-dropdown' },
        { value: 'Salary', className: 'rs-header-cell' },
        { value: 'Start Date', className: 'rs-header-cell' },
        { value: 'Status', className: 'rs-header-cell rs-header-dropdown' },
      ],
      [
        { value: 'John Doe' },
        {
          value: 'Engineering',
          className: 'rs-dropdown-cell',
          dropdownOptions: departments,
          DataEditor: DropdownDataEditor, DataViewer: DropdownDataViewer,
        },
        { value: 75000, className: 'cell-number' },
        { value: '2020-01-15' },
        {
          value: 'Active',
          className: 'rs-dropdown-cell',
          dropdownOptions: statuses,
          DataEditor: DropdownDataEditor, DataViewer: DropdownDataViewer,
        },
      ],
      [
        { value: 'Jane Smith' },
        {
          value: 'Marketing',
          className: 'rs-dropdown-cell',
          dropdownOptions: departments,
          DataEditor: DropdownDataEditor, DataViewer: DropdownDataViewer,
        },
        { value: 65000, className: 'cell-number' },
        { value: '2021-03-22' },
        {
          value: 'Active',
          className: 'rs-dropdown-cell',
          dropdownOptions: statuses,
          DataEditor: DropdownDataEditor, DataViewer: DropdownDataViewer,
        },
      ],
      [
        { value: 'Bob Johnson' },
        {
          value: 'Sales',
          className: 'rs-dropdown-cell',
          dropdownOptions: departments,
          DataEditor: DropdownDataEditor, DataViewer: DropdownDataViewer,
        },
        { value: 85000, className: 'cell-number' },
        { value: '2018-07-01' },
        {
          value: 'Active',
          className: 'rs-dropdown-cell',
          dropdownOptions: statuses,
          DataEditor: DropdownDataEditor, DataViewer: DropdownDataViewer,
        },
      ],
      [
        { value: 'Alice Brown' },
        {
          value: 'HR',
          className: 'rs-dropdown-cell',
          dropdownOptions: departments,
          DataEditor: DropdownDataEditor, DataViewer: DropdownDataViewer,
        },
        { value: 70000, className: 'cell-number' },
        { value: '2019-11-10' },
        {
          value: 'On Leave',
          className: 'rs-dropdown-cell',
          dropdownOptions: statuses,
          DataEditor: DropdownDataEditor, DataViewer: DropdownDataViewer,
        },
      ],
      [
        { value: 'Charlie Wilson' },
        {
          value: 'Engineering',
          className: 'rs-dropdown-cell',
          dropdownOptions: departments,
          DataEditor: DropdownDataEditor, DataViewer: DropdownDataViewer,
        },
        { value: 72000, className: 'cell-number' },
        { value: '2022-02-28' },
        {
          value: 'Active',
          className: 'rs-dropdown-cell',
          dropdownOptions: statuses,
          DataEditor: DropdownDataEditor, DataViewer: DropdownDataViewer,
        },
      ],
    ];

    const sampleDropdownMap = new Map<number, string[]>();
    sampleDropdownMap.set(1, departments);
    sampleDropdownMap.set(4, statuses);
    dropdownMapRef.current = [sampleDropdownMap];

    setData(sampleData);
    setColumnWidths(computeColumnWidths(sampleData));
    setFileName('sample-data.xlsx');
    setSheetNames(['Employee Data']);
    setAllSheetsData([sampleData]);
    setActiveSheet(0);
    setShowSpreadsheet(true);
  };

  return (
    <div className="demo-page">
      <header className="demo-header">
        <h1>React Spreadsheet</h1>
        <span className="license-tag">MIT</span>
      </header>

      <div className="demo-info">
        <p>
          React Spreadsheet is a lightweight, customizable spreadsheet component
          for React. It provides Excel-like functionality with a clean API,
          dropdown cell support, and auto-fit column widths.
        </p>
        <div className="features-list">
          <span className="feature">âœ“ Editable Cells</span>
          <span className="feature">âœ“ Dropdown Validation</span>
          <span className="feature">âœ“ Auto-fit Columns</span>
          <span className="feature">âœ“ Copy/Paste</span>
          <span className="feature">âœ“ Keyboard Navigation</span>
          <span className="feature">âœ“ Searchable Dropdowns</span>
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

      {showSpreadsheet && data.length > 0 && (
        <div className="result-section">
          <div className="result-header">
            <h2>
              <span className="file-icon">ðŸ“Š</span>
              {fileName || 'Spreadsheet'}
              <span className="editable-badge">Editable</span>
            </h2>
            <button
              className="reset-btn"
              onClick={() => {
                setShowSpreadsheet(false);
                setData([]);
                setAllSheetsData([]);
                setSheetNames([]);
                setFileName('');
                setColumnWidths([]);
              }}
            >
              Upload New File
            </button>
          </div>

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

          <div className="toolbar-toggles">
            <label className="toggle-label">
              <input
                type="checkbox"
                checked={hideRowIndicators}
                onChange={(e) => setHideRowIndicators(e.target.checked)}
              />
              Hide Row Numbers
            </label>
            <label className="toggle-label">
              <input
                type="checkbox"
                checked={hideColumnIndicators}
                onChange={(e) => setHideColumnIndicators(e.target.checked)}
              />
              Hide Column Headers
            </label>
            <div className="toolbar-buttons">
              <button className="add-row-col-btn" onClick={addRow}>
                + Row
              </button>
              <button className="delete-row-btn" onClick={deleteSelectedRow}>
                − Delete Row{selectedRow !== null && selectedRow > 0 ? ` (${selectedRow})` : ''}
              </button>
            </div>
          </div>

          <div className="react-spreadsheet-wrapper rs-auto-widths">
            <style>{colWidthCSS}</style>
            <Spreadsheet
              data={data}
              onChange={handleDataChange}
              columnLabels={columnLabels}
              rowLabels={rowLabels}
              hideRowIndicators={hideRowIndicators}
              hideColumnIndicators={hideColumnIndicators}
              onSelect={handleSelect}
            />
          </div>

          <div className="data-info">
            <h4>Interactive Features</h4>
            <p>â€¢ Click any cell to edit â€” dropdown cells show a searchable list</p>
            <p>â€¢ Use arrow keys to navigate between cells</p>
            <p>â€¢ Select multiple cells and copy/paste</p>
            <p>â€¢ Tab to move to next cell</p>
            <p>â€¢ Columns auto-fit to content width</p>
            <p>â€¢ Dropdown cells show â–¾ indicator</p>
          </div>
        </div>
      )}

      <div className="code-example">
        <div className="code-header">Usage Example</div>
        <pre>
          <code>{`import Spreadsheet, { CellBase } from 'react-spreadsheet';

// Custom cell with dropdown support
interface DropdownCell extends CellBase {
  dropdownOptions?: string[];
}

const data = [
  [{ value: 'Name' }, { value: 'Department' }],
  [
    { value: 'John' },
    {
      value: 'Engineering',
      dropdownOptions: ['Engineering', 'Marketing', 'Sales'],
      DataViewer: DropdownViewer,   // Custom viewer with â–¾
      DataEditor: DropdownEditor,   // Searchable dropdown
    },
  ],
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
