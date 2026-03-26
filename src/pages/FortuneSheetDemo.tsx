import React, { useState, useCallback, useRef, useEffect } from 'react';
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
  config?: { columnlen?: Record<number, number> };
  defaultRowHeight?: number;
  defaultColWidth?: number;
  dataVerification?: { [key: string]: any };
}

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

/**
 * Extract ALL sheet names referenced anywhere in a formula string.
 * Handles patterns like 'Sheet Name'!A1, SheetName!A1, OFFSET(Sheet!A1,...), etc.
 */
const extractAllSheetNamesFromFormula = (formula: string): string[] => {
  const names: string[] = [];
  const regex = /'([^']+)'!|(?<![A-Za-z_])([A-Za-z_]\w*)!/g;
  let m;
  while ((m = regex.exec(formula)) !== null) {
    names.push(m[1] || m[2]);
  }
  return Array.from(new Set(names));
};

/**
 * Resolve a direct range reference like 'SheetName'!$A$1:$A$10 to string values.
 */
const resolveRangeReference = (
  formula: string,
  workbook: ExcelJS.Workbook
): string[] => {
  const clean = formula.replace(/^=/, '');
  const m = clean.match(/^'?([^'!]+)'?!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/i);
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
      if (typeof cellValue === 'object' && 'result' in cellValue) cellValue = cellValue.result;
      else if (typeof cellValue === 'object' && 'richText' in cellValue)
        cellValue = (cellValue as ExcelJS.CellRichTextValue).richText.map((rt) => rt.text).join('');
      const str = String(cellValue).trim();
      if (str) values.push(str);
    }
  }
  return values;
};

/**
 * Parse an address string (e.g. "A1", "$B$2:$D$50") into 0-based row/col bounds.
 */
const parseAddressRange = (
  address: string
): { startRow: number; startCol: number; endRow: number; endCol: number } | null => {
  const rangeMatch = address.match(
    /^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/i
  );
  if (rangeMatch) {
    const s = parseCellRef(`${rangeMatch[1]}${rangeMatch[2]}`);
    const e = parseCellRef(`${rangeMatch[3]}${rangeMatch[4]}`);
    if (s && e) {
      return { startRow: s.row - 1, startCol: s.col - 1, endRow: e.row - 1, endCol: e.col - 1 };
    }
  }
  const cellMatch = address.match(/^\$?([A-Z]+)\$?(\d+)$/i);
  if (cellMatch) {
    const ref = parseCellRef(`${cellMatch[1]}${cellMatch[2]}`);
    if (ref) {
      return { startRow: ref.row - 1, startCol: ref.col - 1, endRow: ref.row - 1, endCol: ref.col - 1 };
    }
  }
  return null;
};

/**
 * Extract all non-empty string values from a specific column in a worksheet.
 * Used as a fallback for complex formulas (OFFSET, INDIRECT, etc.).
 */
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
      v = (v as ExcelJS.CellRichTextValue).richText.map((rt) => rt.text).join('');
    const str = String(v).trim();
    if (str) values.push(str);
  });
  return values;
};

/**
 * Extract the "field name" portion from a Dynamics-style hidden sheet name.
 * e.g. "Options_msdyn_wastediversionsit" → "wastediversionsit"
 *      "Lookup_hmsm_wastecode"           → "wastecode"
 */
const extractFieldName = (sheetName: string): string => {
  let name = sheetName.replace(/^(Options|Lookup)_/i, '');
  name = name.replace(/^[a-z]+_/i, '');
  return name.toLowerCase();
};

/**
 * Score how well a hidden-sheet field name matches a column header.
 * Higher = better.  0 = no meaningful match.
 */
const columnMatchScore = (fieldName: string, headerNorm: string): number => {
  if (fieldName === headerNorm) return 10000;
  if (headerNorm.includes(fieldName)) return fieldName.length * 100;
  if (fieldName.includes(headerNorm)) return headerNorm.length * 100;

  // Word-overlap fallback: split both into >=3-char substrings and check overlap
  const fw = fieldName.match(/[a-z]{3,}/g) || [];
  const hw = headerNorm.match(/[a-z]{3,}/g) || [];
  let score = 0;
  for (const f of fw) {
    for (const h of hw) {
      if (f === h) score += f.length * 10;
      else if (h.includes(f) || f.includes(h)) score += Math.min(f.length, h.length) * 5;
    }
  }
  return score;
};

const FortuneSheetDemo: React.FC = () => {
  const [sheets, setSheets] = useState<SheetData[]>([]);
  const [fileName, setFileName] = useState('');
  const [showSpreadsheet, setShowSpreadsheet] = useState(false);
  const [showToolbar, setShowToolbar] = useState(false);
  const [showFormulaBar, setShowFormulaBar] = useState(false);
  const workbookRef = useRef<WorkbookInstance>(null);

  // Inject a search input into the Fortune Sheet dropdown list when it appears
  useEffect(() => {
    let pending = false;
    const observer = new MutationObserver(() => {
      if (pending) return;
      pending = true;
      requestAnimationFrame(() => {
        pending = false;
        const dropdown = document.getElementById('luckysheet-dataVerification-dropdown-List');
        if (!dropdown || dropdown.style.display === 'none') return;
        if (dropdown.querySelector('.dropdown-search-input')) return;

        const searchInput = document.createElement('input');
        searchInput.type = 'text';
        searchInput.placeholder = 'Search...';
        searchInput.className = 'dropdown-search-input';

        ['mousedown', 'mouseup', 'click', 'keydown', 'change'].forEach((evt) => {
          searchInput.addEventListener(evt, (e) => e.stopPropagation());
        });

        searchInput.addEventListener('input', () => {
          const query = searchInput.value.toLowerCase();
          const items = dropdown.querySelectorAll('.dropdown-List-item');
          items.forEach((item) => {
            const text = (item as HTMLElement).textContent?.toLowerCase() || '';
            (item as HTMLElement).style.display = text.includes(query) ? '' : 'none';
          });
        });

        dropdown.insertBefore(searchInput, dropdown.firstChild);
      });
    });

    observer.observe(document.body, { childList: true, subtree: true });
    return () => observer.disconnect();
  }, []);

  const handleFileUpload = useCallback(async (file: File) => {
    setFileName(file.name);

    const arrayBuffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);

    // ── Step 1: Build a map of workbook-level defined names ──
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
    } catch (_) { /* definedNames may not exist */ }

    // ── Step 2: Identify every sheet that should be hidden ──
    const sheetsToExclude = new Set<string>();

    // 2a – Hidden / very-hidden sheets
    workbook.eachSheet((ws) => {
      if (ws.state === 'hidden' || ws.state === 'veryHidden') {
        sheetsToExclude.add(ws.name);
      }
    });

    // 2b – Sheets referenced (directly or via named ranges) by data-validation formulas.
    //       Uses worksheet.dataValidations.model so we catch validations on empty cells too.
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

    // 2c – Sheets referenced by ANY defined name that are not themselves data-entry sheets.
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

    // Helper: resolve dropdown options from any formula shape
    const resolveDropdownOptions = (
      formula: string,
      currentSheet: ExcelJS.Worksheet
    ): string[] => {
      const clean = formula.replace(/^=/, '');

      // 1) Direct range reference:  'Sheet'!$A$1:$A$10
      if (/^'?[^(]+!\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+$/i.test(clean)) {
        return resolveRangeReference(clean, workbook);
      }

      // 2) Named range
      if (definedNameMap.has(clean)) {
        const values: string[] = [];
        for (const range of definedNameMap.get(clean)!) {
          values.push(...resolveRangeReference(range, workbook));
        }
        if (values.length > 0) return values;
      }

      // 3) Same-sheet range reference:  $A$1:$A$10  (no sheet prefix)
      if (/^\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+$/i.test(clean)) {
        return resolveRangeReference(`'${currentSheet.name}'!${clean}`, workbook);
      }

      // 4) Complex formulas with a sheet reference (OFFSET, INDIRECT, etc.)
      //    Fall back to pulling every value from the first referenced column.
      const refSheetNames = extractAllSheetNamesFromFormula(clean);
      if (refSheetNames.length > 0) {
        const colMatch = clean.match(/!\$?([A-Z]+)/i);
        if (colMatch) {
          const refSheet = workbook.getWorksheet(refSheetNames[0]);
          if (refSheet) return extractColumnValues(refSheet, colMatch[1]);
        }
      }

      // 5) Inline comma-separated values:  "Option1,Option2,Option3"
      return clean
        .replace(/^"/, '')
        .replace(/"$/, '')
        .split(',')
        .map((s) => s.trim())
        .filter(Boolean);
    };

    // ── Step 3: Build Fortune Sheet data for visible, non-lookup sheets ──
    const fortuneSheets: SheetData[] = [];

    workbook.eachSheet((worksheet) => {
      if (sheetsToExclude.has(worksheet.name)) return;

      const celldata: any[] = [];
      const dataVerification: { [key: string]: any } = {};
      let maxRow = 0;
      let maxCol = 0;

      worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        const r = rowNumber - 1;
        maxRow = Math.max(maxRow, r);

        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
          const c = colNumber - 1;
          maxCol = Math.max(maxCol, c);

          let value = cell.value;
          let displayValue = '';

          if (value === null || value === undefined) {
            return;
          } else if (typeof value === 'object' && 'result' in value) {
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

          const cellValue: any = {
            r,
            c,
            v: { v: value, m: displayValue },
          };

          if (typeof value === 'number') {
            cellValue.v.ct = { fa: 'General', t: 'n' };
          }

          celldata.push(cellValue);
        });
      });

      // ── Data validation from worksheet-level model ──
      // This catches validations on ALL cells including empty ones.
      const dvModel = (worksheet as any).dataValidations?.model as
        | Record<string, any>
        | undefined;
      if (dvModel) {
        Object.keys(dvModel).forEach((addr) => {
          const dv = dvModel[addr];
          if (!dv) return;
          const bounds = parseAddressRange(addr);
          if (!bounds) return;

          for (let r = bounds.startRow; r <= bounds.endRow; r++) {
            for (let c = bounds.startCol; c <= bounds.endCol; c++) {
              const key = `${r}_${c}`;
              if (dataVerification[key]) continue;

              if (dv.type === 'list') {
                const options =
                  dv.formulae && dv.formulae[0]
                    ? resolveDropdownOptions(dv.formulae[0], worksheet)
                    : [];
                if (options.length > 0) {
                  dataVerification[key] = {
                    type: 'dropdown',
                    type2: 'false',
                    value1: options.join(','),
                    value2: '',
                    checked: false,
                    remote: false,
                    prohibitInput: false,
                    hintShow: false,
                    hintText: '',
                  };
                }
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
          }
        });
      }

      // Auto-fit column widths based on content
      const columnlen: Record<number, number> = {};
      const colMaxChars: Record<number, number> = {};
      celldata.forEach((cd: any) => {
        const text = String(cd.v?.m ?? cd.v?.v ?? '');
        const len = text.length;
        if (!colMaxChars[cd.c] || len > colMaxChars[cd.c]) {
          colMaxChars[cd.c] = len;
        }
      });
      for (const col in colMaxChars) {
        const chars = colMaxChars[col];
        columnlen[col] = Math.max(chars * 7 + 10, 60);
      }

      fortuneSheets.push({
        name: worksheet.name,
        celldata,
        row: maxRow + 1,
        column: maxCol + 1,
        config: { columnlen },
        defaultRowHeight: 40,
        dataVerification: Object.keys(dataVerification).length > 0 ? dataVerification : undefined,
      });
    });

    // ── Step 4: Auto-map hidden lookup sheets to columns ──
    // Dynamics 365 templates (and similar) store option values in hidden sheets
    // named Options_<field> or Lookup_<field> but have NO standard data validation.
    // Detect this pattern and synthesise dropdown validations.
    const hiddenLookups: { name: string; fieldName: string; values: string[] }[] = [];
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
          v = (v as ExcelJS.CellRichTextValue).richText.map((rt) => rt.text).join('');
        const s = String(v).trim();
        if (s) values.push(s);
      });
      if (values.length > 0) {
        hiddenLookups.push({ name: ws.name, fieldName: extractFieldName(ws.name), values });
      }
    });

    if (hiddenLookups.length > 0) {
      fortuneSheets.forEach((sheet) => {
        if (sheet.dataVerification && Object.keys(sheet.dataVerification).length > 0) return;

        // Gather header cells (row 0) with normalised text
        const headers: { col: number; norm: string }[] = [];
        sheet.celldata.forEach((cd: any) => {
          if (cd.r === 0 && cd.v) {
            const text = String(cd.v.v ?? cd.v.m ?? '');
            headers.push({ col: cd.c, norm: text.toLowerCase().replace(/[^a-z]/g, '') });
          }
        });
        if (headers.length === 0) return;

        // Build all (lookup, column, score) candidates
        const candidates: { idx: number; col: number; score: number }[] = [];
        hiddenLookups.forEach((lookup, idx) => {
          headers.forEach((h) => {
            const score = columnMatchScore(lookup.fieldName, h.norm);
            if (score > 0) candidates.push({ idx, col: h.col, score });
          });
        });

        // Greedy best-match assignment
        candidates.sort((a, b) => b.score - a.score);
        const usedCols = new Set<number>();
        const usedSheets = new Set<number>();
        const dv: { [key: string]: any } = sheet.dataVerification || {};

        for (const c of candidates) {
          if (usedCols.has(c.col) || usedSheets.has(c.idx)) continue;
          usedCols.add(c.col);
          usedSheets.add(c.idx);

          const opts = hiddenLookups[c.idx].values.join(',');
          const maxDataRow = Math.min(sheet.row || 50, 200);
          for (let r = 1; r < maxDataRow; r++) {
            const key = `${r}_${c.col}`;
            if (!dv[key]) {
              dv[key] = {
                type: 'dropdown',
                type2: 'false',
                value1: opts,
                value2: '',
                checked: false,
                remote: false,
                prohibitInput: false,
                hintShow: false,
                hintText: '',
              };
            }
          }
        }

        if (Object.keys(dv).length > 0) {
          sheet.dataVerification = dv;
        }
      });
    }

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
        type2: 'false',
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
        type2: 'false',
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
        row: 7,
        column: 5,
        config: {
          columnlen: {
            0: 100,  // Name
            1: 90,   // Department
            2: 80,   // Salary
            3: 90,   // Start Date
            4: 80,   // Status
          },
        },
        defaultRowHeight: 40,
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
            <button
              className="add-row-col-btn"
              onClick={() => {
                const wb = workbookRef.current;
                if (!wb) return;
                const sheet = wb.getSheet();
                const rowCount = sheet?.row ?? 0;
                wb.insertRowOrColumn('row', rowCount - 1, 1, 'rightbottom');
              }}
            >
              + Add Row
            </button>
            <button
              className="delete-row-btn"
              onClick={() => {
                const wb = workbookRef.current;
                if (!wb) return;
                const selection = wb.getSelection();
                if (!selection || selection.length === 0) {
                  alert('Please select a row or cell first');
                  return;
                }
                // Get the row range from selection
                const sel = selection[0];
                const startRow = sel.row[0];
                const endRow = sel.row[1];
                // Don't allow deleting the header row (row 0)
                if (startRow === 0 && endRow === 0) {
                  alert('Cannot delete the header row');
                  return;
                }
                const actualStart = Math.max(startRow, 0);
                wb.deleteRowOrColumn('row', actualStart, endRow);
              }}
            >
              − Delete Row
            </button>
            <button
              className="add-row-col-btn"
              onClick={() => {
                const wb = workbookRef.current;
                if (!wb) return;
                const sheet = wb.getSheet();
                const colCount = sheet?.column ?? 0;
                wb.insertRowOrColumn('column', colCount - 1, 1, 'rightbottom');
              }}
            >
              + Add Column
            </button>
          </div>

          <div className="fortune-sheet-wrapper">
            <Workbook
              key={fileName}
              ref={workbookRef}
              data={sheets}
              allowEdit
              showToolbar={showToolbar}
              showFormulaBar={showFormulaBar}
              rowHeaderWidth={46}
              columnHeaderHeight={20}
              defaultFontSize={13}
              defaultRowHeight={28}
              onChange={(data) => {
                console.log('Sheet data changed:', data);
              }}
              onOp={(ops) => {
                // After a cell value changes, auto-fit the affected column width
                const wb = workbookRef.current;
                if (!wb) return;
                const colsToFit = new Set<number>();
                for (const op of ops) {
                  // path like ["data", row, col, "m"] or ["data", row, col, "v"]
                  if (op.op === 'replace' && op.path[0] === 'data' && op.path.length >= 3) {
                    const col = Number(op.path[2]);
                    if (!isNaN(col)) colsToFit.add(col);
                  }
                }
                if (colsToFit.size > 0) {
                  const sheet = wb.getSheet();
                  if (sheet?.data) {
                    const data2d = sheet.data;
                    const rowCount = data2d.length;
                    const colWidthUpdate: Record<string, number> = {};
                    colsToFit.forEach((col) => {
                      let maxChars = 0;
                      for (let r = 0; r < rowCount; r++) {
                        const cell = data2d[r]?.[col];
                        if (!cell) continue;
                        const text = String(cell.m ?? cell.v ?? '');
                        if (text.length > maxChars) maxChars = text.length;
                      }
                      const newWidth = Math.max(maxChars * 7 + 10, 60);
                      colWidthUpdate[String(col)] = newWidth;
                    });
                    wb.setColumnWidth(colWidthUpdate);
                  }
                }
                // Hide the dropdown button so Fortune Sheet repositions it correctly
                const btn = document.getElementById('luckysheet-dataVerification-dropdown-btn');
                if (btn) btn.style.display = 'none';
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
