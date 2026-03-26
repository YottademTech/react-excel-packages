# XLSX (SheetJS) - Detailed Package Report

## Overview

| Property | Value |
|----------|-------|
| **Package Name** | `xlsx` |
| **Version** | ^0.18.5 |
| **License** | Apache 2.0 |
| **NPM Weekly Downloads** | 2.5M+ |
| **Category** | Excel File Parser/Writer (No UI) |
| **Primary Use** | Universal spreadsheet parsing and generation |

## Description

SheetJS (XLSX) is the most widely-used spreadsheet parser and writer for JavaScript. It supports an extensive range of file formats and provides utilities for converting between formats. Known for its reliability and broad format support.

## Core Capabilities

### 1. Multi-Format Support
- **Excel Formats**: XLSX, XLSM, XLSB, XLS, BIFF5
- **Other Formats**: CSV, TSV, ODS, SYLK, DIF, HTML
- **Output Formats**: JSON, CSV, HTML, Markdown

### 2. Reading Features
- **Cell Data**: Values, types, formulas
- **Cell Metadata**: Number formats, cell references
- **Worksheet Properties**: Ranges, merges, filters
- **Date Handling**: Automatic date parsing with format preservation

### 3. Writing Features
- **Format Conversion**: Convert between any supported formats
- **Data Export**: JSON arrays to Excel
- **Multiple Sheets**: Create multi-sheet workbooks

### 4. Utilities
- **sheet_to_json**: Convert worksheet to JSON array
- **json_to_sheet**: Create worksheet from JSON
- **encode_cell/decode_cell**: Cell address utilities
- **encode_range/decode_range**: Range utilities

## Implementation in Demo

```typescript
import * as XLSX from 'xlsx';

// Read Excel file
const data = new Uint8Array(arrayBuffer);
const workbook = XLSX.read(data, { 
  type: 'array',
  cellDates: true,    // Parse dates as Date objects
  cellStyles: true,   // Preserve styles
  cellNF: true,       // Preserve number formats
});

// Access sheets
workbook.SheetNames.forEach((sheetName) => {
  const worksheet = workbook.Sheets[sheetName];
  
  // Convert to JSON array (header: 1 for 2D array)
  const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
    header: 1, 
    raw: false,
    dateNF: 'yyyy-mm-dd'
  });
});

// Access individual cells
const cell = worksheet['A1'] as XLSX.CellObject;
// cell.t = type ('s', 'n', 'b', 'd', 'e')
// cell.v = raw value
// cell.w = formatted text
// cell.z = number format

// Get cell range
const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
```

## Cell Type System

| Type Code | Meaning | Example |
|-----------|---------|---------|
| `s` | String | "Hello" |
| `n` | Number | 42, 3.14 |
| `b` | Boolean | true, false |
| `d` | Date | Date object |
| `e` | Error | #N/A, #REF! |

## Key Strengths

1. **Industry Standard**: Most popular spreadsheet library (2.5M+ weekly downloads)
2. **Format Support**: Reads and writes almost every spreadsheet format
3. **Lightweight**: Minimal dependencies
4. **Browser & Node.js**: Works in both environments
5. **Fast Parsing**: Optimized for performance
6. **Well Documented**: Extensive documentation and examples

## Limitations

1. **Styling (Free Version)**: Limited style support in open-source version
2. **No UI Component**: Parser only - needs separate UI
3. **Data Validation**: Basic support, complex formulas may not resolve
4. **Pro Features**: Some advanced features require SheetJS Pro license

## Data Validation Handling

```typescript
// Access data validations (if available)
const dataValidations = (worksheet as any)['!dataValidation'];
if (dataValidations) {
  dataValidations.forEach((dv: any) => {
    if (dv.type === 'list') {
      // dv.sqref = cell reference
      // dv.formula1 = dropdown options or formula
      const options = dv.formula1?.split(',');
    }
  });
}
```

## Comparison with ExcelJS

| Feature | XLSX (SheetJS) | ExcelJS |
|---------|----------------|---------|
| Format Support | ⭐⭐⭐⭐⭐ (20+ formats) | ⭐⭐⭐ (XLSX, CSV) |
| Styling | ⭐⭐ (basic free) | ⭐⭐⭐⭐⭐ (full) |
| Data Validation | ⭐⭐⭐ | ⭐⭐⭐⭐⭐ |
| Streaming | ⭐⭐⭐ | ⭐⭐⭐⭐ |
| API Simplicity | ⭐⭐⭐⭐⭐ | ⭐⭐⭐ |
| Download Volume | 2.5M+ | 800K+ |

## Use Cases

| Use Case | Suitability |
|----------|-------------|
| Quick CSV/Excel parsing | ⭐⭐⭐⭐⭐ Excellent |
| Format conversion | ⭐⭐⭐⭐⭐ Excellent |
| Data export to Excel | ⭐⭐⭐⭐ Good |
| Complex styled reports | ⭐⭐ Limited (use ExcelJS) |
| Interactive UI | ❌ Not applicable |

## Demo Features Demonstrated

1. **File Loading**: Multi-format Excel file reading
2. **Sheet Navigation**: Tab-based multi-sheet display
3. **Cell Type Detection**: Type-specific rendering (string, number, date)
4. **Data Validation**: Basic dropdown detection
5. **Cell Reference Utilities**: Column letter encoding

## Sample Output Rendering

```typescript
// Render cells with type-specific styling
const getCellClassName = (rowIdx: number, colIdx: number): string => {
  const cellAddress = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
  const info = sheet.cellInfo[cellAddress];
  
  const classes = ['cell'];
  if (info) {
    classes.push(`cell-type-${info.type}`); // cell-type-string, cell-type-number, etc.
    if (info.validation?.type === 'dropdown') {
      classes.push('cell-dropdown');
    }
  }
  return classes.join(' ');
};
```

## Conclusion

SheetJS is the go-to library for spreadsheet parsing when you need broad format support and reliable performance. While it lacks the rich styling capabilities of ExcelJS in its free version, its simplicity and format support make it ideal for data import/export scenarios. Best used when you need to quickly read Excel/CSV data without complex styling requirements.
