# Fortune Sheet - Detailed Package Report

## Overview

| Property | Value |
|----------|-------|
| **Package Name** | `@fortune-sheet/react` |
| **Version** | ^1.0.4 |
| **License** | MIT |
| **NPM Weekly Downloads** | 10K+ |
| **Category** | Full-Featured Spreadsheet UI Component |
| **Primary Use** | Excel-like interactive spreadsheet in React |

## Description

Fortune Sheet is a comprehensive, Excel-like spreadsheet component for React, forked from the popular Luckysheet project with modern TypeScript support and improved React integration. It provides a full-featured spreadsheet experience with toolbar, formula bar, and native Excel-like UI.

## Core Capabilities

### 1. Excel-Like Interface
- **Toolbar**: Full formatting toolbar (font, colors, borders, alignment)
- **Formula Bar**: Cell reference display and formula editing
- **Sheet Tabs**: Multi-sheet navigation with tab management
- **Context Menus**: Right-click menus for cell operations
- **Cell Selection**: Single, multi-cell, and range selection

### 2. Cell Features
- **Rich Cell Types**: Text, numbers, dates, formulas
- **Cell Styling**: Background colors, fonts, borders, alignment
- **Data Validation**: Dropdown lists with custom options
- **Merged Cells**: Cell spanning support
- **Comments**: Cell annotations

### 3. Formula Support
- **Built-in Functions**: 400+ Excel-compatible functions
- **Cross-Sheet References**: Reference cells across sheets
- **Array Formulas**: Multi-cell formula results
- **Auto-calculation**: Real-time formula updates

### 4. Advanced Features
- **Charts**: Embedded chart support
- **Pivot Tables**: Data summarization
- **Conditional Formatting**: Rules-based styling
- **Data Filtering**: Column filters
- **Find & Replace**: Search functionality

## Implementation in Demo

```typescript
import { Workbook, WorkbookInstance } from '@fortune-sheet/react';
import '@fortune-sheet/react/dist/index.css';

const FortuneSheetDemo: React.FC = () => {
  const workbookRef = useRef<WorkbookInstance>(null);
  const [sheets, setSheets] = useState<SheetData[]>([]);

  return (
    <Workbook
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
        // Handle cell operations
      }}
    />
  );
};
```

## Sheet Data Structure

```typescript
interface SheetData {
  name: string;              // Sheet name
  celldata: CellData[];      // Array of cell objects
  row?: number;              // Total rows
  column?: number;           // Total columns
  config?: {
    columnlen?: Record<number, number>;  // Column widths
  };
  defaultRowHeight?: number;
  defaultColWidth?: number;
  dataVerification?: {       // Dropdown validations
    [key: string]: {
      type: string;
      value1: string;        // Options list
    };
  };
}

interface CellData {
  r: number;      // Row index (0-based)
  c: number;      // Column index (0-based)
  v: {
    v?: any;      // Raw value
    m?: string;   // Display value
    ct?: {        // Cell type
      fa: string; // Format
      t: string;  // Type
    };
    bg?: string;  // Background color
    fc?: string;  // Font color
    bl?: number;  // Bold (1 = true)
    it?: number;  // Italic (1 = true)
  };
}
```

## Data Validation (Dropdowns)

```typescript
// Creating dropdown validation in Fortune Sheet format
const dataVerification = {
  '1_2': {  // Row 1, Column 2
    type: 'dropdown',
    type2: null,
    value1: 'Option1,Option2,Option3',  // Comma-separated
    value2: '',
    checked: false,
    remote: false,
    prohibitInput: false,
    hintShow: false,
    hintText: '',
  },
};
```

## Component Props

| Prop | Type | Description |
|------|------|-------------|
| `data` | `SheetData[]` | Array of sheet data |
| `ref` | `WorkbookInstance` | Ref for imperative API |
| `allowEdit` | `boolean` | Enable cell editing |
| `showToolbar` | `boolean` | Show/hide toolbar |
| `showFormulaBar` | `boolean` | Show/hide formula bar |
| `rowHeaderWidth` | `number` | Row header column width |
| `columnHeaderHeight` | `number` | Column header height |
| `defaultFontSize` | `number` | Default font size (px) |
| `defaultRowHeight` | `number` | Default row height (px) |
| `onChange` | `function` | Data change callback |
| `onOp` | `function` | Operation callback |

## WorkbookInstance API

```typescript
const wb = workbookRef.current;

// Get current sheet data
const sheet = wb.getSheet();

// Get all sheets
const allSheets = wb.getAllSheets();

// Set cell value
wb.setCellValue(row, col, value, options);

// Get cell value
const val = wb.getCellValue(row, col);
```

## Key Strengths

1. **Complete Excel Experience**: Full toolbar, formula bar, sheet tabs
2. **React Native**: Built specifically for React with hooks support
3. **TypeScript**: Full type definitions
4. **Canvas Rendering**: High-performance rendering via HTML5 Canvas
5. **Formula Engine**: Powerful built-in formula support
6. **Data Validation**: Built-in dropdown support with visual indicators
7. **Active Development**: Regular updates and improvements

## Limitations

1. **Canvas-Based**: Some CSS customizations require understanding internal structure
2. **Bundle Size**: Larger than simpler alternatives (~500KB+)
3. **Learning Curve**: Complex API for advanced customizations
4. **Documentation**: Some documentation only available in Chinese
5. **Cell Padding**: Canvas cells require special handling for padding/styling

## CSS Customization

```css
/* Override internal styles */
.fortune-sheet-wrapper .luckysheet-input-box-inner {
  padding: 4px 8px !important;
}

.fortune-sheet-wrapper .luckysheet-cell-input {
  padding: 4px 8px !important;
}

/* Header styling */
.fortune-sheet-wrapper .fortune-col-header {
  background: #f1f5f9 !important;
  font-weight: 500 !important;
}

/* Dropdown button */
#luckysheet-dataVerification-dropdown-btn {
  background: #0369a1 !important;
  border-radius: 3px;
}
```

## Use Cases

| Use Case | Suitability |
|----------|-------------|
| Interactive Excel-like editor | ⭐⭐⭐⭐⭐ Excellent |
| Complex spreadsheet application | ⭐⭐⭐⭐⭐ Excellent |
| Data entry forms with validation | ⭐⭐⭐⭐ Good |
| Simple data display | ⭐⭐ Overkill |
| Embedded mini-spreadsheet | ⭐⭐⭐ Moderate |
| Server-side rendering | ❌ Not applicable |

## Integration with ExcelJS

The demo uses ExcelJS to parse uploaded files and convert to Fortune Sheet format:

```typescript
// Convert ExcelJS cell to Fortune Sheet format
const fortuneCell = {
  r: rowIndex,
  c: colIndex,
  v: {
    v: cell.value,      // Raw value
    m: displayValue,    // Formatted value
    ct: { fa: 'General', t: 's' },
    bg: cell.fill?.fgColor?.argb,
    fc: cell.font?.color?.argb,
    bl: cell.font?.bold ? 1 : 0,
    it: cell.font?.italic ? 1 : 0,
  },
};
```

## Performance Considerations

- Virtual rendering for large datasets
- Canvas-based rendering minimizes DOM nodes
- Efficient cell data structure (sparse array)
- Auto column width calculation available
- Lazy loading for large sheets recommended

## Conclusion

Fortune Sheet is the most complete Excel-like spreadsheet solution for React applications. It provides the familiar Excel experience users expect, including toolbar, formulas, and data validation. While it has a larger footprint than simpler alternatives, it's the best choice when you need a full-featured spreadsheet editor rather than just a data grid.
