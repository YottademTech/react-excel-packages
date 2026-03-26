# React Spreadsheet - Detailed Package Report

## Overview

| Property | Value |
|----------|-------|
| **Package Name** | `react-spreadsheet` |
| **Version** | ^0.10.1 |
| **License** | MIT |
| **NPM Weekly Downloads** | 50K+ |
| **Category** | Lightweight Spreadsheet UI Component |
| **Primary Use** | Simple, customizable spreadsheet for React |

## Description

React Spreadsheet is a lightweight, customizable spreadsheet component for React. It focuses on simplicity and flexibility, providing a clean API for building basic spreadsheet functionality with full control over cell rendering and editing.

## Core Capabilities

### 1. Basic Spreadsheet Features
- **Cell Editing**: Inline cell editing with custom editors
- **Cell Selection**: Single and multi-cell selection
- **Keyboard Navigation**: Arrow keys, Tab, Enter
- **Copy/Paste**: Standard clipboard operations
- **Row/Column Headers**: Customizable labels

### 2. Customization
- **Custom Cell Renderers**: Full control over cell display
- **Custom Cell Editors**: Custom input components
- **Custom Data Viewers**: Custom display components
- **Column/Row Labels**: Dynamic label generation
- **CSS Styling**: Standard CSS customization

### 3. Data Management
- **Matrix Structure**: 2D array data format
- **Cell Objects**: Extensible cell data structure
- **Controlled Component**: Full state control
- **Change Callbacks**: Cell and selection events

## Implementation in Demo

```typescript
import Spreadsheet, {
  CellBase,
  Matrix,
  DataEditorProps,
  DataViewerProps,
  Selection,
} from 'react-spreadsheet';

interface DropdownCell extends CellBase {
  dropdownOptions?: string[];  // Extended cell with dropdown
}

const ReactSpreadsheetDemo: React.FC = () => {
  const [data, setData] = useState<Matrix<DropdownCell>>([]);
  const [columnWidths, setColumnWidths] = useState<number[]>([]);

  return (
    <Spreadsheet
      data={data}
      onChange={setData}
      columnLabels={columnLabels}
      rowLabels={rowLabels}
      DataViewer={DropdownDataViewer}
      DataEditor={DropdownDataEditor}
      hideRowIndicators={hideRowIndicators}
      hideColumnIndicators={hideColumnIndicators}
    />
  );
};
```

## Data Structure

```typescript
// Basic cell structure
interface CellBase {
  value?: string | number | boolean | null;
  readOnly?: boolean;
  className?: string;
  DataViewer?: React.ComponentType<DataViewerProps<CellBase>>;
  DataEditor?: React.ComponentType<DataEditorProps<CellBase>>;
}

// Extended cell with dropdown support (custom)
interface DropdownCell extends CellBase {
  dropdownOptions?: string[];
}

// Data is a 2D matrix
type Matrix<T extends CellBase> = (T | undefined)[][];

// Example data
const data: Matrix<CellBase> = [
  [{ value: 'Name' }, { value: 'Age' }, { value: 'City' }],
  [{ value: 'Alice' }, { value: 30 }, { value: 'NYC' }],
  [{ value: 'Bob' }, { value: 25 }, { value: 'LA' }],
];
```

## Custom Components

### Custom Data Viewer (Display)

```typescript
const DropdownDataViewer: React.FC<DataViewerProps<DropdownCell>> = ({ cell }) => {
  const hasDropdown = cell?.dropdownOptions && cell.dropdownOptions.length > 0;
  return (
    <span className="rs-cell-viewer">
      <span className="rs-cell-value">{cell?.value ?? ''}</span>
      {hasDropdown && <span className="rs-dropdown-arrow">▾</span>}
    </span>
  );
};
```

### Custom Data Editor (Input)

```typescript
const DropdownDataEditor: React.FC<DataEditorProps<DropdownCell>> = ({
  cell,
  onChange,
  exitEditMode,
}) => {
  const options = cell?.dropdownOptions || [];
  const [search, setSearch] = useState('');
  
  const handleSelect = (opt: string) => {
    onChange({ ...cell!, value: opt } as CellBase);
    exitEditMode();
  };

  return (
    <div className="rs-dropdown-editor">
      <input
        type="text"
        placeholder="Search..."
        value={search}
        onChange={(e) => setSearch(e.target.value)}
      />
      <div className="rs-dropdown-list">
        {filtered.map((opt) => (
          <div key={opt} onClick={() => handleSelect(opt)}>
            {opt}
          </div>
        ))}
      </div>
    </div>
  );
};
```

## Component Props

| Prop | Type | Description |
|------|------|-------------|
| `data` | `Matrix<CellBase>` | 2D array of cell objects |
| `onChange` | `(data: Matrix) => void` | Data change callback |
| `columnLabels` | `string[]` | Column header labels |
| `rowLabels` | `string[]` | Row header labels |
| `DataViewer` | `ComponentType` | Custom cell display component |
| `DataEditor` | `ComponentType` | Custom cell editor component |
| `hideRowIndicators` | `boolean` | Hide row numbers |
| `hideColumnIndicators` | `boolean` | Hide column letters |
| `onSelect` | `(selection) => void` | Selection change callback |
| `selected` | `Selection` | Controlled selection |

## Column Width Calculation

```typescript
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
```

## Row Selection Implementation

```typescript
// State for selected row
const [selectedRow, setSelectedRow] = useState<number | null>(null);

// Handle selection changes
const handleSelection = (selection: Selection) => {
  if (selection && 'range' in selection) {
    setSelectedRow(selection.range.start.row);
  } else if (selection && 'rows' in selection) {
    // Entire row selected
    setSelectedRow(selection.rows[0]);
  }
};

// Delete selected row
const deleteSelectedRow = () => {
  if (selectedRow !== null) {
    const newData = data.filter((_, idx) => idx !== selectedRow);
    setData(newData);
    setSelectedRow(null);
  }
};
```

## Key Strengths

1. **Lightweight**: Small bundle size (~30KB)
2. **Simple API**: Easy to understand and implement
3. **Highly Customizable**: Full control over cell rendering
4. **React Native**: Idiomatic React patterns (controlled components)
5. **TypeScript Support**: Full type definitions
6. **Flexible Data**: Extensible cell object structure
7. **Minimal Dependencies**: Only React peer dependency

## Limitations

1. **No Built-in Features**: No toolbar, formula bar, or advanced Excel features
2. **No Formula Support**: Pure data display (formulas must be computed externally)
3. **No Virtual Scrolling**: May struggle with very large datasets (10K+ rows)
4. **Basic Styling**: Minimal default styles
5. **No Built-in Dropdowns**: Must implement custom editors

## CSS Customization

```css
/* Cell styling */
.Spreadsheet__cell {
  padding: 4px 8px;
  min-height: 32px;
}

/* Active cell */
.Spreadsheet__active-cell {
  border: 2px solid #0284c7;
}

/* Custom dropdown styling */
.rs-dropdown-editor {
  position: fixed;
  background: white;
  border: 1px solid #e2e8f0;
  border-radius: 8px;
  box-shadow: 0 4px 12px rgba(0,0,0,0.1);
}

.rs-dropdown-item {
  padding: 8px 12px;
  cursor: pointer;
}

.rs-dropdown-item:hover {
  background: #f1f5f9;
}
```

## Use Cases

| Use Case | Suitability |
|----------|-------------|
| Simple data entry grids | ⭐⭐⭐⭐⭐ Excellent |
| Customized spreadsheet UI | ⭐⭐⭐⭐⭐ Excellent |
| Small to medium datasets | ⭐⭐⭐⭐ Good |
| Full Excel replacement | ⭐⭐ Limited |
| Large datasets (10K+ rows) | ⭐⭐ Limited |
| Complex formulas | ❌ Not supported |

## Integration with ExcelJS

```typescript
// Convert ExcelJS data to react-spreadsheet Matrix
const convertToMatrix = (worksheet: ExcelJS.Worksheet): Matrix<DropdownCell> => {
  const matrix: Matrix<DropdownCell> = [];
  
  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const rowData: (DropdownCell | undefined)[] = [];
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      rowData[colNumber - 1] = {
        value: extractCellValue(cell),
        dropdownOptions: resolveDropdownOptions(cell),
      };
    });
    matrix[rowNumber - 1] = rowData;
  });
  
  return matrix;
};
```

## Performance Considerations

- Best for small to medium datasets (< 5,000 cells)
- No virtualization - all cells rendered in DOM
- Consider pagination for larger datasets
- Custom renderers add overhead per cell
- Memoize custom components when possible

## Comparison with Other Libraries

| Feature | React Spreadsheet | Fortune Sheet | React Data Grid |
|---------|------------------|---------------|-----------------|
| Bundle Size | ~30KB | ~500KB+ | ~200KB |
| Learning Curve | Low | High | Medium |
| Customization | High | Medium | High |
| Built-in Features | Low | High | Medium |
| Performance (large data) | Limited | Good | Excellent |
| Formula Support | No | Yes | No |

## Conclusion

React Spreadsheet is the ideal choice when you need a simple, lightweight spreadsheet component with maximum customization control. It's perfect for data entry forms, small datasets, and scenarios where you want to build your own spreadsheet features. For large datasets or full Excel functionality, consider Fortune Sheet or React Data Grid instead.
