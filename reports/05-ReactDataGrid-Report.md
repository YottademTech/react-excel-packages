# React Data Grid - Detailed Package Report

## Overview

| Property | Value |
|----------|-------|
| **Package Name** | `react-data-grid` |
| **Version** | ^7.0.0-beta.59 |
| **License** | MIT |
| **NPM Weekly Downloads** | 200K+ |
| **Category** | High-Performance Data Grid |
| **Primary Use** | Large dataset display with Excel-like editing |

## Description

React Data Grid is a feature-rich, high-performance data grid component for React. It focuses on efficient rendering of large datasets through virtualization while providing Excel-like editing capabilities, sorting, filtering, and row selection.

## Core Capabilities

### 1. Performance Features
- **Virtual Scrolling**: Renders only visible rows/columns
- **Large Dataset Support**: Handles 100K+ rows efficiently
- **Optimized Re-renders**: Smart update batching
- **Lazy Loading**: Load data on demand

### 2. Data Grid Features
- **Column Sorting**: Single and multi-column sorting
- **Column Resizing**: Drag-to-resize columns
- **Row Selection**: Checkbox selection with Select All
- **Cell Editing**: Inline editing with custom editors
- **Column Freezing**: Freeze left columns

### 3. Cell Customization
- **Custom Renderers**: Full control over cell display
- **Custom Editors**: Text, dropdown, date editors
- **Formatter Functions**: Transform display values
- **Cell Classes**: Dynamic styling

### 4. Built-in Components
- **SelectColumn**: Checkbox selection column
- **TextEditor**: Default text input editor
- **Custom Editor Pattern**: Easy editor creation

## Implementation in Demo

```typescript
import { DataGrid, Column, SortColumn, RenderEditCellProps, SelectColumn } from 'react-data-grid';
import 'react-data-grid/lib/styles.css';

interface Row {
  [key: string]: any;
  _rowIndex?: number;
}

const ReactDataGridDemo: React.FC = () => {
  const [rows, setRows] = useState<Row[]>([]);
  const [columns, setColumns] = useState<Column<Row>[]>([]);
  const [sortColumns, setSortColumns] = useState<readonly SortColumn[]>([]);
  const [selectedRows, setSelectedRows] = useState<ReadonlySet<number>>(new Set());

  return (
    <DataGrid
      columns={columns}
      rows={sortedRows}
      rowKeyGetter={rowKeyGetter}
      onRowsChange={setRows}
      sortColumns={sortColumns}
      onSortColumnsChange={setSortColumns}
      selectedRows={selectedRows}
      onSelectedRowsChange={setSelectedRows}
      className="rdg-light"
    />
  );
};
```

## Column Definition

```typescript
interface Column<R> {
  key: string;              // Unique column identifier
  name: string;             // Column header text
  width?: number | string;  // Column width (px or %)
  minWidth?: number;        // Minimum width
  maxWidth?: number;        // Maximum width
  resizable?: boolean;      // Allow resizing
  sortable?: boolean;       // Enable sorting
  frozen?: boolean;         // Freeze column
  renderCell?: (props) => ReactNode;       // Custom cell renderer
  renderEditCell?: (props) => ReactNode;   // Custom cell editor
  editable?: boolean | ((row) => boolean); // Enable editing
}

// Example column definitions
const columns: Column<Row>[] = [
  SelectColumn,  // Built-in checkbox column
  {
    key: 'name',
    name: 'Name',
    width: 200,
    sortable: true,
    resizable: true,
    editable: true,
    renderEditCell: TextEditor,
  },
  {
    key: 'status',
    name: 'Status',
    width: 150,
    renderEditCell: (props) => <DropdownEditor {...props} options={statusOptions} />,
  },
];
```

## Custom Dropdown Editor

```typescript
function DropdownEditor({
  row,
  column,
  onRowChange,
  onClose,
  options,
}: RenderEditCellProps<Row> & { options: string[] }) {
  const [search, setSearch] = useState('');
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    setTimeout(() => inputRef.current?.focus(), 0);
  }, []);

  const filtered = options.filter((opt) =>
    opt.toLowerCase().includes(search.toLowerCase())
  );

  const handleSelect = (value: string) => {
    onRowChange({ ...row, [column.key]: value }, true);
  };

  return (
    <div className="rdg-dropdown-editor">
      <input
        ref={inputRef}
        type="text"
        placeholder="Search..."
        value={search}
        onChange={(e) => setSearch(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Escape') onClose();
          if (e.key === 'Enter' && filtered.length === 1) {
            handleSelect(filtered[0]);
          }
        }}
      />
      <div className="rdg-dropdown-list">
        {filtered.map((opt) => (
          <div
            key={opt}
            className={`rdg-dropdown-item${row[column.key] === opt ? ' selected' : ''}`}
            onMouseDown={(e) => {
              e.preventDefault();
              handleSelect(opt);
            }}
          >
            {opt}
          </div>
        ))}
      </div>
    </div>
  );
}
```

## Sorting Implementation

```typescript
const [sortColumns, setSortColumns] = useState<readonly SortColumn[]>([]);

// Sort rows based on sort columns
const sortedRows = useMemo(() => {
  if (sortColumns.length === 0) return rows;

  return [...rows].sort((a, b) => {
    for (const sort of sortColumns) {
      const { columnKey, direction } = sort;
      const aVal = a[columnKey];
      const bVal = b[columnKey];

      if (aVal == null && bVal == null) continue;
      if (aVal == null) return direction === 'ASC' ? -1 : 1;
      if (bVal == null) return direction === 'ASC' ? 1 : -1;

      let comparison = 0;
      if (typeof aVal === 'number' && typeof bVal === 'number') {
        comparison = aVal - bVal;
      } else {
        comparison = String(aVal).localeCompare(String(bVal));
      }

      if (comparison !== 0) {
        return direction === 'ASC' ? comparison : -comparison;
      }
    }
    return 0;
  });
}, [rows, sortColumns]);
```

## Row Selection

```typescript
// Include SelectColumn as first column
const columns = [
  SelectColumn,
  // ... other columns
];

// State for selected rows
const [selectedRows, setSelectedRows] = useState<ReadonlySet<number>>(new Set());

// Row key getter for selection tracking
const rowKeyGetter = (row: Row) => row._rowIndex ?? 0;

// Add row index to each row
const rowsWithIndex = useMemo(() => 
  rows.map((row, idx) => ({ ...row, _rowIndex: idx })), 
  [rows]
);

// Delete selected rows
const deleteSelectedRows = () => {
  if (selectedRows.size === 0) return;
  const newRows = rows.filter((_, idx) => !selectedRows.has(idx));
  setRows(newRows);
  setSelectedRows(new Set());
};
```

## Component Props

| Prop | Type | Description |
|------|------|-------------|
| `columns` | `Column<R>[]` | Column definitions |
| `rows` | `R[]` | Row data array |
| `rowKeyGetter` | `(row) => Key` | Unique row identifier |
| `onRowsChange` | `(rows) => void` | Row data change callback |
| `sortColumns` | `SortColumn[]` | Current sort state |
| `onSortColumnsChange` | `(cols) => void` | Sort change callback |
| `selectedRows` | `Set<Key>` | Selected row keys |
| `onSelectedRowsChange` | `(set) => void` | Selection callback |
| `className` | `string` | CSS class name |
| `style` | `CSSProperties` | Inline styles |
| `rowHeight` | `number` | Row height in pixels |
| `headerRowHeight` | `number` | Header row height |

## Key Strengths

1. **Performance**: Handles 100K+ rows with virtualization
2. **Modern React**: Hooks-based, modern patterns
3. **TypeScript First**: Excellent type support
4. **Sorting Built-in**: Multi-column sorting support
5. **Row Selection**: Built-in checkbox column
6. **Flexible Editing**: Custom editor components
7. **Active Community**: Regular updates and support

## Limitations

1. **No Formula Support**: Data display only
2. **No Excel Export**: Requires additional library
3. **Beta Version**: Still in beta (v7.0.0-beta)
4. **Learning Curve**: Complex API for advanced features
5. **No Built-in Dropdowns**: Must implement custom editors

## CSS Customization

```css
/* Theme override */
.rdg-light {
  --rdg-header-background-color: #f8fafc;
  --rdg-row-hover-background-color: #f1f5f9;
  --rdg-selection-color: #0284c7;
  --rdg-border-color: #e2e8f0;
}

/* Custom editor styling */
.rdg-dropdown-editor {
  position: absolute;
  z-index: 100;
  background: white;
  border: 1px solid #e2e8f0;
  border-radius: 8px;
  box-shadow: 0 4px 12px rgba(0,0,0,0.1);
  min-width: 200px;
}

.rdg-dropdown-item {
  padding: 8px 12px;
  cursor: pointer;
}

.rdg-dropdown-item:hover {
  background: #f1f5f9;
}

.rdg-dropdown-item.selected {
  background: #e0f2fe;
  color: #0369a1;
}
```

## Use Cases

| Use Case | Suitability |
|----------|-------------|
| Large datasets (10K+ rows) | ⭐⭐⭐⭐⭐ Excellent |
| Data tables with sorting | ⭐⭐⭐⭐⭐ Excellent |
| Inline editing | ⭐⭐⭐⭐ Good |
| Row selection workflows | ⭐⭐⭐⭐⭐ Excellent |
| Full Excel replacement | ⭐⭐ Limited |
| Formulas and calculation | ❌ Not supported |

## Integration with ExcelJS

```typescript
// Parse Excel file and build columns/rows
const handleFileUpload = async (file: File) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());

  worksheet.eachSheet((ws) => {
    // Build columns from header row
    const headerRow = ws.getRow(1);
    const columns: Column<Row>[] = [SelectColumn];
    
    headerRow.eachCell((cell, colNumber) => {
      const header = String(cell.value || `Column ${colNumber}`);
      columns.push({
        key: header,
        name: header,
        sortable: true,
        resizable: true,
        editable: true,
        renderEditCell: hasDropdown(colNumber) 
          ? (props) => <DropdownEditor {...props} options={options} />
          : TextEditor,
      });
    });

    // Build rows from data
    const rows: Row[] = [];
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header
      const rowData: Row = {};
      row.eachCell((cell, colNumber) => {
        rowData[columns[colNumber].key] = cell.value;
      });
      rows.push(rowData);
    });
  });
};
```

## Performance Comparison

| Rows | React Data Grid | React Spreadsheet | Fortune Sheet |
|------|-----------------|-------------------|---------------|
| 100 | Instant | Instant | Instant |
| 1,000 | Instant | Smooth | Smooth |
| 10,000 | Smooth | Laggy | Smooth |
| 100,000 | Smooth | Crashes | Moderate |

## Conclusion

React Data Grid is the top choice for applications requiring high-performance data display with large datasets. Its virtual scrolling, built-in sorting, and row selection make it ideal for data-intensive applications. While it lacks spreadsheet-specific features like formulas, it excels at what it does: displaying and editing tabular data efficiently.
