# RevoGrid - Detailed Package Report

## Overview

| Property | Value |
|----------|-------|
| **Package Name** | `@revolist/react-datagrid` |
| **Core Package** | `@revolist/revogrid` |
| **Version** | ^4.21.0 |
| **License** | MIT |
| **NPM Weekly Downloads** | 15K+ |
| **Category** | High-Performance Virtual Data Grid |
| **Primary Use** | Large dataset management with advanced features |

## Description

RevoGrid is a powerful, high-performance virtual data grid built with modern web technologies. It leverages Web Components (Stencil.js) for the core and provides wrappers for React, Vue, and Angular. Known for excellent performance with large datasets and extensible architecture.

## Core Capabilities

### 1. Performance Features
- **Virtual Scrolling**: Renders only visible viewport
- **100K+ Rows**: Handles massive datasets smoothly
- **Smart Caching**: Intelligent row/column caching
- **Lazy Loading**: Load data on scroll

### 2. Data Grid Features
- **Column Sorting**: Single and multi-column
- **Column Filters**: Built-in filter row
- **Column Resize**: Drag-to-resize
- **Column Reorder**: Drag-and-drop columns
- **Range Selection**: Multi-cell selection
- **Auto-size Columns**: Fit content width

### 3. Cell Editing
- **Inline Editing**: Click/double-click to edit
- **Column Types**: Custom editor plugins
- **Select/Dropdown**: Built-in select column type
- **Validation**: Input validation support

### 4. Advanced Features
- **Grouping**: Row grouping
- **Pinned Rows/Columns**: Freeze areas
- **Export**: Data export capabilities
- **Themes**: Multiple built-in themes
- **Plugin Architecture**: Extensible via plugins

## Implementation in Demo

```typescript
import { RevoGrid } from '@revolist/react-datagrid';
import type { ColumnRegular } from '@revolist/revogrid';
import SelectColumnType from '@revolist/revogrid-column-select';

interface SheetData {
  name: string;
  columns: ColumnRegular[];
  source: Record<string, any>[];
  dropdowns: { [key: string]: string[] };
}

const RevoGridDemo: React.FC = () => {
  const [sheets, setSheets] = useState<SheetData[]>([]);
  const [activeSheetIndex, setActiveSheetIndex] = useState(0);

  // Column types including custom select
  const columnTypes = useMemo(() => ({
    select: new SelectColumnType(),
  }), []);

  return (
    <RevoGrid
      columns={columnsWithDropdowns}
      source={activeSheet.source}
      columnTypes={columnTypes}
      theme="compact"
      resize={true}
      autoSizeColumn={true}
      filter={true}
      range={true}
      readonly={false}
    />
  );
};
```

## Column Definition

```typescript
import type { ColumnRegular } from '@revolist/revogrid';

interface ColumnRegular {
  prop: string;              // Property key in row data
  name: string;              // Column header text
  size?: number;             // Column width
  minSize?: number;          // Minimum width
  maxSize?: number;          // Maximum width
  sortable?: boolean;        // Enable sorting
  filter?: boolean;          // Enable filtering
  pin?: 'colPinStart' | 'colPinEnd';  // Pin position
  cellTemplate?: Function;   // Custom cell template
  columnType?: string;       // Column type reference
  source?: string[];         // Options for select type
  readonly?: boolean;        // Disable editing
}

// Example column definitions
const columns: ColumnRegular[] = [
  {
    prop: 'id',
    name: 'ID',
    size: 80,
    sortable: true,
    readonly: true,
  },
  {
    prop: 'name',
    name: 'Name',
    size: 200,
    sortable: true,
    filter: true,
  },
  {
    prop: 'status',
    name: 'Status',
    size: 150,
    columnType: 'select',      // Use select plugin
    source: ['Active', 'Inactive', 'Pending'],
  },
  {
    prop: 'date',
    name: 'Created Date',
    size: 120,
    sortable: true,
  },
];
```

## Select Column Plugin

```typescript
import SelectColumnType from '@revolist/revogrid-column-select';

// Register column type
const columnTypes = {
  select: new SelectColumnType(),
};

// Apply to column
const columns = [
  {
    prop: 'category',
    name: 'Category',
    columnType: 'select',
    source: ['Electronics', 'Clothing', 'Food', 'Books'],
  },
];

// Use in RevoGrid
<RevoGrid
  columns={columns}
  source={data}
  columnTypes={columnTypes}
/>
```

## Row Data Format

```typescript
// Row data is an array of objects
const source: Record<string, any>[] = [
  {
    id: 1,
    name: 'John Doe',
    status: 'Active',
    date: '2024-01-15',
  },
  {
    id: 2,
    name: 'Jane Smith',
    status: 'Pending',
    date: '2024-02-20',
  },
];

// Dynamic column props
const columnsFromHeaders = headers.map((header, idx) => ({
  prop: `col_${idx + 1}`,
  name: header,
  size: Math.max(header.length * 10 + 20, 100),
  sortable: true,
}));
```

## Component Props

| Prop | Type | Description |
|------|------|-------------|
| `columns` | `ColumnRegular[]` | Column definitions |
| `source` | `object[]` | Row data array |
| `columnTypes` | `object` | Column type plugins |
| `theme` | `string` | Theme name ('compact', 'material', etc.) |
| `resize` | `boolean` | Enable column resize |
| `autoSizeColumn` | `boolean` | Auto-fit column widths |
| `filter` | `boolean` | Enable filter row |
| `range` | `boolean` | Enable range selection |
| `readonly` | `boolean` | Disable all editing |
| `rowSize` | `number` | Row height |
| `rowClass` | `function` | Dynamic row classes |
| `cellClass` | `function` | Dynamic cell classes |

## Event Handling

```typescript
<RevoGrid
  columns={columns}
  source={source}
  onBeforeEdit={(event) => {
    // Cancel edit for certain cells
    if (event.detail.prop === 'id') {
      event.preventDefault();
    }
  }}
  onAfterEdit={(event) => {
    const { val, oldVal, prop, model } = event.detail;
    console.log(`Changed ${prop}: ${oldVal} → ${val}`);
  }}
  onBeforeRange={(event) => {
    // Handle range selection
    console.log('Range selected:', event.detail);
  }}
  onBeforeFilter={(event) => {
    // Custom filter logic
    console.log('Filter:', event.detail);
  }}
  onBeforeSort={(event) => {
    // Custom sort logic
    console.log('Sort:', event.detail);
  }}
/>
```

## Key Strengths

1. **Outstanding Performance**: Handles 100K+ rows effortlessly
2. **Virtual Scrolling**: Both rows and columns virtualized
3. **Web Components**: Framework-agnostic core
4. **Plugin Architecture**: Highly extensible
5. **Built-in Filtering**: Column filters included
6. **Range Selection**: Excel-like selection
7. **Modern Stack**: Uses modern web technologies
8. **Multiple Themes**: Ready-to-use themes

## Limitations

1. **Web Component Overhead**: Stencil.js wrapper adds complexity
2. **Documentation**: Some docs are sparse
3. **Learning Curve**: Plugin system requires understanding
4. **React Patterns**: Less idiomatic than pure React libraries
5. **No Formula Support**: Data display only

## CSS Customization

```css
/* Theme variables */
revo-grid {
  --rv-gap: 1px;
  --rv-header-bg: #f8fafc;
  --rv-header-color: #334155;
  --rv-border-color: #e2e8f0;
  --rv-row-height: 36px;
  --rv-hover-bg: #f1f5f9;
  --rv-selection-color: #0284c7;
}

/* Cell styling */
.revogrid-cell {
  padding: 8px 12px;
  font-size: 13px;
}

/* Header styling */
.revogrid-header-cell {
  font-weight: 600;
  background: var(--rv-header-bg);
}

/* Filter input */
.revogrid-filter-input {
  border: 1px solid #e2e8f0;
  border-radius: 4px;
  padding: 4px 8px;
}
```

## Use Cases

| Use Case | Suitability |
|----------|-------------|
| Large datasets (100K+ rows) | ⭐⭐⭐⭐⭐ Excellent |
| Data analysis dashboards | ⭐⭐⭐⭐⭐ Excellent |
| Filtering and sorting | ⭐⭐⭐⭐⭐ Excellent |
| Dropdown columns | ⭐⭐⭐⭐ Good (with plugin) |
| Full Excel replacement | ⭐⭐ Limited |
| Formula calculations | ❌ Not supported |

## Integration with ExcelJS

```typescript
// Parse Excel and create RevoGrid config
const handleFileSelect = async (file: File) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());

  workbook.eachSheet((worksheet) => {
    const columns: ColumnRegular[] = [];
    const source: Record<string, any>[] = [];
    const dropdowns: { [key: string]: string[] } = {};

    // Build columns from header row
    const headerRow = worksheet.getRow(1);
    for (let c = 1; c <= worksheet.columnCount; c++) {
      const headerText = String(headerRow.getCell(c).value || `Column ${c}`);
      columns.push({
        prop: `col_${c}`,
        name: headerText,
        size: Math.max(headerText.length * 10 + 20, 100),
        sortable: true,
      });
    }

    // Build source data from rows 2+
    for (let r = 2; r <= worksheet.rowCount; r++) {
      const row = worksheet.getRow(r);
      const rowData: Record<string, any> = {};
      for (let c = 1; c <= worksheet.columnCount; c++) {
        rowData[`col_${c}`] = row.getCell(c).value ?? '';
      }
      source.push(rowData);
    }

    // Parse data validations for dropdowns
    const validations = (worksheet as any).dataValidations?.model;
    if (validations) {
      Object.entries(validations).forEach(([address, rule]: [string, any]) => {
        if (rule.type === 'list') {
          const options = resolveDropdownOptions(rule.formulae[0]);
          const colIndex = parseColumnIndex(address);
          dropdowns[`col_${colIndex + 1}`] = options;
        }
      });
    }

    // Apply dropdown column types
    const finalColumns = columns.map((col) => {
      if (dropdowns[col.prop]) {
        return {
          ...col,
          columnType: 'select',
          source: dropdowns[col.prop],
        };
      }
      return col;
    });
  });
};
```

## Performance Comparison

| Feature | RevoGrid | React Data Grid | Fortune Sheet |
|---------|----------|-----------------|---------------|
| 100K rows | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐ | ⭐⭐⭐ |
| Virtual scroll | Rows + Cols | Rows only | Partial |
| Initial render | Fast | Fast | Moderate |
| Memory usage | Low | Low | Higher |
| Re-render speed | Excellent | Good | Good |

## Built-in Themes

- `default` - Standard theme
- `compact` - Reduced spacing
- `material` - Material Design style
- `darkMaterial` - Dark mode Material

```typescript
<RevoGrid
  theme="compact"  // or "material", "darkMaterial"
/>
```

## Conclusion

RevoGrid is the premier choice for applications requiring maximum performance with large datasets. Its virtual scrolling for both rows and columns, combined with the plugin architecture, makes it highly capable and extensible. While it requires the select column plugin for dropdowns and has a steeper learning curve than simpler alternatives, its performance characteristics make it unmatched for data-intensive applications.
