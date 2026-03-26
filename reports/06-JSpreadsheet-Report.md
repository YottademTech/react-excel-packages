# JSpreadsheet CE - Detailed Package Report

## Overview

| Property | Value |
|----------|-------|
| **Package Name** | `jspreadsheet-ce` |
| **Version** | ^5.0.4 |
| **License** | MIT |
| **NPM Weekly Downloads** | 25K+ |
| **Category** | JavaScript Spreadsheet Component |
| **Primary Use** | Interactive spreadsheet with Excel-like features |

## Description

JSpreadsheet CE (Community Edition) is a lightweight vanilla JavaScript spreadsheet component that can be integrated with React. It provides a clean spreadsheet interface with built-in features like dropdowns, formulas, and keyboard navigation. The CE version is the free, MIT-licensed community fork.

## Core Capabilities

### 1. Spreadsheet Features
- **Cell Editing**: Double-click to edit cells
- **Dropdowns**: Built-in autocomplete dropdowns
- **Search**: Find values across the sheet
- **Undo/Redo**: Full history support
- **Copy/Paste**: Clipboard integration

### 2. Column Types
- **Text**: Standard text input
- **Dropdown**: Autocomplete selection
- **Calendar**: Date picker
- **Checkbox**: Boolean toggle
- **Numeric**: Number formatting
- **Color**: Color picker
- **Image**: Image cells

### 3. Data Management
- **Multiple Worksheets**: Tab-based sheets
- **Table Overflow**: Scrollable tables
- **Pagination**: Built-in pagination
- **Lazy Loading**: Load data on demand
- **Data Export**: CSV, JSON export

### 4. Customization
- **Column Configuration**: Width, type, title, alignment
- **Row Configuration**: Height, row data
- **Toolbar**: Optional toolbar
- **Context Menu**: Right-click menus
- **Styling**: CSS customization

## Implementation in Demo

```typescript
import jspreadsheet from 'jspreadsheet-ce';
import 'jspreadsheet-ce/dist/jspreadsheet.css';
import 'jsuites/dist/jsuites.css';

interface ColumnConfig {
  type?: string;
  title: string;
  width?: number;
  source?: string[];  // Dropdown options
}

interface SheetData {
  name: string;
  data: any[][];
  columns: ColumnConfig[];
}

const JSpreadsheetDemo: React.FC = () => {
  const containerRef = useRef<HTMLDivElement>(null);
  const jspreadsheetRef = useRef<any>(null);

  useEffect(() => {
    if (!containerRef.current) return;

    // Create jspreadsheet instance
    jspreadsheetRef.current = jspreadsheet(containerRef.current, {
      worksheets: [{
        data: sheetData,
        columns: columns,
        minDimensions: [10, 20],
        tableOverflow: true,
        tableWidth: '100%',
        tableHeight: '450px',
        search: true,
      }],
    });

    return () => {
      jspreadsheetRef.current?.destroy();
    };
  }, [sheetData, columns]);

  return <div ref={containerRef} />;
};
```

## Column Configuration

```typescript
interface ColumnConfig {
  title: string;           // Column header
  type?: string;           // 'text', 'dropdown', 'calendar', 'checkbox', etc.
  width?: number;          // Column width in pixels
  align?: string;          // 'left', 'center', 'right'
  source?: string[];       // Dropdown options array
  autocomplete?: boolean;  // Enable autocomplete for dropdown
  options?: object;        // Additional options
  readOnly?: boolean;      // Make column read-only
  mask?: string;           // Input mask
  decimal?: string;        // Decimal separator
}

// Example column definitions
const columns: ColumnConfig[] = [
  {
    title: 'Name',
    type: 'text',
    width: 200,
  },
  {
    title: 'Status',
    type: 'dropdown',
    width: 150,
    source: ['Active', 'Inactive', 'Pending'],
    autocomplete: true,
  },
  {
    title: 'Date',
    type: 'calendar',
    width: 120,
    options: { format: 'YYYY-MM-DD' },
  },
  {
    title: 'Active',
    type: 'checkbox',
    width: 80,
  },
];
```

## Worksheet Configuration

```typescript
const worksheetConfig = {
  data: [
    ['John', 'Active', '2024-01-15', true],
    ['Jane', 'Pending', '2024-02-20', false],
  ],
  columns: columns,
  minDimensions: [columns.length, 10],  // Min cols, rows
  tableOverflow: true,
  tableWidth: '100%',
  tableHeight: '450px',
  search: true,                          // Enable search
  pagination: 20,                        // Rows per page
  allowInsertRow: true,
  allowDeleteRow: true,
  allowInsertColumn: false,
  columnSorting: true,
  columnDrag: true,
  columnResize: true,
  rowDrag: true,
  rowResize: true,
};
```

## API Methods

```typescript
const js = jspreadsheetRef.current;

// Get data
const data = js.getData();              // All data
const value = js.getValue('A1');        // Cell value

// Set data
js.setData(newData);                    // Set all data
js.setValue('A1', 'New Value');         // Set cell value

// Rows
js.insertRow();                         // Insert at end
js.insertRow(2);                        // Insert at index 2
js.deleteRow(3);                        // Delete row 3

// Columns
js.insertColumn();                      // Insert at end
js.deleteColumn(2);                     // Delete column 2

// Selection
const selected = js.getSelected();      // Get selected cells

// Search
js.search('query');                     // Search cells

// Undo/Redo
js.undo();
js.redo();

// Download
js.download();                          // Download as CSV
```

## Event Handlers

```typescript
const config = {
  // ... other config
  onchange: (instance, cell, x, y, value) => {
    console.log('Cell changed:', x, y, value);
  },
  onselection: (instance, x1, y1, x2, y2) => {
    console.log('Selection:', x1, y1, x2, y2);
  },
  oninsertrow: (instance, rowNumber, numOfRows, insertBefore) => {
    console.log('Row inserted:', rowNumber);
  },
  ondeleterow: (instance, rowNumber, numOfRows) => {
    console.log('Row deleted:', rowNumber);
  },
  onsort: (instance, column, order) => {
    console.log('Sorted column:', column, order);
  },
};
```

## Key Strengths

1. **Vanilla JS**: No framework dependency, works everywhere
2. **Built-in Dropdowns**: Native autocomplete support
3. **Column Types**: Various built-in types (calendar, checkbox, etc.)
4. **Search**: Built-in search functionality
5. **Lightweight**: Small bundle size (~50KB)
6. **Easy Setup**: Simple API to get started
7. **Undo/Redo**: Full history support

## Limitations

1. **No Virtual Scrolling**: Table overflow instead of virtualization
2. **jQuery Heritage**: API patterns reflect jQuery origins
3. **React Integration**: Requires ref-based wrapper
4. **Formula Support**: Basic formulas only in CE
5. **Styling**: Some CSS quirks to work around
6. **Pro Features**: Advanced features need commercial license

## CSS Customization

```css
/* Table styling */
.jexcel tbody td {
  padding: 6px 10px;
  font-size: 13px;
}

/* Header styling */
.jexcel thead td {
  background: #f8fafc;
  font-weight: 600;
  color: #475569;
}

/* Selected cells */
.jexcel tbody td.highlight {
  background: #e0f2fe;
}

/* Dropdown styling */
.jexcel_dropdown {
  border: 1px solid #e2e8f0;
  border-radius: 6px;
  box-shadow: 0 4px 12px rgba(0,0,0,0.1);
}

/* Search box */
.jexcel_search {
  border: 1px solid #e2e8f0;
  border-radius: 4px;
  padding: 6px 10px;
}
```

## Use Cases

| Use Case | Suitability |
|----------|-------------|
| Simple data entry forms | ⭐⭐⭐⭐⭐ Excellent |
| Dropdown-heavy interfaces | ⭐⭐⭐⭐⭐ Excellent |
| Date entry applications | ⭐⭐⭐⭐ Good |
| Small to medium datasets | ⭐⭐⭐⭐ Good |
| Large datasets (10K+) | ⭐⭐ Limited |
| Complex formulas | ⭐⭐ Limited |

## Integration with ExcelJS

```typescript
// Parse Excel and create jspreadsheet config
const handleFileUpload = async (file: File) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());

  workbook.eachSheet((worksheet) => {
    const columns: ColumnConfig[] = [];
    const data: any[][] = [];

    // Build columns from header row
    const headerRow = worksheet.getRow(1);
    for (let c = 1; c <= worksheet.columnCount; c++) {
      const headerValue = String(headerRow.getCell(c).value || `Column ${c}`);
      columns.push({
        title: headerValue,
        width: Math.max(headerValue.length * 10 + 20, 100),
        type: determineColumnType(c), // Custom logic
        source: getDropdownOptions(c), // From data validation
        autocomplete: true,
      });
    }

    // Build data from rows 2+
    for (let r = 2; r <= worksheet.rowCount; r++) {
      const row = worksheet.getRow(r);
      const rowData: any[] = [];
      for (let c = 1; c <= worksheet.columnCount; c++) {
        rowData.push(row.getCell(c).value ?? '');
      }
      data.push(rowData);
    }

    // Create jspreadsheet
    jspreadsheet(container, {
      worksheets: [{
        data,
        columns,
        tableOverflow: true,
        search: true,
      }],
    });
  });
};
```

## Pro vs CE Comparison

| Feature | CE (Free) | Pro |
|---------|-----------|-----|
| Basic editing | ✅ | ✅ |
| Dropdowns | ✅ | ✅ |
| Calendar | ✅ | ✅ |
| Search | ✅ | ✅ |
| Formulas | Basic | Advanced |
| Filters | ❌ | ✅ |
| Frozen rows/cols | ❌ | ✅ |
| Merged cells | Basic | Advanced |
| Nested headers | ❌ | ✅ |
| Import/Export | Basic | Advanced |

## Performance Considerations

- Good for datasets up to ~5,000 rows
- Table overflow provides smooth scrolling
- No DOM virtualization (all cells rendered)
- Consider pagination for larger datasets
- Dropdown autocomplete is efficient

## Conclusion

JSpreadsheet CE is an excellent choice for applications needing built-in dropdown support and basic spreadsheet features without the overhead of larger libraries. Its vanilla JS implementation makes it framework-agnostic, though React integration requires manual ref management. Best suited for data entry forms with moderate dataset sizes.
