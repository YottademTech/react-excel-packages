import React, { useState, useCallback, useMemo } from 'react';
import { DataGrid, Column, SortColumn } from 'react-data-grid';
import * as XLSX from 'xlsx';
import FileUpload from '../components/FileUpload';
import './DemoPage.css';
import 'react-data-grid/lib/styles.css';

interface Row {
  [key: string]: any;
}

const ReactDataGridDemo: React.FC = () => {
  const [rows, setRows] = useState<Row[]>([]);
  const [columns, setColumns] = useState<Column<Row>[]>([]);
  const [fileName, setFileName] = useState('');
  const [sortColumns, setSortColumns] = useState<readonly SortColumn[]>([]);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [activeSheet, setActiveSheet] = useState(0);
  const [allSheetsData, setAllSheetsData] = useState<{ rows: Row[]; columns: Column<Row>[] }[]>([]);

  const sortedRows = useMemo(() => {
    if (sortColumns.length === 0) return rows;

    return [...rows].sort((a, b) => {
      for (const sort of sortColumns) {
        const { columnKey, direction } = sort;
        const aVal = a[columnKey];
        const bVal = b[columnKey];

        // Handle undefined/null
        if (aVal == null && bVal == null) continue;
        if (aVal == null) return direction === 'ASC' ? -1 : 1;
        if (bVal == null) return direction === 'ASC' ? 1 : -1;

        // Compare
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

  const handleFileUpload = useCallback((file: File) => {
    setFileName(file.name);
    const reader = new FileReader();

    reader.onload = (e) => {
      const arrayBuffer = e.target?.result as ArrayBuffer;
      const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });

      const sheets: { rows: Row[]; columns: Column<Row>[] }[] = [];
      const names: string[] = [];

      workbook.SheetNames.forEach((sheetName) => {
        names.push(sheetName);
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: false, dateNF: 'yyyy-mm-dd' }) as Row[];

        // Create columns from first row keys
        const cols: Column<Row>[] = [];
        if (jsonData.length > 0) {
          Object.keys(jsonData[0]).forEach((key) => {
            cols.push({
              key,
              name: key,
              resizable: true,
              sortable: true,
              width: 150,
              renderCell: ({ row }) => {
                const value = row[key];
                // Format numbers
                if (typeof value === 'number') {
                  return <span style={{ color: '#1d4ed8' }}>{value.toLocaleString()}</span>;
                }
                return <span>{value}</span>;
              },
            });
          });
        }

        sheets.push({ rows: jsonData, columns: cols });
      });

      setSheetNames(names);
      setAllSheetsData(sheets);
      if (sheets.length > 0) {
        setRows(sheets[0].rows);
        setColumns(sheets[0].columns);
      }
      setActiveSheet(0);
      setSortColumns([]);
    };

    reader.readAsArrayBuffer(file);
  }, []);

  const handleSheetChange = (index: number) => {
    setActiveSheet(index);
    setRows(allSheetsData[index]?.rows || []);
    setColumns(allSheetsData[index]?.columns || []);
    setSortColumns([]);
  };

  const handleRowsChange = (newRows: Row[]) => {
    setRows(newRows);
  };

  const createSampleData = () => {
    const sampleRows: Row[] = [
      { id: 1, name: 'John Doe', department: 'Engineering', salary: 75000, startDate: '2020-01-15', status: 'Active' },
      { id: 2, name: 'Jane Smith', department: 'Marketing', salary: 65000, startDate: '2021-03-22', status: 'Active' },
      { id: 3, name: 'Bob Johnson', department: 'Sales', salary: 85000, startDate: '2018-07-01', status: 'Active' },
      { id: 4, name: 'Alice Brown', department: 'HR', salary: 70000, startDate: '2019-11-10', status: 'On Leave' },
      { id: 5, name: 'Charlie Wilson', department: 'Engineering', salary: 72000, startDate: '2022-02-28', status: 'Active' },
      { id: 6, name: 'Diana Miller', department: 'Finance', salary: 78000, startDate: '2020-06-15', status: 'Active' },
      { id: 7, name: 'Edward Davis', department: 'Engineering', salary: 82000, startDate: '2019-04-20', status: 'Active' },
      { id: 8, name: 'Fiona Garcia', department: 'Marketing', salary: 68000, startDate: '2021-09-01', status: 'Remote' },
    ];

    const sampleColumns: Column<Row>[] = [
      { key: 'id', name: 'ID', width: 60, sortable: true },
      { 
        key: 'name', 
        name: 'Name', 
        width: 150, 
        sortable: true,
        resizable: true,
      },
      { 
        key: 'department', 
        name: 'Department', 
        width: 120, 
        sortable: true,
        resizable: true,
      },
      { 
        key: 'salary', 
        name: 'Salary', 
        width: 100, 
        sortable: true,
        resizable: true,
        renderCell: ({ row }) => (
          <span style={{ color: '#1d4ed8' }}>${row.salary.toLocaleString()}</span>
        ),
      },
      { 
        key: 'startDate', 
        name: 'Start Date', 
        width: 120, 
        sortable: true,
        resizable: true,
        renderCell: ({ row }) => (
          <span style={{ color: '#7c3aed' }}>{row.startDate}</span>
        ),
      },
      { 
        key: 'status', 
        name: 'Status', 
        width: 100, 
        sortable: true,
        resizable: true,
        renderCell: ({ row }) => {
          const colors: { [key: string]: string } = {
            'Active': '#10b981',
            'On Leave': '#f59e0b',
            'Remote': '#3b82f6',
          };
          return (
            <span 
              style={{ 
                background: colors[row.status] || '#6b7280',
                color: 'white',
                padding: '2px 8px',
                borderRadius: '4px',
                fontSize: '0.75rem',
              }}
            >
              {row.status}
            </span>
          );
        },
      },
    ];

    setRows(sampleRows);
    setColumns(sampleColumns);
    setFileName('sample-data.xlsx');
    setSheetNames(['Employees']);
    setAllSheetsData([{ rows: sampleRows, columns: sampleColumns }]);
    setActiveSheet(0);
  };

  return (
    <div className="demo-page">
      <header className="demo-header">
        <h1>React Data Grid</h1>
        <span className="license-tag">MIT</span>
      </header>

      <div className="demo-info">
        <p>
          React Data Grid is a feature-rich grid component for React. It provides excellent performance
          with virtual scrolling and supports sorting, filtering, cell editing, and column resizing.
        </p>
        <div className="features-list">
          <span className="feature">✓ Virtual Scrolling</span>
          <span className="feature">✓ Column Sorting</span>
          <span className="feature">✓ Column Resizing</span>
          <span className="feature">✓ Cell Editing</span>
          <span className="feature">✓ Row Selection</span>
        </div>
      </div>

      <div className="upload-section">
        <FileUpload onFileUpload={handleFileUpload} />
        <div className="sample-data-option">
          <p>or</p>
          <button className="sample-data-btn" onClick={createSampleData}>
            Load Sample Data
          </button>
        </div>
      </div>

      {rows.length > 0 && (
        <div className="result-section">
          <div className="result-header">
            <h2>
              <span className="file-icon">📊</span>
              {fileName || 'Data Grid'}
              <span className="editable-badge">Sortable</span>
            </h2>
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
          </div>

          <div className="rdg-wrapper">
            <DataGrid
              columns={columns}
              rows={sortedRows}
              onRowsChange={handleRowsChange}
              sortColumns={sortColumns}
              onSortColumnsChange={setSortColumns}
              className="rdg-light"
              style={{ height: '100%' }}
            />
          </div>

          <div className="data-info">
            <h4>Interactive Features</h4>
            <p>• Click column headers to sort (click again to reverse)</p>
            <p>• Drag column edges to resize</p>
            <p>• {rows.length} rows × {columns.length} columns</p>
            <p>• Virtual scrolling for performance with large datasets</p>
          </div>
        </div>
      )}

      <div className="code-example">
        <div className="code-header">Usage Example</div>
        <pre>
          <code>{`import DataGrid from 'react-data-grid';
import 'react-data-grid/lib/styles.css';

const columns = [
  { key: 'id', name: 'ID' },
  { key: 'name', name: 'Name', sortable: true },
  { key: 'salary', name: 'Salary', sortable: true },
];

const rows = [
  { id: 1, name: 'John', salary: 75000 },
  { id: 2, name: 'Jane', salary: 65000 },
];

function MyGrid() {
  return (
    <DataGrid
      columns={columns}
      rows={rows}
      sortColumns={sortColumns}
      onSortColumnsChange={setSortColumns}
    />
  );
}`}</code>
        </pre>
      </div>
    </div>
  );
};

export default ReactDataGridDemo;
