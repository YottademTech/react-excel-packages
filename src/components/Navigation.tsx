import React from 'react';
import { Link, useLocation } from 'react-router-dom';
import './Navigation.css';

const Navigation: React.FC = () => {
  const location = useLocation();

  const navItems = [
    { path: '/', label: 'Home', description: 'Overview' },
    { path: '/xlsx', label: 'XLSX (SheetJS)', description: 'Apache 2.0' },
    { path: '/exceljs', label: 'ExcelJS', description: 'MIT' },
    { path: '/react-spreadsheet', label: 'React Spreadsheet', description: 'MIT' },
    { path: '/react-data-grid', label: 'React Data Grid', description: 'MIT' },
    { path: '/fortune-sheet', label: 'Fortune Sheet', description: 'MIT' },
    { path: '/revogrid', label: 'RevoGrid', description: 'MIT' },
    { path: '/jspreadsheet', label: 'jspreadsheet CE', description: 'MIT' },
  ];

  return (
    <nav className="navigation">
      <div className="nav-header">
        <h1>Excel Packages Demo</h1>
        <p>React + TypeScript</p>
      </div>
      <ul className="nav-list">
        {navItems.map((item) => (
          <li key={item.path}>
            <Link
              to={item.path}
              className={location.pathname === item.path ? 'active' : ''}
            >
              <span className="nav-label">{item.label}</span>
              <span className="nav-license">{item.description}</span>
            </Link>
          </li>
        ))}
      </ul>
    </nav>
  );
};

export default Navigation;
