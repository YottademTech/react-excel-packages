import React from 'react';
import './Home.css';

const Home: React.FC = () => {
  const packages = [
    {
      name: 'XLSX (SheetJS)',
      license: 'Apache 2.0',
      description: 'The most popular spreadsheet parser and writer. Reads and writes Excel files with full support for formulas, styles, and data validation.',
      features: ['Read/Write XLSX, XLS, CSV', 'Formula support', 'Data validation (dropdowns)', 'Date handling', 'Multiple sheets'],
      npmWeekly: '2.5M+',
    },
    {
      name: 'ExcelJS',
      license: 'MIT',
      description: 'Read, manipulate and write spreadsheet data in XLSX and JSON. Great for server-side Excel generation with rich styling.',
      features: ['Rich cell styling', 'Images support', 'Data validation', 'Conditional formatting', 'Streaming for large files'],
      npmWeekly: '800K+',
    },
    {
      name: 'React Spreadsheet',
      license: 'MIT',
      description: 'Simple, customizable spreadsheet component for React. Lightweight and easy to integrate.',
      features: ['Editable cells', 'Custom cell renderers', 'Formula support', 'Copy/paste', 'Keyboard navigation'],
      npmWeekly: '50K+',
    },
    {
      name: 'React Data Grid',
      license: 'MIT',
      description: 'Feature-rich and customizable data grid for React. Excel-like grid editing experience.',
      features: ['Virtual scrolling', 'Column resizing', 'Sorting & filtering', 'Cell editing', 'Row selection'],
      npmWeekly: '200K+',
    },
    {
      name: 'Fortune Sheet',
      license: 'MIT',
      description: 'An Excel-like spreadsheet component for React. Fork of Luckysheet with TypeScript support.',
      features: ['Full Excel UI', 'Formulas', 'Charts', 'Pivot tables', 'Conditional formatting'],
      npmWeekly: '10K+',
    },
  ];

  return (
    <div className="home">
      <header className="home-header">
        <h1>React Excel Packages Comparison</h1>
        <p>
          This demo app showcases different open-source Excel/spreadsheet packages for React.
          All packages shown here have <strong>permissive licenses</strong> (MIT or Apache 2.0).
        </p>
      </header>

      <section className="packages-grid">
        {packages.map((pkg) => (
          <div key={pkg.name} className="package-card">
            <div className="package-header">
              <h2>{pkg.name}</h2>
              <span className="license-badge">{pkg.license}</span>
            </div>
            <p className="package-description">{pkg.description}</p>
            <div className="package-features">
              <h4>Key Features:</h4>
              <ul>
                {pkg.features.map((feature) => (
                  <li key={feature}>{feature}</li>
                ))}
              </ul>
            </div>
            <div className="package-stats">
              <span>Weekly Downloads: {pkg.npmWeekly}</span>
            </div>
          </div>
        ))}
      </section>

      <section className="skipped-packages">
        <h2>Packages Skipped (Non-Permissive Licenses)</h2>
        <ul>
          <li>
            <strong>Handsontable</strong> - Requires commercial license for commercial use
          </li>
          <li>
            <strong>AG Grid Enterprise</strong> - Enterprise features require commercial license
          </li>
          <li>
            <strong>SpreadJS</strong> - Commercial license only
          </li>
          <li>
            <strong>Syncfusion Spreadsheet</strong> - Commercial license
          </li>
        </ul>
      </section>
    </div>
  );
};

export default Home;
