import React from 'react';
import { BrowserRouter as Router, Routes, Route } from 'react-router-dom';
import Navigation from './components/Navigation';
import Home from './pages/Home';
import XlsxDemo from './pages/XlsxDemo';
import ExcelJsDemo from './pages/ExcelJsDemo';
import ReactSpreadsheetDemo from './pages/ReactSpreadsheetDemo';
import ReactDataGridDemo from './pages/ReactDataGridDemo';
import FortuneSheetDemo from './pages/FortuneSheetDemo';
import './App.css';

function App() {
  return (
    <Router>
      <div className="app">
        <Navigation />
        <main className="main-content">
          <Routes>
            <Route path="/" element={<Home />} />
            <Route path="/xlsx" element={<XlsxDemo />} />
            <Route path="/exceljs" element={<ExcelJsDemo />} />
            <Route path="/react-spreadsheet" element={<ReactSpreadsheetDemo />} />
            <Route path="/react-data-grid" element={<ReactDataGridDemo />} />
            <Route path="/fortune-sheet" element={<FortuneSheetDemo />} />
          </Routes>
        </main>
      </div>
    </Router>
  );
}

export default App;
