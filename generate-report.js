const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  AlignmentType,
  ShadingType,
  PageBreak,
} = require('docx');
const fs = require('fs');
const path = require('path');

async function generateReport() {
  const doc = new Document({
    creator: 'Excel Packages Research Team',
    title: 'Fortune Sheet & XLSX Package Research Report',
    description: 'Comprehensive analysis of Fortune Sheet and XLSX packages for React applications',
    styles: {
      paragraphStyles: [
        {
          id: 'Normal',
          name: 'Normal',
          run: {
            font: 'Calibri',
            size: 24,
          },
          paragraph: {
            spacing: { after: 120 },
          },
        },
      ],
    },
    sections: [
      {
        properties: {},
        children: [
          // Title
          new Paragraph({
            children: [
              new TextRun({
                text: 'Excel Spreadsheet Packages for React Applications',
                bold: true,
                size: 56,
                color: '2563EB',
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: 'Technical Research Report: Fortune Sheet & XLSX (SheetJS)',
                size: 28,
                color: '64748B',
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 },
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: `Prepared: ${new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })}`,
                size: 22,
                color: '94A3B8',
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 600 },
          }),

          // Executive Summary
          new Paragraph({
            text: 'Executive Summary',
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 400, after: 200 },
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: 'This report presents a comprehensive analysis of two leading open-source packages for handling Excel spreadsheets in React applications: ',
              }),
              new TextRun({
                text: 'XLSX (SheetJS)',
                bold: true,
              }),
              new TextRun({
                text: ' for file parsing and ',
              }),
              new TextRun({
                text: 'Fortune Sheet',
                bold: true,
              }),
              new TextRun({
                text: ' for interactive spreadsheet rendering. Together, these packages provide a complete solution for uploading, displaying, editing, and exporting Excel files in web applications.',
              }),
            ],
            spacing: { after: 200 },
          }),

          // Section 1: XLSX (SheetJS)
          new Paragraph({
            text: '1. XLSX (SheetJS) - The Parser',
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 400, after: 200 },
          }),

          new Paragraph({
            text: '1.1 Overview',
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 200, after: 100 },
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: 'XLSX, also known as SheetJS, is the most widely adopted spreadsheet parsing library in the JavaScript ecosystem. With over ',
              }),
              new TextRun({
                text: '2.5 million weekly downloads',
                bold: true,
              }),
              new TextRun({
                text: ' on npm, it has become the de facto standard for reading and writing Excel files in JavaScript applications.',
              }),
            ],
            spacing: { after: 200 },
          }),

          new Paragraph({
            text: '1.2 License Information',
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 200, after: 100 },
          }),

          // License Table for XLSX
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Attribute', bold: true })] })],
                    shading: { fill: 'E2E8F0', type: ShadingType.SOLID },
                    width: { size: 30, type: WidthType.PERCENTAGE },
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Details', bold: true })] })],
                    shading: { fill: 'E2E8F0', type: ShadingType.SOLID },
                    width: { size: 70, type: WidthType.PERCENTAGE },
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('License Type')] }),
                  new TableCell({ children: [new Paragraph('Apache License 2.0')] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('Source Code')] }),
                  new TableCell({ children: [new Paragraph('Open Source (Community Edition)')] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('Commercial Use')] }),
                  new TableCell({ children: [new Paragraph('Permitted without royalties')] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('Modification')] }),
                  new TableCell({ children: [new Paragraph('Permitted with attribution')] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('Distribution')] }),
                  new TableCell({ children: [new Paragraph('Permitted')] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('Repository')] }),
                  new TableCell({ children: [new Paragraph('https://github.com/SheetJS/sheetjs')] }),
                ],
              }),
            ],
          }),

          new Paragraph({
            text: '',
            spacing: { after: 200 },
          }),

          new Paragraph({
            text: '1.3 Key Capabilities',
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 200, after: 100 },
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '• File Format Support: ', bold: true }),
              new TextRun('XLSX, XLS, XLSB, XLSM, CSV, TSV, ODS, and more'),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '• Formula Parsing: ', bold: true }),
              new TextRun('Reads and interprets Excel formulas'),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '• Data Validation: ', bold: true }),
              new TextRun('Extracts dropdown lists, number constraints, and validation rules'),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '• Date Handling: ', bold: true }),
              new TextRun('Proper parsing of Excel date serial numbers'),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '• Multi-Sheet Support: ', bold: true }),
              new TextRun('Handles workbooks with multiple worksheets'),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '• Cell Styling: ', bold: true }),
              new TextRun('Basic style information extraction (Pro version has full support)'),
            ],
            spacing: { after: 200 },
          }),

          new Paragraph({
            text: '1.4 Limitations',
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 200, after: 100 },
          }),
          new Paragraph({
            text: 'The community (free) version has some limitations:',
            spacing: { after: 100 },
          }),
          new Paragraph({ text: '• Limited cell styling extraction (colors, borders)' }),
          new Paragraph({ text: '• No image extraction support' }),
          new Paragraph({ text: '• Data validation extraction requires additional processing' }),
          new Paragraph({
            text: '• Pro version ($) required for advanced features like streaming large files',
            spacing: { after: 200 },
          }),

          // Section 2: Fortune Sheet
          new Paragraph({
            text: '2. Fortune Sheet - The Renderer',
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 400, after: 200 },
          }),

          new Paragraph({
            text: '2.1 Overview',
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 200, after: 100 },
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: 'Fortune Sheet is a modern, TypeScript-based spreadsheet component for React. It is a fork of the popular Luckysheet project, rebuilt with better TypeScript support, React integration, and ongoing maintenance. It provides a ',
              }),
              new TextRun({
                text: 'complete Excel-like user interface',
                bold: true,
              }),
              new TextRun({
                text: ' including toolbars, formula bars, cell editing, and context menus.',
              }),
            ],
            spacing: { after: 200 },
          }),

          new Paragraph({
            text: '2.2 License Information',
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 200, after: 100 },
          }),

          // License Table for Fortune Sheet
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Attribute', bold: true })] })],
                    shading: { fill: 'E2E8F0', type: ShadingType.SOLID },
                    width: { size: 30, type: WidthType.PERCENTAGE },
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Details', bold: true })] })],
                    shading: { fill: 'E2E8F0', type: ShadingType.SOLID },
                    width: { size: 70, type: WidthType.PERCENTAGE },
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('License Type')] }),
                  new TableCell({ children: [new Paragraph('MIT License')] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('Source Code')] }),
                  new TableCell({ children: [new Paragraph('Fully Open Source')] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('Commercial Use')] }),
                  new TableCell({ children: [new Paragraph('Permitted without restrictions')] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('Modification')] }),
                  new TableCell({ children: [new Paragraph('Permitted')] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('Distribution')] }),
                  new TableCell({ children: [new Paragraph('Permitted with license inclusion')] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('Repository')] }),
                  new TableCell({ children: [new Paragraph('https://github.com/ruilisi/fortune-sheet')] }),
                ],
              }),
            ],
          }),

          new Paragraph({
            text: '',
            spacing: { after: 200 },
          }),

          new Paragraph({
            text: '2.3 Key Capabilities',
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 200, after: 100 },
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '• Full Excel UI: ', bold: true }),
              new TextRun('Toolbar, formula bar, sheet tabs, context menus'),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '• Formula Support: ', bold: true }),
              new TextRun('400+ built-in Excel formulas'),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '• Cell Formatting: ', bold: true }),
              new TextRun('Colors, borders, fonts, number formats'),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '• Data Validation: ', bold: true }),
              new TextRun('Dropdown lists, number constraints, custom rules'),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '• Cell Merging: ', bold: true }),
              new TextRun('Merge and unmerge cells'),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '• Freeze Panes: ', bold: true }),
              new TextRun('Lock rows and columns'),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '• Conditional Formatting: ', bold: true }),
              new TextRun('Apply styles based on cell values'),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '• Configurable UI: ', bold: true }),
              new TextRun('Hide/show toolbar, formula bar, sheet tabs'),
            ],
            spacing: { after: 200 },
          }),

          // Section 3: Combined Usage
          new Paragraph({
            text: '3. Combined Architecture: XLSX + Fortune Sheet',
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 400, after: 200 },
          }),

          new Paragraph({
            text: '3.1 Why Combine These Packages?',
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 200, after: 100 },
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: 'Neither package alone provides a complete solution. XLSX excels at parsing Excel files but has no UI components. Fortune Sheet provides a rich UI but cannot directly read .xlsx files. ',
              }),
              new TextRun({
                text: 'Together, they form a complete Excel handling solution.',
                bold: true,
              }),
            ],
            spacing: { after: 200 },
          }),

          // Architecture Table
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Component', bold: true })] })],
                    shading: { fill: '2563EB', type: ShadingType.SOLID },
                    width: { size: 25, type: WidthType.PERCENTAGE },
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Package', bold: true })] })],
                    shading: { fill: '2563EB', type: ShadingType.SOLID },
                    width: { size: 25, type: WidthType.PERCENTAGE },
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Responsibility', bold: true })] })],
                    shading: { fill: '2563EB', type: ShadingType.SOLID },
                    width: { size: 50, type: WidthType.PERCENTAGE },
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('File Parser')] }),
                  new TableCell({ children: [new Paragraph('XLSX / ExcelJS')] }),
                  new TableCell({ children: [new Paragraph('Read .xlsx files, extract cells, formulas, data validations, and styling')] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('Data Transformer')] }),
                  new TableCell({ children: [new Paragraph('Custom Code')] }),
                  new TableCell({ children: [new Paragraph('Convert parsed data to Fortune Sheet\'s expected format')] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('UI Renderer')] }),
                  new TableCell({ children: [new Paragraph('Fortune Sheet')] }),
                  new TableCell({ children: [new Paragraph('Display interactive spreadsheet with editing capabilities')] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('File Exporter')] }),
                  new TableCell({ children: [new Paragraph('XLSX')] }),
                  new TableCell({ children: [new Paragraph('Convert edited data back to .xlsx for download')] }),
                ],
              }),
            ],
          }),

          new Paragraph({
            text: '',
            spacing: { after: 200 },
          }),

          new Paragraph({
            text: '3.2 Data Flow',
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 200, after: 100 },
          }),
          new Paragraph({
            text: 'The typical workflow when using both packages together:',
            spacing: { after: 100 },
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '1. Upload: ', bold: true }),
              new TextRun('User uploads an Excel file (.xlsx, .xls)'),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '2. Parse: ', bold: true }),
              new TextRun('XLSX reads the file and extracts all data'),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '3. Transform: ', bold: true }),
              new TextRun('Data is converted to Fortune Sheet format (celldata, dataVerification)'),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '4. Display: ', bold: true }),
              new TextRun('Fortune Sheet renders the interactive spreadsheet'),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '5. Edit: ', bold: true }),
              new TextRun('User makes changes using Fortune Sheet\'s UI'),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '6. Export: ', bold: true }),
              new TextRun('XLSX writes the modified data back to an Excel file'),
            ],
            spacing: { after: 200 },
          }),

          // Section 4: Comparison with Alternatives
          new Paragraph({
            text: '4. Comparison with Commercial Alternatives',
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 400, after: 200 },
          }),

          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Package', bold: true })] })],
                    shading: { fill: 'E2E8F0', type: ShadingType.SOLID },
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'License', bold: true })] })],
                    shading: { fill: 'E2E8F0', type: ShadingType.SOLID },
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Cost', bold: true })] })],
                    shading: { fill: 'E2E8F0', type: ShadingType.SOLID },
                  }),
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Recommendation', bold: true })] })],
                    shading: { fill: 'E2E8F0', type: ShadingType.SOLID },
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('XLSX + Fortune Sheet')] }),
                  new TableCell({ children: [new Paragraph('Apache 2.0 + MIT')] }),
                  new TableCell({ children: [new Paragraph('Free')] }),
                  new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: '✓ RECOMMENDED', bold: true, color: '10B981' })] })] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('Handsontable')] }),
                  new TableCell({ children: [new Paragraph('Commercial')] }),
                  new TableCell({ children: [new Paragraph('$1,199+/dev/year')] }),
                  new TableCell({ children: [new Paragraph('Skip - License cost')] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('AG Grid Enterprise')] }),
                  new TableCell({ children: [new Paragraph('Commercial')] }),
                  new TableCell({ children: [new Paragraph('$999+/dev/year')] }),
                  new TableCell({ children: [new Paragraph('Skip - License cost')] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('SpreadJS')] }),
                  new TableCell({ children: [new Paragraph('Commercial')] }),
                  new TableCell({ children: [new Paragraph('$999+/dev')] }),
                  new TableCell({ children: [new Paragraph('Skip - Proprietary')] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('Syncfusion Spreadsheet')] }),
                  new TableCell({ children: [new Paragraph('Commercial')] }),
                  new TableCell({ children: [new Paragraph('$995+/dev/year')] }),
                  new TableCell({ children: [new Paragraph('Skip - License cost')] }),
                ],
              }),
            ],
          }),

          new Paragraph({
            text: '',
            spacing: { after: 200 },
          }),

          // Section 5: Implementation
          new Paragraph({
            text: '5. Implementation Considerations',
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 400, after: 200 },
          }),

          new Paragraph({
            text: '5.1 Bundle Size',
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 200, after: 100 },
          }),
          new Paragraph({ text: '• XLSX (minified): ~300 KB' }),
          new Paragraph({ text: '• Fortune Sheet: ~500 KB' }),
          new Paragraph({
            text: '• Total: ~800 KB (acceptable for enterprise applications)',
            spacing: { after: 200 },
          }),

          new Paragraph({
            text: '5.2 Browser Support',
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 200, after: 100 },
          }),
          new Paragraph({ text: '• Chrome 60+' }),
          new Paragraph({ text: '• Firefox 55+' }),
          new Paragraph({ text: '• Safari 12+' }),
          new Paragraph({ text: '• Edge 79+' }),
          new Paragraph({
            text: '• IE11 (with polyfills)',
            spacing: { after: 200 },
          }),

          new Paragraph({
            text: '5.3 Key Configuration Options',
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 200, after: 100 },
          }),
          new Paragraph({ text: 'Fortune Sheet can be customized to match your application needs:' }),
          new Paragraph({ text: '• showToolbar: Show/hide the formatting toolbar' }),
          new Paragraph({ text: '• showFormulaBar: Show/hide the formula input bar' }),
          new Paragraph({ text: '• showSheetTabs: Show/hide sheet tabs for multi-sheet workbooks' }),
          new Paragraph({ text: '• allowEdit: Enable/disable cell editing' }),
          new Paragraph({
            text: '• data: The spreadsheet data in Fortune Sheet format',
            spacing: { after: 200 },
          }),

          // Section 6: Conclusion
          new Paragraph({
            text: '6. Conclusion & Recommendation',
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 400, after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: 'The combination of XLSX and Fortune Sheet provides an enterprise-grade solution for Excel file handling in React applications. Key benefits include:',
              }),
            ],
            spacing: { after: 100 },
          }),

          new Paragraph({
            children: [
              new TextRun({ text: '✓ Cost Effective: ', bold: true, color: '10B981' }),
              new TextRun('Both packages are free and open source with permissive licenses'),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '✓ Full Featured: ', bold: true, color: '10B981' }),
              new TextRun('Complete Excel-like experience including dropdowns, formulas, and formatting'),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '✓ Actively Maintained: ', bold: true, color: '10B981' }),
              new TextRun('Both projects have active communities and regular updates'),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '✓ Production Ready: ', bold: true, color: '10B981' }),
              new TextRun('Used in production by numerous organizations worldwide'),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: '✓ No Vendor Lock-in: ', bold: true, color: '10B981' }),
              new TextRun('Open source means freedom to modify and maintain independently'),
            ],
            spacing: { after: 300 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: 'RECOMMENDATION: ',
                bold: true,
                size: 28,
                color: '2563EB',
              }),
              new TextRun({
                text: 'Proceed with XLSX + Fortune Sheet as the primary solution for Excel file handling in the application.',
                size: 26,
              }),
            ],
            spacing: { after: 400 },
            shading: { fill: 'EFF6FF', type: ShadingType.SOLID },
          }),

          // Appendix
          new Paragraph({
            text: 'Appendix A: Package Links & Resources',
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 400, after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({ text: 'XLSX (SheetJS)', bold: true }),
            ],
          }),
          new Paragraph({ text: '• npm: https://www.npmjs.com/package/xlsx' }),
          new Paragraph({ text: '• GitHub: https://github.com/SheetJS/sheetjs' }),
          new Paragraph({
            text: '• Documentation: https://docs.sheetjs.com/',
            spacing: { after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({ text: 'Fortune Sheet', bold: true }),
            ],
          }),
          new Paragraph({ text: '• npm: https://www.npmjs.com/package/@fortune-sheet/react' }),
          new Paragraph({ text: '• GitHub: https://github.com/ruilisi/fortune-sheet' }),
          new Paragraph({
            text: '• Documentation: https://ruilisi.github.io/fortune-sheet-docs/',
            spacing: { after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({ text: 'ExcelJS (Alternative Parser)', bold: true }),
            ],
          }),
          new Paragraph({ text: '• npm: https://www.npmjs.com/package/exceljs' }),
          new Paragraph({ text: '• GitHub: https://github.com/exceljs/exceljs' }),
          new Paragraph({ text: '• License: MIT' }),
          new Paragraph({
            text: '• Note: Better for extracting cell styles and data validations',
            spacing: { after: 400 },
          }),

          // Footer
          new Paragraph({
            children: [
              new TextRun({
                text: '— End of Report —',
                italics: true,
                color: '94A3B8',
              }),
            ],
            alignment: AlignmentType.CENTER,
          }),
        ],
      },
    ],
  });

  // Generate the document
  const buffer = await Packer.toBuffer(doc);
  const outputPath = path.join(__dirname, 'Excel_Packages_Research_Report.docx');
  fs.writeFileSync(outputPath, buffer);
  
  console.log(`\n✓ Report generated successfully!`);
  console.log(`  Location: ${outputPath}`);
  console.log(`\nReport Contents:`);
  console.log(`  1. Executive Summary`);
  console.log(`  2. XLSX (SheetJS) - Overview, License, Capabilities`);
  console.log(`  3. Fortune Sheet - Overview, License, Capabilities`);
  console.log(`  4. Combined Architecture & Data Flow`);
  console.log(`  5. Comparison with Commercial Alternatives`);
  console.log(`  6. Implementation Considerations`);
  console.log(`  7. Conclusion & Recommendation`);
  console.log(`  8. Appendix: Package Links & Resources`);
}

generateReport().catch(console.error);
