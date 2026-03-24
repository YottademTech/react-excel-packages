const ExcelJS = require('exceljs');
const path = require('path');

async function generateSampleExcel() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Excel Demo App';
  workbook.created = new Date();

  const worksheet = workbook.addWorksheet('Employee Records');

  // Define columns with headers
  worksheet.columns = [
    { header: 'Employee ID', key: 'id', width: 12 },
    { header: 'Full Name', key: 'name', width: 20 },
    { header: 'Department', key: 'department', width: 15 },
    { header: 'Status', key: 'status', width: 12 },
    { header: 'Hire Date', key: 'hireDate', width: 14 },
    { header: 'Salary', key: 'salary', width: 12 },
    { header: 'Rating', key: 'rating', width: 10 },
    { header: 'Active', key: 'active', width: 10 },
  ];

  // Style header row
  worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  worksheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF2563EB' },
  };
  worksheet.getRow(1).alignment = { horizontal: 'center' };

  // Add sample data
  const data = [
    { id: 1001, name: 'John Smith', department: 'Engineering', status: 'Full-Time', hireDate: new Date('2020-03-15'), salary: 75000, rating: 4, active: 'Yes' },
    { id: 1002, name: 'Sarah Johnson', department: 'Marketing', status: 'Full-Time', hireDate: new Date('2019-07-22'), salary: 68000, rating: 5, active: 'Yes' },
    { id: 1003, name: 'Michael Brown', department: 'Sales', status: 'Part-Time', hireDate: new Date('2021-01-10'), salary: 45000, rating: 3, active: 'Yes' },
    { id: 1004, name: 'Emily Davis', department: 'HR', status: 'Contract', hireDate: new Date('2022-05-01'), salary: 55000, rating: 4, active: 'No' },
    { id: 1005, name: 'David Wilson', department: 'Engineering', status: 'Full-Time', hireDate: new Date('2018-11-30'), salary: 92000, rating: 5, active: 'Yes' },
    { id: 1006, name: 'Lisa Anderson', department: 'Finance', status: 'Full-Time', hireDate: new Date('2020-08-14'), salary: 78000, rating: 4, active: 'Yes' },
    { id: 1007, name: 'James Taylor', department: 'Sales', status: 'Part-Time', hireDate: new Date('2023-02-28'), salary: 42000, rating: 3, active: 'Yes' },
    { id: 1008, name: 'Jennifer Martinez', department: 'Marketing', status: 'Contract', hireDate: new Date('2021-09-15'), salary: 52000, rating: 4, active: 'No' },
    { id: 1009, name: 'Robert Garcia', department: 'Engineering', status: 'Full-Time', hireDate: new Date('2019-04-20'), salary: 85000, rating: 5, active: 'Yes' },
    { id: 1010, name: 'Amanda Thomas', department: 'HR', status: 'Full-Time', hireDate: new Date('2022-12-01'), salary: 62000, rating: 3, active: 'Yes' },
  ];

  data.forEach(row => worksheet.addRow(row));

  // Format date column
  worksheet.getColumn('hireDate').numFmt = 'yyyy-mm-dd';
  
  // Format salary column as currency
  worksheet.getColumn('salary').numFmt = '"$"#,##0';

  // Add data validation - Department dropdown (Column C, rows 2-11)
  for (let row = 2; row <= 11; row++) {
    worksheet.getCell(`C${row}`).dataValidation = {
      type: 'list',
      allowBlank: false,
      formulae: ['"Engineering,Marketing,Sales,HR,Finance,Operations"'],
      showDropDown: true,
    };
  }

  // Add data validation - Status dropdown (Column D, rows 2-11)
  for (let row = 2; row <= 11; row++) {
    worksheet.getCell(`D${row}`).dataValidation = {
      type: 'list',
      allowBlank: false,
      formulae: ['"Full-Time,Part-Time,Contract,Intern"'],
      showDropDown: true,
    };
  }

  // Add data validation - Rating (1-5 number) (Column G, rows 2-11)
  for (let row = 2; row <= 11; row++) {
    worksheet.getCell(`G${row}`).dataValidation = {
      type: 'whole',
      operator: 'between',
      allowBlank: false,
      formulae: [1, 5],
      showErrorMessage: true,
      errorTitle: 'Invalid Rating',
      error: 'Rating must be between 1 and 5',
    };
  }

  // Add data validation - Active Yes/No dropdown (Column H, rows 2-11)
  for (let row = 2; row <= 11; row++) {
    worksheet.getCell(`H${row}`).dataValidation = {
      type: 'list',
      allowBlank: false,
      formulae: ['"Yes,No"'],
      showDropDown: true,
    };
  }

  // Add conditional formatting for Status column
  worksheet.addConditionalFormatting({
    ref: 'D2:D11',
    rules: [
      {
        type: 'containsText',
        operator: 'containsText',
        text: 'Full-Time',
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FF10B981' } } },
      },
    ],
  });

  // Add alternating row colors
  for (let row = 2; row <= 11; row++) {
    if (row % 2 === 0) {
      worksheet.getRow(row).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF1F5F9' },
      };
    }
  }

  // Add borders to all cells
  const borderStyle = {
    top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
  };

  for (let row = 1; row <= 11; row++) {
    for (let col = 1; col <= 8; col++) {
      worksheet.getCell(row, col).border = borderStyle;
    }
  }

  // Center align certain columns
  worksheet.getColumn('id').alignment = { horizontal: 'center' };
  worksheet.getColumn('rating').alignment = { horizontal: 'center' };
  worksheet.getColumn('active').alignment = { horizontal: 'center' };
  worksheet.getColumn('hireDate').alignment = { horizontal: 'center' };

  // Save the file
  const outputPath = path.join(__dirname, 'sample-data.xlsx');
  await workbook.xlsx.writeFile(outputPath);
  console.log(`Excel file created: ${outputPath}`);
  console.log('\nFile contains:');
  console.log('- Column A: Employee ID (Numbers)');
  console.log('- Column B: Full Name (Strings)');
  console.log('- Column C: Department (Dropdown: Engineering, Marketing, Sales, HR, Finance, Operations)');
  console.log('- Column D: Status (Dropdown: Full-Time, Part-Time, Contract, Intern)');
  console.log('- Column E: Hire Date (Dates formatted as yyyy-mm-dd)');
  console.log('- Column F: Salary (Currency format)');
  console.log('- Column G: Rating (Number validation 1-5)');
  console.log('- Column H: Active (Dropdown: Yes, No)');
}

generateSampleExcel().catch(console.error);
