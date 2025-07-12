import React, { useState } from 'react';
import * as XLSX from 'xlsx';

// This fillDown is now simplified because the grid construction will be more accurate
const fillDown = (data) => {
  if (!data || !data.length) return [];

  const filledData = [];
  let prevRow = []; // Initialize prevRow

  for (let i = 0; i < data.length; i++) {
    const currentRow = data[i] || []; // Ensure it's an array
    const newRow = [];
    for (let j = 0; j < currentRow.length; j++) { // Iterate up to current row's length
      let cellValue = currentRow[j];
      if (cellValue === undefined || cellValue === null || cellValue === '') {
        // Only fill down if there's a previous value at the same index
        cellValue = (prevRow[j] !== undefined && prevRow[j] !== null) ? prevRow[j] : '';
      }
      newRow.push(cellValue);
    }
    filledData.push(newRow);
    prevRow = newRow;
  }
  return filledData;
};

const ExcelReader = () => {
  const [workbook, setWorkbook] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  // We'll store headers and rows separately to manage the complex structure
  const [headers, setHeaders] = useState([]);
  const [rowData, setRowData] = useState([]);

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = (event) => {
      const arrayBuffer = event.target.result;
      const wb = XLSX.read(arrayBuffer, { type: 'array' });

      setWorkbook(wb);
      setSheetNames(wb.SheetNames);

      if (wb.SheetNames.length > 0) {
        const firstSheetName = wb.SheetNames[0];
        setSelectedSheet(firstSheetName);
        loadSheetData(wb, firstSheetName);
      } else {
        setSelectedSheet('');
        setHeaders([]);
        setRowData([]);
      }
    };

    reader.readAsArrayBuffer(file);
  };

  const loadSheetData = (wb, sheetName) => {
    const sheet = wb.Sheets[sheetName];

    if (!sheet || !sheet['!ref']) {
      console.warn(`Sheet "${sheetName}" is empty or invalid.`);
      setHeaders([]);
      setRowData([]);
      return;
    }

    const range = XLSX.utils.decode_range(sheet['!ref']);
    const maxRow = range.e.r; // Last row index
    const maxCol = range.e.c; // Last column index

    // Create an empty grid based on the full range
    const grid = [];
    for (let r = 0; r <= maxRow; r++) {
      grid.push(Array(maxCol + 1).fill('')); // +1 because columns are 0-indexed
    }

    // Populate the grid with actual cell values
    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cell_address = { c: C, r: R };
        const cell_ref = XLSX.utils.encode_cell(cell_address);
        const cell = sheet[cell_ref];

        // Get the formatted value. If cell doesn't exist or is empty, it will be '' from fill.
        // We use v (raw value) or w (formatted text) depending on preference.
        // For display, 'w' (formatted text) is usually better.
        const cellValue = cell ? (cell.w !== undefined ? cell.w : cell.v) : '';

        grid[R][C] = cellValue;
      }
    }

    // Your Excel sheet's structure:
    // Row 0 (Excel Row 1): Contains job titles starting from column C.
    // Row 1 (Excel Row 2): Contains "Learning and Development" in column B, and data starting from C.
    // This implies a complex header structure.

    // Let's manually define how to extract headers and data based on your image.
    // This part is CUSTOMIZED to your specific Excel layout!
    const extractedHeaders = [];
    if (grid.length > 0) {
      // The main column headers are in Excel Row 1 (grid[0]) starting from column index 2 (C)
      for (let c = 2; c <= maxCol; c++) { // Start from column 'C' (index 2)
        extractedHeaders.push(grid[0][c]);
      }
    }

    const extractedRowData = [];
    // Start from Excel Row 2 (grid[1])
    for (let r = 1; r <= maxRow; r++) {
      const row = [];
      // Column A and B act as row headers/categories
      // Column B is the "Learning and Development" or "Team Collaboration" label
      // The actual data starts from Column C (index 2)
      row.push(grid[r][1]); // This is Excel Column B (e.g., "Learning and Development")

      for (let c = 2; c <= maxCol; c++) { // Data columns from C onwards
        row.push(grid[r][c]);
      }
      extractedRowData.push(row);
    }
    
    // Now apply fillDown to the extractedRowData, specifically for column-wise filling if needed
    // The fillDown function should ensure data consistency within the processed rows
    const processedRowData = fillDown(extractedRowData);

    setHeaders(extractedHeaders);
    setRowData(processedRowData);
  };

  const handleSheetChange = (e) => {
    const name = e.target.value;
    setSelectedSheet(name);
    if (workbook) {
      loadSheetData(workbook, name);
    }
  };

  return (
    <div style={{ padding: 20 }}>
      <h2>Excel File Viewer</h2>

      <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />

      {sheetNames.length > 1 && (
        <div style={{ marginTop: 15 }}>
          <label>Select Sheet:&nbsp;</label>
          <select value={selectedSheet} onChange={handleSheetChange}>
            {sheetNames.map((name, idx) => (
              <option key={idx} value={name}>
                {name}
              </option>
            ))}
          </select>
        </div>
      )}

      {rowData.length > 0 && (
        <table border="1" cellPadding="10" style={{ marginTop: 20 }}>
          <thead>
            <tr>
              {/* Add an empty header for the "Learning and Development" column */}
              <th></th>
              {headers.map((header, idx) => (
                <th key={idx}>{header}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rowData.map((row, i) => (
              <tr key={i}>
                {row.map((cell, j) => (
                  <td key={j}>{cell}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      )}
      {headers.length === 0 && rowData.length === 0 && selectedSheet && workbook && (
          <p style={{ marginTop: 20 }}>No data found for the selected sheet, or the sheet is empty/invalid.</p>
      )}
    </div>
  );
};

export default ExcelReader;
