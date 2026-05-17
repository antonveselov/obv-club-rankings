const XLSX = require('xlsx');

function analyzeExcel() {
    const workbook = XLSX.readFile('VRC Erw Teilnahme Sep-Dez 2025.xlsx');
    console.log("Sheet Names:", workbook.SheetNames);
    
    const sheetName = workbook.SheetNames[1] || workbook.SheetNames[0];
    console.log("Analyzing Sheet:", sheetName);
    
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    console.log("Rows with data (first 200):");
    data.slice(0, 200).forEach((row, i) => {
        if (row.length > 0 && row.some(cell => cell !== null && cell !== '')) {
            console.log(`Row ${i}:`, row);
        }
    });
}

analyzeExcel();
