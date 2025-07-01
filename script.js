let excelData = [];

window.onload = function() {
    fetch('KEYE_Pending_Orders_Report_SIMA.xlsx')
        .then(response => {
            if (!response.ok) throw new Error('Network response was not ok');
            return response.arrayBuffer();
        })
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            excelData = XLSX.utils.sheet_to_json(worksheet);
            document.getElementById('results').innerHTML = '<p style="color:green;">Excel data loaded. Enter a style code to search.</p>';
            // Optional: console.log for debugging
            console.log('Excel data loaded:', excelData);
        })
        .catch(err => {
            document.getElementById('results').innerHTML = '<p style="color:red;">Failed to load Excel file: ' + err + '</p>';
            console.error('Excel load error:', err);
        });
};

document.addEventListener('DOMContentLoaded', function(){
    document.getElementById('searchBtn').addEventListener('click', function() {
        const styleCode = document.getElementById('styleCodeInput').value.trim();
        displayResults(styleCode);
    });
});

function displayResults(styleCode) {
    if (!excelData.length) {
        document.getElementById('results').innerHTML = '<p>Data not loaded yet. Wait a few seconds and try again.</p>';
        return;
    }
    if (!styleCode) {
        document.getElementById('results').innerHTML = '<p>Please enter a style code.</p>';
        return;
    }

    // Update these to EXACT column names in your Excel
    const columnsToShow = [
        'First Allocation Date',
        'Confirmed Allocation Date',
        'Pending Order Qty',
        'Under Packing Qty',
        'Allocation Available Qty',
        'Open Qty'
    ];

    // Update this to the exact column header for style code
    const STYLE_CODE_COL = 'Style Code';

    const filtered = excelData.filter(row => String(row[STYLE_CODE_COL]).toLowerCase() === styleCode.toLowerCase());

    if (filtered.length === 0) {
        document.getElementById('results').innerHTML = '<p>No results found for this style code.</p>';
        return;
    }

    let table = '<table border="1"><thead><tr>';
    columnsToShow.forEach(col => table += `<th>${col}</th>`);
    table += '</tr></thead><tbody>';

    filtered.forEach(row => {
        table += '<tr>';
        columnsToShow.forEach(col => table += `<td>${row[col] || ''}</td>`);
        table += '</tr>';
    });
    table += '</tbody></table>';

    document.getElementById('results').innerHTML = table;
}
