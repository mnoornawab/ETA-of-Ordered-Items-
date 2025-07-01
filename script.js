let excelData = [];

window.onload = function() {
    // Load and parse the Excel file
    fetch('KEYE - Pending Orders Report - SIMA.xlsx')
        .then(response => response.arrayBuffer())
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            // Convert to JSON
            excelData = XLSX.utils.sheet_to_json(worksheet);
        })
        .catch(err => {
            document.getElementById('results').innerHTML = '<p style="color:red;">Failed to load Excel file.</p>';
        });
};

// Search logic
document.addEventListener('DOMContentLoaded', function(){
    document.getElementById('searchBtn').addEventListener('click', function() {
        const styleCode = document.getElementById('styleCodeInput').value.trim();
        displayResults(styleCode);
    });
});

function displayResults(styleCode) {
    if (!excelData.length) {
        document.getElementById('results').innerHTML = '<p>Data not loaded yet.</p>';
        return;
    }
    if (!styleCode) {
        document.getElementById('results').innerHTML = '<p>Please enter a style code.</p>';
        return;
    }

    // Adjust the column names as they appear in your Excel file
    const columnsToShow = [
        'First Allocation Date',
        'Confirmed Allocation Date',
        'Pending Order Qty',
        'Under Packing Qty',
        'Allocation Available Qty',
        'Open Qty'
    ];

    // Replace 'Style Code' below if your Excel column uses a different name
    const filtered = excelData.filter(row => String(row['Style Code']).toLowerCase() === styleCode.toLowerCase());

    if (filtered.length === 0) {
        document.getElementById('results').innerHTML = '<p>No results found for this style code.</p>';
        return;
    }

    // Build table
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
