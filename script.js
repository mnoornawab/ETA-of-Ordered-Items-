let excelData = [];

window.onload = function() {
    // Try both URL-encoded and plain name; update as per your file!
    fetch('KEYE%20-%20Pending%20Orders%20Report%20-%20SIMA.xlsx')
        .then(response => {
            if (!response.ok) throw new Error('Network response was not ok');
            return response.arrayBuffer();
        })
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            excelData = XLSX.utils.sheet_to_json(worksheet);
            // Optional: log data to debug column names
            console.log('Loaded data:', excelData);
            document.getElementById('results').innerHTML = '<p style="color:green;">Excel data loaded. Enter a style code to search.</p>';
        })
        .catch(err => {
            console.error(err);
            document.getElementById('results').innerHTML = '<p style="color:red;">Failed to load Excel file.</p>';
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
        document.getElementById('results').innerHTML = '<p>Data not loaded yet.</p>';
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

    // Also, update this key to your actual style code column name!
    const STYLE_CODE_COL = 'Style Code'; 

    // Show all columns for debugging if needed
    //console.log('First row keys:', Object.keys(excelData[0]));

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
