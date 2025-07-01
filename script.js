let excelData = [];

window.onload = function() {
    fetch('KEYE_Pending_Orders_Report_SIMA.xlsx')
        .then(response => response.arrayBuffer())
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            excelData = XLSX.utils.sheet_to_json(worksheet);
        })
        .catch(err => {
            document.getElementById('results').innerHTML = '<p style="color:red;">Failed to load Excel file.</p>';
        });
};
