let excelData = [];

function loadExcelData(forceReload = false) {
    let fileUrl = 'KEYE_Pending_Orders_Report_SIMA.xlsx';
    if (forceReload) {
        fileUrl += '?t=' + new Date().getTime();
    }
    fetch(fileUrl)
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
            console.log('Excel data loaded:', excelData);
        })
        .catch(err => {
            document.getElementById('results').innerHTML = '<p style="color:red;">Failed to load Excel file: ' + err + '</p>';
            console.error('Excel load error:', err);
        });
}

document.addEventListener('DOMContentLoaded', function(){
    loadExcelData(); // Initial load

    document.getElementById('searchBtn').addEventListener('click', function() {
        const styleCode = document.getElementById('styleCodeInput').value.trim();
        displayResults(styleCode);
    });
    const reloadBtn = document.getElementById('reloadBtn');
    if (reloadBtn) {
        reloadBtn.addEventListener('click', function() {
            loadExcelData(true);
        });
    }
});

// Robust date formatter: handles Excel serials and date strings
function formatExcelDate(value) {
    // If value is a number or numeric string, likely an Excel serial
    if (!isNaN(value) && value !== "" && value !== null) {
        const serial = Number(value);
        if (serial > 25569 && serial < 60000) { // Reasonable Excel serial range
            const utc_days = serial - 25569;
            const utc_value = utc_days * 86400; // seconds
            const date_info = new Date(utc_value * 1000);
            if (!isNaN(date_info.getTime())) {
                // Format as MM/DD/YY
                return `${date_info.getMonth()+1}/${date_info.getDate()}/${String(date_info.getFullYear()).slice(-2)}`;
            }
        }
    }
    // If it's already a string that looks like a date, just return it as is
    if (typeof value === "string" && /\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}/.test(value)) {
        return value;
    }
    return value || "";
}

function displayResults(styleCode) {
    if (!excelData.length) {
        document.getElementById('results').innerHTML = '<p>Data not loaded yet. Wait a few seconds and try again.</p>';
        return;
    }
    if (!styleCode) {
        document.getElementById('results').innerHTML = '<p>Please enter a style code.</p>';
        return;
    }

    // Update these keys to match your Excel file exactly (case and spacing)
    const colPendingOrder = 'Pending Order Qty';
    const colUnderPacking = 'Under Packing Qty';
    const colAllocAvail = 'Allocation Available Qty';
    const colOpenQty = 'Open Qty';
    const colFirstAllocDate = 'First Allocation Date';
    const colConfirmedAllocDate = 'Confirmed Allocation Date';
    const colPORef = 'PO Reference';
    const STYLE_CODE_COL = 'Style Code';

    const filtered = excelData.filter(row => 
        row[STYLE_CODE_COL] && String(row[STYLE_CODE_COL]).toLowerCase() === styleCode.toLowerCase()
    );

    if (filtered.length === 0) {
        document.getElementById('results').innerHTML = '<div class="error-message">Item not on order</div>';
        return;
    }

    // Build a table for multiple orders
    let html = `
    <table class="results-table">
        <thead>
            <tr>
                <th>PO Reference</th>
                <th>Pending Order Qty</th>
                <th>Under Packing Qty</th>
                <th>Allocation Available Qty</th>
                <th>Open Qty</th>
                <th>First Allocation Date</th>
                <th>Confirmed Allocation Date</th>
            </tr>
        </thead>
        <tbody>
    `;

    filtered.forEach(row => {
        html += `
            <tr>
                <td>${row[colPORef] || ""}</td>
                <td>${row[colPendingOrder] || ""}</td>
                <td>${row[colUnderPacking] || ""}</td>
                <td>${row[colAllocAvail] || ""}</td>
                <td>${row[colOpenQty] || ""}</td>
                <td>${formatExcelDate(row[colFirstAllocDate])}</td>
                <td>${formatExcelDate(row[colConfirmedAllocDate])}</td>
            </tr>
        `;
    });

    html += "</tbody></table>";

    document.getElementById('results').innerHTML = html;
}
