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
            // Uncomment to debug column names:
            // console.log('Loaded columns:', Object.keys(excelData[0]));
        })
        .catch(err => {
            document.getElementById('results').innerHTML = '<p style="color:red;">Failed to load Excel file: ' + err + '</p>';
            console.error('Excel load error:', err);
        });
}

// Only attach listeners after DOM is ready
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
        document.getElementById('results').innerHTML = '<p>No results found for this style code.</p>';
        return;
    }

    let html = '';
    filtered.forEach(row => {
        html += `
        <div class="result-card">
            <div class="qty-blocks">
                <div class="qty-card">
                    <div class="qty-label">Pending Order Qty</div>
                    <div class="qty-value">${row[colPendingOrder] || ""}</div>
                </div>
                <div class="qty-card">
                    <div class="qty-label">Under Packing Qty</div>
                    <div class="qty-value">${row[colUnderPacking] || ""}</div>
                </div>
                <div class="qty-card">
                    <div class="qty-label">Allocation Available Qty</div>
                    <div class="qty-value">${row[colAllocAvail] || ""}</div>
                </div>
                <div class="qty-card">
                    <div class="qty-label">Open Qty</div>
                    <div class="qty-value">${row[colOpenQty] || ""}</div>
                </div>
            </div>
            <div class="dates-block">
                <div class="date-card">
                    <span class="date-label">First Allocation Date:</span>
                    <span class="date-value">${row[colFirstAllocDate] || ""}</span>
                </div>
                <div class="date-card">
                    <span class="date-label">Confirmed Allocation Date:</span>
                    <span class="date-value">${row[colConfirmedAllocDate] || ""}</span>
                </div>
            </div>
            <div class="po-block">
                <span class="po-label">PO Reference:</span>
                <span class="po-value">${row[colPORef] || ""}</span>
            </div>
        </div>
        `;
    });

    document.getElementById('results').innerHTML = html;
}
