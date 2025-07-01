function excelDateToJSDate(serial) {
    if (!serial) return "";
    // Excel dates are days since 1900-01-00 (with bug)
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400; // seconds
    const date_info = new Date(utc_value * 1000);
    // Adjust for Excel leap year bug if necessary
    return date_info.toISOString().slice(0,10); // yyyy-mm-dd
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

    // Update these keys to match your Excel file
    const colPendingOrder = 'Pending Order Qty';
    const colUnderPacking = 'Under Packing Qty';
    const colAllocAvail = 'Allocation Available Qty';
    const colOpenQty = 'Open Qty';
    const colFirstAllocDate = 'First Allocation Date';
    const colConfirmedAllocDate = 'Confirmed Allocation Date';
    const colPORef = 'PO Reference';
    const STYLE_CODE_COL = 'Style Code';

    const filtered = excelData.filter(row => String(row[STYLE_CODE_COL]).toLowerCase() === styleCode.toLowerCase());

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
                    <div class="qty-value">${row[colPendingOrder] || 0}</div>
                </div>
                <div class="qty-card">
                    <div class="qty-label">Under Packing Qty</div>
                    <div class="qty-value">${row[colUnderPacking] || 0}</div>
                </div>
                <div class="qty-card">
                    <div class="qty-label">Allocation Available Qty</div>
                    <div class="qty-value">${row[colAllocAvail] || 0}</div>
                </div>
                <div class="qty-card">
                    <div class="qty-label">Open Qty</div>
                    <div class="qty-value">${row[colOpenQty] || 0}</div>
                </div>
            </div>
            <div class="dates-block">
                <div class="date-card">
                    <span class="date-label">First Allocation Date:</span>
                    <span class="date-value">${excelDateToJSDate(row[colFirstAllocDate])}</span>
                </div>
                <div class="date-card">
                    <span class="date-label">Confirmed Allocation Date:</span>
                    <span class="date-value">${excelDateToJSDate(row[colConfirmedAllocDate])}</span>
                </div>
            </div>
            <div class="po-block">
                <span class="po-label">PO Reference:</span>
                <span class="po-value">${row[colPORef] || ''}</span>
            </div>
        </div>
        `;
    });

    document.getElementById('results').innerHTML = html;
}
