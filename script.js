let excelData = [];
let headers = [];

const RAW_EXCEL_URL = "https://raw.githubusercontent.com/mnoornawab/ETA-of-Ordered-Items-/main/KEYE%20-%20Pending%20Orders%20Report%20-%20SIMA%20.xlsx";

// Fields to display
const FIELDS = [
  "First Allocation Date",
  "Confirmed Allocation Date",
  "Pending Order Qty",
  "Under Packing Qty",
  "Allocation Available Qty",
  "Open Qty"
];

document.getElementById('loadFileBtn').addEventListener('click', async () => {
  document.getElementById('result').innerText = "Loading Excel file...";
  try {
    const response = await fetch(RAW_EXCEL_URL);
    if (!response.ok) throw new Error("Failed to fetch file.");
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    // Get all rows as arrays
    excelData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    headers = excelData[0].map(h => h.trim());
    document.getElementById('querySection').style.display = 'block';
    document.getElementById('result').innerHTML = `<b>Loaded sheet:</b> ${sheetName}`;
  } catch (e) {
    document.getElementById('result').innerText = "Error loading Excel file: " + e.message;
  }
});

document.getElementById('searchBtn').addEventListener('click', handleSearch, false);

function handleSearch() {
  const styleCode = document.getElementById('queryInput').value.trim().toLowerCase();
  if (!excelData.length) {
    document.getElementById('result').innerText = 'No Excel file loaded.';
    return;
  }
  const styleColIdx = headers.findIndex(h => h.toLowerCase() === "style code");
  if (styleColIdx === -1) {
    document.getElementById('result').innerText = 'No "Style Code" column found in sheet.';
    return;
  }

  // Find requested field indices
  const fieldIndices = FIELDS.map(field => headers.findIndex(h => h.toLowerCase() === field.toLowerCase()));
  if (fieldIndices.some(idx => idx === -1)) {
    document.getElementById('result').innerText = "One or more required fields are missing from the sheet.";
    return;
  }

  // Filter rows by style code (skip header row)
  const matches = excelData.slice(1).filter(row =>
    String(row[styleColIdx] ?? "").trim().toLowerCase() === styleCode
  );

  if (!matches.length) {
    document.getElementById('result').innerText = 'No matching Style Code found.';
    return;
  }

  // Build results table
  let html = '<table border="1" cellpadding="5"><tr>';
  html += `<th>Style Code</th>`;
  FIELDS.forEach(field => html += `<th>${field}</th>`);
  html += '</tr>';

  matches.forEach(row => {
    html += `<tr><td>${row[styleColIdx] ?? ""}</td>`;
    fieldIndices.forEach(idx => html += `<td>${row[idx] ?? ""}</td>`);
    html += '</tr>';
  });
  html += '</table>';
  document.getElementById('result').innerHTML = html;
}
