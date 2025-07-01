let excelData = [];

// UPDATE THIS with your repo details:
const RAW_EXCEL_URL = "https://raw.githubusercontent.com/<username>/<repo>/<branch>/KEYE%20-%20Pending%20Orders%20Report%20-%20SIMA%20.xlsx";

document.getElementById('loadFileBtn').addEventListener('click', async () => {
  document.getElementById('result').innerText = "Loading Excel file...";
  try {
    const response = await fetch(RAW_EXCEL_URL);
    if (!response.ok) throw new Error("Failed to fetch file.");
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    excelData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    document.getElementById('querySection').style.display = 'block';
    document.getElementById('result').innerHTML = `<b>Loaded sheet:</b> ${sheetName}`;
  } catch (e) {
    document.getElementById('result').innerText = "Error loading Excel file: " + e.message;
  }
});

document.getElementById('searchBtn').addEventListener('click', handleSearch, false);

function handleSearch() {
  const query = document.getElementById('queryInput').value.toLowerCase();
  if (!excelData.length) {
    document.getElementById('result').innerText = 'No Excel file loaded.';
    return;
  }
  let results = [];
  for (let row of excelData) {
    if (row.join(' ').toLowerCase().includes(query)) {
      results.push(row);
    }
  }
  if (results.length) {
    let html = '<table border="1" cellpadding="5">';
    results.forEach(row => {
      html += '<tr>' + row.map(cell => `<td>${cell}</td>`).join('') + '</tr>';
    });
    html += '</table>';
    document.getElementById('result').innerHTML = html;
  } else {
    document.getElementById('result').innerText = 'No match found.';
  }
}
