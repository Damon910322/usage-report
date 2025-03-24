let usageData = [];

document.getElementById("excelFileInput").addEventListener("change", function(event) {
  const file = event.target.files[0];
  if (!file) return;

  const loader = document.getElementById("loader");
  loader.style.display = "block";

  const reader = new FileReader();
reader.onload = function(e) {
  const data = new Uint8Array(e.target.result);
  const workbook = XLSX.read(data, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  // Read sheet as raw array of rows
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

  if (rows.length < 2 || rows[0].length < 45) {
    document.getElementById("loader").style.display = "none";
    Swal.fire({
      icon: 'error',
      title: 'Invalid Format',
      text: "The file doesn't have enough rows or columns (expected at least 45 columns).",
    });
    return;
  }

  // Fixed indices: 0 = Item Number, 28 = Merchandise Amt, 44 = Each Quantity
  usageData = rows.slice(1).map(row => ({
    item: String(row[0] || "").trim(),             // Column A
    spend: parseFloat(row[28]) || 0,               // Column AC
    usage: parseFloat(row[44]) || 0                // Column AS
  }));

  document.getElementById("loader").style.display = "none";

  Swal.fire({
    icon: 'success',
    title: 'File Loaded!',
    text: 'Now enter a PeopleSoft Number to search.',
    confirmButtonColor: '#3085d6'
  });

  document.getElementById("peopleSoftSection").style.display = "block";
};


  reader.readAsArrayBuffer(file);
});

document.getElementById("lookupBtn").addEventListener("click", function() {
  const input = document.getElementById("peopleSoftInput").value;
  const itemNumbers = input.split(/[\s,]+/).map(num => num.trim()).filter(num => num !== "");

  if (itemNumbers.length === 0) {
    Swal.fire({
      icon: 'warning',
      title: 'Missing Input',
      text: 'Please enter at least one valid PeopleSoft Number.'
    });
    return;
  }

  let resultHTML = `
    <table id="resultTable" class="table table-striped table-bordered mt-4">
      <thead class="table-dark">
        <tr>
          <th>PeopleSoft Number</th>
          <th>Total Usage</th>
          <th>Total Spend</th>
        </tr>
      </thead>
      <tbody>
  `;

  let anyMatch = false;

  itemNumbers.forEach(number => {
    const matches = usageData.filter(row => row.item.toLowerCase() === number.toLowerCase());
    if (matches.length > 0) {
      const totalUsage = matches.reduce((sum, row) => sum + row.usage, 0);
      const totalSpend = matches.reduce((sum, row) => sum + row.spend, 0);
      const formattedSpend = totalSpend.toLocaleString('en-US', {
        style: 'currency',
        currency: 'USD'
      });

      resultHTML += `
        <tr>
          <td>${number}</td>
          <td>${totalUsage}</td>
          <td>${formattedSpend}</td>
        </tr>
      `;
      anyMatch = true;
    }
  });

  resultHTML += "</tbody></table>";

  if (anyMatch) {
    document.getElementById("result").innerHTML = resultHTML;
    $('#resultTable').DataTable();

    Swal.fire({
      icon: 'success',
      title: 'Results Ready!',
      text: 'Your usage and spend summary is shown below.',
      confirmButtonColor: '#28a745'
    });
  } else {
    document.getElementById("result").innerHTML = "";
    Swal.fire({
      icon: 'info',
      title: 'No Matches Found',
      text: 'None of the entered PeopleSoft Numbers matched the data.'
    });
  }
});
