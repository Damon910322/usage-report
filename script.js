let usageData = [];

document.getElementById("excelFileInput").addEventListener("change", function(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();

  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    const sheetName = workbook.SheetNames[0]; // Only one sheet expected
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "", raw: true });

    // Validate headers
    if (!jsonData[0] || !("Item Numbers" in jsonData[0]) || !("Usage" in jsonData[0]) || !("Spend" in jsonData[0])) {
      Swal.fire({
        icon: 'error',
        title: 'Invalid Format',
        text: "Missing columns: 'Item Numbers', 'Usage', or 'Spend'",
      });
      return;
    }

    // Store formatted data
    usageData = jsonData.map(row => ({
      item: String(row["Item Numbers"]).trim(),
      usage: parseFloat(row["Usage"]) || 0,
      spend: parseFloat(row["Spend"]) || 0
    }));

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
    const matches = usageData.filter(row => row.item === number);
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

    // ✅ Initialize DataTable after table is added to DOM
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
