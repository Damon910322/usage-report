let peopleSoftNumbers = [];
let resultMap = {};

document.getElementById("excelFileInput").addEventListener("change", function(event) {
  const file = event.target.files[0];
  if (!file) return;

  Swal.fire({
    icon: 'info',
    title: 'File Loaded!',
    text: 'Now enter a PeopleSoft Number to search.',
    confirmButtonColor: '#3085d6'
  });

  document.getElementById("peopleSoftSection").style.display = "block";

  // Save the file reference for when we search
  window.uploadedExcelFile = file;
});

document.getElementById("lookupBtn").addEventListener("click", function () {
  const input = document.getElementById("peopleSoftInput").value;
  peopleSoftNumbers = input.split(/[\s,]+/).map(num => num.trim()).filter(num => num !== "");

  if (peopleSoftNumbers.length === 0) {
    Swal.fire({
      icon: 'warning',
      title: 'Missing Input',
      text: 'Please enter at least one valid PeopleSoft Number.'
    });
    return;
  }

  const loader = document.getElementById("loader");
  loader.style.display = "block";

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    resultMap = {};

    const range = XLSX.utils.decode_range(worksheet["!ref"]);
    for (let rowNum = range.s.r + 1; rowNum <= range.e.r; rowNum++) {
      const itemCell = worksheet[XLSX.utils.encode_cell({ c: 0, r: rowNum })];    // Column A - PeopleSoft Number
      const spendCell = worksheet[XLSX.utils.encode_cell({ c: 1, r: rowNum })];   // Column B - Merchandise Amt (Spend)
      const usageCell = worksheet[XLSX.utils.encode_cell({ c: 2, r: rowNum })];   // Column C - Each Quantity (Usage)


      const itemNumber = itemCell ? String(itemCell.v).trim() : "";
      if (!peopleSoftNumbers.includes(itemNumber)) continue;

      const spend = parseFloat(spendCell ? spendCell.v : 0) || 0;
      const usage = parseFloat(usageCell ? usageCell.v : 0) || 0;

      if (!resultMap[itemNumber]) {
        resultMap[itemNumber] = { usage: 0, spend: 0 };
      }

      resultMap[itemNumber].usage += usage;
      resultMap[itemNumber].spend += spend;
    }

    loader.style.display = "none";
    displayResults();
  };

  reader.readAsArrayBuffer(window.uploadedExcelFile);
});

function displayResults() {
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

  peopleSoftNumbers.forEach(number => {
    const result = resultMap[number];
    if (result) {
      const formattedSpend = result.spend.toLocaleString('en-US', {
        style: 'currency',
        currency: 'USD'
      });

      resultHTML += `
        <tr>
          <td>${number}</td>
          <td>${result.usage}</td>
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
}
