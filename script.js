let usageData = [];

document.getElementById("csvFileInput").addEventListener("change", function(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();

  reader.onload = function(e) {
    const text = e.target.result;
    const lines = text.split(/\r?\n/).filter(line => line.trim() !== "");

    const headers = lines[0].split(",").map(h => h.trim());
    const itemIndex = headers.indexOf("Item Numbers");
    const usageIndex = headers.indexOf("Usage");
    const spendIndex = headers.indexOf("Spend");

    if (itemIndex === -1 || usageIndex === -1 || spendIndex === -1) {
      document.getElementById("result").textContent =
        "Invalid CSV headers. Make sure you have 'Item Numbers', 'Usage', and 'Spend'.";
      return;
    }

    usageData = lines.slice(1).map(line => {
      const parts = line.split(",");
      return {
        item: parts[itemIndex]?.trim(),
        usage: parseFloat(parts[usageIndex]) || 0,
        spend: parseFloat(parts[spendIndex]) || 0
      };
    });

    document.getElementById("peopleSoftSection").style.display = "block";
  };

  reader.readAsText(file);
});

document.getElementById("lookupBtn").addEventListener("click", function() {
  const input = document.getElementById("peopleSoftInput").value;
  const itemNumbers = input.split(/[\s,]+/).map(num => num.trim()).filter(num => num !== "");

  if (itemNumbers.length === 0) {
    document.getElementById("result").textContent = "Please enter at least one valid PeopleSoft Number.";
    return;
  }

  let resultHTML = `
    <table border="1" cellpadding="8" cellspacing="0">
      <thead>
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

  document.getElementById("result").innerHTML = anyMatch
    ? resultHTML
    : "No matches found for the provided PeopleSoft Numbers.";
});
