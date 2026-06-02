let workbook = null;
let currentRows = [];
let currentHeaders = [];
let filteredRows = [];

const excelFileInput = document.getElementById("excelFile");
const sheetSelect = document.getElementById("sheetSelect");
const searchInput = document.getElementById("searchInput");
const columnSelect = document.getElementById("columnSelect");
const downloadCsvBtn = document.getElementById("downloadCsvBtn");

const controlsSection = document.getElementById("controlsSection");
const statusSection = document.getElementById("statusSection");
const statusText = document.getElementById("statusText");
const resultsSection = document.getElementById("resultsSection");
const detailsSection = document.getElementById("detailsSection");

const resultsTable = document.getElementById("resultsTable");
const rowDetails = document.getElementById("rowDetails");

/*
  Upload flow:
  User selects an .xlsx file.
  File is read locally in the browser.
  Nothing is uploaded to a server.
*/
excelFileInput.addEventListener("change", handleFileUpload);

sheetSelect.addEventListener("change", () => {
  loadSelectedSheet();
  runSearch();
});

searchInput.addEventListener("input", runSearch);
columnSelect.addEventListener("change", runSearch);
downloadCsvBtn.addEventListener("click", downloadFilteredCsv);

function showStatus(message) {
  statusSection.classList.remove("hidden");
  statusText.textContent = message;
}

function hideStatus() {
  statusSection.classList.add("hidden");
  statusText.textContent = "";
}

function handleFileUpload(event) {
  const file = event.target.files[0];

  if (!file) return;

  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    showStatus("Please upload only a .xlsx Excel file.");
    return;
  }

  const reader = new FileReader();

  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);

      /*
        SheetJS reads the Excel workbook locally.
        workbook.SheetNames gives all available sheets.
      */
      workbook = XLSX.read(data, { type: "array" });

      populateSheetDropdown(workbook.SheetNames);

      controlsSection.classList.remove("hidden");
      resultsSection.classList.remove("hidden");
      detailsSection.classList.remove("hidden");

      loadSelectedSheet();
      runSearch();

      showStatus("Excel file loaded successfully.");
    } catch (error) {
      console.error(error);
      showStatus("Could not read this Excel file. Please check the file format.");
    }
  };

  reader.onerror = function () {
    showStatus("Error reading file.");
  };

  reader.readAsArrayBuffer(file);
}

function populateSheetDropdown(sheetNames) {
  sheetSelect.innerHTML = "";

  sheetNames.forEach((sheetName) => {
    const option = document.createElement("option");
    option.value = sheetName;
    option.textContent = sheetName;
    sheetSelect.appendChild(option);
  });
}

/*
  Sheet loading:
  Converts selected Excel sheet into a 2D array.
  Then detects the best header row and creates row objects.
*/
function loadSelectedSheet() {
  const sheetName = sheetSelect.value;
  const worksheet = workbook.Sheets[sheetName];

  const rawData = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
    raw: false
  });

  const cleanedData = removeEmptyRows(rawData);

  if (cleanedData.length === 0) {
    currentRows = [];
    currentHeaders = [];
    populateColumnDropdown();
    return;
  }

  const headerIndex = detectHeaderRow(cleanedData);
  currentHeaders = makeUniqueHeaders(cleanedData[headerIndex]);

  const dataRows = cleanedData.slice(headerIndex + 1);

  currentRows = dataRows
    .filter((row) => !isRowEmpty(row))
    .map((row, index) => {
      const obj = {};

      obj["Excel Row"] = headerIndex + index + 2;

      currentHeaders.forEach((header, colIndex) => {
        obj[header] = row[colIndex] ?? "";
      });

      return obj;
    });

  populateColumnDropdown();
  clearRowDetails();
}

/*
  Removes fully empty rows.
*/
function removeEmptyRows(rows) {
  return rows.filter((row) => !isRowEmpty(row));
}

function isRowEmpty(row) {
  return row.every((cell) => String(cell ?? "").trim() === "");
}

/*
  Header detection:
  Your assumption says columns are in the first row.
  But many real price lists have a title row first.
  So this chooses the row with the most filled cells among the first 10 rows.
*/
function detectHeaderRow(rows) {
  const maxScan = Math.min(rows.length, 10);
  let bestIndex = 0;
  let bestCount = 0;

  for (let i = 0; i < maxScan; i++) {
    const count = rows[i].filter((cell) => String(cell ?? "").trim() !== "").length;

    if (count > bestCount) {
      bestCount = count;
      bestIndex = i;
    }
  }

  return bestIndex;
}

/*
  Creates safe unique column names.
  Handles blank headers and duplicate headers.
*/
function makeUniqueHeaders(headerRow) {
  const headers = [];
  const seen = {};

  headerRow.forEach((cell, index) => {
    let base = String(cell ?? "").trim();

    if (!base) {
      base = `Column ${index + 1}`;
    }

    if (seen[base]) {
      seen[base] += 1;
      headers.push(`${base}_${seen[base]}`);
    } else {
      seen[base] = 1;
      headers.push(base);
    }
  });

  return headers;
}

function populateColumnDropdown() {
  columnSelect.innerHTML = "";

  const allOption = document.createElement("option");
  allOption.value = "__ALL__";
  allOption.textContent = "All columns";
  columnSelect.appendChild(allOption);

  const allHeaders = ["Excel Row", ...currentHeaders];

  allHeaders.forEach((header) => {
    const option = document.createElement("option");
    option.value = header;
    option.textContent = header;
    columnSelect.appendChild(option);
  });
}

/*
  Search logic:
  - Case-insensitive
  - Partial match
  - Can search all columns or selected column
*/
function runSearch() {
  const query = searchInput.value.trim().toLowerCase();
  const selectedColumn = columnSelect.value;

  clearRowDetails();

  if (!query) {
    filteredRows = [...currentRows];
  } else {
    filteredRows = currentRows.filter((row) => {
      if (selectedColumn === "__ALL__") {
        return Object.values(row).some((value) =>
          String(value ?? "").toLowerCase().includes(query)
        );
      }

      return String(row[selectedColumn] ?? "").toLowerCase().includes(query);
    });
  }

  renderResultsTable();

  if (filteredRows.length === 0) {
    showStatus("No results found. Try another product code, name, price, or column.");
  } else {
    showStatus(`Found ${filteredRows.length} matching row(s).`);
  }
}

/*
  Results table:
  Shows all matching rows.
  Clicking a row opens full row details.
*/
function renderResultsTable() {
  resultsTable.innerHTML = "";

  if (filteredRows.length === 0) {
    return;
  }

  const headers = Object.keys(filteredRows[0]);

  const thead = document.createElement("thead");
  const headerTr = document.createElement("tr");

  headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headerTr.appendChild(th);
  });

  thead.appendChild(headerTr);
  resultsTable.appendChild(thead);

  const tbody = document.createElement("tbody");

  filteredRows.forEach((row) => {
    const tr = document.createElement("tr");

    headers.forEach((header) => {
      const td = document.createElement("td");
      td.textContent = row[header] ?? "";
      tr.appendChild(td);
    });

    tr.addEventListener("click", () => {
      renderRowDetails(row);
    });

    tbody.appendChild(tr);
  });

  resultsTable.appendChild(tbody);
}

/*
  Row Details:
  Shows every column and value from selected row.
  This solves the main problem: user can see data to the left and right
  of the matched Excel cell.
*/
function renderRowDetails(row) {
  rowDetails.innerHTML = "";

  Object.entries(row).forEach(([key, value]) => {
    const detailRow = document.createElement("div");
    detailRow.className = "detail-row";

    const keyDiv = document.createElement("div");
    keyDiv.className = "detail-key";
    keyDiv.textContent = key;

    const valueDiv = document.createElement("div");
    valueDiv.innerHTML = highlightMatch(String(value ?? ""));

    detailRow.appendChild(keyDiv);
    detailRow.appendChild(valueDiv);

    rowDetails.appendChild(detailRow);
  });
}

function clearRowDetails() {
  rowDetails.innerHTML = "";
}

/*
  Highlight matched text in row details.
*/
function highlightMatch(text) {
  const query = searchInput.value.trim();

  if (!query) {
    return escapeHtml(text);
  }

  const escapedText = escapeHtml(text);
  const escapedQuery = escapeRegExp(escapeHtml(query));

  const regex = new RegExp(`(${escapedQuery})`, "gi");

  return escapedText.replace(regex, "<mark>$1</mark>");
}

function escapeHtml(text) {
  return text.replace(/[&<>"']/g, function (char) {
    return {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#039;"
    }[char];
  });
}

function escapeRegExp(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/*
  CSV download:
  Downloads only the filtered search results.
*/
function downloadFilteredCsv() {
  if (filteredRows.length === 0) {
    showStatus("No filtered results to download.");
    return;
  }

  const worksheet = XLSX.utils.json_to_sheet(filteredRows);
  const csv = XLSX.utils.sheet_to_csv(worksheet);

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = url;
  link.download = "filtered-results.csv";
  document.body.appendChild(link);
  link.click();

  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}