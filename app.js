let skuDatabase = [];

let pendingRows = [];
let pendingFileName = "";
let pendingHeaders = [];

document.getElementById("excelFile").addEventListener("change", handleExcelUpload);

function normalizeCode(code) {
  return String(code || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "");
}

function cleanPrice(price) {
  return String(price || "")
    .replace("₹", "")
    .replace(/,/g, "")
    .trim();
}

function escapeHTML(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function handleExcelUpload(event) {
  const file = event.target.files[0];

  if (!file) {
    alert("Please select an Excel file.");
    return;
  }

  const reader = new FileReader();

  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);

    const workbook = XLSX.read(data, { type: "array" });

    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    const sheetRows = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: "",
      raw: false
    });

    if (sheetRows.length === 0) {
      alert("Excel file is empty.");
      return;
    }

    const headerRowIndex = findHeaderRowIndex(sheetRows);
    const headerRow = sheetRows[headerRowIndex];

    pendingHeaders = makeUniqueHeaders(headerRow);

    const dataRows = sheetRows
      .slice(headerRowIndex + 1)
      .filter(row => row.some(cell => String(cell).trim() !== ""));

    pendingRows = dataRows.map(row => convertArrayRowToObject(row, pendingHeaders));

    if (pendingRows.length === 0) {
      alert("No product rows found after the header row.");
      return;
    }

    pendingFileName = file.name;

    saveExcelDataForFutureMapping();

    showColumnMappingScreen();
  };

  reader.readAsArrayBuffer(file);
}

function findHeaderRowIndex(sheetRows) {
  const knownHeaderWords = [
    "product code",
    "sku",
    "sku code",
    "item code",
    "material code",
    "code",
    "list price",
    "price",
    "mrp",
    "product name",
    "description",
    "item name"
  ];

  let bestIndex = 0;
  let bestScore = -1;

  const maxRowsToCheck = Math.min(sheetRows.length, 15);

  for (let i = 0; i < maxRowsToCheck; i++) {
    const row = sheetRows[i];

    const nonEmptyCells = row.filter(cell => String(cell).trim() !== "").length;

    if (nonEmptyCells < 2) {
      continue;
    }

    let knownMatchCount = 0;

    row.forEach(cell => {
      const cellText = String(cell).toLowerCase().trim();

      if (knownHeaderWords.includes(cellText)) {
        knownMatchCount++;
      }
    });

    const score = knownMatchCount * 10 + nonEmptyCells;

    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }
  }

  return bestIndex;
}

function makeUniqueHeaders(headerRow) {
  const headers = [];
  const usedHeaders = {};

  headerRow.forEach((cell, index) => {
    let header = String(cell || "").trim();

    if (!header) {
      header = "Column " + (index + 1);
    }

    if (usedHeaders[header]) {
      usedHeaders[header]++;
      header = header + " (" + usedHeaders[header] + ")";
    } else {
      usedHeaders[header] = 1;
    }

    headers.push(header);
  });

  return headers;
}

function convertArrayRowToObject(row, headers) {
  const obj = {};

  headers.forEach((header, index) => {
    obj[header] = row[index] !== undefined ? row[index] : "";
  });

  return obj;
}

function saveExcelDataForFutureMapping() {
  try {
    localStorage.setItem("lastExcelRowsForMapping", JSON.stringify(pendingRows));
    localStorage.setItem("lastExcelHeadersForMapping", JSON.stringify(pendingHeaders));
    localStorage.setItem("lastExcelFileNameForMapping", pendingFileName);
  } catch (error) {
    alert("Excel file has too much data to save locally. You can still use it now, but may need to upload again next time.");
  }
}

function loadExcelDataForFutureMapping() {
  const savedRows = localStorage.getItem("lastExcelRowsForMapping");
  const savedHeaders = localStorage.getItem("lastExcelHeadersForMapping");
  const savedFileName = localStorage.getItem("lastExcelFileNameForMapping");

  if (savedRows && savedHeaders) {
    try {
      pendingRows = JSON.parse(savedRows);
      pendingHeaders = JSON.parse(savedHeaders);
      pendingFileName = savedFileName || "Previous Excel File";

      showColumnMappingScreen();

      return true;
    } catch (error) {
      return false;
    }
  }

  return false;
}

function showColumnMappingScreen() {
  const mappingSection = document.getElementById("mappingSection");

  const productCodeSelect = document.getElementById("productCodeSelect");
  const priceSelect = document.getElementById("priceSelect");
  const productNameSelect = document.getElementById("productNameSelect");

  productCodeSelect.innerHTML = "";
  priceSelect.innerHTML = "";
  productNameSelect.innerHTML = "";

  const productCodeAuto = findColumn(pendingHeaders, [
    "product code",
    "sku",
    "sku code",
    "item code",
    "material code",
    "code"
  ]);

  const priceAuto = findColumn(pendingHeaders, [
    "list price",
    "lp",
    "price",
    "mrp",
    "basic price",
    "rate",
    "selling price"
  ]);

  const productNameAuto = findColumn(pendingHeaders, [
    "product name",
    "item name",
    "description",
    "sku description",
    "material description",
    "item description"
  ]);

  addDefaultOption(productCodeSelect, "Select Product Code column");
  addDefaultOption(priceSelect, "Select List Price column");
  addDefaultOption(productNameSelect, "No Product Name column");

  pendingHeaders.forEach(header => {
    addOption(productCodeSelect, header);
    addOption(priceSelect, header);
    addOption(productNameSelect, header);
  });

  const savedProductCodeColumn = localStorage.getItem("lastProductCodeColumn");
  const savedPriceColumn = localStorage.getItem("lastPriceColumn");
  const savedProductNameColumn = localStorage.getItem("lastProductNameColumn");

  if (savedProductCodeColumn && pendingHeaders.includes(savedProductCodeColumn)) {
    productCodeSelect.value = savedProductCodeColumn;
  } else if (productCodeAuto) {
    productCodeSelect.value = productCodeAuto;
  }

  if (savedPriceColumn && pendingHeaders.includes(savedPriceColumn)) {
    priceSelect.value = savedPriceColumn;
  } else if (priceAuto) {
    priceSelect.value = priceAuto;
  }

  if (savedProductNameColumn && pendingHeaders.includes(savedProductNameColumn)) {
    productNameSelect.value = savedProductNameColumn;
  } else if (productNameAuto) {
    productNameSelect.value = productNameAuto;
  }

  mappingSection.classList.remove("hidden");

  document.getElementById("importStatus").innerText =
    "Excel loaded: " + pendingFileName + ". Please check column mapping and confirm import.";

  showPreview();
}

function addDefaultOption(selectElement, text) {
  const option = document.createElement("option");
  option.value = "";
  option.textContent = text;
  selectElement.appendChild(option);
}

function addOption(selectElement, value) {
  const option = document.createElement("option");
  option.value = value;
  option.textContent = value;
  selectElement.appendChild(option);
}

function findColumn(headers, possibleNames) {
  for (const header of headers) {
    const cleanedHeader = header.toLowerCase().trim();

    if (possibleNames.includes(cleanedHeader)) {
      return header;
    }
  }

  return null;
}

function showPreview() {
  const previewBox = document.getElementById("previewBox");

  const firstFiveRows = pendingRows.slice(0, 5);

  let html = "<div class='table-scroll'>";
  html += "<table class='preview-table'>";
  html += "<tr>";

  pendingHeaders.forEach(header => {
    html += "<th>" + escapeHTML(header) + "</th>";
  });

  html += "</tr>";

  firstFiveRows.forEach(row => {
    html += "<tr>";

    pendingHeaders.forEach(header => {
      html += "<td>" + escapeHTML(row[header]) + "</td>";
    });

    html += "</tr>";
  });

  html += "</table>";
  html += "</div>";

  previewBox.innerHTML = html;
}

function confirmColumnMapping() {
  const productCodeColumn = document.getElementById("productCodeSelect").value;
  const priceColumn = document.getElementById("priceSelect").value;
  const productNameColumn = document.getElementById("productNameSelect").value;

  if (!productCodeColumn) {
    alert("Please select the Product Code / SKU column.");
    return;
  }

  if (!priceColumn) {
    alert("Please select the List Price column.");
    return;
  }

  localStorage.setItem("lastProductCodeColumn", productCodeColumn);
  localStorage.setItem("lastPriceColumn", priceColumn);
  localStorage.setItem("lastProductNameColumn", productNameColumn);

  importRows(productCodeColumn, priceColumn, productNameColumn);
}

function importRows(productCodeColumn, priceColumn, productNameColumn) {
  skuDatabase = [];

  let skippedRows = 0;
  let invalidPriceRows = 0;
  let duplicateRows = 0;

  const seenCodes = {};

  pendingRows.forEach(row => {
    const productCode = row[productCodeColumn];
    const listPrice = row[priceColumn];
    const productName = productNameColumn ? row[productNameColumn] : "";

    if (!productCode || !listPrice) {
      skippedRows++;
      return;
    }

    const normalizedCode = normalizeCode(productCode);
    const cleanedPrice = cleanPrice(listPrice);

    if (isNaN(Number(cleanedPrice))) {
      invalidPriceRows++;
      return;
    }

    if (seenCodes[normalizedCode]) {
      duplicateRows++;
      return;
    }

    seenCodes[normalizedCode] = true;

    skuDatabase.push({
      productCodeOriginal: String(productCode).trim(),
      productCodeNormalized: normalizedCode,
      productName: String(productName || "").trim(),
      listPrice: cleanedPrice,
      allProductDetails: row
    });
  });

  localStorage.setItem("skuDatabase", JSON.stringify(skuDatabase));
  localStorage.setItem("importFileName", pendingFileName);

  document.getElementById("mappingSection").classList.remove("hidden");

  document.getElementById("importStatus").innerText =
    "Imported " + skuDatabase.length + " SKUs from " + pendingFileName +
    ". Skipped " + skippedRows + " blank rows, " +
    invalidPriceRows + " invalid price rows, " +
    duplicateRows + " duplicate rows.";

  alert("Price list imported successfully. All product columns will now show in search result.");
}

function searchSku() {
  const input = document.getElementById("searchInput").value;

  if (!input.trim()) {
    alert("Please enter Product Code / SKU.");
    return;
  }

  const normalizedInput = normalizeCode(input);

  const result = skuDatabase.find(item =>
    item.productCodeNormalized === normalizedInput
  );

  const resultBox = document.getElementById("resultBox");

  if (result) {
    resultBox.innerHTML =
      "<h2>Product Found</h2>" +
      "<p><strong>Product Code:</strong> " + escapeHTML(result.productCodeOriginal) + "</p>" +
      "<p><strong>Product Name:</strong> " + escapeHTML(result.productName || "Not available") + "</p>" +
      "<p><strong>List Price:</strong> ₹" + escapeHTML(result.listPrice) + "</p>" +
      "<h3>All Product Details</h3>" +
      createProductDetailsTable(result.allProductDetails);
  } else {
    resultBox.innerHTML =
      "<h2>No Exact Match Found</h2>" +
      "<p>Please check the Product Code or import the latest price list.</p>";
  }
}

function createProductDetailsTable(productDetails) {
  if (!productDetails) {
    return "<p>No extra details available.</p>";
  }

  let html = "<div class='details-table-box'>";
  html += "<table class='details-table'>";

  Object.keys(productDetails).forEach(key => {
    const value = productDetails[key];

    html += "<tr>";
    html += "<th>" + escapeHTML(key) + "</th>";
    html += "<td>" + escapeHTML(value || "-") + "</td>";
    html += "</tr>";
  });

  html += "</table>";
  html += "</div>";

  return html;
}

window.onload = function() {
  const savedData = localStorage.getItem("skuDatabase");
  const fileName = localStorage.getItem("importFileName");

  if (savedData) {
    try {
      skuDatabase = JSON.parse(savedData);

      document.getElementById("importStatus").innerText =
        "Loaded " + skuDatabase.length + " SKUs from " + fileName + ".";
    } catch (error) {
      skuDatabase = [];
    }
  }

  const mappingLoaded = loadExcelDataForFutureMapping();

  if (!mappingLoaded && !savedData) {
    document.getElementById("importStatus").innerText =
      "No Excel file loaded yet. Please upload a price list.";
  }
};

if ("serviceWorker" in navigator) {
  navigator.serviceWorker.register("service-worker.js");
}