let skuDatabase = [];

let pendingRows = [];
let pendingFileName = "";
let pendingHeaders = [];

function normalizeCode(code) {
  return String(code || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "");
}

function cleanPrice(price) {
  return String(price || "")
    .replace("₹", "")
    .replace(",", "")
    .trim();
}

document.getElementById("excelFile").addEventListener("change", handleExcelUpload);

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

    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

    if (rows.length === 0) {
      alert("Excel file is empty.");
      return;
    }

    pendingRows = rows;
    pendingFileName = file.name;
    pendingHeaders = Object.keys(rows[0]);

    showColumnMappingScreen();
  };

  reader.readAsArrayBuffer(file);
}

function showColumnMappingScreen() {
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
    "basic price"
  ]);

  const productNameAuto = findColumn(pendingHeaders, [
    "product name",
    "item name",
    "description",
    "sku description",
    "material description"
  ]);

  addDefaultOption(productCodeSelect, "Select Product Code column");
  addDefaultOption(priceSelect, "Select List Price column");
  addDefaultOption(productNameSelect, "No Product Name column");

  pendingHeaders.forEach(header => {
    addOption(productCodeSelect, header);
    addOption(priceSelect, header);
    addOption(productNameSelect, header);
  });

  if (productCodeAuto) {
    productCodeSelect.value = productCodeAuto;
  }

  if (priceAuto) {
    priceSelect.value = priceAuto;
  }

  if (productNameAuto) {
    productNameSelect.value = productNameAuto;
  }

  document.getElementById("mappingSection").classList.remove("hidden");

  document.getElementById("importStatus").innerText =
    "Excel loaded. Please check column mapping and confirm import.";

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

  let html = "<table class='preview-table'>";
  html += "<tr>";

  pendingHeaders.forEach(header => {
    html += `<th>${header}</th>`;
  });

  html += "</tr>";

  firstFiveRows.forEach(row => {
    html += "<tr>";

    pendingHeaders.forEach(header => {
      html += `<td>${row[header]}</td>`;
    });

    html += "</tr>";
  });

  html += "</table>";

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
      productName: productName,
      listPrice: cleanedPrice
    });
  });

  localStorage.setItem("skuDatabase", JSON.stringify(skuDatabase));
  localStorage.setItem("importFileName", pendingFileName);

  document.getElementById("mappingSection").classList.add("hidden");

  document.getElementById("importStatus").innerText =
    `Imported ${skuDatabase.length} SKUs from ${pendingFileName}. Skipped ${skippedRows} blank rows, ${invalidPriceRows} invalid price rows, ${duplicateRows} duplicate rows.`;

  alert("Price list imported successfully.");
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
    resultBox.innerHTML = `
      <h2>Product Found</h2>
      <p><strong>Product Code:</strong> ${result.productCodeOriginal}</p>
      <p><strong>Product Name:</strong> ${result.productName || "Not available"}</p>
      <p><strong>List Price:</strong> ₹${result.listPrice}</p>
    `;
  } else {
    resultBox.innerHTML = `
      <h2>No Exact Match Found</h2>
      <p>Please check the Product Code or import the latest price list.</p>
    `;
  }
}

window.onload = function() {
  const savedData = localStorage.getItem("skuDatabase");
  const fileName = localStorage.getItem("importFileName");

  if (savedData) {
    skuDatabase = JSON.parse(savedData);

    document.getElementById("importStatus").innerText =
      `Loaded ${skuDatabase.length} SKUs from ${fileName}.`;
  }
};

if ("serviceWorker" in navigator) {
  navigator.serviceWorker.register("service-worker.js");
}