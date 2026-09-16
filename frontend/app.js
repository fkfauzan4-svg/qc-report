// Column metadata mirrors backend/app/schema.py — keep the two in sync.
const ITEM_COLUMNS = [
  { key: "itemNo", label: "Item #", type: "text", width: 60 },
  { key: "drawingRef", label: "Drawing / Item Ref", type: "text", width: 120 },
  { key: "description", label: "Description & Specification", type: "textarea" },
  { key: "dimensions", label: "Dimensions (mm)", type: "text", width: 140 },
  { key: "unit", label: "Unit", type: "text", width: 70 },
  { key: "qty", label: "Qty", type: "number", width: 80 },
  { key: "unitCost", label: "Unit Cost", type: "number", width: 100 },
];

const BOM_COLUMNS = [
  { key: "itemNo", label: "Part #", type: "text", width: 60 },
  { key: "parentRef", label: "Parent Item Ref", type: "text", width: 110 },
  { key: "assembly", label: "Assembly", type: "text", width: 120 },
  { key: "partDescription", label: "Part Description", type: "textarea" },
  { key: "materialSpec", label: "Material & Substrate Spec", type: "textarea" },
  { key: "dimensions", label: "Finished Dimensions (mm)", type: "text", width: 140 },
  { key: "qtyPerUnit", label: "Qty / Unit", type: "number", width: 80 },
  { key: "totalQty", label: "Total Qty", type: "number", width: 80 },
  { key: "unit", label: "Unit", type: "text", width: 60 },
  { key: "notes", label: "Notes", type: "textarea" },
];

const HARDWARE_COLUMNS = [
  { key: "code", label: "Item Code", type: "text", width: 70 },
  { key: "category", label: "Category", type: "text", width: 110 },
  { key: "itemName", label: "Item Name / Model", type: "text", width: 140 },
  { key: "brand", label: "Brand", type: "text", width: 110 },
  { key: "material", label: "Material & Finish", type: "textarea" },
  { key: "applicable", label: "Applicable Items", type: "text", width: 110 },
  { key: "qtyPerUnit", label: "Qty / Unit", type: "number", width: 80 },
  { key: "totalQty", label: "Total Qty", type: "number", width: 80 },
  { key: "unit", label: "Unit", type: "text", width: 60 },
  { key: "unitCost", label: "Unit Cost", type: "number", width: 90 },
  { key: "scope", label: "Scope / Function", type: "textarea" },
];

const PROJECT_FIELDS = [
  { key: "projectName", label: "Project Name" },
  { key: "documentRef", label: "Document Ref" },
  { key: "client", label: "Client" },
  { key: "manufacturer", label: "Manufacturer" },
  { key: "date", label: "Date" },
  { key: "currency", label: "Currency" },
  { key: "vatRate", label: "VAT Rate (e.g. 0.15 = 15%)", type: "number" },
  { key: "notes", label: "Notes", full: true, textarea: true },
];

let state = null;

const $ = (sel) => document.querySelector(sel);
const view = {
  upload: $("#view-upload"),
  loading: $("#view-loading"),
  results: $("#view-results"),
};

function showView(name) {
  Object.values(view).forEach((v) => v.classList.add("hidden"));
  view[name].classList.remove("hidden");
}

function fmt(n) {
  const v = Number(n) || 0;
  return v.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function emptyRow(columns) {
  const row = {};
  columns.forEach((c) => (row[c.key] = c.type === "number" ? 0 : ""));
  return row;
}

// ---------------- Upload flow ----------------

const dropzone = $("#dropzone");
const fileInput = $("#file-input");
const fileChosen = $("#file-chosen");
const fileChosenName = $("#file-chosen-name");
const uploadError = $("#upload-error");
let selectedFile = null;

dropzone.addEventListener("click", () => fileInput.click());
dropzone.addEventListener("dragover", (e) => { e.preventDefault(); dropzone.classList.add("dragover"); });
dropzone.addEventListener("dragleave", () => dropzone.classList.remove("dragover"));
dropzone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropzone.classList.remove("dragover");
  if (e.dataTransfer.files.length) setSelectedFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener("change", () => {
  if (fileInput.files.length) setSelectedFile(fileInput.files[0]);
});

function setSelectedFile(file) {
  selectedFile = file;
  fileChosenName.textContent = `${file.name} (${(file.size / 1024).toFixed(0)} KB)`;
  fileChosen.classList.remove("hidden");
  uploadError.classList.add("hidden");
}

$("#btn-clear-file").addEventListener("click", () => {
  selectedFile = null;
  fileInput.value = "";
  fileChosen.classList.add("hidden");
});

$("#btn-analyze").addEventListener("click", async () => {
  if (!selectedFile) return;
  showView("loading");
  try {
    const form = new FormData();
    form.append("file", selectedFile);
    const res = await fetch("/api/analyze", { method: "POST", body: form });
    const body = await res.json();
    if (!res.ok) throw new Error(body.detail || "Analysis failed.");
    state = body.data;
    renderResults();
    showView("results");
  } catch (err) {
    showView("upload");
    uploadError.textContent = err.message || String(err);
    uploadError.classList.remove("hidden");
  }
});

$("#btn-new").addEventListener("click", () => {
  selectedFile = null;
  fileInput.value = "";
  fileChosen.classList.add("hidden");
  uploadError.classList.add("hidden");
  $("#global-error").classList.add("hidden");
  showView("upload");
});

// ---------------- Results rendering ----------------

function renderResults() {
  renderProjectFields();
  renderTable("table-items", "items", ITEM_COLUMNS, { totalFrom: ["qty", "unitCost"] });
  renderTable("table-bom", "bom", BOM_COLUMNS, {});
  renderTable("table-hardware", "hardware", HARDWARE_COLUMNS, { totalFrom: ["totalQty", "unitCost"] });
  renderAssumptions();
  recalcTotals();
}

function renderProjectFields() {
  const container = $("#project-fields");
  container.innerHTML = "";
  PROJECT_FIELDS.forEach((f) => {
    const wrap = document.createElement("div");
    wrap.className = "field" + (f.full ? " full" : "");
    const label = document.createElement("label");
    label.textContent = f.label;
    wrap.appendChild(label);
    const input = document.createElement(f.textarea ? "textarea" : "input");
    if (!f.textarea) input.type = f.type === "number" ? "number" : "text";
    if (f.type === "number") input.step = "0.01";
    input.value = state.project[f.key] ?? "";
    input.addEventListener("input", () => {
      state.project[f.key] = f.type === "number" ? parseFloat(input.value) || 0 : input.value;
      if (f.key === "vatRate") recalcTotals();
      if (f.key === "currency") renderResults();
    });
    wrap.appendChild(input);
    container.appendChild(wrap);
  });
}

function renderTable(tableId, stateKey, columns, opts) {
  const table = document.getElementById(tableId);
  const rows = state[stateKey];
  const hasTotal = !!opts.totalFrom;

  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  columns.forEach((c) => {
    const th = document.createElement("th");
    th.textContent = c.label;
    headRow.appendChild(th);
  });
  if (hasTotal) {
    const th = document.createElement("th");
    th.textContent = "Total Cost";
    headRow.appendChild(th);
  }
  headRow.appendChild(document.createElement("th")); // remove col
  thead.appendChild(headRow);

  const tbody = document.createElement("tbody");
  rows.forEach((row, idx) => {
    tbody.appendChild(buildRow(stateKey, columns, row, idx, opts));
  });

  table.innerHTML = "";
  table.appendChild(thead);
  table.appendChild(tbody);
}

function buildRow(stateKey, columns, row, idx, opts) {
  const tr = document.createElement("tr");

  columns.forEach((c) => {
    const td = document.createElement("td");
    const el = document.createElement(c.type === "textarea" ? "textarea" : "input");
    if (c.type === "number") {
      el.type = "number";
      el.step = "0.01";
      el.className = "num";
    } else if (c.type === "textarea") {
      el.rows = 2;
    } else {
      el.type = "text";
    }
    el.style.minWidth = c.width ? `${c.width}px` : c.type === "textarea" ? "220px" : "100px";
    el.value = row[c.key] ?? "";
    el.addEventListener("input", () => {
      row[c.key] = c.type === "number" ? parseFloat(el.value) || 0 : el.value;
      if (opts.totalFrom && opts.totalFrom.includes(c.key)) {
        updateRowTotal(tr, row, opts.totalFrom);
        recalcTotals();
      }
    });
    td.appendChild(el);
    tr.appendChild(td);
  });

  if (opts.totalFrom) {
    const totalTd = document.createElement("td");
    totalTd.className = "col-total";
    totalTd.textContent = fmt(rowTotal(row, opts.totalFrom));
    tr.appendChild(totalTd);
  }

  const removeTd = document.createElement("td");
  removeTd.className = "col-remove no-print";
  const btn = document.createElement("button");
  btn.className = "row-remove-btn";
  btn.title = "Remove row";
  btn.textContent = "✕";
  btn.addEventListener("click", () => {
    const arr = state[stateKey];
    arr.splice(arr.indexOf(row), 1);
    renderResults();
  });
  removeTd.appendChild(btn);
  tr.appendChild(removeTd);

  return tr;
}

function rowTotal(row, fromKeys) {
  return fromKeys.reduce((acc, k) => (acc === null ? Number(row[k]) || 0 : acc * (Number(row[k]) || 0)), null);
}

function updateRowTotal(tr, row, fromKeys) {
  const cell = tr.querySelector("td.col-total");
  if (cell) cell.textContent = fmt(rowTotal(row, fromKeys));
}

function renderAssumptions() {
  const list = $("#assumptions-list");
  list.innerHTML = "";
  state.assumptions.forEach((note, idx) => {
    const li = document.createElement("li");
    const input = document.createElement("input");
    input.value = note;
    input.addEventListener("input", () => (state.assumptions[idx] = input.value));
    li.appendChild(input);
    const btn = document.createElement("button");
    btn.className = "row-remove-btn no-print";
    btn.textContent = "✕";
    btn.addEventListener("click", () => {
      state.assumptions.splice(idx, 1);
      renderAssumptions();
    });
    li.appendChild(btn);
    list.appendChild(li);
  });
}

document.querySelectorAll("[data-add]").forEach((btn) => {
  btn.addEventListener("click", () => {
    const key = btn.getAttribute("data-add");
    if (key === "assumptions") {
      state.assumptions.push("");
      renderAssumptions();
      return;
    }
    const columns = { items: ITEM_COLUMNS, bom: BOM_COLUMNS, hardware: HARDWARE_COLUMNS }[key];
    state[key].push(emptyRow(columns));
    renderResults();
  });
});

function recalcTotals() {
  const subtotal = state.items.reduce((acc, r) => acc + (Number(r.qty) || 0) * (Number(r.unitCost) || 0), 0);
  const vatRate = Number(state.project.vatRate) || 0;
  const vat = subtotal * vatRate;
  const grand = subtotal + vat;
  const hardwareTotal = state.hardware.reduce(
    (acc, r) => acc + (Number(r.totalQty) || 0) * (Number(r.unitCost) || 0),
    0
  );
  const currency = state.project.currency || "";

  $("#total-subtotal").textContent = `${fmt(subtotal)} ${currency}`;
  $("#vat-label").textContent = `VAT (${(vatRate * 100).toFixed(1)}%)`;
  $("#total-vat").textContent = `${fmt(vat)} ${currency}`;
  $("#total-grand").textContent = `${fmt(grand)} ${currency}`;
  $("#total-hardware").textContent = `${fmt(hardwareTotal)} ${currency}`;
}

// ---------------- Export / Print ----------------

$("#btn-print").addEventListener("click", () => window.print());

$("#btn-download").addEventListener("click", async () => {
  const globalError = $("#global-error");
  globalError.classList.add("hidden");
  try {
    const res = await fetch("/api/export/xlsx", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(state),
    });
    if (!res.ok) {
      const body = await res.json().catch(() => ({}));
      throw new Error(body.detail || "Export failed.");
    }
    const blob = await res.blob();
    const disposition = res.headers.get("Content-Disposition") || "";
    const match = /filename="?([^"]+)"?/.exec(disposition);
    const filename = match ? match[1] : "BOM_Cost_Estimation.xlsx";

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  } catch (err) {
    globalError.textContent = err.message || String(err);
    globalError.classList.remove("hidden");
  }
});
