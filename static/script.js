const ERROR_VALUES = [
  "#REF!", "#N/A", "#VALUE!", "#DIV/0!", "#NAME?", "#NUM!",
  "#NULL!", "#SPILL!", "#CALC!", "#GETTING_DATA", "#FIELD!",
  "#BLOCKED!", "#UNKNOWN!"
];

let selectedFiles = [];       // File objects
let fileSheetsInfo = [];      // [{filename, sheets, size}]
let currentJobId = null;
let pollTimer = null;

const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("fileInput");
const folderInput = document.getElementById("folderInput");
const btnPickFiles = document.getElementById("btnPickFiles");
const btnPickFolder = document.getElementById("btnPickFolder");
const fileListEl = document.getElementById("fileList");

const sheetsCard = document.getElementById("sheetsCard");
const sheetsContainer = document.getElementById("sheetsContainer");

const filterCard = document.getElementById("filterCard");
const errorFilterGrid = document.getElementById("errorFilterGrid");
const btnSelectAllErr = document.getElementById("btnSelectAllErr");
const btnClearAllErr = document.getElementById("btnClearAllErr");

const scanCard = document.getElementById("scanCard");
const btnScan = document.getElementById("btnScan");
const btnExport = document.getElementById("btnExport");
const progressWrap = document.getElementById("progressWrap");
const progressBar = document.getElementById("progressBar");
const progressText = document.getElementById("progressText");
const statusMsg = document.getElementById("statusMsg");
const resultSummary = document.getElementById("resultSummary");
const resultCount = document.getElementById("resultCount");
const tableWrap = document.getElementById("tableWrap");
const resultTableBody = document.getElementById("resultTableBody");

const ALLOWED_EXT = [".xlsx", ".xlsm", ".xltx", ".xltm"];

function isAllowedFile(name) {
  const lower = name.toLowerCase();
  return ALLOWED_EXT.some(ext => lower.endsWith(ext));
}

function formatBytes(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / (1024 * 1024)).toFixed(1) + " MB";
}

/* ---------------- FILE PICKING ---------------- */

btnPickFiles.addEventListener("click", () => fileInput.click());
btnPickFolder.addEventListener("click", () => folderInput.click());

fileInput.addEventListener("change", (e) => addFiles(Array.from(e.target.files)));
folderInput.addEventListener("change", (e) => addFiles(Array.from(e.target.files)));

["dragenter", "dragover"].forEach(evt => {
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.add("dragover");
  });
});
["dragleave", "drop"].forEach(evt => {
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
  });
});
dropzone.addEventListener("drop", (e) => {
  const items = e.dataTransfer.items;
  if (items && items.length && items[0].webkitGetAsEntry) {
    // Hỗ trợ kéo thả cả folder
    const entries = Array.from(items).map(it => it.webkitGetAsEntry());
    readEntriesRecursively(entries).then(files => addFiles(files));
  } else {
    addFiles(Array.from(e.dataTransfer.files));
  }
});

function readEntriesRecursively(entries) {
  const promises = entries.map(entry => readEntry(entry));
  return Promise.all(promises).then(results => results.flat());
}

function readEntry(entry) {
  return new Promise((resolve) => {
    if (entry.isFile) {
      entry.file(file => resolve([file]));
    } else if (entry.isDirectory) {
      const reader = entry.createReader();
      const allEntries = [];
      const readBatch = () => {
        reader.readEntries((batch) => {
          if (!batch.length) {
            readEntriesRecursively(allEntries).then(resolve);
          } else {
            allEntries.push(...batch);
            readBatch();
          }
        });
      };
      readBatch();
    } else {
      resolve([]);
    }
  });
}

function addFiles(newFiles) {
  const valid = newFiles.filter(f => isAllowedFile(f.name));
  const skipped = newFiles.length - valid.length;

  for (const f of valid) {
    // tránh trùng tên file đã có
    if (!selectedFiles.some(sf => sf.name === f.name && sf.size === f.size)) {
      selectedFiles.push(f);
    }
  }

  if (skipped > 0) {
    statusMsg.textContent = `Đã bỏ qua ${skipped} file không đúng định dạng Excel.`;
  }

  renderFileList();
  fetchSheetsInfo();
}

function removeFile(index) {
  selectedFiles.splice(index, 1);
  renderFileList();
  if (selectedFiles.length > 0) {
    fetchSheetsInfo();
  } else {
    sheetsCard.style.display = "none";
    filterCard.style.display = "none";
    scanCard.style.display = "none";
    fileSheetsInfo = [];
  }
}

function renderFileList() {
  fileListEl.innerHTML = "";
  selectedFiles.forEach((f, idx) => {
    const item = document.createElement("div");
    item.className = "file-item";
    item.innerHTML = `
      <div class="file-item-left">
        <span class="file-icon">✓</span>
        <span class="file-name" title="${f.name}">${f.name}</span>
        <span class="file-size">${formatBytes(f.size)}</span>
      </div>
      <button class="file-remove" data-idx="${idx}">✕</button>
    `;
    fileListEl.appendChild(item);
  });

  fileListEl.querySelectorAll(".file-remove").forEach(btn => {
    btn.addEventListener("click", () => removeFile(parseInt(btn.dataset.idx)));
  });
}

/* ---------------- SHEETS FETCH ---------------- */

async function fetchSheetsInfo() {
  if (selectedFiles.length === 0) return;

  statusMsg.textContent = "Đang đọc danh sách sheet...";

  const formData = new FormData();
  selectedFiles.forEach(f => formData.append("files", f));

  try {
    const res = await fetch("/api/list-sheets", { method: "POST", body: formData });
    const data = await res.json();

    if (data.error) {
      statusMsg.textContent = data.error;
      return;
    }

    fileSheetsInfo = data.files;
    renderSheetsSelector();
    renderErrorFilter();

    sheetsCard.style.display = "block";
    filterCard.style.display = "block";
    scanCard.style.display = "block";
    statusMsg.textContent = "";
  } catch (err) {
    statusMsg.textContent = "Lỗi khi đọc file: " + err.message;
  }
}

function renderSheetsSelector() {
  sheetsContainer.innerHTML = "";

  fileSheetsInfo.forEach(fileInfo => {
    const group = document.createElement("div");
    group.className = "sheet-file-group";

    const title = document.createElement("div");
    title.className = "sheet-file-title";
    title.innerHTML = `📄 ${fileInfo.filename} <span class="badge-count">(${fileInfo.sheets.length} sheet)</span>`;
    group.appendChild(title);

    const grid = document.createElement("div");
    grid.className = "sheet-checkbox-grid";

    fileInfo.sheets.forEach(sheetName => {
      const label = document.createElement("label");
      label.className = "sheet-checkbox";
      label.innerHTML = `
        <input type="checkbox" checked
               data-file="${encodeURIComponent(fileInfo.filename)}"
               data-sheet="${encodeURIComponent(sheetName)}">
        <span>${sheetName}</span>
      `;
      grid.appendChild(label);
    });

    group.appendChild(grid);
    sheetsContainer.appendChild(group);
  });
}

function getSelectedSheetsMap() {
  const map = {};
  sheetsContainer.querySelectorAll('input[type="checkbox"]').forEach(cb => {
    const filename = decodeURIComponent(cb.dataset.file);
    const sheetName = decodeURIComponent(cb.dataset.sheet);
    if (!map[filename]) map[filename] = [];
    if (cb.checked) map[filename].push(sheetName);
  });
  return map;
}

/* ---------------- ERROR FILTER ---------------- */

function renderErrorFilter() {
  errorFilterGrid.innerHTML = "";
  ERROR_VALUES.forEach(err => {
    const label = document.createElement("label");
    label.className = "error-checkbox";
    label.innerHTML = `
      <input type="checkbox" checked value="${err}">
      <span>${err}</span>
    `;
    errorFilterGrid.appendChild(label);
  });
}

btnSelectAllErr.addEventListener("click", () => {
  errorFilterGrid.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = true);
});
btnClearAllErr.addEventListener("click", () => {
  errorFilterGrid.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
});

function getSelectedErrors() {
  return Array.from(errorFilterGrid.querySelectorAll('input[type="checkbox"]:checked'))
    .map(cb => cb.value);
}

/* ---------------- SCAN ---------------- */

btnScan.addEventListener("click", startScan);

async function startScan() {
  if (selectedFiles.length === 0) {
    statusMsg.textContent = "Vui lòng chọn ít nhất 1 file Excel.";
    return;
  }

  const selectedErrors = getSelectedErrors();
  if (selectedErrors.length === 0) {
    statusMsg.textContent = "Vui lòng chọn ít nhất 1 loại lỗi cần tìm.";
    return;
  }

  const selectedSheetsMap = getSelectedSheetsMap();
  const hasAnySheet = Object.values(selectedSheetsMap).some(arr => arr.length > 0);
  if (!hasAnySheet) {
    statusMsg.textContent = "Vui lòng chọn ít nhất 1 sheet cần quét.";
    return;
  }

  btnScan.disabled = true;
  btnExport.disabled = true;
  progressWrap.style.display = "flex";
  progressBar.style.width = "0%";
  progressText.textContent = "0%";
  resultSummary.style.display = "none";
  tableWrap.style.display = "none";
  resultTableBody.innerHTML = "";
  statusMsg.textContent = "Đang tải file lên...";

  const formData = new FormData();
  selectedFiles.forEach(f => formData.append("files", f));
  formData.append("selected_sheets", JSON.stringify(selectedSheetsMap));
  formData.append("selected_errors", JSON.stringify(selectedErrors));

  try {
    const res = await fetch("/api/scan", { method: "POST", body: formData });
    const data = await res.json();

    if (data.error) {
      statusMsg.textContent = data.error;
      btnScan.disabled = false;
      return;
    }

    currentJobId = data.job_id;
    pollStatus();
  } catch (err) {
    statusMsg.textContent = "Lỗi khi bắt đầu quét: " + err.message;
    btnScan.disabled = false;
  }
}

function pollStatus() {
  if (pollTimer) clearInterval(pollTimer);

  pollTimer = setInterval(async () => {
    try {
      const res = await fetch(`/api/status/${currentJobId}`);
      const data = await res.json();

      if (data.error) {
        clearInterval(pollTimer);
        statusMsg.textContent = data.error;
        btnScan.disabled = false;
        return;
      }

      progressBar.style.width = data.progress + "%";
      progressText.textContent = data.progress + "%";
      statusMsg.textContent = data.message;

      renderResults(data.results);

      if (data.status === "finished") {
        clearInterval(pollTimer);
        btnScan.disabled = false;
        if (data.results.length > 0) {
          btnExport.disabled = false;
        }
        resultSummary.style.display = "block";
        resultCount.textContent = `Lỗi tìm thấy: ${data.error_count}`;
      } else if (data.status === "error") {
        clearInterval(pollTimer);
        btnScan.disabled = false;
      }
    } catch (err) {
      clearInterval(pollTimer);
      statusMsg.textContent = "Mất kết nối tới server: " + err.message;
      btnScan.disabled = false;
    }
  }, 1000);
}

function renderResults(results) {
  if (!results || results.length === 0) return;

  tableWrap.style.display = "block";
  resultTableBody.innerHTML = "";

  results.forEach(r => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(r.file)}</td>
      <td>${escapeHtml(r.sheet)}</td>
      <td>${r.row}</td>
      <td>${escapeHtml(r.column)}</td>
      <td>${escapeHtml(r.cell)}</td>
      <td><span class="error-tag">${escapeHtml(r.error)}</span></td>
      <td class="formula-cell">${escapeHtml(r.formula)}</td>
    `;
    resultTableBody.appendChild(tr);
  });

  resultCount.textContent = `Lỗi tìm thấy: ${results.length}`;
  resultSummary.style.display = "block";
}

function escapeHtml(str) {
  if (str === null || str === undefined) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/* ---------------- EXPORT ---------------- */

btnExport.addEventListener("click", () => {
  if (!currentJobId) return;
  window.location.href = `/api/export/${currentJobId}`;
});
