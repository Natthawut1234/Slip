"use strict";

const elements = {
  slipInput: document.getElementById("slipInput"),
  scanBtn: document.getElementById("scanBtn"),
  clearBtn: document.getElementById("clearBtn"),
  importExcelBtn: document.getElementById("importExcelBtn"),
  importExcelInput: document.getElementById("importExcelInput"),
  exportBtn: document.getElementById("exportBtn"),
  selectAllBtn: document.getElementById("selectAllBtn"),
  deleteSelectedBtn: document.getElementById("deleteSelectedBtn"),
  statusText: document.getElementById("statusText"),
  progressBar: document.getElementById("progressBar"),
  fileCountText: document.getElementById("fileCountText"),
  resultBody: document.getElementById("resultBody"),
  dropZone: document.getElementById("dropZone"),
  sumRow: document.getElementById("sumRow"),
  sumAmount: document.getElementById("sumAmount"),
  receiptBtn: document.getElementById("receiptBtn")
};

const state = {
  files: [],
  busy: false,
  workerPool: [],
  workerPoolReady: false
};

const IMAGE_MAX_WIDTH = 700;
const OCR_PRIMARY_BOTTOM_RATIO = 0.52;
const WORKER_POOL_SIZE = 3;

const THAI_DIGITS = {
  "๐": "0",
  "๑": "1",
  "๒": "2",
  "๓": "3",
  "๔": "4",
  "๕": "5",
  "๖": "6",
  "๗": "7",
  "๘": "8",
  "๙": "9"
};

const MEMO_LABEL_HINTS = [
  "บันทึกช่วยจำ",
  "ช่วยจำ",
  "หมายเหตุ",
  "memo",
  "note",
  "บนทกชวยจา",
  "บนทกชวยจำ",
  "บนทกชวยจํา"
];

const MEMO_BLOCKLIST_HINTS = [
  "จำนวน",
  "ค่าธรรมเนียม",
  "เลขที่รายการ",
  "สแกนตรวจสอบสลิป",
  "โอนเงินสำเร็จ",
  "verified by"
];

const MEMO_LABEL_HINTS_NORMALIZED = MEMO_LABEL_HINTS.map((label) => normalizeForMatch(label));
const MEMO_BLOCKLIST_HINTS_NORMALIZED = MEMO_BLOCKLIST_HINTS.map((label) => normalizeForMatch(label));

elements.slipInput.addEventListener("change", onFileSelected);
elements.scanBtn.addEventListener("click", onScanClicked);
elements.clearBtn.addEventListener("click", confirmClearAll);
elements.importExcelBtn.addEventListener("click", onImportExcelClicked);
elements.importExcelInput.addEventListener("change", onImportExcelSelected);
elements.exportBtn.addEventListener("click", exportTableToExcel);
elements.selectAllBtn.addEventListener("click", toggleSelectAllRows);
elements.deleteSelectedBtn.addEventListener("click", confirmDeleteSelected);
elements.receiptBtn.addEventListener("click", exportReceipt);
elements.resultBody.addEventListener("change", onResultBodyChanged);
elements.resultBody.addEventListener("dblclick", onCellDoubleClick);

// Drag & Drop
setupDragAndDrop();

// Fix mobile table scroll — ให้เลื่อนแนวตั้งผ่านตารางได้เสมอ
setupTableTouchScroll();

setStatus("รออัปโหลดรูปสลิป", 0);
updateFileCount(0);
resetTable();
updateSelectionActions();
initWorkerPool();

function onFileSelected(event) {
  const files = Array.from(event.target.files || []).filter((file) => file.type.startsWith("image/"));
  state.files = files;
  updateFileCount(files.length);

  if (files.length === 0) {
    setStatus("ยังไม่ได้เลือกไฟล์", 0);
    return;
  }

  setStatus(`พร้อมอ่าน ${files.length} สลิป (ผลใหม่จะต่อท้ายตาราง)`, 0);
}

// ─── Worker Pool ────────────────────────────────────────────
async function initWorkerPool() {
  if (typeof Tesseract === "undefined") return;
  try {
    setStatus("กำลังเตรียม OCR engine…", 5);
    const promises = [];
    for (let i = 0; i < WORKER_POOL_SIZE; i++) {
      promises.push(createWorker());
    }
    state.workerPool = await Promise.all(promises);
    state.workerPoolReady = true;
    setStatus("รออัปโหลดรูปสลิป (OCR พร้อมแล้ว)", 0);
  } catch (err) {
    console.error("Worker pool init failed", err);
    state.workerPoolReady = false;
    setStatus("รออัปโหลดรูปสลิป", 0);
  }
}

async function createWorker() {
  const worker = await Tesseract.createWorker("eng+tha", 1, {
    logger: () => {}
  });
  return worker;
}

async function getWorker() {
  if (state.workerPool.length > 0) {
    return state.workerPool.pop();
  }
  return createWorker();
}

function returnWorker(worker) {
  state.workerPool.push(worker);
}

// ─── Drag & Drop ────────────────────────────────────────────
function setupDragAndDrop() {
  const dropZone = elements.dropZone;
  if (!dropZone) return;

  ["dragenter", "dragover"].forEach((ev) => {
    dropZone.addEventListener(ev, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.add("drag-over");
    });
  });

  ["dragleave", "drop"].forEach((ev) => {
    dropZone.addEventListener(ev, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.remove("drag-over");
    });
  });

  dropZone.addEventListener("drop", (e) => {
    const dt = e.dataTransfer;
    if (!dt?.files?.length) return;
    const files = Array.from(dt.files).filter((f) => f.type.startsWith("image/"));
    if (files.length === 0) return;
    state.files = files;
    updateFileCount(files.length);
    setStatus(`พร้อมอ่าน ${files.length} สลิป (ผลใหม่จะต่อท้ายตาราง)`, 0);
  });
}

// ─── Touch direction detection สำหรับ .table-wrap ───────────────
// ถ้าผู้ใช้เลื่อนแนวตั้ง → ปล่อยให้ page scroll ปกติ
// ถ้าเลื่อนแนวนอน → ให้ table scroll ตามปกติ
function setupTableTouchScroll() {
  const wrap = document.querySelector(".table-wrap");
  if (!wrap) return;

  let startX = 0;
  let startY = 0;
  let direction = null; // "h" | "v"

  wrap.addEventListener("touchstart", (e) => {
    const t = e.touches[0];
    startX = t.clientX;
    startY = t.clientY;
    direction = null;
    // คืน touch-action เป็นค่าเดิมทุกครั้งเริ่มสัมผัสใหม่
    wrap.style.touchAction = "";
  }, { passive: true });

  wrap.addEventListener("touchmove", (e) => {
    if (direction) return; // ตัดสินแล้ว

    const t = e.touches[0];
    const dx = Math.abs(t.clientX - startX);
    const dy = Math.abs(t.clientY - startY);

    // ต้องเลื่อนอย่างน้อย 8px ก่อนตัดสินทิศทาง
    if (dx < 8 && dy < 8) return;

    if (dy > dx) {
      // แนวตั้ง → ให้ page scroll ได้ (ล็อกแนวนอนของ table)
      direction = "v";
      wrap.style.touchAction = "pan-y";
    } else {
      // แนวนอน → ให้ table scroll ได้
      direction = "h";
      wrap.style.touchAction = "pan-x";
    }
  }, { passive: true });

  wrap.addEventListener("touchend", () => {
    direction = null;
    wrap.style.touchAction = "";
  }, { passive: true });
}

async function onScanClicked() {
  if (state.busy) {
    return;
  }

  if (state.files.length === 0) {
    setStatus("กรุณาเลือกไฟล์สลิปก่อน", 0);
    return;
  }

  if (typeof Tesseract === "undefined") {
    setStatus("ไม่พบ tesseract.js", 0);
    return;
  }

  state.busy = true;
  elements.slipInput.disabled = true;
  elements.scanBtn.disabled = true;
  elements.clearBtn.disabled = true;
  setTableActionDisabled(true);
  showScanSpinner(true);

  const total = state.files.length;
  const startOrder = getNextOrder();
  let completed = 0;

  try {
    // Parallel processing with worker pool
    const concurrency = Math.min(WORKER_POOL_SIZE, total);
    const results = new Array(total).fill(null);
    let nextIndex = 0;

    const runNext = async () => {
      while (nextIndex < total) {
        const idx = nextIndex++;
        const file = state.files[idx];
        try {
          results[idx] = await processSlipFile(file, idx, total);
        } catch (error) {
          console.error(error);
          results[idx] = { amount: "-", memo: "อ่านไม่สำเร็จ" };
        }
        completed++;
        const percent = Math.round((completed / total) * 100);
        setStatus(`อ่านแล้ว ${completed}/${total} สลิป`, percent);
      }
    };

    const workers = [];
    for (let i = 0; i < concurrency; i++) {
      workers.push(runNext());
    }
    await Promise.all(workers);

    // Append results in order
    for (let i = 0; i < total; i++) {
      const parsed = results[i] || { amount: "-", memo: "-" };
      appendResultRow(startOrder + i, parsed.amount || "-", parsed.memo || "-");
    }

    setStatus(`อ่านครบ ${total} สลิปแล้ว`, 100);
    state.files = [];
    elements.slipInput.value = "";
    updateFileCount(0);
    updateSumRow();
  } finally {
    state.busy = false;
    elements.slipInput.disabled = false;
    elements.scanBtn.disabled = false;
    elements.clearBtn.disabled = false;
    setTableActionDisabled(false);
    updateSelectionActions();
    showScanSpinner(false);
  }
}

async function processSlipFile(file, fileIndex, totalFiles) {
  const image = await fileToImage(file);

  const originalCanvas = document.createElement("canvas");
  drawImageFit(originalCanvas, image, IMAGE_MAX_WIDTH);

  const processedCanvas = document.createElement("canvas");
  preprocessCanvas(originalCanvas, processedCanvas);

  const primaryCanvas = buildBottomCropCanvas(processedCanvas, OCR_PRIMARY_BOTTOM_RATIO) || processedCanvas;

  const primaryText = await runOcr(primaryCanvas);

  const parsed = parseSlipText(primaryText);

  if (!parsed.amount || !parsed.memo) {
    const fallbackText = await runOcr(processedCanvas);

    const fallbackParsed = parseSlipText(fallbackText);
    if (!parsed.amount) {
      parsed.amount = fallbackParsed.amount;
    }
    if (!parsed.memo) {
      parsed.memo = fallbackParsed.memo;
    }
  }

  return parsed;
}

function mapFileProgress(fileIndex, totalFiles, localProgress) {
  const each = 100 / Math.max(1, totalFiles);
  const safeLocal = clamp(localProgress, 0, 1);
  return Math.round(fileIndex * each + safeLocal * each);
}

async function runOcr(canvas) {
  const worker = await getWorker();
  try {
    const result = await worker.recognize(canvas);
    return normalizeText(result?.data?.text || "");
  } finally {
    returnWorker(worker);
  }
}

// ─── Confirm Dialogs ────────────────────────────────────────
function confirmClearAll() {
  if (state.busy) return;
  const rows = getDataRows();
  if (rows.length === 0) {
    clearAll();
    return;
  }
  showConfirmModal("ต้องการเคลียร์ตารางทั้งหมดใช่ไหม?", clearAll);
}

function confirmDeleteSelected() {
  if (state.busy) return;
  const rows = getDataRows();
  const selectedRows = rows.filter((row) => {
    const check = row.querySelector(".row-check");
    return check instanceof HTMLInputElement && check.checked;
  });
  if (selectedRows.length === 0) {
    setStatus("ยังไม่ได้เลือกแถวที่จะลบ", getCurrentProgressValue());
    return;
  }
  showConfirmModal(`ต้องการลบ ${selectedRows.length} แถวที่เลือกใช่ไหม?`, deleteSelectedRows);
}

function showConfirmModal(message, onConfirm) {
  const existing = document.getElementById("confirmModal");
  if (existing) existing.remove();

  const overlay = document.createElement("div");
  overlay.id = "confirmModal";
  overlay.className = "modal-overlay";
  overlay.innerHTML = `
    <div class="modal-box">
      <p class="modal-message">${message}</p>
      <div class="modal-actions">
        <button class="modal-confirm" type="button">ยืนยัน</button>
        <button class="modal-cancel secondary" type="button">ยกเลิก</button>
      </div>
    </div>
  `;

  const close = () => overlay.remove();
  overlay.querySelector(".modal-confirm").addEventListener("click", () => { close(); onConfirm(); });
  overlay.querySelector(".modal-cancel").addEventListener("click", close);
  overlay.addEventListener("click", (e) => { if (e.target === overlay) close(); });
  document.body.appendChild(overlay);
  overlay.querySelector(".modal-confirm").focus();
}

// ─── Spinner ────────────────────────────────────────────────
function showScanSpinner(show) {
  const el = document.getElementById("scanSpinner");
  if (el) el.style.display = show ? "flex" : "none";
}

function clearAll() {
  state.files = [];
  state.busy = false;
  elements.slipInput.value = "";
  elements.importExcelInput.value = "";
  resetTable();
  updateFileCount(0);
  setStatus("ล้างข้อมูลแล้ว", 0);
  updateSelectionActions();
  updateSumRow();
}

function resetTable() {
  elements.resultBody.innerHTML = "";
  const row = document.createElement("tr");
  row.className = "empty-row";

  const cell = document.createElement("td");
  cell.colSpan = 4;
  cell.textContent = "ยังไม่มีผลลัพธ์";

  row.appendChild(cell);
  elements.resultBody.appendChild(row);
}

function appendResultRow(order, amount, memo) {
  const empty = elements.resultBody.querySelector(".empty-row");
  if (empty) {
    empty.remove();
  }

  const row = document.createElement("tr");
  row.className = "data-row";

  const selectCell = document.createElement("td");
  selectCell.className = "check-col";
  const checkbox = document.createElement("input");
  checkbox.type = "checkbox";
  checkbox.className = "row-check";
  checkbox.setAttribute("aria-label", `เลือกแถวสลิป ${order}`);
  selectCell.appendChild(checkbox);

  const orderCell = document.createElement("td");
  orderCell.textContent = String(order);

  const amountCell = document.createElement("td");
  amountCell.className = "editable-cell";
  amountCell.textContent = amount;
  amountCell.title = "ดับเบิลคลิกเพื่อแก้ไข";

  const memoCell = document.createElement("td");
  memoCell.className = "editable-cell";
  memoCell.textContent = memo;
  memoCell.title = "ดับเบิลคลิกเพื่อแก้ไข";

  row.appendChild(selectCell);
  row.appendChild(orderCell);
  row.appendChild(amountCell);
  row.appendChild(memoCell);
  elements.resultBody.appendChild(row);
  updateSelectionActions();
  updateSumRow();
}

// ─── Editable Cells ─────────────────────────────────────────
function onCellDoubleClick(event) {
  const cell = event.target.closest(".editable-cell");
  if (!cell || cell.querySelector("input")) return;

  const currentText = cell.textContent;
  const input = document.createElement("input");
  input.type = "text";
  input.className = "inline-edit";
  input.value = currentText;

  cell.textContent = "";
  cell.appendChild(input);
  input.focus();
  input.select();

  const save = () => {
    const newValue = input.value.trim() || "-";
    cell.textContent = newValue;
    updateSumRow();
  };

  input.addEventListener("blur", save);
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); input.blur(); }
    if (e.key === "Escape") { input.value = currentText; input.blur(); }
  });
}

function getDataRows() {
  return Array.from(elements.resultBody.querySelectorAll("tr.data-row"));
}

function getNextOrder() {
  return getDataRows().length + 1;
}

function onResultBodyChanged(event) {
  if (!(event.target instanceof HTMLInputElement)) {
    return;
  }
  if (!event.target.classList.contains("row-check")) {
    return;
  }
  updateSelectionActions();
}

function toggleSelectAllRows() {
  const rows = getDataRows();
  if (rows.length === 0) {
    return;
  }

  const checks = rows
    .map((row) => row.querySelector(".row-check"))
    .filter((el) => el instanceof HTMLInputElement);
  const allChecked = checks.length > 0 && checks.every((el) => el.checked);
  const nextValue = !allChecked;

  checks.forEach((el) => {
    el.checked = nextValue;
  });
  updateSelectionActions();
}

function deleteSelectedRows() {
  if (state.busy) {
    return;
  }

  const rows = getDataRows();
  if (rows.length === 0) {
    return;
  }

  const selectedRows = rows.filter((row) => {
    const check = row.querySelector(".row-check");
    return check instanceof HTMLInputElement && check.checked;
  });

  if (selectedRows.length === 0) {
    setStatus("ยังไม่ได้เลือกแถวที่จะลบ", clamp(parseInt(elements.progressBar.style.width, 10) || 0, 0, 100));
    return;
  }

  selectedRows.forEach((row) => row.remove());
  renumberRows();

  if (getDataRows().length === 0) {
    resetTable();
  }

  setStatus(`ลบแล้ว ${selectedRows.length} แถว`, clamp(parseInt(elements.progressBar.style.width, 10) || 0, 0, 100));
  updateSelectionActions();
  updateSumRow();
}

function renumberRows() {
  const rows = getDataRows();
  rows.forEach((row, index) => {
    const orderCell = row.children[1];
    if (orderCell) {
      orderCell.textContent = String(index + 1);
    }
    const check = row.querySelector(".row-check");
    if (check instanceof HTMLInputElement) {
      check.setAttribute("aria-label", `เลือกแถวสลิป ${index + 1}`);
    }
  });
}

function setTableActionDisabled(disabled) {
  elements.importExcelBtn.disabled = disabled;
  elements.exportBtn.disabled = disabled;
  elements.receiptBtn.disabled = disabled;
  elements.selectAllBtn.disabled = disabled;
  elements.deleteSelectedBtn.disabled = disabled;
}

function updateSelectionActions() {
  const rows = getDataRows();
  const checks = rows
    .map((row) => row.querySelector(".row-check"))
    .filter((el) => el instanceof HTMLInputElement);

  const hasRows = rows.length > 0;
  const selectedCount = checks.filter((el) => el.checked).length;
  const allChecked = hasRows && selectedCount === checks.length;

  elements.importExcelBtn.disabled = state.busy;
  elements.exportBtn.disabled = state.busy || selectedCount === 0;
  elements.receiptBtn.disabled = state.busy || selectedCount === 0;
  elements.selectAllBtn.textContent = allChecked ? "ยกเลิกเลือกทั้งหมด" : "เลือกทั้งหมด";
  elements.selectAllBtn.disabled = state.busy || !hasRows;
  elements.deleteSelectedBtn.disabled = state.busy || selectedCount === 0;
}

// ─── Sum Row ────────────────────────────────────────────────
function updateSumRow() {
  const rows = getDataRows();
  let total = 0;
  let count = 0;

  rows.forEach((row) => {
    const amountText = (row.children[2]?.textContent || "")
      .replace(/[^\d.,]/g, "")
      .replace(/,/g, "");
    const n = Number.parseFloat(amountText);
    if (Number.isFinite(n) && n > 0) {
      total += n;
      count++;
    }
  });

  if (elements.sumRow) {
    elements.sumRow.style.display = count > 0 ? "" : "none";
  }
  if (elements.sumAmount) {
    elements.sumAmount.textContent = count > 0
      ? `${total.toLocaleString("th-TH", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} บาท (${count} รายการ)`
      : "";
  }
}

function onImportExcelClicked() {
  if (state.busy) {
    return;
  }
  elements.importExcelInput.click();
}

async function onImportExcelSelected(event) {
  const file = event.target.files?.[0];
  event.target.value = "";

  if (!file) {
    return;
  }

  if (typeof XLSX === "undefined") {
    setStatus("ไม่พบไลบรารี Excel Import", getCurrentProgressValue());
    return;
  }

  try {
    setStatus(`กำลังนำเข้า Excel: ${file.name}`, getCurrentProgressValue());
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const firstSheetName = workbook.SheetNames?.[0];

    if (!firstSheetName) {
      setStatus("ไม่พบชีตในไฟล์ Excel", getCurrentProgressValue());
      return;
    }

    const sheet = workbook.Sheets[firstSheetName];
    const importedRows = parseImportedRows(sheet);

    if (importedRows.length === 0) {
      setStatus("ไฟล์ Excel ไม่มีข้อมูลที่นำเข้าได้", getCurrentProgressValue());
      return;
    }

    const startOrder = getNextOrder();
    importedRows.forEach((item, index) => {
      appendResultRow(startOrder + index, item.amount, item.memo);
    });

    setStatus(`นำเข้า Excel สำเร็จ (${importedRows.length} แถว)`, getCurrentProgressValue());
    updateSelectionActions();
    updateSumRow();
  } catch (error) {
    console.error(error);
    setStatus("นำเข้า Excel ไม่สำเร็จ", getCurrentProgressValue());
  }
}

function parseImportedRows(sheet) {
  const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  const amountAliases = ["จำนวนเงิน", "amount"].map((v) => normalizeImportKey(v));
  const memoAliases = ["บันทึกช่วยจำ", "memo", "note"].map((v) => normalizeImportKey(v));
  const out = [];

  rawRows.forEach((row) => {
    const amount = cleanImportCell(getCellByAliases(row, amountAliases));
    const memo = cleanImportCell(getCellByAliases(row, memoAliases));

    if (!amount && !memo) {
      return;
    }

    out.push({
      amount: amount || "-",
      memo: memo || "-"
    });
  });

  return out;
}

function getCellByAliases(row, aliasKeys) {
  const entries = Object.entries(row || {});
  for (const [key, value] of entries) {
    const normalizedKey = normalizeImportKey(key);
    const matched = aliasKeys.some((alias) => normalizedKey === alias || normalizedKey.includes(alias));
    if (matched) {
      return value;
    }
  }
  return "";
}

function cleanImportCell(value) {
  return String(value ?? "").trim();
}

function normalizeImportKey(value) {
  return normalizeForMatch(String(value || "")).replace(/[:/ ]+/g, "");
}

function exportTableToExcel() {
  if (state.busy) {
    return;
  }

  const rows = getDataRows();
  if (rows.length === 0) {
    setStatus("ยังไม่มีข้อมูลในตารางสำหรับ Export", getCurrentProgressValue());
    return;
  }

  if (typeof XLSX === "undefined") {
    setStatus("ไม่พบไลบรารี Excel Export", getCurrentProgressValue());
    return;
  }

  const selectedRows = rows.filter((row) => {
    const check = row.querySelector(".row-check");
    return check instanceof HTMLInputElement && check.checked;
  });

  if (selectedRows.length === 0) {
    setStatus("กรุณาเลือกแถวก่อน Export", getCurrentProgressValue());
    return;
  }

  const excelRows = selectedRows.map((row) => {
    const cells = row.querySelectorAll("td");
    return {
      ลำดับ: (cells[1]?.textContent || "").trim(),
      จำนวนเงิน: (cells[2]?.textContent || "").trim(),
      บันทึกช่วยจำ: (cells[3]?.textContent || "").trim()
    };
  });

  const sheet = XLSX.utils.json_to_sheet(excelRows);
  sheet["!cols"] = [{ wch: 10 }, { wch: 18 }, { wch: 48 }];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "SlipResults");

  const filename = `slip-results-${buildTimestampForFilename()}.xlsx`;
  XLSX.writeFile(workbook, filename, { compression: true });
  setStatus(`Export Excel สำเร็จ (${excelRows.length} แถว)`, getCurrentProgressValue());
}

// ─── Receipt Export ──────────────────────────────────────────
function exportReceipt() {
  if (state.busy) return;

  const rows = getDataRows();
  const selectedRows = rows.filter((row) => {
    const check = row.querySelector(".row-check");
    return check instanceof HTMLInputElement && check.checked;
  });

  if (selectedRows.length === 0) {
    setStatus("กรุณาเลือกแถวก่อนพิมพ์ใบเสร็จ", getCurrentProgressValue());
    return;
  }

  // สร้างข้อมูลใบเสร็จแยกแต่ละรายการ
  const receipts = selectedRows.map((row) => {
    const cells = row.querySelectorAll("td");
    return {
      order: (cells[1]?.textContent || "").trim(),
      amount: (cells[2]?.textContent || "").trim(),
      memo: (cells[3]?.textContent || "").trim()
    };
  });

  showReceiptModal(receipts, 0);
}

function buildReceiptHtml(item, index, total) {
  const now = new Date();
  const dateStr = now.toLocaleDateString("th-TH", {
    year: "numeric", month: "long", day: "numeric"
  });
  const timeStr = now.toLocaleTimeString("th-TH", {
    hour: "2-digit", minute: "2-digit"
  });
  const receiptNo = `RCP-${buildTimestampForFilename()}-${String(index + 1).padStart(3, "0")}`;

  const amountNum = parseNumber(item.amount);
  const amountFormatted = Number.isFinite(amountNum)
    ? amountNum.toLocaleString("th-TH", { minimumFractionDigits: 2, maximumFractionDigits: 2 })
    : item.amount;

  return `
    <div class="receipt-content">
      <div class="receipt-header">
        <div class="receipt-logo">
          <span class="logo-text">The Best Village</span>
        </div>
        <h3>เดอะเบสท์วิลเลจ</h3>
        <p>ใบเสร็จรับเงิน / Receipt</p>
        <p style="margin-top:4px;">${dateStr} เวลา ${timeStr}</p>
        <p>เลขที่: ${receiptNo}</p>
      </div>
      <table class="receipt-table">
        <thead>
          <tr>
            <th>รายการ</th>
            <th>จำนวนเงิน (฿)</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td class="receipt-memo-col">${escapeHtml(item.memo) || "-"}</td>
            <td>${escapeHtml(item.amount)}</td>
          </tr>
        </tbody>
      </table>
      <div class="receipt-total">
        <span>ยอดรวม</span>
        <span>฿ ${amountFormatted}</span>
      </div>
      <div class="receipt-footer">
        <p>ขอบคุณที่ใช้บริการ</p>
      </div>
    </div>
  `;
}
// <p>เอกสารนี้ออกโดย Slip Manager</p>

function showReceiptModal(receipts, currentIndex) {
  const existing = document.getElementById("receiptModal");
  if (existing) existing.remove();

  const total = receipts.length;
  const item = receipts[currentIndex];
  const contentHtml = buildReceiptHtml(item, currentIndex, total);

  const hasPrev = currentIndex > 0;
  const hasNext = currentIndex < total - 1;

  const overlay = document.createElement("div");
  overlay.id = "receiptModal";
  overlay.className = "receipt-overlay";
  overlay.innerHTML = `
    <div class="receipt-modal">
      <div class="receipt-toolbar">
        ${total > 1 ? `
        <button class="receipt-prev ghost-btn" type="button" ${!hasPrev ? "disabled" : ""} style="flex:0;padding:10px 12px;font-size:1.1rem;">
          ◀
        </button>
        <span class="receipt-counter">${currentIndex + 1} / ${total}</span>
        <button class="receipt-next ghost-btn" type="button" ${!hasNext ? "disabled" : ""} style="flex:0;padding:10px 12px;font-size:1.1rem;">
          ▶
        </button>
        ` : ""}
        <button class="receipt-print btn-scan" type="button">
          <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><path d="M4 6V1h8v5"/><path d="M4 12H2a1 1 0 01-1-1V7a1 1 0 011-1h12a1 1 0 011 1v4a1 1 0 01-1 1h-2"/><rect x="4" y="9" width="8" height="5" rx="1"/></svg>
          พิมพ์
        </button>
        <button class="receipt-save ghost-btn" type="button">
          <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><path d="M8 10V1M4 6l4 4 4-4"/><path d="M1 12v2a1 1 0 001 1h12a1 1 0 001-1v-2"/></svg>
          บันทึก PNG
        </button>
        ${total > 1 ? `
        <button class="receipt-save-all ghost-btn" type="button">
          <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><path d="M8 10V1M4 6l4 4 4-4"/><path d="M1 12v2a1 1 0 001 1h12a1 1 0 001-1v-2"/></svg>
          บันทึกทั้งหมด
        </button>
        ` : ""}
        <button class="receipt-close" type="button" style="flex:0;padding:10px 14px;background:transparent;color:var(--text-mid);border:1px solid var(--line);">
          ✕
        </button>
      </div>
      <div class="receipt-body">
        ${contentHtml}
      </div>
    </div>
  `;

  const close = () => overlay.remove();
  overlay.querySelector(".receipt-close").addEventListener("click", close);
  overlay.addEventListener("click", (e) => { if (e.target === overlay) close(); });

  // Navigation
  if (total > 1) {
    const prevBtn = overlay.querySelector(".receipt-prev");
    const nextBtn = overlay.querySelector(".receipt-next");
    if (prevBtn) prevBtn.addEventListener("click", () => showReceiptModal(receipts, currentIndex - 1));
    if (nextBtn) nextBtn.addEventListener("click", () => showReceiptModal(receipts, currentIndex + 1));

    // Save all — download all receipts as PNG one by one
    const saveAllBtn = overlay.querySelector(".receipt-save-all");
    if (saveAllBtn) {
      saveAllBtn.addEventListener("click", async () => {
        saveAllBtn.disabled = true;
        saveAllBtn.textContent = "กำลังบันทึก…";
        await saveAllReceiptsAsImages(receipts);
        saveAllBtn.textContent = "✓ เสร็จแล้ว";
        setTimeout(() => {
          saveAllBtn.disabled = false;
          saveAllBtn.textContent = "บันทึกทั้งหมด";
        }, 1500);
      });
    }
  }

  overlay.querySelector(".receipt-print").addEventListener("click", () => {
    printReceiptContent(overlay.querySelector(".receipt-content"));
  });

  overlay.querySelector(".receipt-save").addEventListener("click", () => {
    saveReceiptAsImage(receipts[currentIndex], currentIndex);
  });

  document.body.appendChild(overlay);

  // Keyboard navigation
  const onKey = (e) => {
    if (e.key === "Escape") { close(); document.removeEventListener("keydown", onKey); }
    if (e.key === "ArrowLeft" && hasPrev) showReceiptModal(receipts, currentIndex - 1);
    if (e.key === "ArrowRight" && hasNext) showReceiptModal(receipts, currentIndex + 1);
  };
  document.addEventListener("keydown", onKey);
}

async function saveAllReceiptsAsImages(receipts) {
  for (let i = 0; i < receipts.length; i++) {
    await saveReceiptAsImagePromise(receipts[i], i);
    await new Promise((r) => setTimeout(r, 400));
  }
}

function saveReceiptAsImagePromise(item, index) {
  return new Promise((resolve) => {
    saveReceiptAsImage(item, index, resolve);
  });
}

// ─── Print: เปิด window ใหม่มี inline style ครบ ─────────────
function printReceiptContent(contentEl) {
  if (!contentEl) return;
  const html = contentEl.outerHTML;
  const printWin = window.open("", "_blank", "width=450,height=650");
  if (!printWin) { alert("กรุณาอนุญาต popup เพื่อพิมพ์ใบเสร็จ"); return; }
  printWin.document.write(`<!DOCTYPE html>
<html lang="th">
<head>
  <meta charset="utf-8">
  <title>ใบเสร็จ</title>
  <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap" rel="stylesheet">
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'Sarabun', sans-serif; background: #fff; color: #1a1a1a; padding: 20px; }
    .receipt-content { max-width: 400px; margin: 0 auto; }
    .receipt-header { text-align: center; margin-bottom: 18px; padding-bottom: 14px; border-bottom: 2px dashed #ccc; }
    .receipt-logo { display: flex; flex-direction: column; align-items: center; gap: 4px; margin-bottom: 6px; }

    .receipt-logo .logo-text { font-family: Georgia, 'Times New Roman', serif; font-size: 1.15rem; font-weight: 700; font-style: italic; color: #0e0e0e; letter-spacing: 2px; text-shadow: 0 1px 1px rgba(0, 0, 0, 0.18); }
    .receipt-header h3 { font-size: 1.3rem; font-weight: 700; color: #0e0e0e; margin-bottom: 1px; }
    .receipt-header p { font-size: 0.85rem; color: #888; margin-top: 2px; }
    .receipt-table { width: 100%; border-collapse: collapse; margin-bottom: 16px; }
    .receipt-table thead th { text-align: left; font-size: 0.8rem; font-weight: 700; text-transform: uppercase; color: #666; border-bottom: 1px solid #ddd; padding: 8px 6px; }
    .receipt-table thead th:last-child { text-align: right; }
    .receipt-table tbody td { padding: 10px 6px; border-bottom: 1px solid #eee; font-size: 0.92rem; }
    .receipt-table tbody td:last-child { text-align: right; font-weight: 600; font-variant-numeric: tabular-nums; }
    .receipt-memo-col { max-width: 200px; word-break: break-word; }
    .receipt-total { display: flex; justify-content: space-between; padding: 14px 0; border-top: 2px dashed #ccc; margin-top: 4px; }
    .receipt-total span:first-child { font-weight: 700; font-size: 1rem; }
    .receipt-total span:last-child { font-weight: 700; font-size: 1.15rem; color: #0ea5a4; }
    .receipt-footer { text-align: center; margin-top: 16px; padding-top: 12px; border-top: 1px solid #eee; font-size: 0.8rem; color: #aaa; }
    .receipt-footer p { margin: 2px 0; }
    @page { margin: 0; size: auto; }
    @media print {
      html, body { margin: 0; padding: 15mm; }
    }
  </style>
</head>
<body>${html}</body>
</html>`);
  printWin.document.close();
  // รอโหลด font ก่อนพิมพ์
  setTimeout(() => {
    printWin.focus();
    printWin.print();
    setTimeout(() => printWin.close(), 500);
  }, 600);
}

// ─── Save PNG: วาด canvas โดยตรง (ไม่ใช้ foreignObject) ─────
function saveReceiptAsImage(item, index, onDone) {
  const canvas = document.createElement("canvas");
  const scale = 2;
  const W = 380;
  const pad = 28;
  const lineH = 22;
  const ctx = canvas.getContext("2d");

  // Pre-calculate layout
  const amountNum = parseNumber(item.amount);
  const amountStr = Number.isFinite(amountNum)
    ? amountNum.toLocaleString("th-TH", { minimumFractionDigits: 2, maximumFractionDigits: 2 })
    : item.amount;
  const memo = item.memo || "-";
  const now = new Date();
  const dateStr = now.toLocaleDateString("th-TH", { year: "numeric", month: "long", day: "numeric" });
  const timeStr = now.toLocaleTimeString("th-TH", { hour: "2-digit", minute: "2-digit" });
  const receiptNo = `RCP-${buildTimestampForFilename()}-${String((index || 0) + 1).padStart(3, "0")}`;

  // Wrap memo text
  const maxMemoW = W - pad * 2 - 10;
  const memoLines = wrapText(ctx, memo, maxMemoW, "15px Sarabun, sans-serif");

  // Calculate height
  let H = 0;
  H += pad;       // top padding
  H += 28;        // logo text
  H += 8;         // gap after logo
  H += 20;        // subtitle
  H += 18;        // date
  H += 18;        // receipt no
  H += 18;        // dashed line gap
  H += 4;         // spacing
  H += 30;        // table header
  H += Math.max(1, memoLines.length) * lineH + 16; // table row
  H += 4;         // spacing
  H += 18;        // dashed line
  H += 36;        // total
  H += 18;        // dashed line
  H += 26;        // footer
  H += pad;       // bottom padding

  canvas.width = W * scale;
  canvas.height = H * scale;
  ctx.scale(scale, scale);

  // Background
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, W, H);

  let y = pad;

  // Logo text "The Best Village"
  ctx.fillStyle = "#0e0e0e";
  ctx.font = "italic bold 18px Georgia, 'Times New Roman', serif";
  ctx.textAlign = "center";
  ctx.letterSpacing = "2px";
  ctx.fillText("The Best Village", W / 2, y + 18);
  ctx.letterSpacing = "0px";
  y += 28;
  y += 8;

  ctx.fillStyle = "#0e0e0e";
  ctx.font = "bold 18px Sarabun, sans-serif";
  ctx.textAlign = "center";
  ctx.letterSpacing = "2px";
  ctx.fillText("เดอะเบสท์วิลเลจ", W / 2, y + 18);
  ctx.letterSpacing = "0px";
  y += 28;
  y += 8;

  // Subtitle
  ctx.fillStyle = "#888888";
  ctx.font = "14px Sarabun, sans-serif";
  ctx.fillText("\u0e43\u0e1a\u0e40\u0e2a\u0e23\u0e47\u0e08\u0e23\u0e31\u0e1a\u0e40\u0e07\u0e34\u0e19 / Receipt", W / 2, y + 14);
  y += 20;

  // Date
  ctx.fillText(`${dateStr} \u0e40\u0e27\u0e25\u0e32 ${timeStr}`, W / 2, y + 14);
  y += 18;

  // Receipt no
  ctx.fillText(`\u0e40\u0e25\u0e02\u0e17\u0e35\u0e48: ${receiptNo}`, W / 2, y + 14);
  y += 18;

  // Dashed line
  y += 8;
  drawDashedLine(ctx, pad, y, W - pad, y, "#cccccc");
  y += 10;

  // Table header
  ctx.textAlign = "left";
  ctx.fillStyle = "#666666";
  ctx.font = "bold 12px Sarabun, sans-serif";
  ctx.fillText("\u0e23\u0e32\u0e22\u0e01\u0e32\u0e23", pad, y + 18);
  ctx.textAlign = "right";
  ctx.fillText("\u0e08\u0e33\u0e19\u0e27\u0e19\u0e40\u0e07\u0e34\u0e19 (\u0e3f)", W - pad, y + 18);
  y += 24;
  ctx.strokeStyle = "#dddddd";
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(pad, y);
  ctx.lineTo(W - pad, y);
  ctx.stroke();
  y += 6;

  // Table row — memo + amount
  ctx.font = "15px Sarabun, sans-serif";
  ctx.fillStyle = "#1a1a1a";
  ctx.textAlign = "left";
  memoLines.forEach((line, i) => {
    ctx.fillText(line, pad, y + 16 + i * lineH);
  });

  ctx.textAlign = "right";
  ctx.font = "bold 15px Sarabun, sans-serif";
  ctx.fillText(item.amount, W - pad, y + 16);

  y += Math.max(1, memoLines.length) * lineH + 10;

  // Bottom border
  ctx.strokeStyle = "#eeeeee";
  ctx.beginPath();
  ctx.moveTo(pad, y);
  ctx.lineTo(W - pad, y);
  ctx.stroke();
  y += 4;

  // Dashed total line
  drawDashedLine(ctx, pad, y, W - pad, y, "#cccccc");
  y += 14;

  // Total
  ctx.textAlign = "left";
  ctx.fillStyle = "#333333";
  ctx.font = "bold 16px Sarabun, sans-serif";
  ctx.fillText("\u0e22\u0e2d\u0e14\u0e23\u0e27\u0e21", pad, y + 16);
  ctx.textAlign = "right";
  ctx.fillStyle = "#0ea5a4";
  ctx.font = "bold 18px Sarabun, sans-serif";
  ctx.fillText(`\u0e3f ${amountStr}`, W - pad, y + 16);
  y += 22;

  // Footer dashed
  drawDashedLine(ctx, pad, y, W - pad, y, "#eeeeee");
  y += 14;

  // Footer
  ctx.textAlign = "center";
  ctx.fillStyle = "#aaaaaa";
  ctx.font = "12px Sarabun, sans-serif";
  ctx.fillText("\u0e02\u0e2d\u0e1a\u0e04\u0e38\u0e13\u0e17\u0e35\u0e48\u0e43\u0e0a\u0e49\u0e1a\u0e23\u0e34\u0e01\u0e32\u0e23", W / 2, y + 12);

  // Download
  canvas.toBlob((b) => {
    if (!b) { if (onDone) onDone(); return; }
    const a = document.createElement("a");
    a.href = URL.createObjectURL(b);
    const suffix = typeof index === "number" ? `-${String(index + 1).padStart(3, "0")}` : "";
    a.download = `receipt-${buildTimestampForFilename()}${suffix}.png`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    if (onDone) onDone();
  }, "image/png");
}

function wrapText(ctx, text, maxWidth, font) {
  ctx.font = font;
  const words = text.split("");
  const lines = [];
  let current = "";
  for (const char of words) {
    const test = current + char;
    if (ctx.measureText(test).width > maxWidth && current.length > 0) {
      lines.push(current);
      current = char;
    } else {
      current = test;
    }
  }
  if (current) lines.push(current);
  return lines.length > 0 ? lines : [""];
}

function drawDashedLine(ctx, x1, y1, x2, y2, color) {
  ctx.save();
  ctx.strokeStyle = color;
  ctx.lineWidth = 1;
  ctx.setLineDash([4, 3]);
  ctx.beginPath();
  ctx.moveTo(x1, y1);
  ctx.lineTo(x2, y2);
  ctx.stroke();
  ctx.restore();
}

function escapeHtml(text) {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

function buildTimestampForFilename() {
  const now = new Date();
  const yyyy = now.getFullYear();
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const dd = String(now.getDate()).padStart(2, "0");
  const hh = String(now.getHours()).padStart(2, "0");
  const mi = String(now.getMinutes()).padStart(2, "0");
  const ss = String(now.getSeconds()).padStart(2, "0");
  return `${yyyy}${mm}${dd}-${hh}${mi}${ss}`;
}

function getCurrentProgressValue() {
  const width = (elements.progressBar.style.width || "").replace("%", "");
  const n = Number.parseInt(width, 10);
  return Number.isFinite(n) ? clamp(n, 0, 100) : 0;
}

function updateFileCount(count) {
  if (count <= 0) {
    elements.fileCountText.textContent = "ยังไม่ได้เลือกไฟล์ใหม่ (ข้อมูลในตารางยังอยู่)";
    return;
  }
  elements.fileCountText.textContent = `เลือกแล้ว ${count} ไฟล์`;
}

function fileToImage(file) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    const objectUrl = URL.createObjectURL(file);

    image.onload = () => {
      URL.revokeObjectURL(objectUrl);
      resolve(image);
    };

    image.onerror = () => {
      URL.revokeObjectURL(objectUrl);
      reject(new Error("Cannot load image"));
    };

    image.src = objectUrl;
  });
}

function drawImageFit(canvas, image, maxWidth) {
  const ratio = image.width > maxWidth ? maxWidth / image.width : 1;
  const width = Math.max(1, Math.round(image.width * ratio));
  const height = Math.max(1, Math.round(image.height * ratio));

  canvas.width = width;
  canvas.height = height;

  const ctx = canvas.getContext("2d");
  ctx.clearRect(0, 0, width, height);
  ctx.drawImage(image, 0, 0, width, height);
}

function preprocessCanvas(inputCanvas, outputCanvas) {
  outputCanvas.width = inputCanvas.width;
  outputCanvas.height = inputCanvas.height;

  const ctx = outputCanvas.getContext("2d", { willReadFrequently: true });
  ctx.drawImage(inputCanvas, 0, 0);

  const imageData = ctx.getImageData(0, 0, outputCanvas.width, outputCanvas.height);
  const data = imageData.data;
  const contrast = 1.75;
  const bias = 8;

  for (let i = 0; i < data.length; i += 4) {
    const gray = data[i] * 0.299 + data[i + 1] * 0.587 + data[i + 2] * 0.114;
    let enhanced = (gray - 128) * contrast + 128 + bias;

    if (enhanced > 170) {
      enhanced = 255;
    } else if (enhanced < 75) {
      enhanced = 0;
    }

    const v = clamp(enhanced, 0, 255);
    data[i] = v;
    data[i + 1] = v;
    data[i + 2] = v;
  }

  ctx.putImageData(imageData, 0, 0);
}

function parseSlipText(rawText) {
  const lines = splitTextToLines(rawText);
  const amount = extractAmount(rawText);
  const memo = extractMemo(lines);

  return {
    amount,
    memo
  };
}

function extractAmount(text) {
  const keywordPatterns = [
    /(?:จำนวนเงิน|จำนวน|จํานวน|ยอดโอน|amount|total)\s*[:\-]?\s*([0-9][0-9,]*(?:\.[0-9]{1,2})?)/i,
    /([0-9][0-9,]*(?:\.[0-9]{1,2})?)\s*(?:บาท|baht)/i,
    /฿\s*([0-9][0-9,]*(?:\.[0-9]{1,2})?)/i
  ];

  for (const pattern of keywordPatterns) {
    const match = pattern.exec(text);
    if (!match) {
      continue;
    }

    const n = parseNumber(match[1]);
    if (n !== null) {
      return `${n.toLocaleString("th-TH", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
      })} บาท`;
    }
  }

  const allMatches = [...text.matchAll(/\b([0-9][0-9,]*\.[0-9]{2})\b/g)];
  let best = null;

  for (const match of allMatches) {
    const n = parseNumber(match[1]);
    if (n === null) {
      continue;
    }
    if (best === null || n > best) {
      best = n;
    }
  }

  if (best !== null) {
    return `${best.toLocaleString("th-TH", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    })} บาท`;
  }

  return "";
}

function extractMemo(lines) {
  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i] || "";
    const nextLine = lines[i + 1] || "";
    const memo = extractMemoAroundLabel(line, nextLine, MEMO_LABEL_HINTS_NORMALIZED);
    if (memo) {
      return memo;
    }
  }

  for (let i = lines.length - 1; i >= 0; i -= 1) {
    const rawLine = cleanMemo(lines[i]);
    if (!rawLine) {
      continue;
    }

    const normalizedLine = normalizeForMatch(rawLine);
    if (!isLikelyMemoContent(rawLine, normalizedLine)) {
      continue;
    }
    return rawLine;
  }

  return "";
}

function cleanMemo(value) {
  const memo = value
    .replace(/\s+/g, " ")
    .trim()
    .replace(/^[:：\-\s]+/, "");

  const normalizedMemo = normalizeThaiMemoSpacing(memo);
  const calendarFixedMemo = normalizeThaiCalendarInMemo(normalizedMemo);

  if (!calendarFixedMemo) {
    return "";
  }
  if (/^[-.]+$/.test(calendarFixedMemo)) {
    return "";
  }
  return calendarFixedMemo;
}

function normalizeThaiMemoSpacing(value) {
  if (!value) {
    return "";
  }

  let output = value
    .replace(/\u200b/g, "")
    .replace(/\s+([,./])/g, "$1")
    .replace(/([/])\s+/g, "$1");

  let previous = "";
  while (output !== previous) {
    previous = output;
    output = output.replace(/([\u0E00-\u0E7F])\s+(?=[\u0E00-\u0E7F])/g, "$1");
  }

  return output.trim();
}

function normalizeThaiCalendarInMemo(value) {
  if (!value) {
    return "";
  }

  let out = value
    .replace(/เดือน(?=[ก-๙])/g, "เดือน ")
    .replace(/\s*,\s*/g, ", ")
    .replace(/\.\s+(\d{4})/g, ".$1");

  // Handle common OCR confusion where "ก.พ." is read as "ท.พ.", "N.พ." or "N.W.".
  out = replaceThaiMonthToken(
    out,
    "[กทฑตดNnHhMmWw]\\s*\\.?\\s*[พPwW]\\s*\\.?",
    "ก.พ."
  );

  const monthRules = [
    { token: "ม\\s*\\.?\\s*ค\\s*\\.?", canonical: "ม.ค." },
    { token: "มี\\s*\\.?\\s*ค\\s*\\.?", canonical: "มี.ค." },
    { token: "เม\\s*\\.?\\s*ย\\s*\\.?", canonical: "เม.ย." },
    { token: "พ\\s*\\.?\\s*ย\\s*\\.?", canonical: "พ.ย." },
    { token: "พ\\s*\\.?\\s*ค\\s*\\.?", canonical: "พ.ค." },
    { token: "มิ\\s*\\.?\\s*ย\\s*\\.?", canonical: "มิ.ย." },
    { token: "ก\\s*\\.?\\s*ค\\s*\\.?", canonical: "ก.ค." },
    { token: "ส\\s*\\.?\\s*ค\\s*\\.?", canonical: "ส.ค." },
    { token: "ก\\s*\\.?\\s*ย\\s*\\.?", canonical: "ก.ย." },
    { token: "ต\\s*\\.?\\s*ค\\s*\\.?", canonical: "ต.ค." },
    { token: "ธ\\s*\\.?\\s*ค\\s*\\.?", canonical: "ธ.ค." }
  ];

  monthRules.forEach((rule) => {
    out = replaceThaiMonthToken(out, rule.token, rule.canonical);
  });

  return out.replace(/\s{2,}/g, " ").trim();
}

function replaceThaiMonthToken(value, tokenPattern, canonical) {
  const pattern = new RegExp(
    `(^|[^ก-๙A-Za-z0-9])(?:${tokenPattern})(?=(?:\\d{2,4})?(?:[^ก-๙A-Za-z0-9]|$))`,
    "g"
  );
  return value.replace(pattern, `$1${canonical}`);
}

function extractMemoAroundLabel(line, nextLine, hints) {
  const rawLine = cleanMemo(line);
  if (!rawLine) {
    return "";
  }

  const normalizedLine = normalizeForMatch(rawLine);
  const hasLabel = hints.some((hint) => hint && normalizedLine.includes(hint));
  if (!hasLabel) {
    return "";
  }

  const afterColon = rawLine.match(/[:：]\s*(.+)$/);
  if (afterColon?.[1]) {
    const candidate = cleanMemo(afterColon[1]);
    if (candidate && !isBlockedMemoLine(normalizeForMatch(candidate)) && !looksLikeAmount(candidate) && !looksLikeTimeFragment(candidate) && !looksLikeUiNoiseLine(candidate)) {
      return candidate;
    }
  }

  const compactTail = cleanMemo(rawLine.replace(/^[^:：\s]+[\s]*/, ""));
  if (compactTail && compactTail !== rawLine && !isBlockedMemoLine(normalizeForMatch(compactTail)) && !looksLikeAmount(compactTail) && !looksLikeTimeFragment(compactTail) && !looksLikeUiNoiseLine(compactTail)) {
    return compactTail;
  }

  const fromNextLine = cleanMemo(nextLine);
  if (fromNextLine && !isBlockedMemoLine(normalizeForMatch(fromNextLine)) && !looksLikeAmount(fromNextLine) && !looksLikeTimeFragment(fromNextLine) && !looksLikeUiNoiseLine(fromNextLine)) {
    return fromNextLine;
  }

  return "";
}

function isBlockedMemoLine(normalizedLine) {
  return MEMO_BLOCKLIST_HINTS_NORMALIZED.some((hint) => hint && normalizedLine.includes(hint));
}

function looksLikeAmount(text) {
  const value = (text || "").replace(/,/g, "").trim();
  return /^[0-9]+(?:\.[0-9]{1,2})?(?:\s*(บาท|baht))?$/i.test(value);
}

function looksLikeTimeFragment(text) {
  const value = normalizeThaiMemoSpacing((text || "").toLowerCase());
  if (!value) {
    return false;
  }

  if (/^([01]?\d|2[0-3])[:.][0-5]\d(?:\s*น\.?)?$/i.test(value)) {
    return true;
  }

  if (/^[0-5]?\d\s*น\.?$/i.test(value)) {
    return true;
  }

  return /(?:^|\s)([01]?\d|2[0-3])[:.][0-5]\d(?:\s*น\.?)?(?:$|\s)/i.test(value);
}

function looksLikeUiNoiseLine(text) {
  const compact = (text || "").toLowerCase().replace(/\s+/g, "");
  if (!compact) {
    return false;
  }
  return (
    compact.includes("ik+") ||
    compact.includes("k+") ||
    compact.includes("kplus") ||
    compact.includes("verified")
  );
}

function isLikelyMemoContent(rawLine, normalizedLine) {
  if (!rawLine) {
    return false;
  }

  if (isBlockedMemoLine(normalizedLine)) {
    return false;
  }

  if (looksLikeAmount(rawLine) || looksLikeTimeFragment(rawLine) || looksLikeUiNoiseLine(rawLine)) {
    return false;
  }

  const hasDigits = /\d/.test(rawLine);
  const hasThai = /[ก-๙]/.test(rawLine);
  const hasLatin = /[a-z]/i.test(rawLine);
  const hasSlashPattern = /\d+\/\d+/.test(rawLine);
  const hasMonth = normalizedLine.includes("เดอน") || normalizedLine.includes("month");

  if (hasMonth && hasDigits) {
    return true;
  }

  if (hasSlashPattern && hasDigits && (hasThai || hasLatin)) {
    return true;
  }

  return false;
}

function splitTextToLines(rawText) {
  return (rawText || "")
    .split(/\r?\n/)
    .map((line) => line.replace(/\u00a0/g, " ").trim())
    .filter(Boolean);
}

function buildBottomCropCanvas(sourceCanvas, ratio) {
  const width = sourceCanvas.width;
  const height = sourceCanvas.height;

  if (!width || !height) {
    return null;
  }

  const cropHeight = Math.max(140, Math.round(height * ratio));
  const startY = Math.max(0, height - cropHeight);

  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = cropHeight;

  const ctx = canvas.getContext("2d");
  ctx.drawImage(sourceCanvas, 0, startY, width, cropHeight, 0, 0, width, cropHeight);
  return canvas;
}

function setStatus(text, progress) {
  elements.statusText.textContent = text;
  const p = clamp(progress, 0, 100);
  elements.progressBar.style.width = `${p}%`;
}

function parseNumber(value) {
  if (typeof value !== "string") {
    return null;
  }
  const normalized = value.replace(/,/g, "").trim();
  const number = Number.parseFloat(normalized);
  if (!Number.isFinite(number)) {
    return null;
  }
  return number;
}

function normalizeText(value) {
  let text = value || "";
  for (const [thai, arabic] of Object.entries(THAI_DIGITS)) {
    text = text.split(thai).join(arabic);
  }
  return text
    .replace(/\u0E4D\u0E32/g, "ำ")
    .replace(/\u00a0/g, " ");
}

function normalizeForMatch(value) {
  return (value || "")
    .toLowerCase()
    .normalize("NFKD")
    .replace(/[\u0E31\u0E34-\u0E3A\u0E47-\u0E4E]/g, "")
    .replace(/\u0E33/g, "า")
    .replace(/[^a-z0-9ก-๙:\/ ]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}
