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
  resultBody: document.getElementById("resultBody")
};

const state = {
  files: [],
  busy: false
};

const IMAGE_MAX_WIDTH = 960;
const OCR_PRIMARY_BOTTOM_RATIO = 0.52;

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
elements.clearBtn.addEventListener("click", clearAll);
elements.importExcelBtn.addEventListener("click", onImportExcelClicked);
elements.importExcelInput.addEventListener("change", onImportExcelSelected);
elements.exportBtn.addEventListener("click", exportTableToExcel);
elements.selectAllBtn.addEventListener("click", toggleSelectAllRows);
elements.deleteSelectedBtn.addEventListener("click", deleteSelectedRows);
elements.resultBody.addEventListener("change", onResultBodyChanged);

setStatus("รออัปโหลดรูปสลิป", 0);
updateFileCount(0);
resetTable();
updateSelectionActions();

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

  const total = state.files.length;
  const startOrder = getNextOrder();

  try {
    for (let i = 0; i < total; i += 1) {
      const file = state.files[i];
      const order = startOrder + i;

      try {
        const parsed = await processSlipFile(file, i, total);
        appendResultRow(order, parsed.amount || "-", parsed.memo || "-");
      } catch (error) {
        console.error(error);
        appendResultRow(order, "-", "อ่านไม่สำเร็จ");
      }

      const percent = Math.round(((i + 1) / total) * 100);
      setStatus(`อ่านแล้ว ${i + 1}/${total} สลิป`, percent);
    }

    setStatus(`อ่านครบ ${total} สลิปแล้ว`, 100);
    state.files = [];
    elements.slipInput.value = "";
    updateFileCount(0);
  } finally {
    state.busy = false;
    elements.slipInput.disabled = false;
    elements.scanBtn.disabled = false;
    elements.clearBtn.disabled = false;
    setTableActionDisabled(false);
    updateSelectionActions();
  }
}

async function processSlipFile(file, fileIndex, totalFiles) {
  const image = await fileToImage(file);

  const originalCanvas = document.createElement("canvas");
  drawImageFit(originalCanvas, image, IMAGE_MAX_WIDTH);

  const processedCanvas = document.createElement("canvas");
  preprocessCanvas(originalCanvas, processedCanvas);

  const primaryCanvas = buildBottomCropCanvas(processedCanvas, OCR_PRIMARY_BOTTOM_RATIO) || processedCanvas;

  const primaryText = await runOcr(primaryCanvas, (progress) => {
    setStatus(
      `กำลังอ่านสลิป ${fileIndex + 1}/${totalFiles}: ${file.name}`,
      mapFileProgress(fileIndex, totalFiles, 0.04 + progress * 0.78)
    );
  });

  const parsed = parseSlipText(primaryText);

  if (!parsed.amount || !parsed.memo) {
    const fallbackText = await runOcr(processedCanvas, (progress) => {
      setStatus(
        `กำลังเก็บข้อมูลเพิ่ม ${fileIndex + 1}/${totalFiles}: ${file.name}`,
        mapFileProgress(fileIndex, totalFiles, 0.84 + progress * 0.14)
      );
    });

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

async function runOcr(canvas, onProgress) {
  const result = await Tesseract.recognize(canvas, "eng+tha", {
    logger: (message) => {
      if (
        message &&
        typeof message === "object" &&
        message.status === "recognizing text" &&
        typeof message.progress === "number" &&
        typeof onProgress === "function"
      ) {
        onProgress(clamp(message.progress, 0, 1));
      }
    }
  });

  return normalizeText(result?.data?.text || "");
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
  amountCell.textContent = amount;

  const memoCell = document.createElement("td");
  memoCell.textContent = memo;

  row.appendChild(selectCell);
  row.appendChild(orderCell);
  row.appendChild(amountCell);
  row.appendChild(memoCell);
  elements.resultBody.appendChild(row);
  updateSelectionActions();
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
  elements.selectAllBtn.textContent = allChecked ? "ยกเลิกเลือกทั้งหมด" : "เลือกทั้งหมด";
  elements.selectAllBtn.disabled = state.busy || !hasRows;
  elements.deleteSelectedBtn.disabled = state.busy || selectedCount === 0;
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
