import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const DEFAULT_CONFIG = {
  blankBuyerText: "รอกรอกข้อมูล",
  blankLotText: "รอกรอกข้อมูล",
  blankManufacturerText: "รอกรอกข้อมูล",
  blankNoteText: "",
  maxRowsPerSheet: 45,
  demoRandomNames: false,
  customerNames: [],
  maxQtyPerSale: 2,
};

function cellText(value) {
  if (value === null || value === undefined) return "";
  return String(value).replace(/\s+/g, " ").trim();
}

function cleanDrugName(raw) {
  return cellText(raw).replace(/^[\s　]+/, "");
}

function safeSheetName(name, usedNames) {
  const base = cleanDrugName(name)
    .replace(/[\[\]\*\/\\\?:]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 28) || "KY11";
  let candidate = base;
  let i = 2;
  while (usedNames.has(candidate)) {
    const suffix = ` ${i}`;
    candidate = `${base.slice(0, 31 - suffix.length)}${suffix}`;
    i += 1;
  }
  usedNames.add(candidate);
  return candidate;
}

function formatDateForFile(date = new Date()) {
  return date.toISOString().slice(0, 10);
}

function buildOutputFileName(records, printDate = new Date()) {
  const firstRange = records[0]?.dateRange?.label || "";
  const rangeText = firstRange ? firstRange.replace(/^ระหว่างวันที่\s*/u, "") : "unknown-range";
  const printText = formatDateForFile(printDate);
  return `รายงาน ขย 11 ${rangeText} print date ${printText}.xlsx`;
}

function resolveOutputPath(outputPath, records) {
  const dir = path.dirname(outputPath);
  const base = path.basename(outputPath);
  if (/demo-customer/i.test(base)) {
    return path.join(dir, buildOutputFileName(records));
  }
  return outputPath;
}

function extractDateRange(raw) {
  const text = cellText(raw);
  const match = text.match(/(\d{4}-\d{2}-\d{2})\s*-\s*(\d{4}-\d{2}-\d{2})/);
  if (!match) return { start: "", end: "", label: text };
  return { start: match[1], end: match[2], label: `ระหว่างวันที่ ${match[1]} - ${match[2]}` };
}

function normalizeSourceWorkbookTables(summaryNdjson) {
  return summaryNdjson
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => JSON.parse(line))
    .filter((item) => item.kind === "table" && Array.isArray(item.values));
}

function parseTable(table) {
  const rows = table.values;
  const dateRange = extractDateRange(rows[2]?.[2]);
  const drugName = cleanDrugName(rows[3]?.[1]);
  const manufacturer = cellText(rows[4]?.[1]);
  const lot = cellText(rows[4]?.[3]);
  const packageSize = cellText(rows[4]?.[5]);
  const source = cellText(rows[5]?.[1]);
  const receivedQty = cellText(rows[5]?.[3]);
  const receivedDate = cellText(rows[5]?.[5]);
  const remainingQty = cellText(rows[6]?.[1]);
  const sales = rows.slice(8).filter((row) => cellText(row[0]) || cellText(row[1]) || cellText(row[2])).map((row, index) => ({
    no: Number(row[0]) || index + 1,
    saleDate: cellText(row[1]),
    qty: cellText(row[2]),
    buyer: cellText(row[3]),
    practitioner: cellText(row[4]),
    note: cellText(row[5]),
  }));

  return {
    sourceSheet: table.sheet,
    storeName: cellText(rows[1]?.[1]),
    dateRange,
    drugName,
    manufacturer,
    lot,
    packageSize,
    source,
    receivedQty,
    receivedDate,
    remainingQty,
    sales,
  };
}

function issueList(record, config) {
  const issues = [];
  if (!record.drugName) issues.push("ไม่พบชื่อยา");
  if (!record.manufacturer) issues.push("ไม่มีชื่อผู้ผลิต/ผู้นำเข้า");
  if (!record.lot) issues.push("ไม่มีเลข lot");
  if (!record.packageSize) issues.push("ไม่มีขนาดบรรจุ");
  if (!record.source) issues.push("ไม่มีแหล่งที่ได้มา");
  if (!record.receivedQty) issues.push("ไม่มีจำนวนรับ");
  if (!record.receivedDate) issues.push("ไม่มีวันที่รับ");
  for (const sale of record.sales) {
    if (!sale.saleDate) issues.push(`รายการที่ ${sale.no}: ไม่มีวันที่ขาย`);
    if (!sale.qty) issues.push(`รายการที่ ${sale.no}: ไม่มีจำนวนขาย`);
    if (!sale.buyer) issues.push(`รายการที่ ${sale.no}: ไม่มีชื่อผู้ซื้อ`);
    else if (!looksLikePersonName(sale.buyer)) issues.push(`รายการที่ ${sale.no}: ชื่อผู้ซื้อไม่สมบูรณ์`);
    if (!sale.practitioner) issues.push(`รายการที่ ${sale.no}: ไม่มีผู้มีหน้าที่ปฏิบัติการ`);
  }
  return [...new Set(issues)];
}

function displayOr(value, fallback) {
  return cellText(value) || fallback;
}

function parseQuantityText(qtyText) {
  const text = cellText(qtyText);
  const match = text.match(/^(\d+(?:\.\d+)?)\s*(.*)$/);
  if (!match) return null;
  const amount = Number(match[1]);
  if (!Number.isFinite(amount)) return null;
  return { amount, unit: match[2].trim() };
}

function formatQuantity(amount, unit) {
  const displayAmount = Number.isInteger(amount) ? String(amount) : String(amount);
  return `${displayAmount}${unit ? ` ${unit}` : ""}`.trim();
}

function saleNoteText(sale, config) {
  const notes = [];
  if (cellText(sale.note)) notes.push(cellText(sale.note));
  if (sale.splitFromQty) notes.push(`แยกจากรายการเดิม ${sale.splitFromQty} (${sale.splitPart}/${sale.splitTotal})`);
  return notes.join("; ") || config.blankNoteText;
}

function splitOversizedSales(records, config) {
  const maxQty = Number(config.maxQtyPerSale) || 2;
  if (maxQty <= 0) return;
  for (const record of records) {
    const splitSales = [];
    for (const sale of record.sales) {
      const parsed = parseQuantityText(sale.qty);
      if (!parsed || parsed.amount <= maxQty) {
        splitSales.push(sale);
        continue;
      }

      let remaining = parsed.amount;
      let splitPart = 1;
      const splitTotal = Math.ceil(parsed.amount / maxQty);
      while (remaining > 0) {
        const partAmount = Math.min(maxQty, remaining);
        splitSales.push({
          ...sale,
          no: splitSales.length + 1,
          qty: formatQuantity(partAmount, parsed.unit),
          sourceNo: sale.sourceNo ?? sale.no,
          splitFromQty: sale.qty,
          splitPart,
          splitTotal,
        });
        remaining -= partAmount;
        splitPart += 1;
      }
    }
    record.sales = splitSales.map((sale, index) => ({ ...sale, no: index + 1 }));
  }
}

function seededIndex(seed, length) {
  let hash = 2166136261;
  for (let i = 0; i < seed.length; i += 1) {
    hash ^= seed.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return Math.abs(hash >>> 0) % length;
}

function applyDemoRandomBuyerNames(records, config) {
  if (!config.demoRandomNames || !config.customerNames?.length) return;
  for (const record of records) {
    for (const sale of record.sales) {
      if (sale.buyer && looksLikePersonName(sale.buyer)) continue;
      const seed = `${record.drugName}|${sale.saleDate}|${sale.qty}|${sale.no}`;
      sale.buyer = config.customerNames[seededIndex(seed, config.customerNames.length)];
      sale.demoGeneratedBuyer = true;
    }
  }
}

function looksLikePersonName(value) {
  const text = cellText(value);
  if (!text) return false;
  if (text.includes("?")) return false;
  if (/^\d+$/.test(text)) return false;
  if (text.length < 3 || text.length > 80) return false;
  if (/โทร|เบอร์|phone|tel|date|วันที่|drug|product|qty|จำนวน|ราคา|price|หมายเหตุ|note|demo|ชื่อ$|ชื่อลูกค้า|name$/i.test(text)) return false;
  if (/^(ยา|ชื่อยา|รายการยา)$/i.test(text)) return false;
  return /[\u0E00-\u0E7Fa-zA-Z]/.test(text);
}

function parseCsvLine(line) {
  const cells = [];
  let cell = "";
  let quoted = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    if (ch === '"') {
      if (quoted && line[i + 1] === '"') {
        cell += '"';
        i += 1;
      } else {
        quoted = !quoted;
      }
    } else if (ch === "," && !quoted) {
      cells.push(cell);
      cell = "";
    } else {
      cell += ch;
    }
  }
  cells.push(cell);
  return cells.map(cellText);
}

async function loadCustomerNamesFromCsv(customerPath) {
  const bytes = await fs.readFile(customerPath);
  const text = new TextDecoder("utf-8").decode(bytes).replace(/^\uFEFF/, "");
  const rows = text.split(/\r?\n/).map(parseCsvLine).filter((row) => row.some(Boolean));
  const nameColumn = findNameColumn(rows);
  const names = [];
  if (nameColumn) {
    for (const row of rows.slice(nameColumn.row + 1)) {
      if (looksLikePersonName(row[nameColumn.col])) names.push(cellText(row[nameColumn.col]));
    }
  } else {
    for (const row of rows) {
      const firstNameLikeCell = row.find(looksLikePersonName);
      if (firstNameLikeCell) names.push(cellText(firstNameLikeCell));
    }
  }
  return [...new Set(names)];
}

function findNameColumn(values) {
  const maxRows = Math.min(values.length, 10);
  for (let r = 0; r < maxRows; r += 1) {
    const row = values[r] || [];
    for (let c = 0; c < row.length; c += 1) {
      const header = cellText(row[c]).toLowerCase();
      const isNameHeader = /(ชื่อ|name|customer)/i.test(header);
      const isOtherHeader = /(ยา|drug|product|โทร|phone|tel|date|วันที่|qty|จำนวน|ราคา|price|หมายเหตุ|note)/i.test(header);
      if (isNameHeader && !isOtherHeader) return { row: r, col: c };
    }
  }
  return null;
}

async function loadCustomerNames(customerPath) {
  if (!customerPath) return [];
  if (customerPath.toLowerCase().endsWith(".csv")) {
    return loadCustomerNamesFromCsv(customerPath);
  }
  const input = await FileBlob.load(customerPath);
  const workbook = await SpreadsheetFile.importXlsx(input);
  const summary = await workbook.inspect({
    kind: "table",
    maxChars: 200000,
    tableMaxRows: 3000,
    tableMaxCols: 20,
    tableMaxCellChars: 120,
  });
  const tables = normalizeSourceWorkbookTables(summary.ndjson);
  const names = [];
  for (const table of tables) {
    const nameColumn = findNameColumn(table.values);
    if (nameColumn) {
      for (const row of table.values.slice(nameColumn.row + 1)) {
        const value = row[nameColumn.col];
        if (looksLikePersonName(value)) names.push(cellText(value));
      }
    } else {
      for (const row of table.values) {
        for (const value of row) {
          if (looksLikePersonName(value)) names.push(cellText(value));
        }
      }
    }
  }
  return [...new Set(names)];
}

function writeKy11Sheet(sheet, record, issues, config) {
  sheet.showGridLines = false;

  sheet.getRange("A1:E1").merge();
  sheet.getRange("A2:E2").merge();
  sheet.getRange("A3:E3").merge();
  sheet.getRange("A1:F1").values = [["บัญชีการขายยาอันตราย เฉพาะรายการยาที่เลขาธิการคณะกรรมการอาหารและยากำหนด", null, null, null, null, "แบบ ขย.๑๑"]];
  sheet.getRange("A2:F2").values = [[displayOr(record.storeName, "รอกรอกชื่อสถานที่ขายยา"), null, null, null, null, null]];
  sheet.getRange("A3:F3").values = [[record.dateRange.label, null, null, null, null, null]];
  sheet.getRange("A4:F7").values = [
    ["ชื่อยา", displayOr(record.drugName, "รอกรอกข้อมูล"), null, null, null, null],
    ["ชื่อผู้ผลิต/ผู้นำเข้า", displayOr(record.manufacturer, config.blankManufacturerText), "เลขที่หรืออักษรของครั้งที่ผลิต", displayOr(record.lot, config.blankLotText), "ขนาดบรรจุ", displayOr(record.packageSize, "รอกรอกข้อมูล")],
    ["ได้มาจาก", displayOr(record.source, "รอกรอกข้อมูล"), "จำนวนรับเปรียบเทียบหน่วยเล็กสุด", displayOr(record.receivedQty, "รอกรอกข้อมูล"), "วันที่รับ", displayOr(record.receivedDate, "รอกรอกข้อมูล")],
    ["จำนวนคงเหลือ", displayOr(record.remainingQty, "รอกรอกข้อมูล"), null, null, null, null],
  ];
  sheet.getRange("A8:F8").values = [["ลำดับที่", "วันเดือนปีที่ขาย", "จำนวน/ปริมาณที่ขาย", "ชื่อ - สกุล ผู้ซื้อ", "ลายมือชื่อผู้มีหน้าที่ปฏิบัติการ", "หมายเหตุ"]];

  const minRows = Math.max(record.sales.length, Math.min(25, config.maxRowsPerSheet));
  const salesRows = Array.from({ length: minRows }, (_, idx) => {
    const sale = record.sales[idx];
    if (!sale) return [idx + 1, "", "", "", "", ""];
    return [
      idx + 1,
      sale.saleDate,
      sale.qty,
      displayOr(sale.buyer, config.blankBuyerText),
      sale.practitioner,
      saleNoteText(sale, config),
    ];
  });
  const endRow = 8 + salesRows.length;
  sheet.getRange(`A9:F${endRow}`).values = salesRows;

  sheet.getRange("A1:F3").format = {
    font: { bold: true, name: "Aptos", size: 11 },
    horizontalAlignment: "center",
    verticalAlignment: "middle",
    wrapText: true,
  };
  sheet.getRange("A1:E1").format.font = { bold: true, name: "Aptos", size: 12 };
  sheet.getRange("F1").format.horizontalAlignment = "right";
  sheet.getRange("A4:F7").format = {
    font: { name: "Aptos", size: 10 },
    verticalAlignment: "middle",
  };
  sheet.getRange("A8:F8").format = {
    font: { bold: true, name: "Aptos", size: 10 },
    fill: "#F3F4F6",
    horizontalAlignment: "center",
    verticalAlignment: "middle",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: "#000000" },
  };
  sheet.getRange(`A9:F${endRow}`).format = {
    font: { name: "Aptos", size: 10 },
    verticalAlignment: "top",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: "#000000" },
  };
  sheet.getRange(`A8:F${endRow}`).format.borders = { preset: "all", style: "thin", color: "#000000" };
  sheet.getRange(`A9:A${endRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`B9:C${endRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`E9:E${endRow}`).format.horizontalAlignment = "center";

  for (const address of ["B4", "B5", "D5", "F5", "B6", "D6", "F6", "B7"]) {
    sheet.getRange(address).format.font = { bold: true };
  }

  const warningCells = [];
  if (!record.drugName) warningCells.push("B4");
  if (!record.manufacturer) warningCells.push("B5");
  if (!record.lot) warningCells.push("D5");
  if (!record.packageSize) warningCells.push("F5");
  if (!record.source) warningCells.push("B6");
  if (!record.receivedQty) warningCells.push("D6");
  if (!record.receivedDate) warningCells.push("F6");
  if (!record.remainingQty) warningCells.push("B7");
  for (let i = 0; i < record.sales.length; i += 1) {
    const sale = record.sales[i];
    if (!sale.saleDate) warningCells.push(`B${9 + i}`);
    if (!sale.qty) warningCells.push(`C${9 + i}`);
    if (!sale.buyer || !looksLikePersonName(sale.buyer)) warningCells.push(`D${9 + i}`);
    if (!sale.practitioner) warningCells.push(`E${9 + i}`);
  }
  for (const cell of warningCells) {
    sheet.getRange(cell).format = { fill: "#FEF3C7" };
  }

  if (issues.length) {
    sheet.getRange("A" + (endRow + 2) + ":F" + (endRow + 2)).merge();
    sheet.getRange("A" + (endRow + 2)).values = [[`มี ${issues.length} จุดที่ต้องตรวจสอบ ดูรายละเอียดในชีต "รายการที่ต้องแก้"`]];
    sheet.getRange("A" + (endRow + 2)).format = {
      fill: "#FFF7ED",
      font: { color: "#9A3412", size: 9 },
      wrapText: true,
    };
  }

  sheet.getRange("A:F").format.autofitColumns();
  sheet.getRange("A:A").format.columnWidth = 9;
  sheet.getRange("B:B").format.columnWidth = 16;
  sheet.getRange("C:C").format.columnWidth = 23;
  sheet.getRange("D:D").format.columnWidth = 28;
  sheet.getRange("E:E").format.columnWidth = 24;
  sheet.getRange("F:F").format.columnWidth = 18;
  sheet.getRange(`A1:F${Math.max(endRow + 2, 35)}`).format.autofitRows();
}

function writeSummary(sheet, records, recordIssues, config) {
  sheet.showGridLines = false;
  sheet.getRange("A1:H1").merge();
  sheet.getRange("A1").values = [["ตรวจสอบรายงาน ข.ย.11"]];
  sheet.getRange("A2:H2").values = [[
    config.demoRandomNames
      ? "โหมดข้อมูลจำลองเพื่อการศึกษา: ชื่อผู้ซื้อที่ว่างถูกสุ่มจากไฟล์รายชื่อ และไม่ใช่ข้อมูลสำหรับส่งจริง"
      : "ไฟล์นี้สร้างจากข้อมูล POS และไฮไลต์ช่องที่ยังต้องตรวจสอบก่อนส่งจริง",
    null,
    null,
    null,
    null,
    null,
    null,
    null,
  ]];
  sheet.getRange("A4:H4").values = [["ลำดับ", "ชื่อยา", "ช่วงวันที่", "จำนวนรายการขาย", "ช่องที่ต้องตรวจสอบ", "ชื่อแท็บต้นทาง", "สถานะ", "ตัวอย่างที่ต้องแก้"]];

  const rows = records.map((record, idx) => {
    const issues = recordIssues[idx];
    return [
      idx + 1,
      record.drugName,
      record.dateRange.label,
      record.sales.length,
      issues.length,
      record.sourceSheet,
      issues.length ? "ต้องตรวจสอบ" : "พร้อมใช้",
      issues.slice(0, 3).join("; ") + (issues.length > 3 ? `; และอีก ${issues.length - 3} รายการ` : ""),
    ];
  });
  if (rows.length) sheet.getRangeByIndexes(4, 0, rows.length, 8).values = rows;

  sheet.getRange("A1:H1").format = {
    fill: "#111827",
    font: { bold: true, color: "#FFFFFF", size: 14 },
    horizontalAlignment: "center",
  };
  sheet.getRange("A2:H2").format = { font: { color: "#4B5563" } };
  sheet.getRange("A4:H4").format = {
    fill: "#E5E7EB",
    font: { bold: true },
    horizontalAlignment: "center",
    borders: { preset: "all", style: "thin", color: "#9CA3AF" },
  };
  const endRow = 4 + Math.max(rows.length, 1);
  sheet.getRange(`A5:H${endRow}`).format = {
    borders: { preset: "all", style: "thin", color: "#D1D5DB" },
    verticalAlignment: "top",
    wrapText: true,
  };
  sheet.getRange(`A5:A${endRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`D5:G${endRow}`).format.horizontalAlignment = "center";
  sheet.getRange("A:H").format.autofitColumns();
  sheet.getRange("B:B").format.columnWidth = 32;
  sheet.getRange("C:C").format.columnWidth = 34;
  sheet.getRange("H:H").format.columnWidth = 58;
  sheet.freezePanes.freezeRows(4);
}

function writeIssueDetails(sheet, records, recordIssues) {
  sheet.showGridLines = false;
  sheet.getRange("A1:F1").merge();
  sheet.getRange("A1").values = [["รายการที่ต้องแก้ก่อนส่งรายงาน"]];
  sheet.getRange("A3:F3").values = [["ลำดับ", "ชื่อยา", "แท็บต้นทาง", "ประเภท", "รายละเอียด", "คำแนะนำ"]];

  const rows = [];
  for (let i = 0; i < records.length; i += 1) {
    const record = records[i];
    for (const issue of recordIssues[i]) {
      const type = issue.includes("รายการที่") ? "ข้อมูลขาย" : "ข้อมูลหัวฟอร์ม";
      rows.push([
        rows.length + 1,
        record.drugName,
        record.sourceSheet,
        type,
        issue,
        issue.includes("ชื่อผู้ซื้อ")
          ? "ตรวจจากใบเสร็จ/ข้อมูลลูกค้า แล้วกรอกชื่อจริง"
          : "ตรวจเอกสารรับยา/ฉลาก/ข้อมูลสินค้า แล้วกรอกให้ครบ",
      ]);
    }
  }
  if (rows.length) sheet.getRangeByIndexes(3, 0, rows.length, 6).values = rows;

  sheet.getRange("A1:F1").format = {
    fill: "#7C2D12",
    font: { bold: true, color: "#FFFFFF", size: 14 },
    horizontalAlignment: "center",
  };
  sheet.getRange("A3:F3").format = {
    fill: "#FED7AA",
    font: { bold: true },
    horizontalAlignment: "center",
    borders: { preset: "all", style: "thin", color: "#9A3412" },
  };
  const endRow = 3 + Math.max(rows.length, 1);
  sheet.getRange(`A4:F${endRow}`).format = {
    borders: { preset: "all", style: "thin", color: "#FDBA74" },
    verticalAlignment: "top",
    wrapText: true,
  };
  sheet.getRange(`A4:A${endRow}`).format.horizontalAlignment = "center";
  sheet.getRange("A:F").format.autofitColumns();
  sheet.getRange("B:B").format.columnWidth = 34;
  sheet.getRange("E:E").format.columnWidth = 38;
  sheet.getRange("F:F").format.columnWidth = 42;
  sheet.freezePanes.freezeRows(3);
}

async function convert(inputPath, outputPath, config = {}) {
  config = { ...DEFAULT_CONFIG, ...config };
  const input = await FileBlob.load(inputPath);
  const sourceWorkbook = await SpreadsheetFile.importXlsx(input);
  const summary = await sourceWorkbook.inspect({
    kind: "table",
    maxChars: 200000,
    tableMaxRows: 2000,
    tableMaxCols: 12,
    tableMaxCellChars: 200,
  });
  const records = normalizeSourceWorkbookTables(summary.ndjson).map(parseTable);
  splitOversizedSales(records, config);
  applyDemoRandomBuyerNames(records, config);
  const recordIssues = records.map((record) => issueList(record, config));

  const workbook = Workbook.create();
  const summarySheet = workbook.worksheets.add("ตรวจสอบ");
  writeSummary(summarySheet, records, recordIssues, config);
  const issueSheet = workbook.worksheets.add("รายการที่ต้องแก้");
  writeIssueDetails(issueSheet, records, recordIssues);

  const usedNames = new Set(["ตรวจสอบ", "รายการที่ต้องแก้"]);
  for (let i = 0; i < records.length; i += 1) {
    const record = records[i];
    const sheet = workbook.worksheets.add(safeSheetName(record.drugName || record.sourceSheet, usedNames));
    writeKy11Sheet(sheet, record, recordIssues[i], config);
  }

  const resolvedOutputPath = resolveOutputPath(outputPath, records);
  await fs.mkdir(path.dirname(resolvedOutputPath), { recursive: true });
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(resolvedOutputPath);
  return { records, recordIssues, outputPath: resolvedOutputPath };
}

const isMain = fileURLToPath(import.meta.url) === process.argv[1];
if (isMain) {
  const [, , inputArg, outputArg, customerArg, ...flags] = process.argv;
  if (!inputArg || !outputArg) {
    console.error("Usage: node ky11_converter.mjs <input.xlsx> <output.xlsx> [customers.xlsx] [--demo-random-names]");
    process.exit(1);
  }
  const demoRandomNames = flags.includes("--demo-random-names") || customerArg === "--demo-random-names";
  const customerPath = customerArg && customerArg !== "--demo-random-names" ? path.resolve(customerArg) : "";
  const customerNames = await loadCustomerNames(customerPath);
  if (demoRandomNames && !customerNames.length) {
    console.error("Demo random name mode needs a customer list workbook with at least one name.");
    process.exit(1);
  }
  const result = await convert(path.resolve(inputArg), path.resolve(outputArg), {
    demoRandomNames,
    customerNames,
  });
  const totalIssues = result.recordIssues.reduce((sum, issues) => sum + issues.length, 0);
  console.log(`Converted ${result.records.length} sheets. Review items: ${totalIssues}. Customer names loaded: ${customerNames.length}. Output: ${result.outputPath}`);
}

export { convert, loadCustomerNames };
