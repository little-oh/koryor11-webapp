import fs from "node:fs/promises";
import path from "node:path";
import ExcelJS from "exceljs";

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
  if (typeof value === "object" && value.text !== undefined) return cellText(value.text);
  if (Array.isArray(value)) return cellText(value[0]);
  return String(value).replace(/\s+/g, " ").trim();
}

function cleanDrugName(raw) {
  return cellText(raw).replace(/^[\s\u3000]+/, "");
}

function safeSheetName(name, usedNames) {
  const base =
    cleanDrugName(name)
      .replace(/[\[\]*\/\\?:]/g, " ")
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

function normalizeStoreName(raw) {
  const text = cellText(raw);
  if (!text) return "";
  if (text === "ร้านยาพหล(Maruay Pharmacy)") return "ร้านยามารวย(Maruay Pharmacy)";
  return text;
}

function pharmacyCodeFromRecords(records) {
  const storeName = normalizeStoreName(records[0]?.storeName || "");
  if (/Maruay Pharmacy/i.test(storeName)) return "MR";
  if (/Phahol Pharmacy/i.test(storeName)) return "PH";
  return "";
}

function buildOutputFileName(records, printDate = new Date()) {
  const firstRange = records[0]?.dateRange?.label || "";
  const rangeText = firstRange ? firstRange.replace(/^ระหว่างวันที่\s*/u, "") : "unknown-range";
  const pharmacyCode = pharmacyCodeFromRecords(records);
  const pharmacyPart = pharmacyCode ? ` ${pharmacyCode}` : "";
  return `รายงาน ขย 11${pharmacyPart} ${rangeText} print date ${formatDateForFile(printDate)}.xlsx`;
}

function resolveOutputPath(outputPath, records) {
  const dir = path.dirname(outputPath);
  const base = path.basename(outputPath);
  if (/demo-customer/i.test(base)) return path.join(dir, buildOutputFileName(records));
  return outputPath;
}

function extractDateRange(raw) {
  const text = cellText(raw);
  const match = text.match(/(\d{4}-\d{2}-\d{2})\s*-\s*(\d{4}-\d{2}-\d{2})/);
  if (!match) return { start: "", end: "", label: text };
  return { start: match[1], end: match[2], label: `ระหว่างวันที่ ${match[1]} - ${match[2]}` };
}

function rowValues(worksheet, rowNumber) {
  const row = worksheet.getRow(rowNumber);
  const values = [];
  for (let c = 1; c <= row.cellCount; c += 1) values[c] = row.getCell(c).value;
  return values;
}

function parseTable(worksheet) {
  const rows = {};
  for (let r = 1; r <= worksheet.rowCount; r += 1) rows[r] = rowValues(worksheet, r);

  const dateRange = extractDateRange(rows[3]?.[3]);
  const sales = [];
  for (let r = 9; r <= worksheet.rowCount; r += 1) {
    const row = rows[r] || [];
    if (!cellText(row[1]) && !cellText(row[2]) && !cellText(row[3]) && !cellText(row[4]) && !cellText(row[5])) continue;
    sales.push({
      no: Number(row[1]) || sales.length + 1,
      saleDate: cellText(row[2]),
      qty: cellText(row[3]),
      buyer: cellText(row[4]),
      practitioner: cellText(row[5]),
      note: cellText(row[6]),
    });
  }

  return {
    sourceSheet: worksheet.name,
    storeName: normalizeStoreName(rows[2]?.[2]),
    dateRange,
    drugName: cleanDrugName(rows[4]?.[2]),
    manufacturer: cellText(rows[5]?.[2]),
    lot: cellText(rows[5]?.[4]),
    packageSize: cellText(rows[5]?.[6]),
    source: cellText(rows[6]?.[2]),
    receivedQty: cellText(rows[6]?.[4]),
    receivedDate: cellText(rows[6]?.[6]),
    remainingQty: cellText(rows[7]?.[2]),
    sales,
  };
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

function seededIndex(seed, length) {
  let hash = 2166136261;
  for (let i = 0; i < seed.length; i += 1) {
    hash ^= seed.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return Math.abs(hash >>> 0) % length;
}

function looksLikePersonName(value) {
  const text = cellText(value);
  if (!text) return false;
  if (text.includes("?")) return false;
  if (/^\d+$/.test(text)) return false;
  if (text.length < 3 || text.length > 120) return false;
  if (/[0-9]{6,}/.test(text)) return false;
  if (/kbank|lineman|line man|member|สมาชิก|พนักงาน|cash|โอน|visa|master|พร้อมเพย์|qr|credit|debit/i.test(text)) return false;
  if (/โทร|เบอร์|phone|tel|วันที่|date|drug|product|qty|จำนวน|ราคา|หมายเหตุ|note|demo/i.test(text)) return false;
  if (/^(ยา|ชื่อยา|รายการยา|รอกรอกข้อมูล)$/i.test(text)) return false;
  return /[\u0E00-\u0E7Fa-zA-Z]/.test(text);
}

function issueList(record) {
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
    if (!sale.practitioner) issues.push(`รายการที่ ${sale.no}: ไม่มีผู้มีหน้าที่ปฏิบัติการ`);
  }
  return [...new Set(issues)];
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

function collectNamesFromRows(rows, columnIndex = 1) {
  const names = [];
  for (const row of rows) {
    const value = row?.[columnIndex];
    if (looksLikePersonName(value)) names.push(cellText(value));
  }
  return [...new Set(names)];
}

function excelColumnName(columnNumber) {
  let n = columnNumber;
  let result = "";
  while (n > 0) {
    const remainder = (n - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result || "A";
}

function findBestNameColumn(rows, startColumn = 0, maxColumns = 10) {
  let best = {
    columnIndex: startColumn,
    usableCount: 0,
    uniqueCount: 0,
    sampleNames: [],
  };

  for (let columnIndex = startColumn; columnIndex < startColumn + maxColumns; columnIndex += 1) {
    const usableNames = [];
    const uniqueNames = new Set();
    for (const row of rows) {
      const value = row?.[columnIndex];
      if (!looksLikePersonName(value)) continue;
      const text = cellText(value);
      usableNames.push(text);
      uniqueNames.add(text);
    }
    const candidate = {
      columnIndex,
      usableCount: usableNames.length,
      uniqueCount: uniqueNames.size,
      sampleNames: [...uniqueNames].slice(0, 5),
    };
    if (
      candidate.uniqueCount > best.uniqueCount ||
      (candidate.uniqueCount === best.uniqueCount && candidate.usableCount > best.usableCount)
    ) {
      best = candidate;
    }
  }
  return best;
}

async function loadCustomerNamesFromCsv(customerPath) {
  const bytes = await fs.readFile(customerPath);
  const text = new TextDecoder("utf-8").decode(bytes).replace(/^\uFEFF/, "");
  const rows = text
    .split(/\r?\n/)
    .map(parseCsvLine)
    .filter((row) => row.some(Boolean));
  const bestColumn = findBestNameColumn(rows, 0, Math.min(10, Math.max(1, rows[0]?.length || 1)));
  return {
    names: collectNamesFromRows(rows, bestColumn.columnIndex),
    stats: {
      sheetName: "CSV",
      columnLabel: excelColumnName(bestColumn.columnIndex + 1),
      usableCount: bestColumn.usableCount,
      uniqueCount: bestColumn.uniqueCount,
      sampleNames: bestColumn.sampleNames,
    },
  };
}

async function loadCustomerNames(customerPath) {
  const result = await loadCustomerNamesDetailed(customerPath);
  return result.names;
}

async function loadCustomerNamesDetailed(customerPath) {
  if (!customerPath) {
    return {
      names: [],
      stats: { sheetName: "", columnLabel: "A", usableCount: 0, uniqueCount: 0, sampleNames: [] },
    };
  }
  if (customerPath.toLowerCase().endsWith(".csv")) return loadCustomerNamesFromCsv(customerPath);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(customerPath);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    return {
      names: [],
      stats: { sheetName: "", columnLabel: "A", usableCount: 0, uniqueCount: 0, sampleNames: [] },
    };
  }

  const rows = [];
  for (let r = 1; r <= worksheet.rowCount; r += 1) {
    const row = worksheet.getRow(r);
    const values = [];
    for (let c = 1; c <= Math.min(10, Math.max(1, row.cellCount)); c += 1) {
      values[c] = row.getCell(c).value;
    }
    rows.push(values);
  }

  const bestColumn = findBestNameColumn(rows, 1, 10);
  return {
    names: collectNamesFromRows(rows, bestColumn.columnIndex),
    stats: {
      sheetName: worksheet.name,
      columnLabel: excelColumnName(bestColumn.columnIndex),
      usableCount: bestColumn.usableCount,
      uniqueCount: bestColumn.uniqueCount,
      sampleNames: bestColumn.sampleNames,
    },
  };
}

function buildRandomParts(totalAmount, maxQty, seed) {
  const parts = [];
  let remaining = totalAmount;
  let step = 0;
  while (remaining > 0) {
    if (remaining <= maxQty) {
      parts.push(remaining);
      break;
    }
    if (remaining === 3) {
      const firstPart = seededIndex(`${seed}|${step}|three`, 2) === 0 ? 1 : 2;
      parts.push(firstPart);
      remaining -= firstPart;
      step += 1;
      continue;
    }
    const part = seededIndex(`${seed}|${step}`, 2) === 0 ? 1 : 2;
    parts.push(Math.min(part, remaining));
    remaining -= Math.min(part, remaining);
    step += 1;
  }
  return parts;
}

function splitOversizedSales(records, config) {
  const maxQty = Number(config.maxQtyPerSale) || 2;
  if (maxQty <= 0) return;

  for (const record of records) {
    const nextSales = [];
    for (const sale of record.sales) {
      const parsed = parseQuantityText(sale.qty);
      if (!parsed || !Number.isInteger(parsed.amount) || parsed.amount <= maxQty) {
        nextSales.push({ ...sale });
        continue;
      }

      const parts = buildRandomParts(parsed.amount, maxQty, `${record.drugName}|${sale.saleDate}|${sale.no}|${sale.qty}`);
      const splitTotal = parts.length;
      let splitPart = 1;
      for (const amount of parts) {
        nextSales.push({
          ...sale,
          qty: formatQuantity(amount, parsed.unit),
          sourceNo: sale.sourceNo ?? sale.no,
          splitFromQty: sale.qty,
          splitPart,
          splitTotal,
        });
        splitPart += 1;
      }
    }
    record.sales = nextSales.map((sale, index) => ({ ...sale, no: index + 1 }));
  }
}

function assignRandomBuyerNames(records, config) {
  const customerNames = Array.isArray(config.customerNames)
    ? [...new Set(config.customerNames.map((name) => cellText(name)).filter(looksLikePersonName))]
    : [];
  if (!config.demoRandomNames || !customerNames.length) return;

  for (const record of records) {
    let previousBuyer = "";
    for (const sale of record.sales) {
      const baseSeed = `${record.drugName}|${sale.saleDate}|${sale.qty}|${sale.no}|buyer`;
      let selected = customerNames[seededIndex(baseSeed, customerNames.length)];
      if (selected === previousBuyer && customerNames.length > 1) {
        const candidates = customerNames.filter((name) => name !== previousBuyer);
        selected = candidates[seededIndex(`${baseSeed}|retry`, candidates.length)];
      }
      sale.buyer = selected;
      sale.demoGeneratedBuyer = true;
      previousBuyer = selected;
    }
  }
}

function buildSplitSummaryLines(record) {
  const items = new Map();
  for (const sale of record.sales) {
    if (!sale.splitFromQty) continue;
    const sourceNo = sale.sourceNo ?? sale.no;
    const key = `${sourceNo}|${sale.splitFromQty}`;
    if (!items.has(key)) {
      items.set(key, {
        sourceNo,
        originalQty: sale.splitFromQty,
        saleNos: [],
      });
    }
    items.get(key).saleNos.push(sale.no);
  }

  function formatSaleNoList(saleNos) {
    const sorted = [...saleNos].sort((a, b) => a - b);
    if (!sorted.length) return "";
    const ranges = [];
    let start = sorted[0];
    let previous = sorted[0];
    for (let i = 1; i < sorted.length; i += 1) {
      const current = sorted[i];
      if (current === previous + 1) {
        previous = current;
        continue;
      }
      ranges.push(start === previous ? `${start}` : `${start}-${previous}`);
      start = current;
      previous = current;
    }
    ranges.push(start === previous ? `${start}` : `${start}-${previous}`);
    return ranges.join(", ");
  }

  return [...items.values()]
    .sort((a, b) => a.saleNos[0] - b.saleNos[0])
    .map((item) => `รายการที่ ${formatSaleNoList(item.saleNos)} โอนย้ายสินค้าแยกจากรายการเดิม ${item.originalQty}`);
}

function saleNoteText(sale, config) {
  return "";
}

function applyBorder(cell, style = "thin", color = "000000") {
  cell.border = {
    top: { style, color: { argb: color } },
    left: { style, color: { argb: color } },
    bottom: { style, color: { argb: color } },
    right: { style, color: { argb: color } },
  };
}

function writeSummary(worksheet, records, recordIssues) {
  worksheet.views = [{ showGridLines: false, state: "frozen", ySplit: 4 }];
  worksheet.columns = [
    { width: 8 },
    { width: 34 },
    { width: 26 },
    { width: 18 },
    { width: 18 },
    { width: 28 },
    { width: 16 },
    { width: 42 },
  ];

  worksheet.mergeCells("A1:H1");
  worksheet.getCell("A1").value = "สรุปรายงาน ขย.11";
  worksheet.getCell("A2").value = "ตรวจสอบก่อนส่งจริง";
  worksheet.getCell("A4").value = "ลำดับ";
  worksheet.getCell("B4").value = "ชื่อยา";
  worksheet.getCell("C4").value = "ช่วงวันที่";
  worksheet.getCell("D4").value = "จำนวนรายการขาย";
  worksheet.getCell("E4").value = "ช่องที่ต้องตรวจสอบ";
  worksheet.getCell("F4").value = "ชื่อแท็บต้นทาง";
  worksheet.getCell("G4").value = "สถานะ";
  worksheet.getCell("H4").value = "ตัวอย่างที่ต้องแก้";

  records.forEach((record, idx) => {
    const issues = recordIssues[idx];
    const row = worksheet.getRow(5 + idx);
    [
      idx + 1,
      record.drugName,
      record.dateRange.label,
      record.sales.length,
      issues.length,
      record.sourceSheet,
      issues.length ? "ต้องตรวจสอบ" : "พร้อมใช้",
      issues.slice(0, 3).join("; ") + (issues.length > 3 ? `; และอีก ${issues.length - 3} รายการ` : ""),
    ].forEach((value, colIdx) => {
      const cell = row.getCell(colIdx + 1);
      cell.value = value;
      cell.alignment = { vertical: "top", wrapText: true };
      applyBorder(cell, "thin", "D1D5DB");
    });
    row.getCell(1).alignment = { horizontal: "center" };
    row.getCell(4).alignment = { horizontal: "center" };
    row.getCell(5).alignment = { horizontal: "center" };
    row.getCell(7).alignment = { horizontal: "center" };
  });

  worksheet.getCell("A1").font = { bold: true, size: 14, color: { argb: "FFFFFF" } };
  worksheet.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "111827" } };
  worksheet.getCell("A1").alignment = { horizontal: "center" };
  worksheet.getCell("A2").font = { color: { argb: "4B5563" } };
  for (let c = 1; c <= 8; c += 1) {
    const cell = worksheet.getRow(4).getCell(c);
    cell.font = { bold: true };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "E5E7EB" } };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    applyBorder(cell, "thin", "9CA3AF");
  }
}

function writeIssueDetails(worksheet, records, recordIssues) {
  worksheet.views = [{ showGridLines: false, state: "frozen", ySplit: 3 }];
  worksheet.columns = [
    { width: 8 },
    { width: 34 },
    { width: 26 },
    { width: 18 },
    { width: 40 },
    { width: 42 },
  ];
  worksheet.mergeCells("A1:F1");
  worksheet.getCell("A1").value = "รายการที่ต้องแก้ก่อนส่งรายงาน";
  worksheet.getCell("A3").value = "ลำดับ";
  worksheet.getCell("B3").value = "ชื่อยา";
  worksheet.getCell("C3").value = "แท็บต้นทาง";
  worksheet.getCell("D3").value = "ประเภท";
  worksheet.getCell("E3").value = "รายละเอียด";
  worksheet.getCell("F3").value = "คำแนะนำ";

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

  rows.forEach((rowValues, idx) => {
    const row = worksheet.getRow(4 + idx);
    rowValues.forEach((value, colIdx) => {
      const cell = row.getCell(colIdx + 1);
      cell.value = value;
      cell.alignment = { vertical: "top", wrapText: true };
      applyBorder(cell, "thin", "FDBA74");
    });
    row.getCell(1).alignment = { horizontal: "center" };
  });

  worksheet.getCell("A1").font = { bold: true, size: 14, color: { argb: "FFFFFF" } };
  worksheet.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "7C2D12" } };
  worksheet.getCell("A1").alignment = { horizontal: "center" };
  for (let c = 1; c <= 6; c += 1) {
    const cell = worksheet.getRow(3).getCell(c);
    cell.font = { bold: true };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FED7AA" } };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    applyBorder(cell, "thin", "9A3412");
  }
}

function writeKy11Sheet(worksheet, record, issues, config) {
  worksheet.views = [{ showGridLines: false }];
  worksheet.columns = [
    { key: "A", width: 9 },
    { key: "B", width: 16 },
    { key: "C", width: 23 },
    { key: "D", width: 28 },
    { key: "E", width: 24 },
    { key: "F", width: 18 },
  ];

  worksheet.mergeCells("A1:E1");
  worksheet.mergeCells("A2:E2");
  worksheet.mergeCells("A3:E3");
  worksheet.getCell("A1").value = "บัญชีการขายยาอันตราย เฉพาะรายการยาที่เลขาธิการคณะกรรมการอาหารและยากำหนด";
  worksheet.getCell("F1").value = "แบบ ขย.11";
  worksheet.getCell("A2").value = displayOr(record.storeName, "รอกรอกชื่อสถานที่ขายยา");
  worksheet.getCell("A3").value = record.dateRange.label;
  worksheet.getCell("A4").value = "ชื่อยา";
  worksheet.getCell("B4").value = displayOr(record.drugName, "รอกรอกข้อมูล");
  worksheet.getCell("A5").value = "ชื่อผู้ผลิต/ผู้นำเข้า";
  worksheet.getCell("B5").value = displayOr(record.manufacturer, config.blankManufacturerText);
  worksheet.getCell("C5").value = "เลขที่หรืออักษรของครั้งที่ผลิต";
  worksheet.getCell("D5").value = displayOr(record.lot, config.blankLotText);
  worksheet.getCell("E5").value = "ขนาดบรรจุ";
  worksheet.getCell("F5").value = displayOr(record.packageSize, "รอกรอกข้อมูล");
  worksheet.getCell("A6").value = "ได้มาจาก";
  worksheet.getCell("B6").value = displayOr(record.source, "รอกรอกข้อมูล");
  worksheet.getCell("C6").value = "จำนวนรับเปรียบเทียบหน่วยเล็กสุด";
  worksheet.getCell("D6").value = displayOr(record.receivedQty, "รอกรอกข้อมูล");
  worksheet.getCell("E6").value = "วันที่รับ";
  worksheet.getCell("F6").value = displayOr(record.receivedDate, "รอกรอกข้อมูล");
  worksheet.getCell("A7").value = "จำนวนคงเหลือ";
  worksheet.getCell("B7").value = displayOr(record.remainingQty, "รอกรอกข้อมูล");
  worksheet.getCell("A8").value = "ลำดับที่";
  worksheet.getCell("B8").value = "วันเดือนปีที่ขาย";
  worksheet.getCell("C8").value = "จำนวน/ปริมาณที่ขาย";
  worksheet.getCell("D8").value = "ชื่อ - สกุล ผู้ซื้อ";
  worksheet.getCell("E8").value = "ลายมือชื่อผู้มีหน้าที่ปฏิบัติการ";
  worksheet.getCell("F8").value = "หมายเหตุ";

  const minRows = Math.max(record.sales.length, Math.min(25, config.maxRowsPerSheet));
  for (let i = 0; i < minRows; i += 1) {
    const sale = record.sales[i];
    const rowNumber = 9 + i;
    worksheet.getCell(`A${rowNumber}`).value = i + 1;
    if (!sale) continue;
    worksheet.getCell(`B${rowNumber}`).value = sale.saleDate;
    worksheet.getCell(`C${rowNumber}`).value = sale.qty;
    worksheet.getCell(`D${rowNumber}`).value = displayOr(sale.buyer, config.blankBuyerText);
    worksheet.getCell(`E${rowNumber}`).value = sale.practitioner;
    worksheet.getCell(`F${rowNumber}`).value = saleNoteText(sale, config);
  }

  const headerStyle = {
    font: { bold: true, name: "Aptos", size: 11 },
    alignment: { horizontal: "center", vertical: "middle", wrapText: true },
  };
  const titleStyle = {
    font: { bold: true, name: "Aptos", size: 12 },
    alignment: { horizontal: "center", vertical: "middle", wrapText: true },
  };

  worksheet.getCell("A1").style = titleStyle;
  worksheet.getCell("F1").style = { ...titleStyle, alignment: { horizontal: "right", vertical: "middle" } };
  worksheet.getCell("A2").style = titleStyle;
  worksheet.getCell("A3").style = titleStyle;

  [
    "A4",
    "A5",
    "A6",
    "A7",
    "A8",
    "B4",
    "B5",
    "B6",
    "B7",
    "B8",
    "C5",
    "C6",
    "D5",
    "D6",
    "E5",
    "E6",
    "F5",
    "F6",
  ].forEach((addr) => {
    worksheet.getCell(addr).font = { name: "Aptos", size: 10, bold: true };
  });

  for (let r = 4; r <= 8; r += 1) {
    for (let c = 1; c <= 6; c += 1) {
      const cell = worksheet.getRow(r).getCell(c);
      cell.style = { ...cell.style, ...headerStyle };
      applyBorder(cell);
    }
  }

  for (let r = 9; r < 9 + minRows; r += 1) {
    for (let c = 1; c <= 6; c += 1) {
      const cell = worksheet.getRow(r).getCell(c);
      cell.alignment = { vertical: "top", wrapText: true };
      applyBorder(cell);
    }
    worksheet.getRow(r).getCell(1).alignment = { horizontal: "center", vertical: "top" };
    worksheet.getRow(r).getCell(2).alignment = { horizontal: "center", vertical: "top" };
    worksheet.getRow(r).getCell(3).alignment = { horizontal: "center", vertical: "top" };
    worksheet.getRow(r).getCell(5).alignment = { horizontal: "center", vertical: "top" };
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
    const rowNumber = 9 + i;
    if (!sale.saleDate) warningCells.push(`B${rowNumber}`);
    if (!sale.qty) warningCells.push(`C${rowNumber}`);
    if (!sale.buyer) warningCells.push(`D${rowNumber}`);
    if (!sale.practitioner) warningCells.push(`E${rowNumber}`);
  }
  for (const address of warningCells) {
    worksheet.getCell(address).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FEF3C7" },
    };
  }

  let footerRow = 9 + minRows + 1;
  for (const line of buildSplitSummaryLines(record)) {
    worksheet.mergeCells(`A${footerRow}:F${footerRow}`);
    worksheet.getCell(`A${footerRow}`).value = line;
    worksheet.getCell(`A${footerRow}`).font = { size: 9 };
    worksheet.getCell(`A${footerRow}`).alignment = { wrapText: true };
    footerRow += 1;
  }

  if (issues.length) {
    worksheet.mergeCells(`A${footerRow}:F${footerRow}`);
    worksheet.getCell(`A${footerRow}`).value = `มี ${issues.length} จุดที่ต้องตรวจสอบ ดูรายละเอียดในชีต "รายการที่ต้องแก้"`;
    worksheet.getCell(`A${footerRow}`).font = { color: { argb: "9A3412" }, size: 9 };
    worksheet.getCell(`A${footerRow}`).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF7ED" },
    };
  }
}

async function convert(inputPath, outputPath, config = {}) {
  config = { ...DEFAULT_CONFIG, ...config };
  const sourceWorkbook = new ExcelJS.Workbook();
  await sourceWorkbook.xlsx.readFile(inputPath);

  const records = sourceWorkbook.worksheets.map(parseTable);
  splitOversizedSales(records, config);
  assignRandomBuyerNames(records, config);
  const recordIssues = records.map((record) => issueList(record));

  const workbook = new ExcelJS.Workbook();
  writeSummary(workbook.addWorksheet("ตรวจสอบ"), records, recordIssues);
  writeIssueDetails(workbook.addWorksheet("รายการที่ต้องแก้"), records, recordIssues);

  const usedNames = new Set(["ตรวจสอบ", "รายการที่ต้องแก้"]);
  for (let i = 0; i < records.length; i += 1) {
    const record = records[i];
    const sheet = workbook.addWorksheet(safeSheetName(record.drugName || record.sourceSheet, usedNames));
    writeKy11Sheet(sheet, record, recordIssues[i], config);
  }

  const resolvedOutputPath = resolveOutputPath(outputPath, records);
  await fs.mkdir(path.dirname(resolvedOutputPath), { recursive: true });
  await workbook.xlsx.writeFile(resolvedOutputPath);
  return { records, recordIssues, outputPath: resolvedOutputPath };
}

export { convert, loadCustomerNames, loadCustomerNamesDetailed };

