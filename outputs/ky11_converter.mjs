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
  return cellText(raw).replace(/^[\s　]+/, "");
}

function safeSheetName(name, usedNames) {
  const base =
    cleanDrugName(name)
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
  return `รายงาน ขย 11 ${rangeText} print date ${formatDateForFile(printDate)}.xlsx`;
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
  const maxRow = worksheet.rowCount;
  const rows = {};
  for (let r = 1; r <= maxRow; r += 1) rows[r] = rowValues(worksheet, r);

  const dateRange = extractDateRange(rows[3]?.[3]);
  const drugName = cleanDrugName(rows[4]?.[2]);
  const manufacturer = cellText(rows[5]?.[2]);
  const lot = cellText(rows[5]?.[4]);
  const packageSize = cellText(rows[5]?.[6]);
  const source = cellText(rows[6]?.[2]);
  const receivedQty = cellText(rows[6]?.[4]);
  const receivedDate = cellText(rows[6]?.[6]);
  const remainingQty = cellText(rows[7]?.[2]);

  const sales = [];
  for (let r = 9; r <= maxRow; r += 1) {
    const row = rows[r] || [];
    if (!cellText(row[1]) && !cellText(row[2]) && !cellText(row[3])) continue;
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
    storeName: cellText(rows[2]?.[2]),
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
  if (/โทร|เบอร์|phone|tel|date|วันที่|drug|product|qty|จำนวน|ราคา|หมายเหตุ|note|demo|ชื่อ$/i.test(text))
    return false;
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
  if (!customerPath) return [];
  if (customerPath.toLowerCase().endsWith(".csv")) {
    const result = await loadCustomerNamesFromCsv(customerPath);
    return result.names;
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(customerPath);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) return [];

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
  return collectNamesFromRows(rows, bestColumn.columnIndex);
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

function applyBorder(cell, style = "thin", color = "000000") {
  cell.border = {
    top: { style, color: { argb: color } },
    left: { style, color: { argb: color } },
    bottom: { style, color: { argb: color } },
    right: { style, color: { argb: color } },
  };
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
  worksheet.getCell("F1").value = "แบบ ขย.๑๑";
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
  worksheet.getCell("D6").value = displayOr(record.receivedQty, "รอกกรอกข้อมูล");
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
    if (sale) {
      worksheet.getCell(`B${rowNumber}`).value = sale.saleDate;
      worksheet.getCell(`C${rowNumber}`).value = sale.qty;
      worksheet.getCell(`D${rowNumber}`).value = displayOr(sale.buyer, config.blankBuyerText);
      worksheet.getCell(`E${rowNumber}`).value = sale.practitioner;
      worksheet.getCell(`F${rowNumber}`).value = saleNoteText(sale, config);
    }
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
  ["A4", "A5", "A6", "A7", "A8", "B4", "B5", "B6", "B7", "B8", "C5", "C6", "D5", "D6", "E5", "E6", "F5", "F6"]
    .forEach((addr) => {
      worksheet.getCell(addr).font = { name: "Aptos", size: 10, bold: true };
    });
  ["A4:F8"].forEach((range) => {
    const [start, end] = range.split(":");
    const startCol = worksheet.getCell(start).col;
    const startRow = worksheet.getCell(start).row;
    const endCol = worksheet.getCell(end).col;
    const endRow = worksheet.getCell(end).row;
    for (let r = startRow; r <= endRow; r += 1) {
      for (let c = startCol; c <= endCol; c += 1) {
        const cell = worksheet.getRow(r).getCell(c);
        cell.style = { ...cell.style, ...headerStyle };
        applyBorder(cell);
      }
    }
  });

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
    if (!sale.buyer || !looksLikePersonName(sale.buyer)) warningCells.push(`D${rowNumber}`);
    if (!sale.practitioner) warningCells.push(`E${rowNumber}`);
  }
  for (const address of warningCells) {
    worksheet.getCell(address).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FEF3C7" },
    };
  }

  if (issues.length) {
    const row = 9 + minRows + 1;
    worksheet.mergeCells(`A${row}:F${row}`);
    worksheet.getCell(`A${row}`).value = `มี ${issues.length} จุดที่ต้องตรวจสอบ ดูรายละเอียดในชีต "รายการที่ต้องแก้"`;
    worksheet.getCell(`A${row}`).font = { color: { argb: "9A3412" }, size: 9 };
    worksheet.getCell(`A${row}`).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF7ED" },
    };
  }
}

function writeSummary(worksheet, records, recordIssues, config) {
  worksheet.views = [{ showGridLines: false, state: "frozen", ySplit: 4 }];
  worksheet.columns = [
    { width: 8 },
    { width: 32 },
    { width: 34 },
    { width: 16 },
    { width: 18 },
    { width: 22 },
    { width: 12 },
    { width: 58 },
  ];
  worksheet.mergeCells("A1:H1");
  worksheet.getCell("A1").value = "ตรวจสอบรายงาน ข.ย.11";
  worksheet.getCell("A2").value = config.demoRandomNames
    ? "โหมดข้อมูลจำลองเพื่อการศึกษา: ชื่อผู้ซื้อที่ว่างถูกสุ่มจากไฟล์รายชื่อ และไม่ใช่ข้อมูลสำหรับส่งจริง"
    : "ไฟล์นี้สร้างจากข้อมูล POS และไฮไลต์ช่องที่ยังต้องตรวจสอบก่อนส่งจริง";
  worksheet.getCell("A4").value = "ลำดับ";
  worksheet.getCell("B4").value = "ชื่อยา";
  worksheet.getCell("C4").value = "ช่วงวันที่";
  worksheet.getCell("D4").value = "จำนวนรายการขาย";
  worksheet.getCell("E4").value = "ช่องที่ต้องตรวจสอบ";
  worksheet.getCell("F4").value = "ชื่อแท็บต้นทาง";
  worksheet.getCell("G4").value = "สถานะ";
  worksheet.getCell("H4").value = "ตัวอย่างที่ต้องแก้";

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

  rows.forEach((row, idx) => {
    const target = worksheet.getRow(5 + idx);
    row.forEach((value, colIdx) => {
      target.getCell(colIdx + 1).value = value;
      target.getCell(colIdx + 1).alignment = { vertical: "top", wrapText: true };
      applyBorder(target.getCell(colIdx + 1), "thin", "D1D5DB");
    });
    target.getCell(1).alignment = { horizontal: "center" };
    target.getCell(4).alignment = { horizontal: "center" };
    target.getCell(5).alignment = { horizontal: "center" };
    target.getCell(7).alignment = { horizontal: "center" };
  });

  worksheet.getCell("A1").font = { bold: true, size: 14, color: { argb: "FFFFFF" } };
  worksheet.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "111827" } };
  worksheet.getCell("A1").alignment = { horizontal: "center" };
  worksheet.getCell("A2").font = { color: { argb: "4B5563" } };
  worksheet.getCell("A4").font = { bold: true };
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

  rows.forEach((row, idx) => {
    const target = worksheet.getRow(4 + idx);
    row.forEach((value, colIdx) => {
      target.getCell(colIdx + 1).value = value;
      target.getCell(colIdx + 1).alignment = { vertical: "top", wrapText: true };
      applyBorder(target.getCell(colIdx + 1), "thin", "FDBA74");
    });
    target.getCell(1).alignment = { horizontal: "center" };
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

async function convert(inputPath, outputPath, config = {}) {
  config = { ...DEFAULT_CONFIG, ...config };
  const sourceWorkbook = new ExcelJS.Workbook();
  await sourceWorkbook.xlsx.readFile(inputPath);

  const records = sourceWorkbook.worksheets.map(parseTable);
  splitOversizedSales(records, config);
  applyDemoRandomBuyerNames(records, config);
  const recordIssues = records.map((record) => issueList(record, config));

  const workbook = new ExcelJS.Workbook();
  const summarySheet = workbook.addWorksheet("ตรวจสอบ");
  writeSummary(summarySheet, records, recordIssues, config);
  const issueSheet = workbook.addWorksheet("รายการที่ต้องแก้");
  writeIssueDetails(issueSheet, records, recordIssues);

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
