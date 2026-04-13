const MAIN_SHEET = "2025";
const PREV_SHEET = "전년재고_DB";
const PREV2_SHEET = "전전년재고_DB";

const REQUIRED_MAIN_COLUMNS = ["지점명", "날짜", "품목", "판매수량", "판매금액", "최종폐기"];

const ITEM_ORDER = [
  "의류",
  "잡화",
  "생활",
  "문화",
  "건강미용",
  "식품",
  "기증파트너",
];

const INVENTORY_HEADER_CANDIDATES = {
  지점명: ["지점명", "지점", "매장명", "점포명"],
  품목: ["품목", "품목명", "품목군"],
  수량: ["수량", "재고수량", "이월수량", "잔량"],
  금액: ["금액", "재고금액", "이월금액", "잔액"],
};

let latestResult = null;

function toNumber(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  const cleaned = String(v).replace(/,/g, "").trim();
  if (cleaned === "-" || cleaned === "—") return 0;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
}

function normalizeText(v) {
  return String(v ?? "").trim();
}

function normalizeHeaderText(v) {
  return String(v ?? "").replace(/\s+/g, "").trim();
}

function excelDateToJS(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  if (typeof value === "number") {
    const utcDays = Math.floor(value - 25569);
    const utcValue = utcDays * 86400;
    return new Date(utcValue * 1000);
  }
  const d = new Date(value);
  return Number.isNaN(d.getTime()) ? null : d;
}

function formatDate(value) {
  const d = excelDateToJS(value);
  if (!d) return normalizeText(value);
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function getCellValue(ws, addr) {
  return ws[addr] ? ws[addr].v : "";
}

function numberToCol(num) {
  let col = "";
  while (num > 0) {
    const rem = (num - 1) % 26;
    col = String.fromCharCode(65 + rem) + col;
    num = Math.floor((num - 1) / 26);
  }
  return col;
}

function itemOrderIndex(item) {
  const normalized = normalizeText(item);
  const idx = ITEM_ORDER.indexOf(normalized);
  return idx === -1 ? 999 : idx;
}

function compareRowsByBusinessOrder(a, b) {
  const da = normalizeText(a["일자"] || a["날짜"]);
  const db = normalizeText(b["일자"] || b["날짜"]);
  if (da !== db) return da.localeCompare(db);

  const ba = normalizeText(a["지점명"]);
  const bb = normalizeText(b["지점명"]);
  if (ba !== bb) return ba.localeCompare(bb);

  const ia = itemOrderIndex(a["품목군"] || a["품목"]);
  const ib = itemOrderIndex(b["품목군"] || b["품목"]);
  if (ia !== ib) return ia - ib;

  const ta = normalizeText(a["품목군"] || a["품목"]);
  const tb = normalizeText(b["품목군"] || b["품목"]);
  return ta.localeCompare(tb);
}

function sortRowsByBusinessOrder(rows) {
  return [...rows].sort(compareRowsByBusinessOrder);
}

function parseHorizontal2025Sheet(workbook) {
  const ws = workbook.Sheets[MAIN_SHEET];
  if (!ws) throw new Error(`시트 '${MAIN_SHEET}' 을(를) 찾을 수 없습니다.`);
  if (!ws["!ref"]) throw new Error(`시트 '${MAIN_SHEET}' 의 범위를 읽을 수 없습니다.`);

  const range = XLSX.utils.decode_range(ws["!ref"]);
  const headerRowBranch = 9;
  const headerRowField = 10;
  const dataStartRow = 11;

  const rows = [];
  const branchBlocks = [];

  for (let c = 1; c <= range.e.c + 1; c++) {
    const col = numberToCol(c);
    const branchName = normalizeText(getCellValue(ws, `${col}${headerRowBranch}`));
    const fieldName = normalizeHeaderText(getCellValue(ws, `${col}${headerRowField}`));

    if (branchName && fieldName === "판매수량") {
      const block = {
        지점명: branchName,
        판매수량Col: col,
        판매금액Col: null,
        폐기수량Col: null,
        폐기금액Col: null,
      };

      for (let c2 = c; c2 <= range.e.c + 1; c2++) {
        const col2 = numberToCol(c2);
        const branch2 = normalizeText(getCellValue(ws, `${col2}${headerRowBranch}`));
        const field2 = normalizeHeaderText(getCellValue(ws, `${col2}${headerRowField}`));

        if (c2 > c && branch2 && field2 === "판매수량") break;

        if (field2 === "판매수량") block.판매수량Col = col2;
        if (field2 === "판매금액") block.판매금액Col = col2;
        if (field2 === "폐기수량" || field2 === "최종폐기") block.폐기수량Col = col2;
        if (field2 === "폐기금액") block.폐기금액Col = col2;
      }

      branchBlocks.push(block);
    }
  }

  if (!branchBlocks.length) {
    throw new Error("2025 시트에서 지점 블록을 찾지 못했습니다. 9행=지점명, 10행=판매수량/판매금액/폐기수량/폐기금액 구조인지 확인해주세요.");
  }

  for (let r = dataStartRow; r <= range.e.r + 1; r++) {
    const 날짜 = formatDate(getCellValue(ws, `A${r}`));
    const 품목 = normalizeText(getCellValue(ws, `B${r}`));

    if (!날짜 && !품목) continue;

    branchBlocks.forEach((block) => {
      const 판매수량 = toNumber(getCellValue(ws, `${block.판매수량Col}${r}`));
      const 판매금액 = block.판매금액Col ? toNumber(getCellValue(ws, `${block.판매금액Col}${r}`)) : 0;
      const 폐기수량 = block.폐기수량Col ? toNumber(getCellValue(ws, `${block.폐기수량Col}${r}`)) : 0;
      const 폐기금액 = block.폐기금액Col ? toNumber(getCellValue(ws, `${block.폐기금액Col}${r}`)) : 0;

      if (판매수량 === 0 && 판매금액 === 0 && 폐기수량 === 0 && 폐기금액 === 0) return;

      rows.push({
        지점명: block.지점명,
        날짜,
        품목,
        판매수량,
        판매금액,
        최종폐기: 폐기수량,
        폐기금액,
      });
    });
  }

  return sortRowsByBusinessOrder(rows);
}

function findHeaderIndex(headerRow, candidates) {
  for (let i = 0; i < headerRow.length; i++) {
    const cell = normalizeHeaderText(headerRow[i]);
    if (candidates.includes(cell)) return i;
  }
  return -1;
}

function parseInventorySheet(workbook, sheetName) {
  const ws = workbook.Sheets[sheetName];
  if (!ws || !ws["!ref"]) return [];

  const sheet = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (!sheet.length) return [];

  let headerRowIndex = -1;
  let idxBranch = -1;
  let idxItem = -1;
  let idxQty = -1;
  let idxAmt = -1;

  for (let i = 0; i < Math.min(sheet.length, 30); i++) {
    const headerRow = (sheet[i] || []).map((x) => normalizeHeaderText(x));

    const branchIdx = findHeaderIndex(headerRow, INVENTORY_HEADER_CANDIDATES.지점명);
    const itemIdx = findHeaderIndex(headerRow, INVENTORY_HEADER_CANDIDATES.품목);
    const qtyIdx = findHeaderIndex(headerRow, INVENTORY_HEADER_CANDIDATES.수량);
    const amtIdx = findHeaderIndex(headerRow, INVENTORY_HEADER_CANDIDATES.금액);

    if (branchIdx !== -1 && itemIdx !== -1 && qtyIdx !== -1 && amtIdx !== -1) {
      headerRowIndex = i;
      idxBranch = branchIdx;
      idxItem = itemIdx;
      idxQty = qtyIdx;
      idxAmt = amtIdx;
      break;
    }
  }

  if (headerRowIndex === -1) return [];

  const rows = [];

  for (let i = headerRowIndex + 1; i < sheet.length; i++) {
    const row = sheet[i] || [];
    const 지점명 = normalizeText(row[idxBranch]);
    const 품목 = normalizeText(row[idxItem]);
    const 수량 = toNumber(row[idxQty]);
    const 금액 = toNumber(row[idxAmt]);

    if (!지점명 && !품목 && 수량 === 0 && 금액 === 0) continue;

    rows.push({ 지점명, 품목, 수량, 금액 });
  }

  return rows.sort((a, b) => {
    const ba = normalizeText(a.지점명);
    const bb = normalizeText(b.지점명);
    if (ba !== bb) return ba.localeCompare(bb);

    const ia = itemOrderIndex(a.품목);
    const ib = itemOrderIndex(b.품목);
    if (ia !== ib) return ia - ib;

    return normalizeText(a.품목).localeCompare(normalizeText(b.품목));
  });
}

function validateMainColumns(rows) {
  if (!rows.length) return [`${MAIN_SHEET} 시트에서 변환된 사용 데이터가 없습니다.`];

  const first = rows[0] || {};
  const cols = Object.keys(first);

  return REQUIRED_MAIN_COLUMNS
    .filter((c) => !cols.includes(c))
    .map((c) => `${MAIN_SHEET} 변환 데이터 필수 컬럼 누락: ${c}`);
}

function validateStockColumns(rows, sheetName) {
  if (!rows.length) return [];

  const first = rows[0] || {};
  const cols = Object.keys(first);
  const errors = [];

  if (!cols.includes("지점명")) errors.push(`${sheetName} 시트 필수 컬럼 누락: 지점명`);
  if (!cols.includes("품목")) errors.push(`${sheetName} 시트 필수 컬럼 누락: 품목`);
  if (!cols.includes("수량")) errors.push(`${sheetName} 시트 수량 컬럼을 찾을 수 없습니다.`);
  if (!cols.includes("금액")) errors.push(`${sheetName} 시트 금액 컬럼을 찾을 수 없습니다.`);

  return errors;
}

function buildOpeningStocks(prev2Rows, prevRows) {
  const result = new Map();

  const ingest = (rows, yearType) => {
    rows.forEach((row) => {
      const branch = normalizeText(row["지점명"]);
      const item = normalizeText(row["품목"]);
      const key = `${branch}__${item}`;
      const qty = toNumber(row["수량"]);
      const amt = toNumber(row["금액"]);

      if (!branch || !item) return;

      if (!result.has(key)) {
        result.set(key, {
          지점명: branch,
          품목: item,
          전전년수량: 0,
          전전년금액: 0,
          전년수량: 0,
          전년금액: 0,
          당해수량: 0,
          당해금액: 0,
        });
      }

      const target = result.get(key);
      if (yearType === "전전년") {
        target.전전년수량 += qty;
        target.전전년금액 += amt;
      } else {
        target.전년수량 += qty;
        target.전년금액 += amt;
      }
    });
  };

  ingest(prev2Rows, "전전년");
  ingest(prevRows, "전년");

  return result;
}

function aggregateCurrentYearInputs(mainRows) {
  const map = new Map();

  mainRows.forEach((row) => {
    const branch = normalizeText(row["지점명"]);
    const item = normalizeText(row["품목"]);
    const saleQty = toNumber(row["판매수량"]);
    const saleAmt = toNumber(row["판매금액"]);
    const key = `${branch}__${item}`;

    if (!branch || !item) return;

    if (!map.has(key)) map.set(key, { qty: 0, amt: 0 });

    const agg = map.get(key);
    agg.qty += saleQty;
    agg.amt += saleAmt;
  });

  return map;
}

function enrichWithCurrentYear(openingStocks, currentInputs) {
  for (const [key, v] of currentInputs.entries()) {
    if (!openingStocks.has(key)) {
      const [branch, item] = key.split("__");
      openingStocks.set(key, {
        지점명: branch,
        품목: item,
        전전년수량: 0,
        전전년금액: 0,
        전년수량: 0,
        전년금액: 0,
        당해수량: 0,
        당해금액: 0,
      });
    }
    const target = openingStocks.get(key);
    target.당해수량 += v.qty;
    target.당해금액 += v.amt;
  }
  return openingStocks;
}

function validateAnomalies(mainRows) {
  const issues = [];

  mainRows.forEach((row, idx) => {
    const rowNo = idx + 1;
    const saleQty = toNumber(row["판매수량"]);
    const saleAmt = toNumber(row["판매금액"]);
    const discardQty = toNumber(row["최종폐기"]);

    if (saleQty === 0 && saleAmt > 0) {
      issues.push({ type: "이상치", message: `행 ${rowNo}: 판매수량 0인데 판매금액이 존재합니다.` });
    }

    if (saleQty > 0 && saleAmt === 0) {
      issues.push({ type: "이상치", message: `행 ${rowNo}: 판매금액 0인데 판매수량이 존재합니다.` });
    }

    if (saleQty < 0 || saleAmt < 0 || discardQty < 0) {
      issues.push({ type: "음수값", message: `행 ${rowNo}: 음수값이 존재합니다.` });
    }
  });

  return issues;
}

function allocateByFIFO(totalQty, totalAmt, layers) {
  let remainQty = toNumber(totalQty);
  let remainAmt = toNumber(totalAmt);

  const output = {
    전전년수량: 0,
    전전년금액: 0,
    전년수량: 0,
    전년금액: 0,
    당해수량: 0,
    당해금액: 0,
    부족수량: 0,
    부족금액: 0,
  };

  const useLayer = (yearKey, qtyAvailable, amtAvailable) => {
    const usedQty = Math.min(remainQty, Math.max(0, qtyAvailable));
    remainQty -= usedQty;

    const usedAmt = Math.min(remainAmt, Math.max(0, amtAvailable));
    remainAmt -= usedAmt;

    output[`${yearKey}수량`] += usedQty;
    output[`${yearKey}금액`] += usedAmt;
  };

  useLayer("전전년", layers.전전년수량, layers.전전년금액);
  useLayer("전년", layers.전년수량, layers.전년금액);
  useLayer("당해", layers.당해수량, layers.당해금액);

  output.부족수량 = remainQty;
  output.부족금액 = remainAmt;

  return output;
}

function createDailyCombinedRow(branch, date, item) {
  return {
    지점명: branch,
    일자: date,
    품목군: item,
    판매수량: 0,
    판매금액: 0,
    "폐기수량(최종)": 0,
    폐기금액: 0,
    총사용수량: 0,
    총사용금액: 0,

    전전년_판매수량: 0,
    전전년_판매금액: 0,
    전전년_폐기수량: 0,
    전전년_폐기금액: 0,

    전년_판매수량: 0,
    전년_판매금액: 0,
    전년_폐기수량: 0,
    전년_폐기금액: 0,

    당해_판매수량: 0,
    당해_판매금액: 0,
    당해_폐기수량: 0,
    당해_폐기금액: 0,
  };
}

function processWorkbook(workbook) {
  const mainRows = parseHorizontal2025Sheet(workbook);
  const prevRows = parseInventorySheet(workbook, PREV_SHEET);
  const prev2Rows = parseInventorySheet(workbook, PREV2_SHEET);

  const schemaErrors = [
    ...validateMainColumns(mainRows),
    ...validateStockColumns(prevRows, PREV_SHEET),
    ...validateStockColumns(prev2Rows, PREV2_SHEET),
  ];

  const anomalyIssues = validateAnomalies(mainRows);

  if (schemaErrors.length) {
    return {
      ok: false,
      schemaErrors,
      anomalyIssues,
      mergedDailyRows: [],
      salesRows: [],
      discardRows: [],
      validationRows: [],
    };
  }

  const openingStocks = buildOpeningStocks(prev2Rows, prevRows);
  const currentInputs = aggregateCurrentYearInputs(mainRows);
  const stockMap = enrichWithCurrentYear(openingStocks, currentInputs);

  const salesRows = [];
  const discardRows = [];
  const validationRows = [];
  const dailyMap = new Map();

  const sortedMain = sortRowsByBusinessOrder(mainRows);

  sortedMain.forEach((row) => {
    const branch = normalizeText(row["지점명"]);
    const date = normalizeText(row["날짜"]);
    const item = normalizeText(row["품목"]);
    const saleQty = toNumber(row["판매수량"]);
    const saleAmt = toNumber(row["판매금액"]);
    const discardQty = toNumber(row["최종폐기"]);
    const discardAmtInput = toNumber(row["폐기금액"]);
    const stockKey = `${branch}__${item}`;
    const dailyKey = `${branch}__${date}__${item}`;

    if (!stockMap.has(stockKey)) {
      stockMap.set(stockKey, {
        지점명: branch,
        품목: item,
        전전년수량: 0,
        전전년금액: 0,
        전년수량: 0,
        전년금액: 0,
        당해수량: 0,
        당해금액: 0,
      });
    }

    if (!dailyMap.has(dailyKey)) {
      dailyMap.set(dailyKey, createDailyCombinedRow(branch, date, item));
    }

    const dailyRow = dailyMap.get(dailyKey);
    const stock = stockMap.get(stockKey);

    let saleResult = {
      전전년수량: 0, 전전년금액: 0,
      전년수량: 0, 전년금액: 0,
      당해수량: 0, 당해금액: 0,
      부족수량: 0, 부족금액: 0,
    };

    if (saleQty > 0 || saleAmt > 0) {
      saleResult = allocateByFIFO(saleQty, saleAmt, stock);

      stock.전전년수량 -= saleResult.전전년수량;
      stock.전전년금액 -= saleResult.전전년금액;
      stock.전년수량 -= saleResult.전년수량;
      stock.전년금액 -= saleResult.전년금액;
      stock.당해수량 -= saleResult.당해수량;
      stock.당해금액 -= saleResult.당해금액;

      salesRows.push({
        구분: "판매",
        지점명: branch,
        날짜: date,
        품목: item,
        총사용수량: saleQty,
        총사용금액: saleAmt,
        전전년사용수량: saleResult.전전년수량,
        전전년사용금액: saleResult.전전년금액,
        전년사용수량: saleResult.전년수량,
        전년사용금액: saleResult.전년금액,
        당해사용수량: saleResult.당해수량,
        당해사용금액: saleResult.당해금액,
        부족수량: saleResult.부족수량,
        부족금액: saleResult.부족금액,
      });
    }

    let discardQtyResult = {
      전전년수량: 0, 전전년금액: 0,
      전년수량: 0, 전년금액: 0,
      당해수량: 0, 당해금액: 0,
      부족수량: 0, 부족금액: 0,
    };

    let 전전년폐기금액 = 0;
    let 전년폐기금액 = 0;
    let 당해폐기금액 = 0;
    let 표시폐기금액 = 0;
    let 부족폐기금액 = 0;

    if (discardQty > 0 || discardAmtInput > 0) {
      discardQtyResult = allocateByFIFO(discardQty, 0, stock);

      if (discardAmtInput > 0) {
        let remainDiscardAmt = discardAmtInput;

        const amtFromPrev2 = Math.min(remainDiscardAmt, Math.max(0, stock.전전년금액));
        remainDiscardAmt -= amtFromPrev2;

        const amtFromPrev = Math.min(remainDiscardAmt, Math.max(0, stock.전년금액));
        remainDiscardAmt -= amtFromPrev;

        const amtFromCurrent = Math.min(remainDiscardAmt, Math.max(0, stock.당해금액));
        remainDiscardAmt -= amtFromCurrent;

        전전년폐기금액 = amtFromPrev2;
        전년폐기금액 = amtFromPrev;
        당해폐기금액 = amtFromCurrent;
        표시폐기금액 = discardAmtInput;
        부족폐기금액 = Math.max(0, remainDiscardAmt);
      } else {
        const totalQtyAvailable = stock.전전년수량 + stock.전년수량 + stock.당해수량;
        const totalAmtAvailable = stock.전전년금액 + stock.전년금액 + stock.당해금액;
        const usedQty = discardQtyResult.전전년수량 + discardQtyResult.전년수량 + discardQtyResult.당해수량;

        if (totalQtyAvailable > 0 && totalAmtAvailable > 0 && usedQty > 0) {
          const unitAmt = totalAmtAvailable / totalQtyAvailable;
          전전년폐기금액 = Math.round(discardQtyResult.전전년수량 * unitAmt);
          전년폐기금액 = Math.round(discardQtyResult.전년수량 * unitAmt);
          당해폐기금액 = Math.round(discardQtyResult.당해수량 * unitAmt);

          const targetTotalAmt = Math.round(usedQty * unitAmt);
          const allocatedAmt = 전전년폐기금액 + 전년폐기금액 + 당해폐기금액;
          const diff = targetTotalAmt - allocatedAmt;
          if (diff !== 0) 당해폐기금액 += diff;
        }

        표시폐기금액 = 전전년폐기금액 + 전년폐기금액 + 당해폐기금액;
        부족폐기금액 = 0;
      }

      stock.전전년수량 -= discardQtyResult.전전년수량;
      stock.전년수량 -= discardQtyResult.전년수량;
      stock.당해수량 -= discardQtyResult.당해수량;

      stock.전전년금액 -= 전전년폐기금액;
      stock.전년금액 -= 전년폐기금액;
      stock.당해금액 -= 당해폐기금액;

      discardRows.push({
        구분: "폐기",
        지점명: branch,
        날짜: date,
        품목: item,
        총사용수량: discardQty,
        총사용금액: 표시폐기금액,
        입력폐기금액: discardAmtInput,
        전전년사용수량: discardQtyResult.전전년수량,
        전전년사용금액: 전전년폐기금액,
        전년사용수량: discardQtyResult.전년수량,
        전년사용금액: 전년폐기금액,
        당해사용수량: discardQtyResult.당해수량,
        당해사용금액: 당해폐기금액,
        부족수량: discardQtyResult.부족수량,
        부족금액: 부족폐기금액,
      });
    }

    validationRows.push({
      지점명: branch,
      날짜: date,
      품목: item,
      판매수량: saleQty,
      판매금액: saleAmt,
      폐기수량: discardQty,
      폐기금액: 표시폐기금액,
      판매_연차합수량: saleResult.전전년수량 + saleResult.전년수량 + saleResult.당해수량,
      판매_연차합금액: saleResult.전전년금액 + saleResult.전년금액 + saleResult.당해금액,
      폐기_연차합수량: discardQtyResult.전전년수량 + discardQtyResult.전년수량 + discardQtyResult.당해수량,
      폐기_연차합금액: 전전년폐기금액 + 전년폐기금액 + 당해폐기금액,
      부족수량: saleResult.부족수량 + discardQtyResult.부족수량,
      부족금액: saleResult.부족금액 + 부족폐기금액,
    });

    dailyRow.판매수량 += saleQty;
    dailyRow.판매금액 += saleAmt;
    dailyRow["폐기수량(최종)"] += discardQty;
    dailyRow.폐기금액 += 표시폐기금액;
    dailyRow.총사용수량 += saleQty + discardQty;
    dailyRow.총사용금액 += saleAmt + 표시폐기금액;

    dailyRow.전전년_판매수량 += saleResult.전전년수량;
    dailyRow.전전년_판매금액 += saleResult.전전년금액;
    dailyRow.전전년_폐기수량 += discardQtyResult.전전년수량;
    dailyRow.전전년_폐기금액 += 전전년폐기금액;

    dailyRow.전년_판매수량 += saleResult.전년수량;
    dailyRow.전년_판매금액 += saleResult.전년금액;
    dailyRow.전년_폐기수량 += discardQtyResult.전년수량;
    dailyRow.전년_폐기금액 += 전년폐기금액;

    dailyRow.당해_판매수량 += saleResult.당해수량;
    dailyRow.당해_판매금액 += saleResult.당해금액;
    dailyRow.당해_폐기수량 += discardQtyResult.당해수량;
    dailyRow.당해_폐기금액 += 당해폐기금액;
  });

  const mergedDailyRows = sortRowsByBusinessOrder(Array.from(dailyMap.values()));

  return {
    ok: true,
    schemaErrors,
    anomalyIssues,
    mergedDailyRows,
    salesRows: sortRowsByBusinessOrder(salesRows.map(x => ({
      ...x,
      일자: x.날짜,
      품목군: x.품목
    }))).map(({ 일자, 품목군, ...rest }) => rest),
    discardRows: sortRowsByBusinessOrder(discardRows.map(x => ({
      ...x,
      일자: x.날짜,
      품목군: x.품목
    }))).map(({ 일자, 품목군, ...rest }) => rest),
    validationRows: sortRowsByBusinessOrder(validationRows.map(x => ({
      ...x,
      일자: x.날짜,
      품목군: x.품목
    }))).map(({ 일자, 품목군, ...rest }) => rest),
  };
}

function renderIssues(result) {
  const container = document.getElementById("issues");
  if (!container) return;
  container.innerHTML = "";

  const issues = [
    ...result.schemaErrors.map((msg) => ({ message: msg })),
    ...result.anomalyIssues,
  ];

  if (!issues.length) {
    container.innerHTML = `<div class="issue ok">오류 및 이상치가 없습니다.</div>`;
    return;
  }

  issues.forEach((issue) => {
    const div = document.createElement("div");
    div.className = "issue";
    div.textContent = issue.message;
    container.appendChild(div);
  });
}

function createTable(rows) {
  if (!rows || !rows.length) return "<p>데이터가 없습니다.</p>";

  const cols = Object.keys(rows[0]);
  let html = "<table><thead><tr>";
  cols.forEach((c) => {
    html += `<th>${c}</th>`;
  });
  html += "</tr></thead><tbody>";

  rows.forEach((row) => {
    html += "<tr>";
    cols.forEach((c) => {
      html += `<td>${row[c] ?? ""}</td>`;
    });
    html += "</tr>";
  });

  html += "</tbody></table>";
  return html;
}

function populateBranchFilter(result) {
  const select = document.getElementById("branchFilter");
  if (!select) return;

  const branches = new Set();
  result.mergedDailyRows.forEach((row) => {
    if (row.지점명) branches.add(row.지점명);
  });

  const currentValue = select.value || "전체";
  select.innerHTML = `<option value="전체">전체</option>`;

  Array.from(branches).sort().forEach((branch) => {
    const option = document.createElement("option");
    option.value = branch;
    option.textContent = branch;
    select.appendChild(option);
  });

  select.value = Array.from(branches).includes(currentValue) ? currentValue : "전체";
}

function renderTables(result) {
  const selectedBranch = document.getElementById("branchFilter")?.value || "전체";

  const mergedDailyRows =
    selectedBranch === "전체"
      ? result.mergedDailyRows
      : result.mergedDailyRows.filter((row) => row.지점명 === selectedBranch);

  const salesRows =
    selectedBranch === "전체"
      ? result.salesRows
      : result.salesRows.filter((row) => row.지점명 === selectedBranch);

  const discardRows =
    selectedBranch === "전체"
      ? result.discardRows
      : result.discardRows.filter((row) => row.지점명 === selectedBranch);

  const validationRows =
    selectedBranch === "전체"
      ? result.validationRows
      : result.validationRows.filter((row) => row.지점명 === selectedBranch);

  const mergedTable = document.getElementById("mergedTable");
  if (mergedTable) mergedTable.innerHTML = createTable(mergedDailyRows);

  const salesTable = document.getElementById("salesTable");
  if (salesTable) salesTable.innerHTML = createTable(salesRows);

  const discardTable = document.getElementById("discardTable");
  if (discardTable) discardTable.innerHTML = createTable(discardRows);

  const validationTable = document.getElementById("validationTable");
  if (validationTable) validationTable.innerHTML = createTable(validationRows);
}

function updateStats(result) {
  document.getElementById("salesCount").textContent = result.salesRows.length;
  document.getElementById("discardCount").textContent = result.discardRows.length;

  const issues = result.schemaErrors.length + result.anomalyIssues.length;
  document.getElementById("issueCount").textContent = issues;

  const shortages = result.validationRows.filter(
    (r) => toNumber(r.부족수량) > 0 || toNumber(r.부족금액) > 0
  ).length;

  document.getElementById("shortageCount").textContent = shortages;
}

function downloadWorkbook(result) {
  const wb = XLSX.utils.book_new();

  const ws0 = XLSX.utils.json_to_sheet(result.mergedDailyRows);
  const ws1 = XLSX.utils.json_to_sheet(result.salesRows);
  const ws2 = XLSX.utils.json_to_sheet(result.discardRows);
  const ws3 = XLSX.utils.json_to_sheet(result.validationRows);

  const issueRows = [
    ...result.schemaErrors.map((msg) => ({ 유형: "스키마오류", 내용: msg })),
    ...result.anomalyIssues.map((x) => ({ 유형: x.type, 내용: x.message })),
  ];
  const ws4 = XLSX.utils.json_to_sheet(issueRows);

  XLSX.utils.book_append_sheet(wb, ws0, "일별통합결과");
  XLSX.utils.book_append_sheet(wb, ws1, "판매자동소진");
  XLSX.utils.book_append_sheet(wb, ws2, "폐기자동소진");
  XLSX.utils.book_append_sheet(wb, ws3, "검증");
  XLSX.utils.book_append_sheet(wb, ws4, "오류및이상치");

  XLSX.writeFile(wb, "FIFO_자동소진_결과.xlsx");
}

document.getElementById("fileInput").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  if (!file) return;

  try {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array" });
    const result = processWorkbook(workbook);
    latestResult = result;

    updateStats(result);
    renderIssues(result);
    populateBranchFilter(result);
    renderTables(result);

    document.getElementById("downloadBtn").disabled = false;
  } catch (error) {
    alert("파일 처리 중 오류가 발생했습니다: " + error.message);
  }
});

document.getElementById("downloadBtn").addEventListener("click", () => {
  if (!latestResult) return;
  downloadWorkbook(latestResult);
});

document.getElementById("branchFilter")?.addEventListener("change", () => {
  if (!latestResult) return;
  renderTables(latestResult);
});
