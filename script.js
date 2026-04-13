const MAIN_SHEET = "2025";
const PREV_SHEET = "전년재고_DB";
const PREV2_SHEET = "전전년재고_DB";

const REQUIRED_MAIN_COLUMNS = ["지점명", "날짜", "품목", "판매수량", "판매금액", "최종폐기"];
const STOCK_CANDIDATE_QTY = ["수량", "재고수량", "이월수량", "잔량"];
const STOCK_CANDIDATE_AMT = ["금액", "재고금액", "이월금액", "잔액"];

let latestResult = null;

function toNumber(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  const cleaned = String(v).replace(/,/g, "").trim();
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
}

function normalizeText(v) {
  return String(v ?? "").trim();
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
  if (!d) return "";
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function findFirstColumn(row, candidates) {
  return candidates.find((key) => Object.prototype.hasOwnProperty.call(row, key));
}

function parseSheet(workbook, sheetName, allowEmpty = false) {
  const ws = workbook.Sheets[sheetName];

  if (!ws) {
    if (allowEmpty) return [];
    throw new Error(`시트 '${sheetName}' 을(를) 찾을 수 없습니다.`);
  }

  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

  if (!rows.length && allowEmpty) {
    return [];
  }

  return rows;
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

function normalizeHeaderText(v) {
  return String(v ?? "").replace(/\s+/g, "").trim();
}

function parseHorizontal2025Sheet(workbook) {
  const ws = workbook.Sheets[MAIN_SHEET];
  if (!ws) throw new Error(`시트 '${MAIN_SHEET}' 을(를) 찾을 수 없습니다.`);

  const range = XLSX.utils.decode_range(ws["!ref"]);
  const headerRowBranch = 9;
  const headerRowField = 10;
  const dataStartRow = 11;

  const rows = [];
  const branchBlocks = [];

  // 9행 = 지점명 / 10행 = 항목명
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

        if (c2 > c && branch2 && field2 === "판매수량") {
          break;
        }

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
    const dateValue = getCellValue(ws, `A${r}`);
    const itemValue = getCellValue(ws, `B${r}`);

    const 날짜 = formatDate(dateValue);
    const 품목 = normalizeText(itemValue);

    if (!날짜 && !품목) continue;

    branchBlocks.forEach((block) => {
      const 판매수량 = toNumber(getCellValue(ws, `${block.판매수량Col}${r}`));
      const 판매금액 = toNumber(getCellValue(ws, `${block.판매금액Col}${r}`));
      const 폐기수량 = block.폐기수량Col ? toNumber(getCellValue(ws, `${block.폐기수량Col}${r}`)) : 0;
      const 폐기금액 = block.폐기금액Col ? toNumber(getCellValue(ws, `${block.폐기금액Col}${r}`)) : 0;

      if (판매수량 === 0 && 판매금액 === 0 && 폐기수량 === 0 && 폐기금액 === 0) {
        return;
      }

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

  return rows;
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
  // 재고 시트가 비어 있으면 0재고로 간주하고 통과
  if (!rows.length) return [];

  const first = rows[0] || {};
  const cols = Object.keys(first);
  const qtyCol = STOCK_CANDIDATE_QTY.find((c) => cols.includes(c));
  const amtCol = STOCK_CANDIDATE_AMT.find((c) => cols.includes(c));
  const errors = [];

  if (!cols.includes("지점명")) errors.push(`${sheetName} 시트 필수 컬럼 누락: 지점명`);
  if (!cols.includes("품목")) errors.push(`${sheetName} 시트 필수 컬럼 누락: 품목`);
  if (!qtyCol) errors.push(`${sheetName} 시트 수량 컬럼을 찾을 수 없습니다.`);
  if (!amtCol) errors.push(`${sheetName} 시트 금액 컬럼을 찾을 수 없습니다.`);

  return errors;
}

function buildOpeningStocks(prev2Rows, prevRows) {
  const result = new Map();

  const ingest = (rows, yearType) => {
    if (!rows.length) return;
    const qtyCol = findFirstColumn(rows[0], STOCK_CANDIDATE_QTY);
    const amtCol = findFirstColumn(rows[0], STOCK_CANDIDATE_AMT);

    rows.forEach((row) => {
      const branch = normalizeText(row["지점명"]);
      const item = normalizeText(row["품목"]);
      const key = `${branch}__${item}`;
      const qty = toNumber(row[qtyCol]);
      const amt = toNumber(row[amtCol]);

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

    if (!map.has(key)) {
      map.set(key, { qty: 0, amt: 0 });
    }

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
    const rowNo = idx + 2;
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
    const usedQty = Math.min(remainQty, qtyAvailable);
    remainQty -= usedQty;

    const usedAmt = Math.min(remainAmt, amtAvailable);
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

function processWorkbook(workbook) {
  const mainRows = parseHorizontal2025Sheet(workbook);
const prevRows = parseSheet(workbook, PREV_SHEET, true);
const prev2Rows = parseSheet(workbook, PREV2_SHEET, true);
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

  const sortedMain = [...mainRows].sort((a, b) => {
    const da = formatDate(a["날짜"]);
    const db = formatDate(b["날짜"]);
    return da.localeCompare(db);
  });

  sortedMain.forEach((row) => {
    const branch = normalizeText(row["지점명"]);
    const date = formatDate(row["날짜"]);
    const item = normalizeText(row["품목"]);
    const saleQty = toNumber(row["판매수량"]);
    const saleAmt = toNumber(row["판매금액"]);
    const discardQty = toNumber(row["최종폐기"]);
    const key = `${branch}__${item}`;

    if (!stockMap.has(key)) {
      stockMap.set(key, {
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

    const stock = stockMap.get(key);

    if (saleQty > 0 || saleAmt > 0) {
      const result = allocateByFIFO(saleQty, saleAmt, stock);

      stock.전전년수량 -= result.전전년수량;
      stock.전전년금액 -= result.전전년금액;
      stock.전년수량 -= result.전년수량;
      stock.전년금액 -= result.전년금액;
      stock.당해수량 -= result.당해수량;
      stock.당해금액 -= result.당해금액;

      salesRows.push({
        구분: "판매",
        지점명: branch,
        날짜: date,
        품목: item,
        총사용수량: saleQty,
        총사용금액: saleAmt,
        전전년사용수량: result.전전년수량,
        전전년사용금액: result.전전년금액,
        전년사용수량: result.전년수량,
        전년사용금액: result.전년금액,
        당해사용수량: result.당해수량,
        당해사용금액: result.당해금액,
        부족수량: result.부족수량,
        부족금액: result.부족금액,
      });

      validationRows.push({
        구분: "판매",
        지점명: branch,
        날짜: date,
        품목: item,
        총사용수량: saleQty,
        연차별사용수량합: result.전전년수량 + result.전년수량 + result.당해수량,
        수량일치: saleQty === (result.전전년수량 + result.전년수량 + result.당해수량),
        총사용금액: saleAmt,
        연차별사용금액합: result.전전년금액 + result.전년금액 + result.당해금액,
        금액일치: saleAmt === (result.전전년금액 + result.전년금액 + result.당해금액),
        부족수량: result.부족수량,
        부족금액: result.부족금액,
      });
    }

    if (discardQty > 0) {
      const totalQtyAvailable = stock.전전년수량 + stock.전년수량 + stock.당해수량;
      const totalAmtAvailable = stock.전전년금액 + stock.전년금액 + stock.당해금액;
      const qtyResult = allocateByFIFO(discardQty, 0, stock);

      const usedQty = qtyResult.전전년수량 + qtyResult.전년수량 + qtyResult.당해수량;

      let 전전년폐기금액 = 0;
      let 전년폐기금액 = 0;
      let 당해폐기금액 = 0;

      if (totalQtyAvailable > 0 && totalAmtAvailable > 0 && usedQty > 0) {
        const unitAmt = totalAmtAvailable / totalQtyAvailable;

        전전년폐기금액 = Math.round(qtyResult.전전년수량 * unitAmt);
        전년폐기금액 = Math.round(qtyResult.전년수량 * unitAmt);
        당해폐기금액 = Math.round(qtyResult.당해수량 * unitAmt);

        const targetTotalAmt = Math.round(usedQty * unitAmt);
        const allocatedAmt = 전전년폐기금액 + 전년폐기금액 + 당해폐기금액;
        const diff = targetTotalAmt - allocatedAmt;

        if (diff !== 0) {
          당해폐기금액 += diff;
        }
      }

      stock.전전년수량 -= qtyResult.전전년수량;
      stock.전년수량 -= qtyResult.전년수량;
      stock.당해수량 -= qtyResult.당해수량;
      stock.전전년금액 -= 전전년폐기금액;
      stock.전년금액 -= 전년폐기금액;
      stock.당해금액 -= 당해폐기금액;

      discardRows.push({
        구분: "폐기",
        지점명: branch,
        날짜: date,
        품목: item,
        총사용수량: discardQty,
        총사용금액: 전전년폐기금액 + 전년폐기금액 + 당해폐기금액,
        전전년사용수량: qtyResult.전전년수량,
        전전년사용금액: 전전년폐기금액,
        전년사용수량: qtyResult.전년수량,
        전년사용금액: 전년폐기금액,
        당해사용수량: qtyResult.당해수량,
        당해사용금액: 당해폐기금액,
        부족수량: qtyResult.부족수량,
        부족금액: 0,
      });

      validationRows.push({
        구분: "폐기",
        지점명: branch,
        날짜: date,
        품목: item,
        총사용수량: discardQty,
        연차별사용수량합: qtyResult.전전년수량 + qtyResult.전년수량 + qtyResult.당해수량,
        수량일치: discardQty === (qtyResult.전전년수량 + qtyResult.전년수량 + qtyResult.당해수량),
        총사용금액: 전전년폐기금액 + 전년폐기금액 + 당해폐기금액,
        연차별사용금액합: 전전년폐기금액 + 전년폐기금액 + 당해폐기금액,
        금액일치: true,
        부족수량: qtyResult.부족수량,
        부족금액: 0,
      });
    }
  });

  return {
    ok: true,
    schemaErrors,
    anomalyIssues,
    salesRows,
    discardRows,
    validationRows,
  };
}

function renderIssues(result) {
  const container = document.getElementById("issues");
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

  rows.slice(0, 300).forEach((row) => {
    html += "<tr>";
    cols.forEach((c) => {
      html += `<td>${row[c] ?? ""}</td>`;
    });
    html += "</tr>";
  });

  html += "</tbody></table>";
  return html;
}

function updateStats(result) {
  document.getElementById("salesCount").textContent = result.salesRows.length;
  document.getElementById("discardCount").textContent = result.discardRows.length;

  const issues = result.schemaErrors.length + result.anomalyIssues.length;
  document.getElementById("issueCount").textContent = issues;

  const shortages = [...result.salesRows, ...result.discardRows].filter(
    (r) => toNumber(r.부족수량) > 0 || toNumber(r.부족금액) > 0
  ).length;

  document.getElementById("shortageCount").textContent = shortages;
}

function renderTables(result) {
  document.getElementById("salesTable").innerHTML = createTable(result.salesRows);
  document.getElementById("discardTable").innerHTML = createTable(result.discardRows);
  document.getElementById("validationTable").innerHTML = createTable(result.validationRows);
}

function downloadWorkbook(result) {
  const wb = XLSX.utils.book_new();

  const ws1 = XLSX.utils.json_to_sheet(result.salesRows);
  const ws2 = XLSX.utils.json_to_sheet(result.discardRows);
  const ws3 = XLSX.utils.json_to_sheet(result.validationRows);

  const issueRows = [
    ...result.schemaErrors.map((msg) => ({ 유형: "스키마오류", 내용: msg })),
    ...result.anomalyIssues.map((x) => ({ 유형: x.type, 내용: x.message })),
  ];
  const ws4 = XLSX.utils.json_to_sheet(issueRows);

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
