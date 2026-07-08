const MAIN_SHEET = "2025";
const PREV2_SHEET = "전전년재고_DB";
const PREV_SHEET = "전년재고_DB";

const ITEM_ORDER = ["의류", "잡화", "생활", "문화", "건강미용", "식품", "기증파트너"];

let latestWorkbook = null;
let lockedWorkbook = null;
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
    return new Date(utcDays * 86400 * 1000);
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

function todayString() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function setDefaultTodayDates() {
  const today = todayString();

  ["prev2CutoffDate", "prevCutoffDate", "finalCutoffDate"].forEach((id) => {
    const input = document.getElementById(id);
    if (input && !input.value) input.value = today;
  });
}

function getDateInput(id) {
  const input = document.getElementById(id);
  return input && input.value ? input.value : null;
}

function getPrev2CutoffDate() {
  return getDateInput("prev2CutoffDate") || getDateInput("cutoffDate");
}

function getPrevCutoffDate() {
  return getDateInput("prevCutoffDate") || getDateInput("cutoffDate");
}

function getFinalCutoffDate() {
  return getDateInput("finalCutoffDate") || getDateInput("cutoffDate");
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

function makeKey(branch, date, item) {
  return `${normalizeText(branch)}__${normalizeText(date)}__${normalizeText(item)}`;
}

function itemOrderIndex(item) {
  const idx = ITEM_ORDER.indexOf(normalizeText(item));
  return idx === -1 ? 999 : idx;
}

function sortRowsByBusinessOrder(rows) {
  return [...rows].sort((a, b) => {
    const da = normalizeText(a.일자 || a.날짜);
    const db = normalizeText(b.일자 || b.날짜);
    if (da !== db) return da.localeCompare(db);

    const ba = normalizeText(a.지점명);
    const bb = normalizeText(b.지점명);
    if (ba !== bb) return ba.localeCompare(bb);

    const ia = itemOrderIndex(a.품목군 || a.품목);
    const ib = itemOrderIndex(b.품목군 || b.품목);
    if (ia !== ib) return ia - ib;

    return normalizeText(a.품목군 || a.품목).localeCompare(normalizeText(b.품목군 || b.품목));
  });
}

function parseHorizontal2025Sheet(workbook) {
  const ws = workbook.Sheets[MAIN_SHEET];
  if (!ws) throw new Error(`시트 '${MAIN_SHEET}' 을(를) 찾을 수 없습니다.`);
  if (!ws["!ref"]) throw new Error(`시트 '${MAIN_SHEET}' 의 범위를 읽을 수 없습니다.`);

  const range = XLSX.utils.decode_range(ws["!ref"]);
  const rows = [];
  const branchBlocks = [];
  let currentBranch = null;

  const headerRowBranch = 9;
  const headerRowField = 10;
  const dataStartRow = 11;

  for (let c = 1; c <= range.e.c + 1; c++) {
    const col = numberToCol(c);
    const branchName = normalizeText(getCellValue(ws, `${col}${headerRowBranch}`));
    const fieldName = normalizeHeaderText(getCellValue(ws, `${col}${headerRowField}`));

    if (branchName) {
      if (currentBranch && currentBranch.지점명 && currentBranch.판매수량Col) {
        branchBlocks.push(currentBranch);
      }

      currentBranch = {
        지점명: branchName,
        판매수량Col: null,
        판매금액Col: null,
        폐기수량Col: null,
        폐기금액Col: null,
        waitingDiscardAmt: false,
      };
    }

    if (!currentBranch) continue;

    if (fieldName === "판매수량") currentBranch.판매수량Col = col;
    if (fieldName === "판매금액(사용명세서)") currentBranch.판매금액Col = col;

    if (fieldName === "최종폐기") {
      currentBranch.폐기수량Col = col;
      currentBranch.waitingDiscardAmt = true;
    }

    if (fieldName === "금액" && currentBranch.waitingDiscardAmt && !currentBranch.폐기금액Col) {
      currentBranch.폐기금액Col = col;
      currentBranch.waitingDiscardAmt = false;
    }
  }

  if (currentBranch && currentBranch.지점명 && currentBranch.판매수량Col) {
    branchBlocks.push(currentBranch);
  }

  const validBlocks = branchBlocks.filter(
    (b) => b.지점명 && b.판매수량Col && b.판매금액Col && b.폐기수량Col && b.폐기금액Col
  );

  if (!validBlocks.length) {
    throw new Error("2025 시트에서 유효한 지점 블록을 찾지 못했습니다.");
  }

  for (let r = dataStartRow; r <= range.e.r + 1; r++) {
    const 날짜 = formatDate(getCellValue(ws, `A${r}`));
    const 품목 = normalizeText(getCellValue(ws, `B${r}`));

    if (!날짜 && !품목) continue;

    validBlocks.forEach((block) => {
      const 판매수량 = toNumber(getCellValue(ws, `${block.판매수량Col}${r}`));
      const 판매금액 = toNumber(getCellValue(ws, `${block.판매금액Col}${r}`));
      const 폐기수량 = toNumber(getCellValue(ws, `${block.폐기수량Col}${r}`));
      const 폐기금액 = toNumber(getCellValue(ws, `${block.폐기금액Col}${r}`));

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

function parseInventorySheetFixed(workbook, sheetName) {
  const ws = workbook.Sheets[sheetName];
  if (!ws || !ws["!ref"]) return [];

  const sheet = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  const rows = [];

  for (let i = 1; i < sheet.length; i++) {
    const row = sheet[i] || [];

    const 지점명 = normalizeText(row[0]);
    const 품목 = normalizeText(row[1]);
    const 수량 = toNumber(row[3]);
    const 금액 = toNumber(row[4]);

    if (!지점명 && !품목 && 수량 === 0 && 금액 === 0) continue;

    rows.push({ 지점명, 품목, 수량, 금액 });
  }

  return rows;
}

function filterRowsByFinalCutoff(rows) {
  const cutoff = getFinalCutoffDate();
  if (!cutoff) return rows;
  return rows.filter((row) => normalizeText(row.날짜) <= cutoff);
}

function validateAnomalies(rows) {
  const issues = [];

  rows.forEach((row, idx) => {
    const rowNo = idx + 1;
    const saleQty = toNumber(row.판매수량);
    const saleAmt = toNumber(row.판매금액);
    const discardQty = toNumber(row.최종폐기);
    const discardAmt = toNumber(row.폐기금액);

    if (saleQty > 0 && saleAmt === 0) {
      issues.push({ type: "이상치", message: `행 ${rowNo}: 판매수량은 있는데 판매금액이 0입니다.` });
    }
    if (saleQty === 0 && saleAmt > 0) {
      issues.push({ type: "이상치", message: `행 ${rowNo}: 판매금액은 있는데 판매수량이 0입니다.` });
    }
    if (discardQty > 0 && discardAmt === 0) {
      issues.push({ type: "이상치", message: `행 ${rowNo}: 폐기수량은 있는데 폐기금액이 0입니다.` });
    }
    if (discardQty === 0 && discardAmt > 0) {
      issues.push({ type: "이상치", message: `행 ${rowNo}: 폐기금액은 있는데 폐기수량이 0입니다.` });
    }
  });

  return issues;
}

function buildOpeningStocks(prev2Rows, prevRows) {
  const map = new Map();

  const ensure = (branch, item) => {
    const key = `${branch}__${item}`;
    if (!map.has(key)) {
      map.set(key, {
        지점명: branch,
        품목: item,
        전전년수량: 0,
        전전년금액: 0,
        전년수량: 0,
        전년금액: 0,
      });
    }
    return map.get(key);
  };

  prev2Rows.forEach((row) => {
    const branch = normalizeText(row.지점명);
    const item = normalizeText(row.품목);
    if (!branch || !item) return;

    const target = ensure(branch, item);
    target.전전년수량 += toNumber(row.수량);
    target.전전년금액 += toNumber(row.금액);
  });

  prevRows.forEach((row) => {
    const branch = normalizeText(row.지점명);
    const item = normalizeText(row.품목);
    if (!branch || !item) return;

    const target = ensure(branch, item);
    target.전년수량 += toNumber(row.수량);
    target.전년금액 += toNumber(row.금액);
  });

  return map;
}

function parseLockedDailyRows(workbook) {
  if (!workbook) return [];

  const ws = workbook.Sheets["일별통합결과"];
  if (!ws || !ws["!ref"]) return [];

  return XLSX.utils.sheet_to_json(ws, { defval: "" }).map((r) => ({
    지점명: normalizeText(r.지점명),
    일자: formatDate(r.일자),
    품목군: normalizeText(r.품목군),

    판매수량: toNumber(r.판매수량),
    판매금액: toNumber(r.판매금액),
    폐기수량: toNumber(r.폐기수량),
    폐기금액: toNumber(r.폐기금액),

    총사용수량: toNumber(r.총사용수량),
    총사용금액: toNumber(r.총사용금액),

    전전년_판매수량: toNumber(r.전전년_판매수량),
    전전년_판매금액: toNumber(r.전전년_판매금액),
    전전년_폐기수량: toNumber(r.전전년_폐기수량),
    전전년_폐기금액: toNumber(r.전전년_폐기금액),

    전년_판매수량: toNumber(r.전년_판매수량),
    전년_판매금액: toNumber(r.전년_판매금액),
    전년_폐기수량: toNumber(r.전년_폐기수량),
    전년_폐기금액: toNumber(r.전년_폐기금액),

    당해_판매수량: toNumber(r.당해_판매수량 ?? r.당월_판매수량),
    당해_판매금액: toNumber(r.당해_판매금액 ?? r.당월_판매금액),
    당해_폐기수량: toNumber(r.당해_폐기수량 ?? r.당월_폐기수량),
    당해_폐기금액: toNumber(r.당해_폐기금액 ?? r.당월_폐기금액),

    부족수량: toNumber(r.부족수량),
    부족금액: toNumber(r.부족금액),
  }));
}

function buildLockedMap(lockedRows) {
  const map = new Map();

  lockedRows.forEach((r) => {
    const key = makeKey(r.지점명, r.일자, r.품목군);
    map.set(key, {
      판매수량: toNumber(r.판매수량),
      판매금액: toNumber(r.판매금액),
      폐기수량: toNumber(r.폐기수량),
      폐기금액: toNumber(r.폐기금액),
    });
  });

  return map;
}

function buildCurrentDailyRows(mainRows) {
  const map = new Map();

  mainRows.forEach((r) => {
    const key = makeKey(r.지점명, r.날짜, r.품목);

    if (!map.has(key)) {
      map.set(key, {
        지점명: r.지점명,
        날짜: r.날짜,
        품목: r.품목,
        판매수량: 0,
        판매금액: 0,
        최종폐기: 0,
        폐기금액: 0,
      });
    }

    const row = map.get(key);
    row.판매수량 += toNumber(r.판매수량);
    row.판매금액 += toNumber(r.판매금액);
    row.최종폐기 += toNumber(r.최종폐기);
    row.폐기금액 += toNumber(r.폐기금액);
  });

  return Array.from(map.values());
}

function buildDeltaRows(currentRows, lockedRows) {
  if (!lockedRows || !lockedRows.length) return currentRows;

  const lockedMap = buildLockedMap(lockedRows);
  const currentDailyRows = buildCurrentDailyRows(currentRows);
  const deltaRows = [];

  currentDailyRows.forEach((r) => {
    const key = makeKey(r.지점명, r.날짜, r.품목);
    const locked = lockedMap.get(key) || {};

    const deltaSaleQty = Math.max(0, toNumber(r.판매수량) - toNumber(locked.판매수량));
    const deltaSaleAmt = Math.max(0, toNumber(r.판매금액) - toNumber(locked.판매금액));
    const deltaDiscardQty = Math.max(0, toNumber(r.최종폐기) - toNumber(locked.폐기수량));
    const deltaDiscardAmt = Math.max(0, toNumber(r.폐기금액) - toNumber(locked.폐기금액));

    if (deltaSaleQty === 0 && deltaSaleAmt === 0 && deltaDiscardQty === 0 && deltaDiscardAmt === 0) return;

    deltaRows.push({
      지점명: r.지점명,
      날짜: r.날짜,
      품목: r.품목,
      판매수량: deltaSaleQty,
      판매금액: deltaSaleAmt,
      최종폐기: deltaDiscardQty,
      폐기금액: deltaDiscardAmt,
      isDeltaOnly: true,
    });
  });

  return deltaRows;
}

function allocateIntegerByCapacity(total, entries, capacityField, resultField, dateLimit) {
  let target = Math.round(toNumber(total));

  const eligible = entries.filter((e) => {
    if (dateLimit && normalizeText(e.date) > dateLimit) return false;
    if (e.isDeltaOnly) return false;
    return toNumber(e[capacityField]) > 0;
  });

  const totalCap = eligible.reduce((sum, e) => sum + Math.max(0, Math.round(toNumber(e[capacityField]))), 0);
  target = Math.min(target, totalCap);

  if (target <= 0 || totalCap <= 0) return 0;

  const temp = eligible.map((e) => {
    const cap = Math.max(0, Math.round(toNumber(e[capacityField])));
    const exact = (target * cap) / totalCap;
    const base = Math.min(cap, Math.floor(exact));
    return { e, cap, exact, base, frac: exact - base };
  });

  let used = temp.reduce((sum, x) => sum + x.base, 0);
  let remain = target - used;

  temp.sort((a, b) => b.frac - a.frac);

  temp.forEach((x) => {
    let add = x.base;
    if (remain > 0 && add < x.cap) {
      add += 1;
      remain -= 1;
    }

    x.e[resultField] = (x.e[resultField] || 0) + add;
  });

  return target;
}

function distributeOneStockAmount(entries, yearKey, stockQty, stockAmt) {
  const qtyField = `${yearKey}Qty`;
  const amtField = `${yearKey}Amt`;
  const exactField = `${yearKey}AmtExact`;

  const targets = entries.filter((e) => toNumber(e[qtyField]) > 0);

  targets.forEach((e) => {
    e[amtField] = 0;
    e[exactField] = 0;
  });

  if (!targets.length || stockQty <= 0 || stockAmt <= 0) return;

  const usedQty = targets.reduce((sum, e) => sum + toNumber(e[qtyField]), 0);

  let targetAmt =
    usedQty >= stockQty
      ? Math.round(stockAmt)
      : Math.round((stockAmt * usedQty) / stockQty);

  const totalAvailableAmt = targets.reduce((sum, e) => {
    const alreadyAllocated = toNumber(e.prev2Amt) + toNumber(e.prevAmt) + toNumber(e.currentAmt);
    return sum + Math.max(0, Math.round(toNumber(e.amt)) - alreadyAllocated);
  }, 0);

  targetAmt = Math.min(targetAmt, totalAvailableAmt);
  if (targetAmt <= 0) return;

  targets.forEach((e) => {
    const alreadyAllocated = toNumber(e.prev2Amt) + toNumber(e.prevAmt) + toNumber(e.currentAmt);
    const availableAmt = Math.max(0, Math.round(toNumber(e.amt)) - alreadyAllocated);
    const exact = (targetAmt * toNumber(e[qtyField])) / usedQty;

    e[exactField] = exact;
    e[amtField] = Math.min(availableAmt, Math.floor(exact));
  });

  let assigned = targets.reduce((sum, e) => sum + toNumber(e[amtField]), 0);
  let remain = targetAmt - assigned;

  targets.sort((a, b) => {
    const fa = a[exactField] - Math.floor(a[exactField]);
    const fb = b[exactField] - Math.floor(b[exactField]);
    return fb - fa;
  });

  let guard = 0;
  while (remain > 0 && targets.length && guard < targets.length * 20) {
    const e = targets[guard % targets.length];

    const alreadyAllocated = toNumber(e.prev2Amt) + toNumber(e.prevAmt) + toNumber(e.currentAmt);
    const availableAmt = Math.max(0, Math.round(toNumber(e.amt)) - alreadyAllocated);

    if (availableAmt > 0) {
      e[amtField] += 1;
      remain--;
    }

    guard++;
  }
}

function allocateGroupPeriod(entries, stock) {
  const prev2Cutoff = getPrev2CutoffDate();
  const prevCutoff = getPrevCutoffDate();

  entries.forEach((e) => {
    e.remainQty = Math.round(toNumber(e.qty));

    e.prev2Qty = 0;
    e.prevQty = 0;
    e.currentQty = 0;
    e.shortageQty = 0;

    e.prev2Amt = 0;
    e.prevAmt = 0;
    e.currentAmt = 0;
    e.shortageAmt = 0;
  });

  allocateIntegerByCapacity(
    toNumber(stock?.전전년수량),
    entries,
    "remainQty",
    "prev2Qty",
    prev2Cutoff
  );

  entries.forEach((e) => {
    e.remainQty -= e.prev2Qty;
  });

  allocateIntegerByCapacity(
    toNumber(stock?.전년수량),
    entries,
    "remainQty",
    "prevQty",
    prevCutoff
  );

  entries.forEach((e) => {
    e.remainQty -= e.prevQty;
  });

  // 남은 수량은 부족이 아니라 당해 사용
  entries.forEach((e) => {
    e.currentQty = Math.max(0, e.remainQty);
    e.remainQty = 0;
    e.shortageQty = 0;
  });

  // 핵심 수정: 한 행 안에서 수량 있는 연차에만 금액 배분
  distributeRowAmountsWithoutMismatch(entries);
}

function distributeRowAmountsWithoutMismatch(entries) {
  entries.forEach((e) => {
    const totalAmt = Math.round(toNumber(e.amt));

    e.prev2Amt = 0;
    e.prevAmt = 0;
    e.currentAmt = 0;
    e.shortageAmt = 0;

    const parts = [
      { key: "prev2", qty: toNumber(e.prev2Qty), amtField: "prev2Amt" },
      { key: "prev", qty: toNumber(e.prevQty), amtField: "prevAmt" },
      { key: "current", qty: toNumber(e.currentQty), amtField: "currentAmt" },
    ].filter((p) => p.qty > 0);

    if (!parts.length || totalAmt <= 0) return;

    if (parts.length === 1) {
      e[parts[0].amtField] = totalAmt;
      return;
    }

    const totalQty = parts.reduce((sum, p) => sum + p.qty, 0);

    let assigned = 0;

    parts.forEach((p) => {
      const exact = (totalAmt * p.qty) / totalQty;
      p.exact = exact;
      p.base = Math.floor(exact);
      e[p.amtField] = p.base;
      assigned += p.base;
    });

    let remain = totalAmt - assigned;

    parts.sort((a, b) => {
      const fa = a.exact - Math.floor(a.exact);
      const fb = b.exact - Math.floor(b.exact);
      return fb - fa;
    });

    let i = 0;
    while (remain > 0) {
      const p = parts[i % parts.length];
      e[p.amtField] += 1;
      remain--;
      i++;
    }

    // 최종 안전장치:
    // 수량이 있는데 금액이 0인 경우, 다른 연차에서 1원 이동
    parts.forEach((p) => {
      if (p.qty > 0 && e[p.amtField] === 0 && totalAmt >= parts.length) {
        const donor = parts.find((x) => e[x.amtField] > 1);
        if (donor) {
          e[donor.amtField] -= 1;
          e[p.amtField] += 1;
        }
      }
    });
  });
}

  distributeOneStockAmount(entries, "prev2", toNumber(stock?.전전년수량), toNumber(stock?.전전년금액));
  distributeOneStockAmount(entries, "prev", toNumber(stock?.전년수량), toNumber(stock?.전년금액));

  entries.forEach((e) => {
    const totalAmt = Math.round(toNumber(e.amt));
    const allocatedAmt = toNumber(e.prev2Amt) + toNumber(e.prevAmt);

    e.currentAmt = Math.max(0, totalAmt - allocatedAmt);
    e.shortageAmt = 0;
  });
}

function createDailyCombinedRow(branch, date, item) {
  return {
    지점명: branch,
    일자: date,
    품목군: item,

    판매수량: 0,
    판매금액: 0,
    폐기수량: 0,
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

    부족수량: 0,
    부족금액: 0,
  };
}

function mergeLockedRowsIntoMergedMap(mergedMap, lockedRows) {
  lockedRows.forEach((r) => {
    const key = makeKey(r.지점명, r.일자, r.품목군);

    if (!mergedMap.has(key)) {
      mergedMap.set(key, createDailyCombinedRow(r.지점명, r.일자, r.품목군));
    }

    const row = mergedMap.get(key);

    row.판매수량 += toNumber(r.판매수량);
    row.판매금액 += toNumber(r.판매금액);
    row.폐기수량 += toNumber(r.폐기수량);
    row.폐기금액 += toNumber(r.폐기금액);

    row.전전년_판매수량 += toNumber(r.전전년_판매수량);
    row.전전년_판매금액 += toNumber(r.전전년_판매금액);
    row.전전년_폐기수량 += toNumber(r.전전년_폐기수량);
    row.전전년_폐기금액 += toNumber(r.전전년_폐기금액);

    row.전년_판매수량 += toNumber(r.전년_판매수량);
    row.전년_판매금액 += toNumber(r.전년_판매금액);
    row.전년_폐기수량 += toNumber(r.전년_폐기수량);
    row.전년_폐기금액 += toNumber(r.전년_폐기금액);

    row.당해_판매수량 += toNumber(r.당해_판매수량);
    row.당해_판매금액 += toNumber(r.당해_판매금액);
    row.당해_폐기수량 += toNumber(r.당해_폐기수량);
    row.당해_폐기금액 += toNumber(r.당해_폐기금액);
  });
}

function buildStockBalanceRows(stockMap, groupMap, lockedRows) {
  const lockedUseMap = new Map();

  lockedRows.forEach((r) => {
    const key = `${normalizeText(r.지점명)}__${normalizeText(r.품목군)}`;

    if (!lockedUseMap.has(key)) {
      lockedUseMap.set(key, {
        prev2Qty: 0,
        prev2Amt: 0,
        prevQty: 0,
        prevAmt: 0,
      });
    }

    const x = lockedUseMap.get(key);
    x.prev2Qty += toNumber(r.전전년_판매수량) + toNumber(r.전전년_폐기수량);
    x.prev2Amt += toNumber(r.전전년_판매금액) + toNumber(r.전전년_폐기금액);
    x.prevQty += toNumber(r.전년_판매수량) + toNumber(r.전년_폐기수량);
    x.prevAmt += toNumber(r.전년_판매금액) + toNumber(r.전년_폐기금액);
  });

  const rows = [];

  for (const [key, stock] of stockMap.entries()) {
    const entries = groupMap.get(key) || [];
    const locked = lockedUseMap.get(key) || {};

    const prev2UsedQty = toNumber(locked.prev2Qty) + entries.reduce((s, e) => s + toNumber(e.prev2Qty), 0);
    const prev2UsedAmt = toNumber(locked.prev2Amt) + entries.reduce((s, e) => s + toNumber(e.prev2Amt), 0);

    const prevUsedQty = toNumber(locked.prevQty) + entries.reduce((s, e) => s + toNumber(e.prevQty), 0);
    const prevUsedAmt = toNumber(locked.prevAmt) + entries.reduce((s, e) => s + toNumber(e.prevAmt), 0);

    rows.push({
      지점명: stock.지점명,
      품목: stock.품목,

      전전년소진기준일: getPrev2CutoffDate(),
      전전년재고수량: toNumber(stock.전전년수량),
      전전년사용수량: prev2UsedQty,
      전전년잔여수량: toNumber(stock.전전년수량) - prev2UsedQty,
      전전년재고금액: Math.round(toNumber(stock.전전년금액)),
      전전년사용금액: Math.round(prev2UsedAmt),
      전전년잔여금액: Math.round(toNumber(stock.전전년금액) - prev2UsedAmt),

      전년소진기준일: getPrevCutoffDate(),
      전년재고수량: toNumber(stock.전년수량),
      전년사용수량: prevUsedQty,
      전년잔여수량: toNumber(stock.전년수량) - prevUsedQty,
      전년재고금액: Math.round(toNumber(stock.전년금액)),
      전년사용금액: Math.round(prevUsedAmt),
      전년잔여금액: Math.round(toNumber(stock.전년금액) - prevUsedAmt),
    });
  }

  return sortRowsByBusinessOrder(rows);
}

function buildOnlyRemainStockRows(stockBalanceRows) {
  return stockBalanceRows.filter((r) => {
    return (
      toNumber(r.전전년잔여수량) !== 0 ||
      toNumber(r.전전년잔여금액) !== 0 ||
      toNumber(r.전년잔여수량) !== 0 ||
      toNumber(r.전년잔여금액) !== 0
    );
  });
}

function processWorkbook(workbook) {
  const rawMainRows = parseHorizontal2025Sheet(workbook);
  const filteredMainRows = filterRowsByFinalCutoff(rawMainRows);

  const lockedRows = parseLockedDailyRows(lockedWorkbook);
  const mainRows = buildDeltaRows(filteredMainRows, lockedRows);

  const prev2Rows = parseInventorySheetFixed(workbook, PREV2_SHEET);
  const prevRows = parseInventorySheetFixed(workbook, PREV_SHEET);

  const anomalyIssues = validateAnomalies(mainRows);
  const stockMap = buildOpeningStocks(prev2Rows, prevRows);

  const groupMap = new Map();

  mainRows.forEach((row) => {
    const branch = normalizeText(row.지점명);
    const item = normalizeText(row.품목);
    const key = `${branch}__${item}`;

    if (!groupMap.has(key)) groupMap.set(key, []);

    const date = normalizeText(row.날짜);

    const saleQty = toNumber(row.판매수량);
    const saleAmt = toNumber(row.판매금액);

    if (saleQty > 0 && saleAmt > 0) {
  groupMap.get(key).push({
    type: "판매",
    branch,
    item,
    date,
    qty: saleQty,
    amt: saleAmt,
    isDeltaOnly: !!row.isDeltaOnly,
  });
}

    const discardQty = toNumber(row.최종폐기);
    const discardAmt = toNumber(row.폐기금액);

    if (discardQty > 0 && discardAmt > 0) {
  groupMap.get(key).push({
    type: "폐기",
    branch,
    item,
    date,
    qty: discardQty,
    amt: discardAmt,
    isDeltaOnly: !!row.isDeltaOnly,
  });
}

  for (const [key, entries] of groupMap.entries()) {
    entries.sort((a, b) => {
      if (a.date !== b.date) return a.date.localeCompare(b.date);
      return a.type.localeCompare(b.type);
    });

    allocateGroupPeriod(entries, stockMap.get(key));
  }

  const mergedMap = new Map();
  mergeLockedRowsIntoMergedMap(mergedMap, lockedRows);

  const salesRows = [];
  const discardRows = [];
  const validationRows = [];

  for (const entries of groupMap.values()) {
    entries.forEach((e) => {
      const dailyKey = makeKey(e.branch, e.date, e.item);

      if (!mergedMap.has(dailyKey)) {
        mergedMap.set(dailyKey, createDailyCombinedRow(e.branch, e.date, e.item));
      }

      const row = mergedMap.get(dailyKey);

      if (e.type === "판매") {
        row.판매수량 += e.qty;
        row.판매금액 += e.amt;

        row.전전년_판매수량 += e.prev2Qty;
        row.전전년_판매금액 += e.prev2Amt;

        row.전년_판매수량 += e.prevQty;
        row.전년_판매금액 += e.prevAmt;

        row.당해_판매수량 += e.currentQty;
        row.당해_판매금액 += e.currentAmt;

        salesRows.push({
          구분: e.isDeltaOnly ? "판매_증가분" : "판매",
          지점명: e.branch,
          날짜: e.date,
          품목: e.item,
          총사용수량: e.qty,
          총사용금액: e.amt,
          전전년사용수량: e.prev2Qty,
          전전년사용금액: e.prev2Amt,
          전년사용수량: e.prevQty,
          전년사용금액: e.prevAmt,
          당해사용수량: e.currentQty,
          당해사용금액: e.currentAmt,
          부족수량: e.shortageQty,
          부족금액: e.shortageAmt,
        });
      } else {
        row.폐기수량 += e.qty;
        row.폐기금액 += e.amt;

        row.전전년_폐기수량 += e.prev2Qty;
        row.전전년_폐기금액 += e.prev2Amt;

        row.전년_폐기수량 += e.prevQty;
        row.전년_폐기금액 += e.prevAmt;

        row.당해_폐기수량 += e.currentQty;
        row.당해_폐기금액 += e.currentAmt;

        discardRows.push({
          구분: e.isDeltaOnly ? "폐기_증가분" : "폐기",
          지점명: e.branch,
          날짜: e.date,
          품목: e.item,
          총사용수량: e.qty,
          총사용금액: e.amt,
          전전년사용수량: e.prev2Qty,
          전전년사용금액: e.prev2Amt,
          전년사용수량: e.prevQty,
          전년사용금액: e.prevAmt,
          당해사용수량: e.currentQty,
          당해사용금액: e.currentAmt,
          부족수량: e.shortageQty,
          부족금액: e.shortageAmt,
        });
      }
    });
  }

  const mergedDailyRows = sortRowsByBusinessOrder(Array.from(mergedMap.values())).map((row) => {
    row.총사용수량 = row.판매수량 + row.폐기수량;
    row.총사용금액 = row.판매금액 + row.폐기금액;

    const periodQtySum =
      row.전전년_판매수량 +
      row.전전년_폐기수량 +
      row.전년_판매수량 +
      row.전년_폐기수량 +
      row.당해_판매수량 +
      row.당해_폐기수량;

    const periodAmtSum =
      row.전전년_판매금액 +
      row.전전년_폐기금액 +
      row.전년_판매금액 +
      row.전년_폐기금액 +
      row.당해_판매금액 +
      row.당해_폐기금액;

    validationRows.push({
      지점명: row.지점명,
      날짜: row.일자,
      품목: row.품목군,
      총사용수량: row.총사용수량,
      총사용금액: row.총사용금액,
      연차배분수량합: periodQtySum,
      수량차이: row.총사용수량 - periodQtySum,
      연차배분금액합: periodAmtSum,
      금액차이: row.총사용금액 - periodAmtSum,
      부족수량: row.부족수량,
      부족금액: row.부족금액,
    });

    return row;
  });

  const stockBalanceRows = buildStockBalanceRows(stockMap, groupMap, lockedRows);
  const remainStockRows = buildOnlyRemainStockRows(stockBalanceRows);

  return {
    ok: true,
    schemaErrors: [],
    anomalyIssues,
    mergedDailyRows,
    salesRows: sortRowsByBusinessOrder(salesRows),
    discardRows: sortRowsByBusinessOrder(discardRows),
    validationRows: sortRowsByBusinessOrder(validationRows),
    stockBalanceRows,
    remainStockRows,
  };
}

function renderIssues(result) {
  const container = document.getElementById("issues");
  if (!container) return;

  container.innerHTML = "";

  const issues = [
    ...result.schemaErrors.map((m) => ({ message: m })),
    ...result.anomalyIssues,
  ];

  if (!issues.length) {
    container.innerHTML = `<div class="issue ok">오류 및 이상치가 없습니다.</div>`;
    return;
  }

  issues.slice(0, 300).forEach((issue) => {
    const div = document.createElement("div");
    div.className = "issue";
    div.textContent = issue.message;
    container.appendChild(div);
  });
}

function createTable(rows, maxRows = 300) {
  if (!rows || !rows.length) return "<p>데이터가 없습니다.</p>";

  const limitedRows = rows.slice(0, maxRows);
  const cols = Object.keys(limitedRows[0]);

  let html = "";

  if (rows.length > maxRows) {
    html += `<p style="margin:0 0 12px; color:#6b7280; font-size:13px;">총 ${rows.length.toLocaleString()}행 중 ${maxRows.toLocaleString()}행만 화면에 표시합니다. 전체 데이터는 엑셀 다운로드에서 확인하세요.</p>`;
  }

  html += "<table><thead><tr>";
  cols.forEach((c) => (html += `<th>${c}</th>`));
  html += "</tr></thead><tbody>";

  limitedRows.forEach((row) => {
    html += "<tr>";
    cols.forEach((c) => (html += `<td>${row[c] ?? ""}</td>`));
    html += "</tr>";
  });

  html += "</tbody></table>";
  return html;
}

function populateBranchFilter(result) {
  const select = document.getElementById("branchFilter");
  if (!select) return;

  const branches = new Set();

  result.mergedDailyRows.forEach((r) => {
    if (r.지점명) branches.add(r.지점명);
  });

  result.stockBalanceRows.forEach((r) => {
    if (r.지점명) branches.add(r.지점명);
  });

  const current = select.value || "전체";
  select.innerHTML = `<option value="전체">전체</option>`;

  Array.from(branches)
    .sort()
    .forEach((branch) => {
      const option = document.createElement("option");
      option.value = branch;
      option.textContent = branch;
      select.appendChild(option);
    });

  select.value = branches.has(current) ? current : "전체";
}

function renderTables(result) {
  const selectedBranch = document.getElementById("branchFilter")?.value || "전체";

  const filter = (rows) =>
    selectedBranch === "전체" ? rows : rows.filter((row) => row.지점명 === selectedBranch);

  const mergedTable = document.getElementById("mergedTable");
  if (mergedTable) mergedTable.innerHTML = createTable(filter(result.mergedDailyRows), 300);

  const salesTable = document.getElementById("salesTable");
  if (salesTable) salesTable.innerHTML = createTable(filter(result.salesRows), 150);

  const discardTable = document.getElementById("discardTable");
  if (discardTable) discardTable.innerHTML = createTable(filter(result.discardRows), 150);

  const validationTable = document.getElementById("validationTable");
  if (validationTable) validationTable.innerHTML = createTable(filter(result.validationRows), 150);

  const stockBalanceTable = document.getElementById("stockBalanceTable");
  if (stockBalanceTable) stockBalanceTable.innerHTML = createTable(filter(result.stockBalanceRows), 300);

  const remainStockTable = document.getElementById("remainStockTable");
  if (remainStockTable) remainStockTable.innerHTML = createTable(filter(result.remainStockRows), 300);
}

function updateStats(result) {
  const salesCount = document.getElementById("salesCount");
  if (salesCount) salesCount.textContent = result.salesRows.length;

  const discardCount = document.getElementById("discardCount");
  if (discardCount) discardCount.textContent = result.discardRows.length;

  const issueCount = document.getElementById("issueCount");
  if (issueCount) {
    issueCount.textContent = result.schemaErrors.length + result.anomalyIssues.length;
  }

  const shortages = result.mergedDailyRows.filter(
    (r) => toNumber(r.부족수량) > 0 || toNumber(r.부족금액) > 0
  ).length;

  const shortageCount = document.getElementById("shortageCount");
  if (shortageCount) shortageCount.textContent = shortages;

  const stockBalanceCount = document.getElementById("stockBalanceCount");
  if (stockBalanceCount) stockBalanceCount.textContent = result.remainStockRows.length;
}

function downloadWorkbook(result) {
  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.mergedDailyRows), "일별통합결과");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.salesRows), "판매자동소진");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.discardRows), "폐기자동소진");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.validationRows), "검증");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.stockBalanceRows), "연차별재고소진검증");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.remainStockRows), "잔여재고만");

  const issueRows = [
    ...result.schemaErrors.map((m) => ({ 유형: "스키마오류", 내용: m })),
    ...result.anomalyIssues.map((x) => ({ 유형: x.type, 내용: x.message })),
  ];

  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(issueRows), "오류및이상치");

  XLSX.writeFile(wb, "FIFO_연차별재고소진_감사용.xlsx");
}

function runProcess() {
  if (!latestWorkbook) {
    alert("먼저 원본 엑셀 파일을 선택해주세요.");
    return;
  }

  try {
    const result = processWorkbook(latestWorkbook);
    latestResult = result;

    updateStats(result);
    renderIssues(result);
    populateBranchFilter(result);
    renderTables(result);

    const downloadBtn = document.getElementById("downloadBtn");
    if (downloadBtn) downloadBtn.disabled = false;
  } catch (error) {
    alert("실행 중 오류가 발생했습니다: " + error.message);
  }
}

document.addEventListener("DOMContentLoaded", () => {
  setDefaultTodayDates();

  const runBtn = document.getElementById("runBtn");
  if (runBtn) runBtn.disabled = false;

  document.getElementById("fileInput")?.addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
      const data = await file.arrayBuffer();
      latestWorkbook = XLSX.read(data, { type: "array" });
    } catch (error) {
      alert("원본 파일 처리 중 오류가 발생했습니다: " + error.message);
    }
  });

  document.getElementById("lockedResultInput")?.addEventListener("change", async (e) => {
    const file = e.target.files[0];

    if (!file) {
      lockedWorkbook = null;
      return;
    }

    try {
      const data = await file.arrayBuffer();
      lockedWorkbook = XLSX.read(data, { type: "array" });
    } catch (error) {
      alert("기존 제출 결과 파일 처리 중 오류가 발생했습니다: " + error.message);
    }
  });

  document.getElementById("runBtn")?.addEventListener("click", () => {
    runProcess();
  });

  document.getElementById("downloadBtn")?.addEventListener("click", () => {
    if (!latestResult) {
      alert("먼저 실행 버튼을 눌러 계산해주세요.");
      return;
    }

    downloadWorkbook(latestResult);
  });

  document.getElementById("branchFilter")?.addEventListener("change", () => {
    if (!latestResult) return;
    renderTables(latestResult);
  });
});
