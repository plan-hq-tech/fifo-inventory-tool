/*************************************************
 * FIFO 연차별 재고 소진 프로그램
 * - 원본 시트: 2025
 * - 재고 시트: 전전년재고_DB, 전년재고_DB
 * - 기존 제출 결과 파일 선택 가능
 * - 기준일 3개: 전전년 / 전년 / 최종
 * - 실행 버튼 방식
 *************************************************/

const MAIN_SHEET = "2025";
const PREV2_SHEET = "전전년재고_DB";
const PREV_SHEET = "전년재고_DB";

const ITEM_ORDER = ["의류", "잡화", "생활", "문화", "건강미용", "식품", "기증파트너"];

let latestWorkbook = null;
let lockedWorkbook = null;
let latestResult = null;

/*************************************************
 * 기본 유틸
 *************************************************/

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

function todayString() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
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

function getDateInput(id) {
  const el = document.getElementById(id);
  return el && el.value ? el.value : null;
}

function getPrev2CutoffDate() {
  return getDateInput("prev2CutoffDate") || getDateInput("cutoffDate") || todayString();
}

function getPrevCutoffDate() {
  return getDateInput("prevCutoffDate") || getDateInput("cutoffDate") || todayString();
}

function getFinalCutoffDate() {
  return getDateInput("finalCutoffDate") || getDateInput("cutoffDate") || todayString();
}

function setDefaultTodayDates() {
  const today = todayString();

  ["prev2CutoffDate", "prevCutoffDate", "finalCutoffDate"].forEach((id) => {
    const el = document.getElementById(id);
    if (el && !el.value) el.value = today;
  });
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

function makeDailyKey(branch, date, item) {
  return `${normalizeText(branch)}__${normalizeText(date)}__${normalizeText(item)}`;
}

function makeStockKey(branch, item) {
  return `${normalizeText(branch)}__${normalizeText(item)}`;
}

function itemOrderIndex(item) {
  const idx = ITEM_ORDER.indexOf(normalizeText(item));
  return idx === -1 ? 999 : idx;
}

function sortRowsByBusinessOrder(rows) {
  return [...rows].sort((a, b) => {
    const da = normalizeText(a.일자 || a.날짜 || "");
    const db = normalizeText(b.일자 || b.날짜 || "");
    if (da !== db) return da.localeCompare(db);

    const ba = normalizeText(a.지점명);
    const bb = normalizeText(b.지점명);
    if (ba !== bb) return ba.localeCompare(bb);

    const ia = itemOrderIndex(a.품목군 || a.품목);
    const ib = itemOrderIndex(b.품목군 || b.품목);
    if (ia !== ib) return ia - ib;

    return normalizeText(a.품목군 || a.품목).localeCompare(
      normalizeText(b.품목군 || b.품목)
    );
  });
}

/*************************************************
 * 2025 가로형 시트 파싱
 *************************************************/

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

/*************************************************
 * 재고 시트 파싱
 * A 지점명 / B 품목 / C 미사용 / D 수량 / E 금액
 *************************************************/

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

function buildOpeningStocks(prev2Rows, prevRows) {
  const map = new Map();

  function ensure(branch, item) {
    const key = makeStockKey(branch, item);

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
  }

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

/*************************************************
 * 기존 제출 결과 파싱
 *************************************************/

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

function parseLockedDetailRows(workbook, sheetName, type) {
  if (!workbook) return [];

  const ws = workbook.Sheets[sheetName];
  if (!ws || !ws["!ref"]) return [];

  return XLSX.utils.sheet_to_json(ws, { defval: "" }).map((r) => ({
    구분: normalizeText(r.구분) || `${type}_기존제출`,
    지점명: normalizeText(r.지점명),
    날짜: formatDate(r.날짜),
    품목: normalizeText(r.품목),

    총사용수량: toNumber(r.총사용수량),
    총사용금액: toNumber(r.총사용금액),

    전전년사용수량: toNumber(r.전전년사용수량),
    전전년사용금액: toNumber(r.전전년사용금액),

    전년사용수량: toNumber(r.전년사용수량),
    전년사용금액: toNumber(r.전년사용금액),

    당해사용수량: toNumber(r.당해사용수량),
    당해사용금액: toNumber(r.당해사용금액),

    부족수량: 0,
    부족금액: 0,
  }));
}

function buildLockedDetailDailyMap(workbook) {
  const map = new Map();

  const lockedSalesRows = parseLockedDetailRows(workbook, "판매자동소진", "판매");
  const lockedDiscardRows = parseLockedDetailRows(workbook, "폐기자동소진", "폐기");

  function ensure(branch, date, item) {
    const key = makeDailyKey(branch, date, item);

    if (!map.has(key)) {
      map.set(key, createDailyCombinedRow(branch, date, item));
    }

    return map.get(key);
  }

  lockedSalesRows.forEach((r) => {
    const row = ensure(r.지점명, r.날짜, r.품목);

    row.판매수량 += toNumber(r.총사용수량);
    row.판매금액 += toNumber(r.총사용금액);

    row.전전년_판매수량 += toNumber(r.전전년사용수량);
    row.전전년_판매금액 += toNumber(r.전전년사용금액);

    row.전년_판매수량 += toNumber(r.전년사용수량);
    row.전년_판매금액 += toNumber(r.전년사용금액);

    row.당해_판매수량 += toNumber(r.당해사용수량);
    row.당해_판매금액 += toNumber(r.당해사용금액);
  });

  lockedDiscardRows.forEach((r) => {
    const row = ensure(r.지점명, r.날짜, r.품목);

    row.폐기수량 += toNumber(r.총사용수량);
    row.폐기금액 += toNumber(r.총사용금액);

    row.전전년_폐기수량 += toNumber(r.전전년사용수량);
    row.전전년_폐기금액 += toNumber(r.전전년사용금액);

    row.전년_폐기수량 += toNumber(r.전년사용수량);
    row.전년_폐기금액 += toNumber(r.전년사용금액);

    row.당해_폐기수량 += toNumber(r.당해사용수량);
    row.당해_폐기금액 += toNumber(r.당해사용금액);
  });

  for (const row of map.values()) {
    row.총사용수량 = row.판매수량 + row.폐기수량;
    row.총사용금액 = row.판매금액 + row.폐기금액;
  }

  return map;
}

function buildLockedMap(lockedRows) {
  const map = new Map();

  lockedRows.forEach((r) => {
    const key = makeDailyKey(r.지점명, r.일자, r.품목군);

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
    const key = makeDailyKey(r.지점명, r.날짜, r.품목);

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
    const key = makeDailyKey(r.지점명, r.날짜, r.품목);
    const locked = lockedMap.get(key) || {};

    const deltaSaleQty = Math.max(0, toNumber(r.판매수량) - toNumber(locked.판매수량));
    const deltaSaleAmt = Math.max(0, toNumber(r.판매금액) - toNumber(locked.판매금액));

    const deltaDiscardQty = Math.max(0, toNumber(r.최종폐기) - toNumber(locked.폐기수량));
    const deltaDiscardAmt = Math.max(0, toNumber(r.폐기금액) - toNumber(locked.폐기금액));

    if (
      deltaSaleQty === 0 &&
      deltaSaleAmt === 0 &&
      deltaDiscardQty === 0 &&
      deltaDiscardAmt === 0
    ) {
      return;
    }

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

/*************************************************
 * 기준일 필터 및 이상치
 *************************************************/

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
      issues.push({
        type: "이상치",
        message: `행 ${rowNo}: 판매수량은 있는데 판매금액이 0입니다. 계산에서는 제외됩니다.`,
      });
    }

    if (saleQty === 0 && saleAmt > 0) {
      issues.push({
        type: "이상치",
        message: `행 ${rowNo}: 판매금액은 있는데 판매수량이 0입니다. 계산에서는 제외됩니다.`,
      });
    }

    if (discardQty > 0 && discardAmt === 0) {
      issues.push({
        type: "이상치",
        message: `행 ${rowNo}: 폐기수량은 있는데 폐기금액이 0입니다. 계산에서는 제외됩니다.`,
      });
    }

    if (discardQty === 0 && discardAmt > 0) {
      issues.push({
        type: "이상치",
        message: `행 ${rowNo}: 폐기금액은 있는데 폐기수량이 0입니다. 계산에서는 제외됩니다.`,
      });
    }
  });

  return issues;
}

/*************************************************
 * FIFO 수량 배분
 *************************************************/

function allocateIntegerByCapacity(total, entries, capacityField, resultField, dateLimit) {
  let target = Math.round(toNumber(total));

  const eligible = entries.filter((e) => {
    if (dateLimit && normalizeText(e.date) > dateLimit) return false;
    return toNumber(e[capacityField]) > 0;
  });

  const totalCap = eligible.reduce(
    (sum, e) => sum + Math.max(0, Math.round(toNumber(e[capacityField]))),
    0
  );

  target = Math.min(target, totalCap);

  if (target <= 0 || totalCap <= 0) return 0;

  const temp = eligible.map((e) => {
    const cap = Math.max(0, Math.round(toNumber(e[capacityField])));
    const exact = (target * cap) / totalCap;
    const base = Math.min(cap, Math.floor(exact));

    return {
      e,
      cap,
      exact,
      base,
      frac: exact - base,
    };
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

/*************************************************
 * 금액 배분
 * 중요:
 * - 행 단위 총금액 안에서만 배분
 * - 수량 있는 연차에는 금액도 있어야 함
 * - 금액 있는 연차에는 수량도 있어야 함
 * - 당해 금액 마이너스 금지
 *************************************************/

function distributeRowAmountByAllocatedQty(entries) {
  entries.forEach((e) => {
    const totalAmt = Math.round(toNumber(e.amt));

    e.prev2Amt = 0;
    e.prevAmt = 0;
    e.currentAmt = 0;
    e.shortageAmt = 0;

    const parts = [
      { name: "전전년", qtyField: "prev2Qty", amtField: "prev2Amt", qty: toNumber(e.prev2Qty) },
      { name: "전년", qtyField: "prevQty", amtField: "prevAmt", qty: toNumber(e.prevQty) },
      { name: "당해", qtyField: "currentQty", amtField: "currentAmt", qty: toNumber(e.currentQty) },
    ].filter((p) => p.qty > 0);

    if (!parts.length) return;

    if (totalAmt <= 0) {
      e.shortageAmt = 0;
      return;
    }

    if (parts.length === 1) {
      e[parts[0].amtField] = totalAmt;
      return;
    }

    const totalQty = parts.reduce((sum, p) => sum + p.qty, 0);

    // 금액이 연차 개수보다 적으면 모든 연차에 1원 이상 줄 수 없음.
    // 이 경우 최소 불일치를 피하기 위해 수량이 가장 큰 연차부터 금액을 배분.
    if (totalAmt < parts.length) {
      parts.sort((a, b) => b.qty - a.qty);
      let remainSmall = totalAmt;

      parts.forEach((p) => {
        if (remainSmall > 0) {
          e[p.amtField] = 1;
          remainSmall -= 1;
        }
      });

      return;
    }

    // 1차: 수량 있는 연차에 최소 1원씩 보장
    parts.forEach((p) => {
      e[p.amtField] = 1;
    });

    let remain = totalAmt - parts.length;

    const temp = parts.map((p) => {
      const exact = (remain * p.qty) / totalQty;
      const base = Math.floor(exact);

      return {
        ...p,
        exact,
        base,
        frac: exact - base,
      };
    });

    let assigned = 0;

    temp.forEach((p) => {
      e[p.amtField] += p.base;
      assigned += p.base;
    });

    let leftover = remain - assigned;

    temp.sort((a, b) => b.frac - a.frac);

    let i = 0;
    while (leftover > 0 && temp.length) {
      const p = temp[i % temp.length];
      e[p.amtField] += 1;
      leftover -= 1;
      i += 1;
    }

    // 최종 안전장치: 총금액과 배분금액 일치 보정
    const allocated = toNumber(e.prev2Amt) + toNumber(e.prevAmt) + toNumber(e.currentAmt);
    const diff = totalAmt - allocated;

    if (diff !== 0) {
      const receiver =
        parts.find((p) => p.amtField === "currentAmt") ||
        parts.find((p) => p.amtField === "prevAmt") ||
        parts[0];

      e[receiver.amtField] += diff;
    }

    // 최종 안전장치: 음수 방지
    ["prev2Amt", "prevAmt", "currentAmt"].forEach((field) => {
      if (e[field] < 0) e[field] = 0;
    });
  });
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

  // 1. 전전년 수량 배분
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

  // 2. 전년 수량 배분
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

  // 3. 남은 수량은 부족이 아니라 당해 사용
  entries.forEach((e) => {
    e.currentQty = Math.max(0, e.remainQty);
    e.remainQty = 0;
    e.shortageQty = 0;
  });

  // 4. 금액은 행 단위 총금액 안에서 수량 비율로 배분
  distributeRowAmountByAllocatedQty(entries);
}

/*************************************************
 * 결과 행 생성
 *************************************************/

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
    const key = makeDailyKey(r.지점명, r.일자, r.품목군);

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

    row.부족수량 += toNumber(r.부족수량);
    row.부족금액 += toNumber(r.부족금액);
  });

  // 중요:
  // 기존제출파일의 일별통합결과에 전년/당해 값이 없더라도
  // 판매자동소진, 폐기자동소진 시트의 상세값을 다시 합산해서 보정한다.
  const detailDailyMap = buildLockedDetailDailyMap(lockedWorkbook);

  for (const [key, detailRow] of detailDailyMap.entries()) {
    if (!mergedMap.has(key)) {
      mergedMap.set(key, detailRow);
      continue;
    }

    const row = mergedMap.get(key);

    // 총 판매/폐기 수량·금액은 일별통합결과 값이 있으면 유지하고,
    // 상세 시트 기준 연차별 배분값만 강제로 보정한다.
    row.전전년_판매수량 = toNumber(detailRow.전전년_판매수량);
    row.전전년_판매금액 = toNumber(detailRow.전전년_판매금액);
    row.전전년_폐기수량 = toNumber(detailRow.전전년_폐기수량);
    row.전전년_폐기금액 = toNumber(detailRow.전전년_폐기금액);

    row.전년_판매수량 = toNumber(detailRow.전년_판매수량);
    row.전년_판매금액 = toNumber(detailRow.전년_판매금액);
    row.전년_폐기수량 = toNumber(detailRow.전년_폐기수량);
    row.전년_폐기금액 = toNumber(detailRow.전년_폐기금액);

    row.당해_판매수량 = toNumber(detailRow.당해_판매수량);
    row.당해_판매금액 = toNumber(detailRow.당해_판매금액);
    row.당해_폐기수량 = toNumber(detailRow.당해_폐기수량);
    row.당해_폐기금액 = toNumber(detailRow.당해_폐기금액);
  }
}

/*************************************************
 * 재고 잔여 검증
 *************************************************/

function buildStockBalanceRows(stockMap, groupMap, lockedRows) {
  const lockedUseMap = new Map();

  lockedRows.forEach((r) => {
    const key = makeStockKey(r.지점명, r.품목군);

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

    const prev2UsedQty =
      toNumber(locked.prev2Qty) +
      entries.reduce((s, e) => s + toNumber(e.prev2Qty), 0);

    const prev2UsedAmt =
      toNumber(locked.prev2Amt) +
      entries.reduce((s, e) => s + toNumber(e.prev2Amt), 0);

    const prevUsedQty =
      toNumber(locked.prevQty) +
      entries.reduce((s, e) => s + toNumber(e.prevQty), 0);

    const prevUsedAmt =
      toNumber(locked.prevAmt) +
      entries.reduce((s, e) => s + toNumber(e.prevAmt), 0);

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

/*************************************************
 * 메인 처리
 *************************************************/

function processWorkbook(workbook) {
  const rawRows = parseHorizontal2025Sheet(workbook);
  const filteredRows = filterRowsByFinalCutoff(rawRows);

  const lockedRows = parseLockedDailyRows(lockedWorkbook);
  const mainRows = buildDeltaRows(filteredRows, lockedRows);

  const prev2Rows = parseInventorySheetFixed(workbook, PREV2_SHEET);
  const prevRows = parseInventorySheetFixed(workbook, PREV_SHEET);
  const stockMap = buildOpeningStocks(prev2Rows, prevRows);

  const anomalyIssues = validateAnomalies(mainRows);

  const groupMap = new Map();

  mainRows.forEach((row) => {
    const branch = normalizeText(row.지점명);
    const item = normalizeText(row.품목);
    const date = normalizeText(row.날짜);

    if (!branch || !item || !date) return;

    const key = makeStockKey(branch, item);
    if (!groupMap.has(key)) groupMap.set(key, []);

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
  });

  for (const [key, entries] of groupMap.entries()) {
    entries.sort((a, b) => {
      if (a.date !== b.date) return a.date.localeCompare(b.date);
      return a.type.localeCompare(b.type);
    });

    allocateGroupPeriod(entries, stockMap.get(key));
  }

  const mergedMap = new Map();
mergeLockedRowsIntoMergedMap(mergedMap, lockedRows);

// 기존 제출파일의 상세 시트도 그대로 다시 가져와야 함
const lockedSalesRows = parseLockedDetailRows(
  lockedWorkbook,
  "판매자동소진",
  "판매"
);

const lockedDiscardRows = parseLockedDetailRows(
  lockedWorkbook,
  "폐기자동소진",
  "폐기"
);

// 기존 제출분은 그대로 보존하고, 이후 증가분만 뒤에 추가
const salesRows = [...lockedSalesRows];
const discardRows = [...lockedDiscardRows];

const validationRows = [];

  for (const entries of groupMap.values()) {
    entries.forEach((e) => {
      const dailyKey = makeDailyKey(e.branch, e.date, e.item);

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
      }

      if (e.type === "폐기") {
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

/*************************************************
 * 화면 출력
 *************************************************/

function renderIssues(result) {
  const container = document.getElementById("issues");
  if (!container) return;

  container.innerHTML = "";

  const issues = [
    ...result.schemaErrors.map((m) => ({ type: "스키마오류", message: m })),
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
    html += `<p style="margin:0 0 12px; color:#6b7280; font-size:13px;">
      총 ${rows.length.toLocaleString()}행 중 ${maxRows.toLocaleString()}행만 화면에 표시합니다.
      전체 데이터는 엑셀 다운로드에서 확인하세요.
    </p>`;
  }

  html += "<table><thead><tr>";
  cols.forEach((c) => {
    html += `<th>${c}</th>`;
  });
  html += "</tr></thead><tbody>";

  limitedRows.forEach((row) => {
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

  const shortageCount = document.getElementById("shortageCount");
  if (shortageCount) shortageCount.textContent = 0;

  const stockBalanceCount = document.getElementById("stockBalanceCount");
  if (stockBalanceCount) stockBalanceCount.textContent = result.remainStockRows.length;
}

/*************************************************
 * 다운로드
 *************************************************/

function downloadWorkbook(result) {
  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(result.mergedDailyRows),
    "일별통합결과"
  );

  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(result.salesRows),
    "판매자동소진"
  );

  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(result.discardRows),
    "폐기자동소진"
  );

  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(result.validationRows),
    "검증"
  );

  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(result.stockBalanceRows),
    "연차별재고소진검증"
  );

  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(result.remainStockRows),
    "잔여재고만"
  );

  const issueRows = [
    ...result.schemaErrors.map((m) => ({ 유형: "스키마오류", 내용: m })),
    ...result.anomalyIssues.map((x) => ({ 유형: x.type, 내용: x.message })),
  ];

  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(issueRows),
    "오류및이상치"
  );

  XLSX.writeFile(wb, "FIFO_연차별재고소진_감사용.xlsx");
}

/*************************************************
 * 실행
 *************************************************/

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
    console.error(error);
  }
}

/*************************************************
 * 이벤트 연결
 *************************************************/

document.addEventListener("DOMContentLoaded", () => {
  setDefaultTodayDates();

  document.getElementById("fileInput")?.addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
      const data = await file.arrayBuffer();
      latestWorkbook = XLSX.read(data, { type: "array" });
    } catch (error) {
      alert("원본 파일 처리 중 오류가 발생했습니다: " + error.message);
      console.error(error);
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
      console.error(error);
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
