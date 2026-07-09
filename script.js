/*************************************************
 * FIFO 연차별 재고 소진 프로그램 - script.js
 *
 * 최종 대전제
 * 1. 기존 제출파일의 전전년 사용수량/금액은 절대 변경하지 않는다.
 * 2. 기존 제출파일의 전년/당해는 수량·금액 짝오류만 보정한다.
 * 3. 기존 제출 이후 증가분은 전부 당해로 처리한다.
 * 4. 전전년/전년 총 사용량은 총 재고 수량·금액을 초과하면 안 된다.
 * 5. 전전년 재고는 기준일까지 소진 목적, 전년 재고는 연말까지 소진 목적이다.
 * 6. 당해 재고는 이월 가능하다.
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

function roundNumber(v) {
  return Math.round(toNumber(v));
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

  ["prev2CutoffDate", "prevCutoffDate", "finalCutoffDate", "cutoffDate"].forEach((id) => {
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
 * 원본 2025 가로형 시트 파싱
 *************************************************/

function parseHorizontal2025Sheet(workbook) {
  const ws = workbook.Sheets[MAIN_SHEET];

  if (!ws) throw new Error(`시트 '${MAIN_SHEET}' 을(를) 찾을 수 없습니다.`);
  if (!ws["!ref"]) throw new Error(`시트 '${MAIN_SHEET}' 의 범위를 읽을 수 없습니다.`);

  const range = XLSX.utils.decode_range(ws["!ref"]);
  const branchBlocks = [];
  const rows = [];

  let currentBranch = null;

  const headerRowBranch = 9;
  const headerRowField = 10;
  const dataStartRow = 11;

  for (let c = 1; c <= range.e.c + 1; c++) {
    const col = numberToCol(c);
    const branchName = normalizeText(getCellValue(ws, `${col}${headerRowBranch}`));
    const fieldName = normalizeHeaderText(getCellValue(ws, `${col}${headerRowField}`));

    if (branchName) {
      if (
        currentBranch &&
        currentBranch.지점명 &&
        currentBranch.판매수량Col &&
        currentBranch.판매금액Col &&
        currentBranch.폐기수량Col &&
        currentBranch.폐기금액Col
      ) {
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

  if (
    currentBranch &&
    currentBranch.지점명 &&
    currentBranch.판매수량Col &&
    currentBranch.판매금액Col &&
    currentBranch.폐기수량Col &&
    currentBranch.폐기금액Col
  ) {
    branchBlocks.push(currentBranch);
  }

  if (!branchBlocks.length) {
    throw new Error("2025 시트에서 유효한 지점 블록을 찾지 못했습니다.");
  }

  for (let r = dataStartRow; r <= range.e.r + 1; r++) {
    const 날짜 = formatDate(getCellValue(ws, `A${r}`));
    const 품목 = normalizeText(getCellValue(ws, `B${r}`));

    if (!날짜 && !품목) continue;

    branchBlocks.forEach((b) => {
      const 판매수량 = roundNumber(getCellValue(ws, `${b.판매수량Col}${r}`));
      const 판매금액 = roundNumber(getCellValue(ws, `${b.판매금액Col}${r}`));
      const 폐기수량 = roundNumber(getCellValue(ws, `${b.폐기수량Col}${r}`));
      const 폐기금액 = roundNumber(getCellValue(ws, `${b.폐기금액Col}${r}`));

      if (판매수량 === 0 && 판매금액 === 0 && 폐기수량 === 0 && 폐기금액 === 0) return;

      rows.push({
        지점명: b.지점명,
        날짜,
        품목,
        판매수량,
        판매금액,
        폐기수량,
        폐기금액,
      });
    });
  }

  return sortRowsByBusinessOrder(rows);
}

/*************************************************
 * 재고 시트 파싱
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
    const 수량 = roundNumber(row[3]);
    const 금액 = roundNumber(row[4]);

    if (!지점명 && !품목 && 수량 === 0 && 금액 === 0) continue;
    if (!지점명 || !품목) continue;

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

  prev2Rows.forEach((r) => {
    const x = ensure(r.지점명, r.품목);
    x.전전년수량 += roundNumber(r.수량);
    x.전전년금액 += roundNumber(r.금액);
  });

  prevRows.forEach((r) => {
    const x = ensure(r.지점명, r.품목);
    x.전년수량 += roundNumber(r.수량);
    x.전년금액 += roundNumber(r.금액);
  });

  return map;
}

/*************************************************
 * 결과 행 기본 구조
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

/*************************************************
 * 금액 배분 유틸
 *************************************************/

function distributeAmountByQty(totalAmt, parts, target) {
  const amt = roundNumber(totalAmt);

  parts.forEach((p) => {
    target[p.amtField] = 0;
  });

  const valid = parts.filter((p) => roundNumber(p.qty) > 0);

  if (amt <= 0 || !valid.length) return;

  if (valid.length === 1) {
    target[valid[0].amtField] = amt;
    return;
  }

  if (amt < valid.length) {
    const sorted = [...valid].sort((a, b) => roundNumber(b.qty) - roundNumber(a.qty));
    let remain = amt;

    sorted.forEach((p) => {
      if (remain > 0) {
        target[p.amtField] = 1;
        remain -= 1;
      }
    });

    return;
  }

  valid.forEach((p) => {
    target[p.amtField] = 1;
  });

  let remainAmt = amt - valid.length;
  const totalQty = valid.reduce((s, p) => s + roundNumber(p.qty), 0);

  const temp = valid.map((p) => {
    const exact = (remainAmt * roundNumber(p.qty)) / totalQty;
    const base = Math.floor(exact);

    return {
      ...p,
      base,
      frac: exact - base,
    };
  });

  let assigned = 0;

  temp.forEach((p) => {
    target[p.amtField] += p.base;
    assigned += p.base;
  });

  let leftover = remainAmt - assigned;

  temp.sort((a, b) => b.frac - a.frac);

  let i = 0;
  while (leftover > 0 && temp.length) {
    target[temp[i % temp.length].amtField] += 1;
    leftover -= 1;
    i += 1;
  }

  const allocated = valid.reduce((s, p) => s + roundNumber(target[p.amtField]), 0);
  const diff = amt - allocated;

  if (diff !== 0) {
    const receiver =
      valid.find((p) => String(p.amtField).includes("당해") || p.amtField === "currentAmt") ||
      valid.find((p) => String(p.amtField).includes("전년") || p.amtField === "prevAmt") ||
      valid[0];

    target[receiver.amtField] += diff;
  }

  parts.forEach((p) => {
    if (roundNumber(p.qty) <= 0) target[p.amtField] = 0;
    if (roundNumber(target[p.amtField]) < 0) target[p.amtField] = 0;
  });
}

/*************************************************
 * 기존 제출 상세행 정리
 * 중요:
 * - 전전년 수량/금액은 절대 변경하지 않음
 * - 전년/당해 금액만 잔여금액 안에서 보정
 *************************************************/

function normalizeLockedDetailRow(row) {
  const fixed = { ...row };

  const totalAmt = roundNumber(fixed.총사용금액);

  const fixedPrev2Qty = roundNumber(fixed.전전년사용수량);
  const fixedPrev2Amt = fixedPrev2Qty > 0 ? roundNumber(fixed.전전년사용금액) : 0;

  fixed.전전년사용금액 = fixedPrev2Amt;

  let remainAmt = totalAmt - fixedPrev2Amt;

  if (remainAmt < 0) {
    remainAmt = 0;
  }

  distributeAmountByQty(
    remainAmt,
    [
      { qty: fixed.전년사용수량, amtField: "전년사용금액" },
      { qty: fixed.당해사용수량, amtField: "당해사용금액" },
    ],
    fixed
  );

  return fixed;
}

function parseLockedDetailRows(workbook, sheetName, type) {
  if (!workbook) return [];

  const ws = workbook.Sheets[sheetName];
  if (!ws || !ws["!ref"]) return [];

  return XLSX.utils
    .sheet_to_json(ws, { defval: "" })
    .map((r) => ({
      구분: normalizeText(r.구분) || `${type}_기존제출`,
      지점명: normalizeText(r.지점명),
      날짜: formatDate(r.날짜),
      품목: normalizeText(r.품목),

      총사용수량: roundNumber(r.총사용수량),
      총사용금액: roundNumber(r.총사용금액),

      전전년사용수량: roundNumber(r.전전년사용수량),
      전전년사용금액: roundNumber(r.전전년사용금액),

      전년사용수량: roundNumber(r.전년사용수량),
      전년사용금액: roundNumber(r.전년사용금액),

      당해사용수량: roundNumber(r.당해사용수량),
      당해사용금액: roundNumber(r.당해사용금액),

      부족수량: 0,
      부족금액: 0,
    }))
    .filter((r) => r.지점명 && r.날짜 && r.품목)
    .map(normalizeLockedDetailRow);
}

function parseLockedDailyRows(workbook) {
  if (!workbook) return [];

  const ws = workbook.Sheets["일별통합결과"];
  if (!ws || !ws["!ref"]) return [];

  return XLSX.utils
    .sheet_to_json(ws, { defval: "" })
    .map((r) => ({
      지점명: normalizeText(r.지점명),
      일자: formatDate(r.일자),
      품목군: normalizeText(r.품목군),

      판매수량: roundNumber(r.판매수량),
      판매금액: roundNumber(r.판매금액),
      폐기수량: roundNumber(r.폐기수량),
      폐기금액: roundNumber(r.폐기금액),

      총사용수량: roundNumber(r.총사용수량),
      총사용금액: roundNumber(r.총사용금액),

      전전년_판매수량: roundNumber(r.전전년_판매수량),
      전전년_판매금액: roundNumber(r.전전년_판매금액),
      전전년_폐기수량: roundNumber(r.전전년_폐기수량),
      전전년_폐기금액: roundNumber(r.전전년_폐기금액),

      전년_판매수량: roundNumber(r.전년_판매수량),
      전년_판매금액: roundNumber(r.전년_판매금액),
      전년_폐기수량: roundNumber(r.전년_폐기수량),
      전년_폐기금액: roundNumber(r.전년_폐기금액),

      당해_판매수량: roundNumber(r.당해_판매수량 ?? r.당월_판매수량),
      당해_판매금액: roundNumber(r.당해_판매금액 ?? r.당월_판매금액),
      당해_폐기수량: roundNumber(r.당해_폐기수량 ?? r.당월_폐기수량),
      당해_폐기금액: roundNumber(r.당해_폐기금액 ?? r.당월_폐기금액),

      부족수량: 0,
      부족금액: 0,
    }))
    .filter((r) => r.지점명 && r.일자 && r.품목군);
}

/*************************************************
 * 기존 제출파일을 일별통합결과로 복원
 *************************************************/

function buildLockedDailyMap(workbook) {
  const map = new Map();
  if (!workbook) return map;

  const sales = parseLockedDetailRows(workbook, "판매자동소진", "판매");
  const discards = parseLockedDetailRows(workbook, "폐기자동소진", "폐기");

  function ensure(branch, date, item) {
    const key = makeDailyKey(branch, date, item);

    if (!map.has(key)) {
      map.set(key, createDailyCombinedRow(branch, date, item));
    }

    return map.get(key);
  }

  sales.forEach((r) => {
    const row = ensure(r.지점명, r.날짜, r.품목);

    row.판매수량 += r.총사용수량;
    row.판매금액 += r.총사용금액;

    row.전전년_판매수량 += r.전전년사용수량;
    row.전전년_판매금액 += r.전전년사용금액;

    row.전년_판매수량 += r.전년사용수량;
    row.전년_판매금액 += r.전년사용금액;

    row.당해_판매수량 += r.당해사용수량;
    row.당해_판매금액 += r.당해사용금액;
  });

  discards.forEach((r) => {
    const row = ensure(r.지점명, r.날짜, r.품목);

    row.폐기수량 += r.총사용수량;
    row.폐기금액 += r.총사용금액;

    row.전전년_폐기수량 += r.전전년사용수량;
    row.전전년_폐기금액 += r.전전년사용금액;

    row.전년_폐기수량 += r.전년사용수량;
    row.전년_폐기금액 += r.전년사용금액;

    row.당해_폐기수량 += r.당해사용수량;
    row.당해_폐기금액 += r.당해사용금액;
  });

  if (map.size === 0) {
    parseLockedDailyRows(workbook).forEach((r) => {
      map.set(makeDailyKey(r.지점명, r.일자, r.품목군), { ...r });
    });
  }

  for (const row of map.values()) {
    finalizeDailyRow(row);
  }

  return map;
}

/*************************************************
 * 현재 원본 일별 합계 및 증가분 산정
 *************************************************/

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
        폐기수량: 0,
        폐기금액: 0,
      });
    }

    const row = map.get(key);

    row.판매수량 += roundNumber(r.판매수량);
    row.판매금액 += roundNumber(r.판매금액);
    row.폐기수량 += roundNumber(r.폐기수량);
    row.폐기금액 += roundNumber(r.폐기금액);
  });

  return Array.from(map.values());
}

function buildLockedCompareMap(workbook) {
  const map = new Map();
  if (!workbook) return map;

  const lockedDailyMap = buildLockedDailyMap(workbook);

  for (const [key, row] of lockedDailyMap.entries()) {
    map.set(key, {
      판매수량: row.판매수량,
      판매금액: row.판매금액,
      폐기수량: row.폐기수량,
      폐기금액: row.폐기금액,
    });
  }

  return map;
}

function buildDeltaRows(currentRows, workbookForLocked) {
  const currentDailyRows = buildCurrentDailyRows(currentRows);
  if (!workbookForLocked) return currentDailyRows;

  const lockedMap = buildLockedCompareMap(workbookForLocked);
  const deltaRows = [];

  currentDailyRows.forEach((r) => {
    const key = makeDailyKey(r.지점명, r.날짜, r.품목);
    const locked = lockedMap.get(key) || {};

    const deltaSaleQty = Math.max(0, r.판매수량 - roundNumber(locked.판매수량));
    const deltaSaleAmt = Math.max(0, r.판매금액 - roundNumber(locked.판매금액));

    const deltaDiscardQty = Math.max(0, r.폐기수량 - roundNumber(locked.폐기수량));
    const deltaDiscardAmt = Math.max(0, r.폐기금액 - roundNumber(locked.폐기금액));

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

      폐기수량: deltaDiscardQty,
      폐기금액: deltaDiscardAmt,
    });
  });

  return deltaRows;
}

function filterRowsByFinalCutoff(rows) {
  const cutoff = getFinalCutoffDate();
  if (!cutoff) return rows;
  return rows.filter((row) => normalizeText(row.날짜) <= cutoff);
}

/*************************************************
 * 증가분은 무조건 당해로 추가
 *************************************************/

function addDeltaRowsAsCurrent(mergedMap, deltaRows, salesRows, discardRows) {
  deltaRows.forEach((r) => {
    const key = makeDailyKey(r.지점명, r.날짜, r.품목);

    if (!mergedMap.has(key)) {
      mergedMap.set(key, createDailyCombinedRow(r.지점명, r.날짜, r.품목));
    }

    const row = mergedMap.get(key);

    if (roundNumber(r.판매수량) > 0 || roundNumber(r.판매금액) > 0) {
      const qty = roundNumber(r.판매수량);
      const amt = roundNumber(r.판매금액);

      row.판매수량 += qty;
      row.판매금액 += amt;

      row.당해_판매수량 += qty;
      row.당해_판매금액 += amt;

      if (qty > 0 && amt > 0) {
        salesRows.push({
          구분: "판매_증가분_당해",
          지점명: r.지점명,
          날짜: r.날짜,
          품목: r.품목,

          총사용수량: qty,
          총사용금액: amt,

          전전년사용수량: 0,
          전전년사용금액: 0,

          전년사용수량: 0,
          전년사용금액: 0,

          당해사용수량: qty,
          당해사용금액: amt,

          부족수량: 0,
          부족금액: 0,
        });
      }
    }

    if (roundNumber(r.폐기수량) > 0 || roundNumber(r.폐기금액) > 0) {
      const qty = roundNumber(r.폐기수량);
      const amt = roundNumber(r.폐기금액);

      row.폐기수량 += qty;
      row.폐기금액 += amt;

      row.당해_폐기수량 += qty;
      row.당해_폐기금액 += amt;

      if (qty > 0 && amt > 0) {
        discardRows.push({
          구분: "폐기_증가분_당해",
          지점명: r.지점명,
          날짜: r.날짜,
          품목: r.품목,

          총사용수량: qty,
          총사용금액: amt,

          전전년사용수량: 0,
          전전년사용금액: 0,

          전년사용수량: 0,
          전년사용금액: 0,

          당해사용수량: qty,
          당해사용금액: amt,

          부족수량: 0,
          부족금액: 0,
        });
      }
    }
  });
}

/*************************************************
 * 최종 일별 행 보정
 * - 전전년 금액은 고정
 * - 전년/당해 금액만 보정
 *************************************************/

function finalizeDailyRow(row) {
  row.판매수량 = roundNumber(row.판매수량);
  row.판매금액 = roundNumber(row.판매금액);
  row.폐기수량 = roundNumber(row.폐기수량);
  row.폐기금액 = roundNumber(row.폐기금액);

  row.총사용수량 = row.판매수량 + row.폐기수량;
  row.총사용금액 = row.판매금액 + row.폐기금액;

  row.부족수량 = 0;
  row.부족금액 = 0;

  fixDailyRowQtyAmountPairs(row);

  return row;
}

function fixDailyRowQtyAmountPairs(row) {
  const fixedPrev2SaleAmt =
    roundNumber(row.전전년_판매수량) > 0
      ? roundNumber(row.전전년_판매금액)
      : 0;

  row.전전년_판매금액 = fixedPrev2SaleAmt;

  let remainSaleAmt = roundNumber(row.판매금액) - fixedPrev2SaleAmt;
  if (remainSaleAmt < 0) remainSaleAmt = 0;

  distributeAmountByQty(
    remainSaleAmt,
    [
      { qty: row.전년_판매수량, amtField: "전년_판매금액" },
      { qty: row.당해_판매수량, amtField: "당해_판매금액" },
    ],
    row
  );

  const fixedPrev2DiscardAmt =
    roundNumber(row.전전년_폐기수량) > 0
      ? roundNumber(row.전전년_폐기금액)
      : 0;

  row.전전년_폐기금액 = fixedPrev2DiscardAmt;

  let remainDiscardAmt = roundNumber(row.폐기금액) - fixedPrev2DiscardAmt;
  if (remainDiscardAmt < 0) remainDiscardAmt = 0;

  distributeAmountByQty(
    remainDiscardAmt,
    [
      { qty: row.전년_폐기수량, amtField: "전년_폐기금액" },
      { qty: row.당해_폐기수량, amtField: "당해_폐기금액" },
    ],
    row
  );

  return row;
}

/*************************************************
 * 전년 재고 초과분 처리
 * - 전전년은 고정값이므로 자동수정하지 않고 검증오류 표시
 * - 전년은 초과 시 초과분을 당해로 이동
 *************************************************/

function enforcePrevStockLimit(dailyRows, stockMap) {
  const grouped = new Map();

  dailyRows.forEach((row) => {
    const key = makeStockKey(row.지점명, row.품목군);

    if (!grouped.has(key)) grouped.set(key, []);

    grouped.get(key).push(row);
  });

  for (const [key, rows] of grouped.entries()) {
    const stock = stockMap.get(key);
    if (!stock) continue;

    let remainPrevQty = roundNumber(stock.전년수량);

    rows.sort((a, b) => {
      if (a.일자 !== b.일자) return a.일자.localeCompare(b.일자);
      return 0;
    });

    rows.forEach((row) => {
      let salePrevQty = roundNumber(row.전년_판매수량);
      let discardPrevQty = roundNumber(row.전년_폐기수량);

      const totalPrevQty = salePrevQty + discardPrevQty;

      if (totalPrevQty <= remainPrevQty) {
        remainPrevQty -= totalPrevQty;
        return;
      }

      let allowed = Math.max(0, remainPrevQty);

      const saleAllowed = Math.min(salePrevQty, allowed);
      allowed -= saleAllowed;

      const discardAllowed = Math.min(discardPrevQty, allowed);

      const saleExcess = salePrevQty - saleAllowed;
      const discardExcess = discardPrevQty - discardAllowed;

      if (saleExcess > 0) {
        row.전년_판매수량 -= saleExcess;
        row.당해_판매수량 += saleExcess;
      }

      if (discardExcess > 0) {
        row.전년_폐기수량 -= discardExcess;
        row.당해_폐기수량 += discardExcess;
      }

      remainPrevQty = 0;

      finalizeDailyRow(row);
    });
  }
}

/*************************************************
 * 재고 사용 검증
 *************************************************/

function buildStockUseMapFromDailyRows(dailyRows) {
  const map = new Map();

  function ensure(branch, item) {
    const key = makeStockKey(branch, item);

    if (!map.has(key)) {
      map.set(key, {
        전전년사용수량: 0,
        전전년사용금액: 0,
        전년사용수량: 0,
        전년사용금액: 0,
      });
    }

    return map.get(key);
  }

  dailyRows.forEach((r) => {
    const x = ensure(r.지점명, r.품목군);

    x.전전년사용수량 += roundNumber(r.전전년_판매수량) + roundNumber(r.전전년_폐기수량);
    x.전전년사용금액 += roundNumber(r.전전년_판매금액) + roundNumber(r.전전년_폐기금액);

    x.전년사용수량 += roundNumber(r.전년_판매수량) + roundNumber(r.전년_폐기수량);
    x.전년사용금액 += roundNumber(r.전년_판매금액) + roundNumber(r.전년_폐기금액);
  });

  return map;
}

function buildStockBalanceRows(stockMap, dailyRows) {
  const useMap = buildStockUseMapFromDailyRows(dailyRows);
  const rows = [];

  for (const [key, stock] of stockMap.entries()) {
    const used = useMap.get(key) || {};

    const prev2UsedQty = roundNumber(used.전전년사용수량);
    const prev2UsedAmt = roundNumber(used.전전년사용금액);

    const prevUsedQty = roundNumber(used.전년사용수량);
    const prevUsedAmt = roundNumber(used.전년사용금액);

    const prev2StockQty = roundNumber(stock.전전년수량);
    const prev2StockAmt = roundNumber(stock.전전년금액);

    const prevStockQty = roundNumber(stock.전년수량);
    const prevStockAmt = roundNumber(stock.전년금액);

    const prev2RemainQty = prev2StockQty - prev2UsedQty;
    const prev2RemainAmt = prev2StockAmt - prev2UsedAmt;

    const prevRemainQty = prevStockQty - prevUsedQty;
    const prevRemainAmt = prevStockAmt - prevUsedAmt;

    rows.push({
      지점명: stock.지점명,
      품목: stock.품목,

      전전년소진기준일: getPrev2CutoffDate(),
      전전년재고수량: prev2StockQty,
      전전년사용수량: prev2UsedQty,
      전전년잔여수량: prev2RemainQty,
      전전년재고금액: prev2StockAmt,
      전전년사용금액: prev2UsedAmt,
      전전년잔여금액: prev2RemainAmt,
      전전년재고초과: prev2UsedQty > prev2StockQty || prev2UsedAmt > prev2StockAmt ? "초과" : "",
      전전년기준일까지미소진: prev2RemainQty > 0 || prev2RemainAmt > 0 ? "미소진" : "",

      전년소진기준일: getPrevCutoffDate(),
      전년재고수량: prevStockQty,
      전년사용수량: prevUsedQty,
      전년잔여수량: prevRemainQty,
      전년재고금액: prevStockAmt,
      전년사용금액: prevUsedAmt,
      전년잔여금액: prevRemainAmt,
      전년재고초과: prevUsedQty > prevStockQty || prevUsedAmt > prevStockAmt ? "초과" : "",
      전년연말미소진: prevRemainQty > 0 || prevRemainAmt > 0 ? "미소진" : "",
    });
  }

  return sortRowsByBusinessOrder(rows);
}

function buildOnlyRemainStockRows(stockBalanceRows) {
  return stockBalanceRows.filter((r) => {
    return (
      roundNumber(r.전전년잔여수량) !== 0 ||
      roundNumber(r.전전년잔여금액) !== 0 ||
      roundNumber(r.전년잔여수량) !== 0 ||
      roundNumber(r.전년잔여금액) !== 0 ||
      r.전전년재고초과 ||
      r.전년재고초과
    );
  });
}

/*************************************************
 * 검증
 *************************************************/

function hasPairError(qty, amt) {
  qty = roundNumber(qty);
  amt = roundNumber(amt);
  return (qty > 0 && amt <= 0) || (qty <= 0 && amt > 0);
}

function buildValidationRows(dailyRows, stockBalanceRows) {
  const stockErrorMap = new Map();

  stockBalanceRows.forEach((r) => {
    stockErrorMap.set(makeStockKey(r.지점명, r.품목), {
      전전년재고초과: r.전전년재고초과,
      전년재고초과: r.전년재고초과,
      전전년기준일까지미소진: r.전전년기준일까지미소진,
      전년연말미소진: r.전년연말미소진,
    });
  });

  return dailyRows.map((row) => {
    const periodQtySum =
      roundNumber(row.전전년_판매수량) +
      roundNumber(row.전전년_폐기수량) +
      roundNumber(row.전년_판매수량) +
      roundNumber(row.전년_폐기수량) +
      roundNumber(row.당해_판매수량) +
      roundNumber(row.당해_폐기수량);

    const periodAmtSum =
      roundNumber(row.전전년_판매금액) +
      roundNumber(row.전전년_폐기금액) +
      roundNumber(row.전년_판매금액) +
      roundNumber(row.전년_폐기금액) +
      roundNumber(row.당해_판매금액) +
      roundNumber(row.당해_폐기금액);

    const stockError = stockErrorMap.get(makeStockKey(row.지점명, row.품목군)) || {};

    return {
      지점명: row.지점명,
      날짜: row.일자,
      품목: row.품목군,

      총사용수량: row.총사용수량,
      총사용금액: row.총사용금액,

      연차배분수량합: periodQtySum,
      수량차이: roundNumber(row.총사용수량) - periodQtySum,

      연차배분금액합: periodAmtSum,
      금액차이: roundNumber(row.총사용금액) - periodAmtSum,

      전전년판매쌍오류: hasPairError(row.전전년_판매수량, row.전전년_판매금액) ? "오류" : "",
      전년판매쌍오류: hasPairError(row.전년_판매수량, row.전년_판매금액) ? "오류" : "",
      당해판매쌍오류: hasPairError(row.당해_판매수량, row.당해_판매금액) ? "오류" : "",

      전전년폐기쌍오류: hasPairError(row.전전년_폐기수량, row.전전년_폐기금액) ? "오류" : "",
      전년폐기쌍오류: hasPairError(row.전년_폐기수량, row.전년_폐기금액) ? "오류" : "",
      당해폐기쌍오류: hasPairError(row.당해_폐기수량, row.당해_폐기금액) ? "오류" : "",

      전전년재고초과: stockError.전전년재고초과 || "",
      전년재고초과: stockError.전년재고초과 || "",
      전전년기준일까지미소진: stockError.전전년기준일까지미소진 || "",
      전년연말미소진: stockError.전년연말미소진 || "",

      부족수량: 0,
      부족금액: 0,
    };
  });
}

function validationHasError(r) {
  return (
    roundNumber(r.수량차이) !== 0 ||
    roundNumber(r.금액차이) !== 0 ||
    r.전전년판매쌍오류 ||
    r.전년판매쌍오류 ||
    r.당해판매쌍오류 ||
    r.전전년폐기쌍오류 ||
    r.전년폐기쌍오류 ||
    r.당해폐기쌍오류 ||
    r.전전년재고초과 ||
    r.전년재고초과
  );
}

function validateAnomalies(rows) {
  const issues = [];

  rows.forEach((row, idx) => {
    const rowNo = idx + 1;

    const saleQty = roundNumber(row.판매수량);
    const saleAmt = roundNumber(row.판매금액);
    const discardQty = roundNumber(row.폐기수량);
    const discardAmt = roundNumber(row.폐기금액);

    if (saleQty > 0 && saleAmt === 0) {
      issues.push({ 유형: "원본이상치", 내용: `행 ${rowNo}: 판매수량은 있는데 판매금액이 0입니다.` });
    }

    if (saleQty === 0 && saleAmt > 0) {
      issues.push({ 유형: "원본이상치", 내용: `행 ${rowNo}: 판매금액은 있는데 판매수량이 0입니다.` });
    }

    if (discardQty > 0 && discardAmt === 0) {
      issues.push({ 유형: "원본이상치", 내용: `행 ${rowNo}: 폐기수량은 있는데 폐기금액이 0입니다.` });
    }

    if (discardQty === 0 && discardAmt > 0) {
      issues.push({ 유형: "원본이상치", 내용: `행 ${rowNo}: 폐기금액은 있는데 폐기수량이 0입니다.` });
    }
  });

  return issues;
}

/*************************************************
 * 메인 처리
 *************************************************/

function processWorkbook(workbook) {
  const rawRows = parseHorizontal2025Sheet(workbook);
  const filteredRows = filterRowsByFinalCutoff(rawRows);

  const stockMap = buildOpeningStocks(
    parseInventorySheetFixed(workbook, PREV2_SHEET),
    parseInventorySheetFixed(workbook, PREV_SHEET)
  );

  const deltaRows = buildDeltaRows(filteredRows, lockedWorkbook);
  const anomalyIssues = validateAnomalies(deltaRows);

  const mergedMap = buildLockedDailyMap(lockedWorkbook);

  const salesRows = parseLockedDetailRows(lockedWorkbook, "판매자동소진", "판매");
  const discardRows = parseLockedDetailRows(lockedWorkbook, "폐기자동소진", "폐기");

  addDeltaRowsAsCurrent(mergedMap, deltaRows, salesRows, discardRows);

  let mergedDailyRows = sortRowsByBusinessOrder(Array.from(mergedMap.values())).map(finalizeDailyRow);

  enforcePrevStockLimit(mergedDailyRows, stockMap);

  mergedDailyRows = sortRowsByBusinessOrder(mergedDailyRows).map(finalizeDailyRow);

  const stockBalanceRows = buildStockBalanceRows(stockMap, mergedDailyRows);
  const remainStockRows = buildOnlyRemainStockRows(stockBalanceRows);
  const validationRows = sortRowsByBusinessOrder(buildValidationRows(mergedDailyRows, stockBalanceRows));

  return {
    ok: true,
    schemaErrors: [],
    anomalyIssues,
    mergedDailyRows,
    salesRows: sortRowsByBusinessOrder(salesRows),
    discardRows: sortRowsByBusinessOrder(discardRows),
    validationRows,
    stockBalanceRows,
    remainStockRows,
  };
}

/*************************************************
 * 화면 출력
 *************************************************/

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

function renderIssues(result) {
  const container = document.getElementById("issues");
  if (!container) return;

  container.innerHTML = "";

  const validationErrors = result.validationRows
    .filter(validationHasError)
    .map((r) => ({
      유형: "검증오류",
      내용: `${r.지점명} / ${r.날짜} / ${r.품목}: 수량차이 ${r.수량차이}, 금액차이 ${r.금액차이}, 전전년초과 ${r.전전년재고초과}, 전년초과 ${r.전년재고초과}`,
    }));

  const issues = [...result.anomalyIssues, ...validationErrors];

  if (!issues.length) {
    container.innerHTML = `<div class="issue ok">오류 및 이상치가 없습니다.</div>`;
    return;
  }

  issues.slice(0, 300).forEach((issue) => {
    const div = document.createElement("div");
    div.className = "issue";
    div.textContent = issue.내용;
    container.appendChild(div);
  });
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
    issueCount.textContent =
      result.anomalyIssues.length + result.validationRows.filter(validationHasError).length;
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

  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.mergedDailyRows), "일별통합결과");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.salesRows), "판매자동소진");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.discardRows), "폐기자동소진");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.validationRows), "검증");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.stockBalanceRows), "연차별재고소진검증");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.remainStockRows), "잔여재고만");

  const issueRows = [...result.anomalyIssues];

  result.validationRows.forEach((r) => {
    if (validationHasError(r)) {
      issueRows.push({
        유형: "검증오류",
        지점명: r.지점명,
        날짜: r.날짜,
        품목: r.품목,
        내용: `수량차이:${r.수량차이}, 금액차이:${r.금액차이}, 전전년재고초과:${r.전전년재고초과}, 전년재고초과:${r.전년재고초과}`,
      });
    }
  });

  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(issueRows), "오류및이상치");

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
