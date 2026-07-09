/*************************************************
 * FIFO 연차별 재고 소진 프로그램 - script.js
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

    branchBlocks.forEach((block) => {
      const 판매수량 = Math.round(toNumber(getCellValue(ws, `${block.판매수량Col}${r}`)));
      const 판매금액 = Math.round(toNumber(getCellValue(ws, `${block.판매금액Col}${r}`)));
      const 폐기수량 = Math.round(toNumber(getCellValue(ws, `${block.폐기수량Col}${r}`)));
      const 폐기금액 = Math.round(toNumber(getCellValue(ws, `${block.폐기금액Col}${r}`)));

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
    const 수량 = Math.round(toNumber(row[3]));
    const 금액 = Math.round(toNumber(row[4]));

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

  prev2Rows.forEach((row) => {
    const target = ensure(row.지점명, row.품목);
    target.전전년수량 += Math.round(toNumber(row.수량));
    target.전전년금액 += Math.round(toNumber(row.금액));
  });

  prevRows.forEach((row) => {
    const target = ensure(row.지점명, row.품목);
    target.전년수량 += Math.round(toNumber(row.수량));
    target.전년금액 += Math.round(toNumber(row.금액));
  });

  return map;
}

/*************************************************
 * 기존 제출 결과 파싱
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

function parseLockedDailyRows(workbook) {
  if (!workbook) return [];

  const ws = workbook.Sheets["일별통합결과"];
  if (!ws || !ws["!ref"]) return [];

  return XLSX.utils.sheet_to_json(ws, { defval: "" }).map((r) => ({
    지점명: normalizeText(r.지점명),
    일자: formatDate(r.일자),
    품목군: normalizeText(r.품목군),

    판매수량: Math.round(toNumber(r.판매수량)),
    판매금액: Math.round(toNumber(r.판매금액)),
    폐기수량: Math.round(toNumber(r.폐기수량)),
    폐기금액: Math.round(toNumber(r.폐기금액)),

    총사용수량: Math.round(toNumber(r.총사용수량)),
    총사용금액: Math.round(toNumber(r.총사용금액)),

    전전년_판매수량: Math.round(toNumber(r.전전년_판매수량)),
    전전년_판매금액: Math.round(toNumber(r.전전년_판매금액)),
    전전년_폐기수량: Math.round(toNumber(r.전전년_폐기수량)),
    전전년_폐기금액: Math.round(toNumber(r.전전년_폐기금액)),

    전년_판매수량: Math.round(toNumber(r.전년_판매수량)),
    전년_판매금액: Math.round(toNumber(r.전년_판매금액)),
    전년_폐기수량: Math.round(toNumber(r.전년_폐기수량)),
    전년_폐기금액: Math.round(toNumber(r.전년_폐기금액)),

    당해_판매수량: Math.round(toNumber(r.당해_판매수량 ?? r.당월_판매수량)),
    당해_판매금액: Math.round(toNumber(r.당해_판매금액 ?? r.당월_판매금액)),
    당해_폐기수량: Math.round(toNumber(r.당해_폐기수량 ?? r.당월_폐기수량)),
    당해_폐기금액: Math.round(toNumber(r.당해_폐기금액 ?? r.당월_폐기금액)),

    부족수량: Math.round(toNumber(r.부족수량)),
    부족금액: Math.round(toNumber(r.부족금액)),
  }));
}

function distributeAmountToParts(totalAmt, parts, target) {
  const amt = Math.round(toNumber(totalAmt));

  parts.forEach((p) => {
    target[p.amtField] = 0;
  });

  if (amt <= 0 || !parts.length) return;

  const validParts = parts.filter((p) => Math.round(toNumber(p.qty)) > 0);
  if (!validParts.length) return;

  if (validParts.length === 1) {
    target[validParts[0].amtField] = amt;
    return;
  }

  if (amt < validParts.length) {
    const sorted = [...validParts].sort((a, b) => toNumber(b.qty) - toNumber(a.qty));
    let remainSmall = amt;

    sorted.forEach((p) => {
      if (remainSmall > 0) {
        target[p.amtField] = 1;
        remainSmall -= 1;
      }
    });

    return;
  }

  validParts.forEach((p) => {
    target[p.amtField] = 1;
  });

  let remainAmt = amt - validParts.length;
  const totalQty = validParts.reduce((sum, p) => sum + Math.round(toNumber(p.qty)), 0);

  const temp = validParts.map((p) => {
    const exact = (remainAmt * Math.round(toNumber(p.qty))) / totalQty;
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
    const p = temp[i % temp.length];
    target[p.amtField] += 1;
    leftover -= 1;
    i += 1;
  }

  const allocated = validParts.reduce((sum, p) => sum + Math.round(toNumber(target[p.amtField])), 0);
  const diff = amt - allocated;

  if (diff !== 0) {
    const receiver =
      validParts.find((p) => p.amtField === "currentAmt" || p.amtField === "당해사용금액") ||
      validParts.find((p) => p.amtField === "prevAmt" || p.amtField === "전년사용금액") ||
      validParts[0];

    target[receiver.amtField] += diff;
  }

  parts.forEach((p) => {
    if (toNumber(p.qty) <= 0) target[p.amtField] = 0;
    if (toNumber(target[p.amtField]) < 0) target[p.amtField] = 0;
  });
}

function normalizeLockedDetailRows(rows) {
  return rows.map((row) => {
    const fixed = { ...row };

    const totalAmt = Math.round(toNumber(fixed.총사용금액));

    const pairs = [
      { qtyField: "전전년사용수량", amtField: "전전년사용금액" },
      { qtyField: "전년사용수량", amtField: "전년사용금액" },
      { qtyField: "당해사용수량", amtField: "당해사용금액" },
    ];

    pairs.forEach((p) => {
      if (toNumber(fixed[p.qtyField]) <= 0) fixed[p.amtField] = 0;
    });

    const qtyParts = pairs
      .map((p) => ({ ...p, qty: Math.round(toNumber(fixed[p.qtyField])) }))
      .filter((p) => p.qty > 0);

    if (!qtyParts.length || totalAmt <= 0) {
      pairs.forEach((p) => {
        fixed[p.amtField] = 0;
      });
      return fixed;
    }

    pairs.forEach((p) => {
      fixed[p.amtField] = 0;
    });

    distributeAmountToParts(
      totalAmt,
      qtyParts.map((p) => ({ ...p, qty: p.qty })),
      fixed
    );

    return fixed;
  });
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

    총사용수량: Math.round(toNumber(r.총사용수량)),
    총사용금액: Math.round(toNumber(r.총사용금액)),

    전전년사용수량: Math.round(toNumber(r.전전년사용수량)),
    전전년사용금액: Math.round(toNumber(r.전전년사용금액)),

    전년사용수량: Math.round(toNumber(r.전년사용수량)),
    전년사용금액: Math.round(toNumber(r.전년사용금액)),

    당해사용수량: Math.round(toNumber(r.당해사용수량)),
    당해사용금액: Math.round(toNumber(r.당해사용금액)),

    부족수량: 0,
    부족금액: 0,
  })).filter((r) => r.지점명 && r.날짜 && r.품목);
}

function fixDailyRowQtyAmountPairs(row) {
  const pairs = [
    ["전전년_판매수량", "전전년_판매금액"],
    ["전년_판매수량", "전년_판매금액"],
    ["당해_판매수량", "당해_판매금액"],

    ["전전년_폐기수량", "전전년_폐기금액"],
    ["전년_폐기수량", "전년_폐기금액"],
    ["당해_폐기수량", "당해_폐기금액"],
  ];

  pairs.forEach(([qtyField, amtField]) => {
    const qty = Math.round(toNumber(row[qtyField]));
    const amt = Math.round(toNumber(row[amtField]));

    // 수량이 없는데 금액만 있으면 당해 금액으로 옮기지 않고,
    // 같은 판매/폐기 안에서 수량 있는 연차로 흡수한다.
    if (qty <= 0 && amt > 0) {
      const isSale = qtyField.includes("판매");
      const candidates = isSale
        ? [
            ["당해_판매수량", "당해_판매금액"],
            ["전년_판매수량", "전년_판매금액"],
            ["전전년_판매수량", "전전년_판매금액"],
          ]
        : [
            ["당해_폐기수량", "당해_폐기금액"],
            ["전년_폐기수량", "전년_폐기금액"],
            ["전전년_폐기수량", "전전년_폐기금액"],
          ];

      const receiver = candidates.find(([q]) => toNumber(row[q]) > 0);

      if (receiver) {
        row[receiver[1]] = Math.round(toNumber(row[receiver[1]]) + amt);
      }

      row[amtField] = 0;
    }
  });

  // 수량은 있는데 금액이 없는 경우:
  // 같은 판매/폐기 총금액과 연차별 금액합을 비교해서
  // 남는 금액이 있으면 해당 연차에 배정한다.
  const saleAmtSum =
    toNumber(row.전전년_판매금액) +
    toNumber(row.전년_판매금액) +
    toNumber(row.당해_판매금액);

  const saleDiff = Math.round(toNumber(row.판매금액) - saleAmtSum);

  if (saleDiff > 0) {
    if (toNumber(row.당해_판매수량) > 0 && toNumber(row.당해_판매금액) <= 0) {
      row.당해_판매금액 += saleDiff;
    } else if (toNumber(row.전년_판매수량) > 0 && toNumber(row.전년_판매금액) <= 0) {
      row.전년_판매금액 += saleDiff;
    } else if (toNumber(row.전전년_판매수량) > 0 && toNumber(row.전전년_판매금액) <= 0) {
      row.전전년_판매금액 += saleDiff;
    } else if (toNumber(row.당해_판매수량) > 0) {
      row.당해_판매금액 += saleDiff;
    } else if (toNumber(row.전년_판매수량) > 0) {
      row.전년_판매금액 += saleDiff;
    } else if (toNumber(row.전전년_판매수량) > 0) {
      row.전전년_판매금액 += saleDiff;
    }
  }

  const discardAmtSum =
    toNumber(row.전전년_폐기금액) +
    toNumber(row.전년_폐기금액) +
    toNumber(row.당해_폐기금액);

  const discardDiff = Math.round(toNumber(row.폐기금액) - discardAmtSum);

  if (discardDiff > 0) {
    if (toNumber(row.당해_폐기수량) > 0 && toNumber(row.당해_폐기금액) <= 0) {
      row.당해_폐기금액 += discardDiff;
    } else if (toNumber(row.전년_폐기수량) > 0 && toNumber(row.전년_폐기금액) <= 0) {
      row.전년_폐기금액 += discardDiff;
    } else if (toNumber(row.전전년_폐기수량) > 0 && toNumber(row.전전년_폐기금액) <= 0) {
      row.전전년_폐기금액 += discardDiff;
    } else if (toNumber(row.당해_폐기수량) > 0) {
      row.당해_폐기금액 += discardDiff;
    } else if (toNumber(row.전년_폐기수량) > 0) {
      row.전년_폐기금액 += discardDiff;
    } else if (toNumber(row.전전년_폐기수량) > 0) {
      row.전전년_폐기금액 += discardDiff;
    }
  }

  return row;
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

function buildLockedMapFromDailyOrDetail(workbook) {
  const map = new Map();
  if (!workbook) return map;

  const detailMap = buildLockedDetailDailyMap(workbook);

  if (detailMap.size > 0) {
    for (const [key, r] of detailMap.entries()) {
      map.set(key, {
        판매수량: Math.round(toNumber(r.판매수량)),
        판매금액: Math.round(toNumber(r.판매금액)),
        폐기수량: Math.round(toNumber(r.폐기수량)),
        폐기금액: Math.round(toNumber(r.폐기금액)),
      });
    }
    return map;
  }

  const lockedRows = parseLockedDailyRows(workbook);
  lockedRows.forEach((r) => {
    const key = makeDailyKey(r.지점명, r.일자, r.품목군);
    map.set(key, {
      판매수량: Math.round(toNumber(r.판매수량)),
      판매금액: Math.round(toNumber(r.판매금액)),
      폐기수량: Math.round(toNumber(r.폐기수량)),
      폐기금액: Math.round(toNumber(r.폐기금액)),
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

    row.판매수량 += Math.round(toNumber(r.판매수량));
    row.판매금액 += Math.round(toNumber(r.판매금액));
    row.최종폐기 += Math.round(toNumber(r.최종폐기));
    row.폐기금액 += Math.round(toNumber(r.폐기금액));
  });

  return Array.from(map.values());
}

function buildDeltaRows(currentRows, workbookForLocked) {
  if (!workbookForLocked) return buildCurrentDailyRows(currentRows);

  const lockedMap = buildLockedMapFromDailyOrDetail(workbookForLocked);
  const currentDailyRows = buildCurrentDailyRows(currentRows);
  const deltaRows = [];

  currentDailyRows.forEach((r) => {
    const key = makeDailyKey(r.지점명, r.날짜, r.품목);
    const locked = lockedMap.get(key) || {};

    const deltaSaleQty = Math.max(0, Math.round(toNumber(r.판매수량) - toNumber(locked.판매수량)));
    const deltaSaleAmt = Math.max(0, Math.round(toNumber(r.판매금액) - toNumber(locked.판매금액)));

    const deltaDiscardQty = Math.max(0, Math.round(toNumber(r.최종폐기) - toNumber(locked.폐기수량)));
    const deltaDiscardAmt = Math.max(0, Math.round(toNumber(r.폐기금액) - toNumber(locked.폐기금액)));

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
        message: `행 ${rowNo}: 판매수량은 있는데 판매금액이 0입니다. 해당 판매건은 계산에서 제외됩니다.`,
      });
    }

    if (saleQty === 0 && saleAmt > 0) {
      issues.push({
        type: "이상치",
        message: `행 ${rowNo}: 판매금액은 있는데 판매수량이 0입니다. 해당 판매건은 계산에서 제외됩니다.`,
      });
    }

    if (discardQty > 0 && discardAmt === 0) {
      issues.push({
        type: "이상치",
        message: `행 ${rowNo}: 폐기수량은 있는데 폐기금액이 0입니다. 해당 폐기건은 계산에서 제외됩니다.`,
      });
    }

    if (discardQty === 0 && discardAmt > 0) {
      issues.push({
        type: "이상치",
        message: `행 ${rowNo}: 폐기금액은 있는데 폐기수량이 0입니다. 해당 폐기건은 계산에서 제외됩니다.`,
      });
    }
  });

  return issues;
}

function allocateSequentialFIFO(entries, stockQty, resultField, dateLimit) {
  let remainStock = Math.max(0, Math.round(toNumber(stockQty)));
  let usedStock = 0;

  if (remainStock <= 0) return 0;

  for (const e of entries) {
    if (dateLimit && normalizeText(e.date) > dateLimit) continue;
    if (toNumber(e.remainQty) <= 0) continue;
    if (remainStock <= 0) break;

    const useQty = Math.min(Math.round(toNumber(e.remainQty)), remainStock);

    e[resultField] = Math.round(toNumber(e[resultField])) + useQty;
    e.remainQty = Math.round(toNumber(e.remainQty)) - useQty;

    remainStock -= useQty;
    usedStock += useQty;
  }

  return usedStock;
}

function distributeRowAmountByAllocatedQty(entries) {
  entries.forEach((e) => {
    const totalAmt = Math.round(toNumber(e.amt));

    e.prev2Amt = 0;
    e.prevAmt = 0;
    e.currentAmt = 0;
    e.shortageAmt = 0;

    const parts = [
      {
        name: "전전년",
        qtyField: "prev2Qty",
        amtField: "prev2Amt",
        qty: Math.round(toNumber(e.prev2Qty)),
      },
      {
        name: "전년",
        qtyField: "prevQty",
        amtField: "prevAmt",
        qty: Math.round(toNumber(e.prevQty)),
      },
      {
        name: "당해",
        qtyField: "currentQty",
        amtField: "currentAmt",
        qty: Math.round(toNumber(e.currentQty)),
      },
    ].filter((p) => p.qty > 0);

    distributeAmountToParts(totalAmt, parts, e);
  });
}

function fixQtyAmountPair(entries) {
  entries.forEach((e) => {
    const totalAmt = Math.round(toNumber(e.amt));

    const pairs = [
      { qtyField: "prev2Qty", amtField: "prev2Amt" },
      { qtyField: "prevQty", amtField: "prevAmt" },
      { qtyField: "currentQty", amtField: "currentAmt" },
    ];

    pairs.forEach((p) => {
      if (toNumber(e[p.qtyField]) <= 0) e[p.amtField] = 0;
    });

    const qtyParts = pairs
      .map((p) => ({ ...p, qty: Math.round(toNumber(e[p.qtyField])) }))
      .filter((p) => p.qty > 0);

    if (!qtyParts.length || totalAmt <= 0) {
      pairs.forEach((p) => {
        e[p.amtField] = 0;
      });
      return;
    }

    const hasPairError = qtyParts.some((p) => toNumber(e[p.amtField]) <= 0);
    const allocated = pairs.reduce((sum, p) => sum + Math.round(toNumber(e[p.amtField])), 0);

    if (hasPairError || allocated !== totalAmt) {
      distributeAmountToParts(totalAmt, qtyParts, e);
    }
  });
}

function allocateGroupPeriod(entries, stock) {
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

    // 기존 제출 이후 증가분은 기존 연차별 배분을 건드리지 않고
    // 감사 일관성을 위해 전부 당해 사용으로 처리한다.
    e.currentQty = Math.max(0, Math.round(toNumber(e.qty)));
    e.currentAmt = Math.max(0, Math.round(toNumber(e.amt)));

    e.remainQty = 0;
    e.shortageQty = 0;
    e.shortageAmt = 0;
  });
}

  allocateSequentialFIFO(
    entries,
    toNumber(stock?.전전년수량),
    "prev2Qty",
    prev2Cutoff
  );

  allocateSequentialFIFO(
    entries,
    toNumber(stock?.전년수량),
    "prevQty",
    prevCutoff
  );

  entries.forEach((e) => {
    e.currentQty = Math.max(0, Math.round(toNumber(e.remainQty)));
    e.remainQty = 0;
    e.shortageQty = 0;
  });

  distributeRowAmountByAllocatedQty(entries);
  fixQtyAmountPair(entries);
}

function mergeLockedRowsIntoMergedMap(mergedMap, workbook) {
  if (!workbook) return;

  const detailDailyMap = buildLockedDetailDailyMap(workbook);

  if (detailDailyMap.size > 0) {
    for (const [key, detailRow] of detailDailyMap.entries()) {
      mergedMap.set(key, { ...detailRow });
    }
    return;
  }

  const lockedRows = parseLockedDailyRows(workbook);

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
  });
}

function hasPairError(qty, amt) {
  qty = Math.round(toNumber(qty));
  amt = Math.round(toNumber(amt));

  return (qty > 0 && amt <= 0) || (qty <= 0 && amt > 0);
}

function buildValidationRows(dailyRows) {
  return dailyRows.map((row) => {
    const periodQtySum =
      toNumber(row.전전년_판매수량) +
      toNumber(row.전전년_폐기수량) +
      toNumber(row.전년_판매수량) +
      toNumber(row.전년_폐기수량) +
      toNumber(row.당해_판매수량) +
      toNumber(row.당해_폐기수량);

    const periodAmtSum =
      toNumber(row.전전년_판매금액) +
      toNumber(row.전전년_폐기금액) +
      toNumber(row.전년_판매금액) +
      toNumber(row.전년_폐기금액) +
      toNumber(row.당해_판매금액) +
      toNumber(row.당해_폐기금액);

    return {
      지점명: row.지점명,
      날짜: row.일자,
      품목: row.품목군,

      총사용수량: row.총사용수량,
      총사용금액: row.총사용금액,

      연차배분수량합: periodQtySum,
      수량차이: row.총사용수량 - periodQtySum,

      연차배분금액합: periodAmtSum,
      금액차이: row.총사용금액 - periodAmtSum,

      전전년판매쌍오류: hasPairError(row.전전년_판매수량, row.전전년_판매금액) ? "오류" : "",
      전년판매쌍오류: hasPairError(row.전년_판매수량, row.전년_판매금액) ? "오류" : "",
      당해판매쌍오류: hasPairError(row.당해_판매수량, row.당해_판매금액) ? "오류" : "",

      전전년폐기쌍오류: hasPairError(row.전전년_폐기수량, row.전전년_폐기금액) ? "오류" : "",
      전년폐기쌍오류: hasPairError(row.전년_폐기수량, row.전년_폐기금액) ? "오류" : "",
      당해폐기쌍오류: hasPairError(row.당해_폐기수량, row.당해_폐기금액) ? "오류" : "",

      부족수량: row.부족수량,
      부족금액: row.부족금액,

      전전년재고초과: "",
      전년재고초과: "",
    };
  });
}

function buildAllocationStockMap(stockMap, lockedWorkbook) {
  const lockedUseMap = buildLockedUseMap(lockedWorkbook);
  const allocationStockMap = new Map();

  for (const [key, stock] of stockMap.entries()) {
    const locked = lockedUseMap.get(key) || {};

    const remainPrev2Qty = Math.max(
      0,
      Math.round(toNumber(stock.전전년수량) - toNumber(locked.prev2Qty))
    );

    const remainPrevQty = Math.max(
      0,
      Math.round(toNumber(stock.전년수량) - toNumber(locked.prevQty))
    );

    allocationStockMap.set(key, {
      ...stock,

      // 새 증가분 계산에는 기존 제출에서 이미 사용한 재고를 차감한 잔여수량만 사용
      전전년수량: remainPrev2Qty,
      전년수량: remainPrevQty,
    });
  }

  return allocationStockMap;
}

function buildLockedUseMap(workbook) {
  const map = new Map();
  if (!workbook) return map;

  const lockedSalesRows = parseLockedDetailRows(workbook, "판매자동소진", "판매");
  const lockedDiscardRows = parseLockedDetailRows(workbook, "폐기자동소진", "폐기");

  function ensure(branch, item) {
    const key = makeStockKey(branch, item);

    if (!map.has(key)) {
      map.set(key, {
        prev2Qty: 0,
        prev2Amt: 0,
        prevQty: 0,
        prevAmt: 0,
      });
    }

    return map.get(key);
  }

  [...lockedSalesRows, ...lockedDiscardRows].forEach((r) => {
    const x = ensure(r.지점명, r.품목);

    x.prev2Qty += toNumber(r.전전년사용수량);
    x.prev2Amt += toNumber(r.전전년사용금액);

    x.prevQty += toNumber(r.전년사용수량);
    x.prevAmt += toNumber(r.전년사용금액);
  });

  if (map.size > 0) return map;

  const lockedDailyRows = parseLockedDailyRows(workbook);

  lockedDailyRows.forEach((r) => {
    const x = ensure(r.지점명, r.품목군);

    x.prev2Qty += toNumber(r.전전년_판매수량) + toNumber(r.전전년_폐기수량);
    x.prev2Amt += toNumber(r.전전년_판매금액) + toNumber(r.전전년_폐기금액);

    x.prevQty += toNumber(r.전년_판매수량) + toNumber(r.전년_폐기수량);
    x.prevAmt += toNumber(r.전년_판매금액) + toNumber(r.전년_폐기금액);
  });

  return map;
}

function buildStockBalanceRows(stockMap, groupMap, workbook) {
  const lockedUseMap = buildLockedUseMap(workbook);
  const rows = [];

  for (const [key, stock] of stockMap.entries()) {
    const entries = groupMap.get(key) || [];
    const locked = lockedUseMap.get(key) || {};

    const newPrev2UsedQty = entries.reduce((s, e) => s + toNumber(e.prev2Qty), 0);
    const newPrev2UsedAmt = entries.reduce((s, e) => s + toNumber(e.prev2Amt), 0);

    const newPrevUsedQty = entries.reduce((s, e) => s + toNumber(e.prevQty), 0);
    const newPrevUsedAmt = entries.reduce((s, e) => s + toNumber(e.prevAmt), 0);

    const prev2UsedQty = toNumber(locked.prev2Qty) + newPrev2UsedQty;
    const prev2UsedAmt = toNumber(locked.prev2Amt) + newPrev2UsedAmt;

    const prevUsedQty = toNumber(locked.prevQty) + newPrevUsedQty;
    const prevUsedAmt = toNumber(locked.prevAmt) + newPrevUsedAmt;

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
  const rawRows = parseHorizontal2025Sheet(workbook);
  const filteredRows = filterRowsByFinalCutoff(rawRows);

  const mainRows = buildDeltaRows(filteredRows, lockedWorkbook);

  const prev2Rows = parseInventorySheetFixed(workbook, PREV2_SHEET);
  const prevRows = parseInventorySheetFixed(workbook, PREV_SHEET);
  const stockMap = buildOpeningStocks(prev2Rows, prevRows);

  // 중요:
  // 기존 제출파일에서 이미 사용한 전전년/전년 재고를 차감한 뒤
  // 새 증가분에 대해서만 FIFO 계산한다.
  const allocationStockMap = buildAllocationStockMap(stockMap, lockedWorkbook);

  const anomalyIssues = validateAnomalies(mainRows);

  const groupMap = new Map();

  mainRows.forEach((row) => {
    const branch = normalizeText(row.지점명);
    const item = normalizeText(row.품목);
    const date = normalizeText(row.날짜);

    if (!branch || !item || !date) return;

    const key = makeStockKey(branch, item);
    if (!groupMap.has(key)) groupMap.set(key, []);

    const saleQty = Math.round(toNumber(row.판매수량));
    const saleAmt = Math.round(toNumber(row.판매금액));

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

    const discardQty = Math.round(toNumber(row.최종폐기));
    const discardAmt = Math.round(toNumber(row.폐기금액));

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
    allocateGroupPeriod(entries, allocationStockMap.get(key));
}

  const mergedMap = new Map();

  mergeLockedRowsIntoMergedMap(mergedMap, lockedWorkbook);

  const lockedSalesRows = parseLockedDetailRows(lockedWorkbook, "판매자동소진", "판매");
  const lockedDiscardRows = parseLockedDetailRows(lockedWorkbook, "폐기자동소진", "폐기");

  const salesRows = [...lockedSalesRows];
  const discardRows = [...lockedDiscardRows];

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
          부족수량: 0,
          부족금액: 0,
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
          부족수량: 0,
          부족금액: 0,
        });
      }
    });
  }

  const mergedDailyRows = sortRowsByBusinessOrder(Array.from(mergedMap.values())).map((row) => {
  row.판매수량 = Math.round(toNumber(row.판매수량));
  row.판매금액 = Math.round(toNumber(row.판매금액));
  row.폐기수량 = Math.round(toNumber(row.폐기수량));
  row.폐기금액 = Math.round(toNumber(row.폐기금액));

  row.총사용수량 = row.판매수량 + row.폐기수량;
  row.총사용금액 = row.판매금액 + row.폐기금액;

  row.부족수량 = 0;
  row.부족금액 = 0;

  fixDailyRowQtyAmountPairs(row);

  return row;
});

  const validationRows = sortRowsByBusinessOrder(buildValidationRows(mergedDailyRows));
  const stockBalanceRows = buildStockBalanceRows(stockMap, groupMap, lockedWorkbook);
  const remainStockRows = buildOnlyRemainStockRows(stockBalanceRows);

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

function renderIssues(result) {
  const container = document.getElementById("issues");
  if (!container) return;

  container.innerHTML = "";

  const validationErrors = [];

  result.validationRows.forEach((r) => {
    const hasError =
      toNumber(r.수량차이) !== 0 ||
      toNumber(r.금액차이) !== 0 ||
      r.전전년판매쌍오류 ||
      r.전년판매쌍오류 ||
      r.당해판매쌍오류 ||
      r.전전년폐기쌍오류 ||
      r.전년폐기쌍오류 ||
      r.당해폐기쌍오류;

    if (hasError) {
      validationErrors.push({
        type: "검증오류",
        message: `${r.지점명} / ${r.날짜} / ${r.품목}: 수량차이 ${r.수량차이}, 금액차이 ${r.금액차이}, 쌍오류 있음`,
      });
    }
  });

  const issues = [
    ...result.schemaErrors.map((m) => ({ type: "스키마오류", message: m })),
    ...result.anomalyIssues,
    ...validationErrors,
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

  const validationIssueCount = result.validationRows.filter((r) => {
    return (
      toNumber(r.수량차이) !== 0 ||
      toNumber(r.금액차이) !== 0 ||
      r.전전년판매쌍오류 ||
      r.전년판매쌍오류 ||
      r.당해판매쌍오류 ||
      r.전전년폐기쌍오류 ||
      r.전년폐기쌍오류 ||
      r.당해폐기쌍오류
    );
  }).length;

  const issueCount = document.getElementById("issueCount");
  if (issueCount) {
    issueCount.textContent =
      result.schemaErrors.length + result.anomalyIssues.length + validationIssueCount;
  }

  const shortageCount = document.getElementById("shortageCount");
  if (shortageCount) shortageCount.textContent = 0;

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

  result.validationRows.forEach((r) => {
    const errors = [];

    if (toNumber(r.수량차이) !== 0) errors.push(`수량차이 ${r.수량차이}`);
    if (toNumber(r.금액차이) !== 0) errors.push(`금액차이 ${r.금액차이}`);

    [
      "전전년판매쌍오류",
      "전년판매쌍오류",
      "당해판매쌍오류",
      "전전년폐기쌍오류",
      "전년폐기쌍오류",
      "당해폐기쌍오류",
    ].forEach((field) => {
      if (r[field]) errors.push(field);
    });

    if (errors.length) {
      issueRows.push({
        유형: "검증오류",
        지점명: r.지점명,
        날짜: r.날짜,
        품목: r.품목,
        내용: errors.join(", "),
      });
    }
  });

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
