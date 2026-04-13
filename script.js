(() => {
  "use strict";

  /******************************************************************
   * 자동소진 웹프로그램 - script.js 전체 교체본
   *
   * 필수 시트명
   * - 2025
   * - 전년재고_DB
   * - 전전년재고_DB
   *
   * 2025 시트 구조
   * - A열 = 날짜
   * - B열 = 품목
   * - 9행 = 지점명
   * - 10행 = 항목명
   * - 11행부터 데이터
   *
   * 지점 블록에서 반드시 읽는 항목명
   * - 판매수량
   * - 판매금액(사용명세서)
   * - 최종폐기
   * - 금액
   *
   * 재고 시트 구조 (고정)
   * - A열 = 지점명
   * - B열 = 품목
   * - C열 = 일자
   * - D열 = 수량
   * - E열 = 금액
   ******************************************************************/

  // =========================
  // 환경 체크
  // =========================
  if (typeof XLSX === "undefined") {
    alert("XLSX 라이브러리(sheetjs)가 먼저 로드되어야 합니다.");
    return;
  }

  // =========================
  // 품목 정렬 순서
  // =========================
  const ITEM_ORDER = [
    "의류",
    "잡화",
    "생활",
    "문화",
    "건강미용",
    "식품",
    "기증파트너",
  ];

  const ITEM_ORDER_MAP = new Map(ITEM_ORDER.map((name, idx) => [name, idx]));

  // =========================
  // 상태 변수
  // =========================
  let workbook = null;
  let finalRows = [];
  let validationErrors = [];
  let previewRows = [];
  let currentPreviewLimit = 300;
  let downloadReady = false;
  let currentFileName = "자동소진_결과.xlsx";

  // =========================
  // DOM 유틸
  // =========================
  function qs(id) {
    return document.getElementById(id);
  }

  function ensureElement(id, tag, parent, html = "") {
    let el = qs(id);
    if (!el) {
      el = document.createElement(tag);
      el.id = id;
      if (html) el.innerHTML = html;
      parent.appendChild(el);
    }
    return el;
  }

  function initUI() {
    const root = document.body;

    const app = ensureElement("appAutoFifo", "div", root);
    app.style.maxWidth = "1400px";
    app.style.margin = "20px auto";
    app.style.padding = "16px";
    app.style.fontFamily = "Arial, sans-serif";

    if (!qs("fileInput")) {
      const wrapper = document.createElement("div");
      wrapper.style.display = "flex";
      wrapper.style.flexWrap = "wrap";
      wrapper.style.gap = "8px";
      wrapper.style.alignItems = "center";
      wrapper.style.marginBottom = "12px";

      wrapper.innerHTML = `
        <input type="file" id="fileInput" accept=".xlsx,.xls" />
        <button id="processBtn" type="button">처리 시작</button>
        <button id="downloadBtn" type="button" disabled>엑셀 다운로드</button>
        <label style="margin-left:8px;">지점 필터</label>
        <select id="branchFilter">
          <option value="">전체</option>
        </select>
        <button id="moreBtn" type="button">더 보기</button>
      `;
      app.appendChild(wrapper);
    }

    const status = ensureElement("status", "div", app);
    status.style.marginBottom = "12px";
    status.style.padding = "10px";
    status.style.background = "#f5f5f5";
    status.style.border = "1px solid #ddd";
    status.style.whiteSpace = "pre-wrap";

    const summary = ensureElement("summaryBox", "div", app);
    summary.style.marginBottom = "12px";
    summary.style.padding = "10px";
    summary.style.background = "#fafafa";
    summary.style.border = "1px solid #ddd";
    summary.style.whiteSpace = "pre-wrap";

    const errorBox = ensureElement("errorBox", "div", app);
    errorBox.style.marginBottom = "12px";
    errorBox.style.padding = "10px";
    errorBox.style.background = "#fff7f7";
    errorBox.style.border = "1px solid #f0caca";
    errorBox.style.whiteSpace = "pre-wrap";
    errorBox.style.maxHeight = "220px";
    errorBox.style.overflow = "auto";

    const previewInfo = ensureElement("previewInfo", "div", app);
    previewInfo.style.margin = "8px 0";

    const tableWrap = ensureElement("tableWrap", "div", app);
    tableWrap.style.overflow = "auto";
    tableWrap.style.border = "1px solid #ddd";
    tableWrap.style.maxHeight = "650px";

    if (!qs("previewTable")) {
      const table = document.createElement("table");
      table.id = "previewTable";
      table.style.width = "100%";
      table.style.borderCollapse = "collapse";
      table.style.fontSize = "12px";

      const thead = document.createElement("thead");
      thead.innerHTML = `
        <tr id="previewHeadRow"></tr>
      `;

      const tbody = document.createElement("tbody");
      tbody.id = "previewBody";

      table.appendChild(thead);
      table.appendChild(tbody);
      tableWrap.appendChild(table);
    }

    const headRow = qs("previewHeadRow");
    if (headRow && !headRow.dataset.ready) {
      const headers = OUTPUT_HEADERS;
      headRow.innerHTML = headers
        .map(
          (h) =>
            `<th style="position:sticky; top:0; background:#eee; border:1px solid #ccc; padding:6px; white-space:nowrap;">${escapeHtml(
              h
            )}</th>`
        )
        .join("");
      headRow.dataset.ready = "1";
    }
  }

  function setStatus(msg) {
    const el = qs("status");
    if (el) el.textContent = msg;
  }

  function setSummary(msg) {
    const el = qs("summaryBox");
    if (el) el.textContent = msg;
  }

  function setErrors(errors) {
    const el = qs("errorBox");
    if (!el) return;
    if (!errors || errors.length === 0) {
      el.textContent = "검증 오류 없음";
      return;
    }

    const preview = errors.slice(0, 300);
    let text = `검증 오류 ${errors.length}건\n\n`;
    text += preview.map((x, i) => `${i + 1}. ${x}`).join("\n");
    if (errors.length > preview.length) {
      text += `\n\n...외 ${errors.length - preview.length}건`;
    }
    el.textContent = text;
  }

  function escapeHtml(str) {
    return String(str ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  // =========================
  // 헤더 정의
  // =========================
  const OUTPUT_HEADERS = [
    "지점명",
    "일자",
    "품목군",
    "판매수량",
    "판매금액",
    "폐기수량",
    "폐기금액",
    "총사용수량",
    "총사용금액",
    "전전년_판매수량",
    "전전년_판매금액",
    "전전년_폐기수량",
    "전전년_폐기금액",
    "전년_판매수량",
    "전년_판매금액",
    "전년_폐기수량",
    "전년_폐기금액",
    "당해_판매수량",
    "당해_판매금액",
    "당해_폐기수량",
    "당해_폐기금액",
    "부족수량",
    "부족금액",
  ];

  // =========================
  // 기본 유틸
  // =========================
  function toNumber(value) {
    if (value === null || value === undefined || value === "") return 0;
    if (typeof value === "number") {
      return Number.isFinite(value) ? value : 0;
    }
    const cleaned = String(value).replace(/,/g, "").trim();
    if (cleaned === "") return 0;
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : 0;
  }

  function toStringSafe(value) {
    return String(value ?? "").trim();
  }

  function normalizeItemName(value) {
    const v = toStringSafe(value);
    return v;
  }

  function excelDateToJSDate(value) {
    if (value instanceof Date) return value;
    if (typeof value === "number") {
      const date = XLSX.SSF.parse_date_code(value);
      if (!date) return null;
      return new Date(date.y, date.m - 1, date.d);
    }
    if (typeof value === "string" && value.trim()) {
      const s = value.trim().replace(/\./g, "-").replace(/\//g, "-");
      const d = new Date(s);
      if (!isNaN(d.getTime())) return d;
    }
    return null;
  }

  function formatDate(value) {
    const d = excelDateToJSDate(value);
    if (!d) return toStringSafe(value);
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  }

  function parseDateValue(value) {
    const d = excelDateToJSDate(value);
    if (!d) return null;
    return d.getTime();
  }

  function safeDiv(num, den) {
    if (!den) return 0;
    return num / den;
  }

  function round2(n) {
    return Math.round((n + Number.EPSILON) * 100) / 100;
  }

  function createKey(branch, date, item) {
    return `${branch}|||${date}|||${item}`;
  }

  function inventoryKey(branch, item) {
    return `${branch}|||${item}`;
  }

  function sortRows(rows) {
    rows.sort((a, b) => {
      const b1 = a["지점명"].localeCompare(b["지점명"], "ko");
      if (b1 !== 0) return b1;

      const d1 = parseDateValue(a["일자"]) ?? Number.MAX_SAFE_INTEGER;
      const d2 = parseDateValue(b["일자"]) ?? Number.MAX_SAFE_INTEGER;
      if (d1 !== d2) return d1 - d2;

      const i1 = ITEM_ORDER_MAP.has(a["품목군"])
        ? ITEM_ORDER_MAP.get(a["품목군"])
        : 999;
      const i2 = ITEM_ORDER_MAP.has(b["품목군"])
        ? ITEM_ORDER_MAP.get(b["품목군"])
        : 999;
      if (i1 !== i2) return i1 - i2;

      return 0;
    });
  }

  // =========================
  // 2025 시트 파싱
  // =========================
  function parse2025Sheet(ws) {
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const resultMap = new Map();

    const BRANCH_ROW = 8; // 엑셀 9행
    const HEADER_ROW = 9; // 엑셀 10행
    const START_DATA_ROW = 10; // 엑셀 11행

    const COL_DATE = 0; // A
    const COL_ITEM = 1; // B

    const REQUIRED_HEADERS = [
      "판매수량",
      "판매금액(사용명세서)",
      "최종폐기",
      "금액",
    ];

    // 1) 지점 블록 분석
    // 9행(인덱스 8)에 지점명이 있고, 10행(인덱스 9)에 항목명이 반복된다고 가정
    // 같은 지점명 아래 여러 열이 있으면 그 범위에서 필요한 헤더를 찾는다.
    const branchColumns = new Map();

    for (let c = 2; c <= range.e.c; c++) {
      const branchCell = ws[XLSX.utils.encode_cell({ r: BRANCH_ROW, c })];
      const headerCell = ws[XLSX.utils.encode_cell({ r: HEADER_ROW, c })];

      const branchName = toStringSafe(branchCell ? branchCell.v : "");
      const headerName = toStringSafe(headerCell ? headerCell.v : "");

      if (!branchName) continue;
      if (!branchColumns.has(branchName)) {
        branchColumns.set(branchName, []);
      }
      branchColumns.get(branchName).push({
        col: c,
        header: headerName,
      });
    }

    // 2) 지점별 헤더 매핑
    const branchHeaderMap = new Map();

    for (const [branchName, cols] of branchColumns.entries()) {
      const map = {};
      for (const need of REQUIRED_HEADERS) {
        const found = cols.find((x) => toStringSafe(x.header) === need);
        if (found) {
          map[need] = found.col;
        }
      }

      const missing = REQUIRED_HEADERS.filter((h) => map[h] === undefined);
      if (missing.length > 0) {
        validationErrors.push(
          `[2025 시트] 지점 '${branchName}' 에서 필수 헤더 누락: ${missing.join(", ")}`
        );
      }

      branchHeaderMap.set(branchName, map);
    }

    // 3) 데이터 행 파싱
    for (let r = START_DATA_ROW; r <= range.e.r; r++) {
      const dateCell = ws[XLSX.utils.encode_cell({ r, c: COL_DATE })];
      const itemCell = ws[XLSX.utils.encode_cell({ r, c: COL_ITEM })];

      const rawDate = dateCell ? dateCell.v : "";
      const rawItem = itemCell ? itemCell.v : "";

      const date = formatDate(rawDate);
      const item = normalizeItemName(rawItem);

      if (!date && !item) continue;
      if (!date || !item) continue;

      for (const [branchName, headerMap] of branchHeaderMap.entries()) {
        const colSaleQty = headerMap["판매수량"];
        const colSaleAmt = headerMap["판매금액(사용명세서)"];
        const colDiscardQty = headerMap["최종폐기"];
        const colDiscardAmt = headerMap["금액"];

        if (
          colSaleQty === undefined ||
          colSaleAmt === undefined ||
          colDiscardQty === undefined ||
          colDiscardAmt === undefined
        ) {
          continue;
        }

        const saleQtyCell = ws[XLSX.utils.encode_cell({ r, c: colSaleQty })];
        const saleAmtCell = ws[XLSX.utils.encode_cell({ r, c: colSaleAmt })];
        const discardQtyCell = ws[XLSX.utils.encode_cell({ r, c: colDiscardQty })];
        const discardAmtCell = ws[XLSX.utils.encode_cell({ r, c: colDiscardAmt })];

        const saleQty = toNumber(saleQtyCell ? saleQtyCell.v : 0);
        const saleAmt = toNumber(saleAmtCell ? saleAmtCell.v : 0);
        const discardQty = toNumber(discardQtyCell ? discardQtyCell.v : 0);
        const discardAmt = toNumber(discardAmtCell ? discardAmtCell.v : 0);

        // 완전 빈 행은 스킵
        if (
          saleQty === 0 &&
          saleAmt === 0 &&
          discardQty === 0 &&
          discardAmt === 0
        ) {
          continue;
        }

        // 검증 - 판매
        if (saleQty > 0 && saleAmt === 0) {
          validationErrors.push(
            `[검증오류] ${branchName} / ${date} / ${item} : 판매수량 > 0 인데 판매금액 = 0`
          );
        }
        if (saleQty === 0 && saleAmt > 0) {
          validationErrors.push(
            `[검증오류] ${branchName} / ${date} / ${item} : 판매수량 = 0 인데 판매금액 > 0`
          );
        }

        // 검증 - 폐기
        if (discardQty > 0 && discardAmt === 0) {
          validationErrors.push(
            `[검증오류] ${branchName} / ${date} / ${item} : 폐기수량 > 0 인데 폐기금액 = 0`
          );
        }
        if (discardQty === 0 && discardAmt > 0) {
          validationErrors.push(
            `[검증오류] ${branchName} / ${date} / ${item} : 폐기수량 = 0 인데 폐기금액 > 0`
          );
        }

        const key = createKey(branchName, date, item);

        if (!resultMap.has(key)) {
          resultMap.set(key, {
            지점명: branchName,
            일자: date,
            품목군: item,
            판매수량: 0,
            판매금액: 0,
            폐기수량: 0,
            폐기금액: 0,
          });
        }

        const row = resultMap.get(key);
        row["판매수량"] += saleQty;
        row["판매금액"] += saleAmt;
        row["폐기수량"] += discardQty;
        row["폐기금액"] += discardAmt;
      }
    }

    return Array.from(resultMap.values());
  }

  // =========================
  // 재고 시트 파싱 (A:E 고정)
  // =========================
  function parseInventorySheet(ws, yearLabel) {
    const rows = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      raw: true,
      defval: "",
    });

    const result = [];

    // 0행부터 훑되, 실제 데이터 여부로 판별
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i] || [];
      const branch = toStringSafe(row[0]);
      const item = normalizeItemName(row[1]);
      const date = formatDate(row[2]);
      const qty = toNumber(row[3]);
      const amt = toNumber(row[4]);

      // 헤더행/빈행 제외
      const joined = row.map((x) => toStringSafe(x)).join("|");
      if (!branch && !item && !date && qty === 0 && amt === 0) continue;
      if (
        joined.includes("지점명") &&
        joined.includes("품목") &&
        joined.includes("수량")
      ) {
        continue;
      }

      if (!branch || !item) continue;

      if (qty < 0 || amt < 0) {
        validationErrors.push(
          `[재고시트오류] ${yearLabel} / ${branch} / ${item} / ${date} : 수량 또는 금액이 음수입니다.`
        );
      }

      result.push({
        yearLabel,
        지점명: branch,
        품목군: item,
        일자: date,
        수량: qty,
        금액: amt,
      });
    }

    return result;
  }

  // =========================
  // 재고 버킷 생성
  // =========================
  function buildInventoryBuckets(prev2Rows, prevRows) {
    const buckets = new Map();

    function pushRow(sourceYear, row) {
      const key = inventoryKey(row["지점명"], row["품목군"]);
      if (!buckets.has(key)) {
        buckets.set(key, {
          prev2: [],
          prev1: [],
        });
      }

      const target = buckets.get(key);
      const arr = sourceYear === "전전년" ? target.prev2 : target.prev1;

      arr.push({
        date: row["일자"],
        sortDate: parseDateValue(row["일자"]) ?? Number.MAX_SAFE_INTEGER,
        qtyRemaining: toNumber(row["수량"]),
        amtRemaining: toNumber(row["금액"]),
      });
    }

    prev2Rows.forEach((r) => pushRow("전전년", r));
    prevRows.forEach((r) => pushRow("전년", r));

    for (const bucket of buckets.values()) {
      bucket.prev2.sort((a, b) => a.sortDate - b.sortDate);
      bucket.prev1.sort((a, b) => a.sortDate - b.sortDate);
    }

    return buckets;
  }

  // =========================
  // FIFO 배분 로직
  // =========================
  function allocateFromBuckets(lines2025, inventoryBuckets) {
    const result = [];

    const sortedLines = [...lines2025];
    sortRows(sortedLines);

    for (const line of sortedLines) {
      const branch = line["지점명"];
      const item = line["품목군"];

      const saleQty = toNumber(line["판매수량"]);
      const saleAmt = toNumber(line["판매금액"]);
      const discardQty = toNumber(line["폐기수량"]);
      const discardAmt = toNumber(line["폐기금액"]);

      const totalQty = saleQty + discardQty;
      const totalAmt = saleAmt + discardAmt;

      const invKey = inventoryKey(branch, item);
      const bucket = inventoryBuckets.get(invKey) || {
        prev2: [],
        prev1: [],
      };

      // 판매/폐기를 순서대로 각각 FIFO 배분
      // 전전년 -> 전년 -> 남는 부분은 당해
      // 그 뒤에도 남으면 부족
      const saleAlloc = allocateUsage(
        saleQty,
        saleAmt,
        bucket.prev2,
        bucket.prev1
      );

      const discardAlloc = allocateUsage(
        discardQty,
        discardAmt,
        bucket.prev2,
        bucket.prev1
      );

      const row = {
        지점명: branch,
        일자: line["일자"],
        품목군: item,
        판매수량: saleQty,
        판매금액: saleAmt,
        폐기수량: discardQty,
        폐기금액: discardAmt,
        총사용수량: totalQty,
        총사용금액: totalAmt,

        전전년_판매수량: round2(saleAlloc.prev2Qty),
        전전년_판매금액: round2(saleAlloc.prev2Amt),
        전전년_폐기수량: round2(discardAlloc.prev2Qty),
        전전년_폐기금액: round2(discardAlloc.prev2Amt),

        전년_판매수량: round2(saleAlloc.prev1Qty),
        전년_판매금액: round2(saleAlloc.prev1Amt),
        전년_폐기수량: round2(discardAlloc.prev1Qty),
        전년_폐기금액: round2(discardAlloc.prev1Amt),

        당해_판매수량: round2(saleAlloc.currentQty),
        당해_판매금액: round2(saleAlloc.currentAmt),
        당해_폐기수량: round2(discardAlloc.currentQty),
        당해_폐기금액: round2(discardAlloc.currentAmt),

        부족수량: round2(saleAlloc.shortQty + discardAlloc.shortQty),
        부족금액: round2(saleAlloc.shortAmt + discardAlloc.shortAmt),
      };

      // 합계 검증
      const qtyCheck =
        row["전전년_판매수량"] +
        row["전전년_폐기수량"] +
        row["전년_판매수량"] +
        row["전년_폐기수량"] +
        row["당해_판매수량"] +
        row["당해_폐기수량"] +
        row["부족수량"];

      const amtCheck =
        row["전전년_판매금액"] +
        row["전전년_폐기금액"] +
        row["전년_판매금액"] +
        row["전년_폐기금액"] +
        row["당해_판매금액"] +
        row["당해_폐기금액"] +
        row["부족금액"];

      if (Math.abs(qtyCheck - totalQty) > 0.01) {
        validationErrors.push(
          `[합계오류-수량] ${branch} / ${line["일자"]} / ${item} : 배분합(${qtyCheck}) != 총사용수량(${totalQty})`
        );
      }

      if (Math.abs(amtCheck - totalAmt) > 0.01) {
        validationErrors.push(
          `[합계오류-금액] ${branch} / ${line["일자"]} / ${item} : 배분합(${amtCheck}) != 총사용금액(${totalAmt})`
        );
      }

      result.push(row);
    }

    return result;
  }

  function allocateUsage(reqQty, reqAmt, prev2Arr, prev1Arr) {
    let remainingQty = toNumber(reqQty);
    let remainingAmt = toNumber(reqAmt);

    const out = {
      prev2Qty: 0,
      prev2Amt: 0,
      prev1Qty: 0,
      prev1Amt: 0,
      currentQty: 0,
      currentAmt: 0,
      shortQty: 0,
      shortAmt: 0,
    };

    // 1) 전전년 배분
    const a1 = consumeFromInventory(prev2Arr, remainingQty, remainingAmt);
    out.prev2Qty += a1.usedQty;
    out.prev2Amt += a1.usedAmt;
    remainingQty = a1.remainingQty;
    remainingAmt = a1.remainingAmt;

    // 2) 전년 배분
    const a2 = consumeFromInventory(prev1Arr, remainingQty, remainingAmt);
    out.prev1Qty += a2.usedQty;
    out.prev1Amt += a2.usedAmt;
    remainingQty = a2.remainingQty;
    remainingAmt = a2.remainingAmt;

    // 3) 당해 배분
    // 규칙:
    // 전전년, 전년을 다 쓰고도 남은 사용분은 우선 당해로 흡수
    // 다만 "재고 총합보다 총사용이 큰 경우"에는 부족이 발생해야 하므로
    // 당해는 무한정 흡수하는 개념이 아니라
    // 현재 사용 요청에서 "전전년+전년으로 커버 못 한 수량/금액" 중
    // 서로 연동되는 범위만 당해로 보고,
    // 끝까지 맞지 않으면 부족 처리
    //
    // 사용자 요구 취지상:
    // - 정상 데이터는 대부분 당해 흡수
    // - 그러나 전전년DB보다 총사용이 더 적은/큰 불일치 등에서 부족이 가려지면 안 됨
    //
    // 따라서 여기서는
    // "남은 수량"과 "남은 금액"을 그대로 당해에 넣되,
    // 수량/금액 일관성을 깨는 경우는 부족으로 돌린다.

    if (remainingQty > 0 || remainingAmt > 0) {
      if (remainingQty > 0 && remainingAmt > 0) {
        out.currentQty = remainingQty;
        out.currentAmt = remainingAmt;
        remainingQty = 0;
        remainingAmt = 0;
      } else {
        // 수량 또는 금액 한쪽만 남는다면 정상 흡수가 아니라 부족으로 처리
        out.shortQty = remainingQty;
        out.shortAmt = remainingAmt;
        remainingQty = 0;
        remainingAmt = 0;
      }
    }

    // 4) 마지막 부족
    // 현재 구조상 위에서 처리되지만 안전장치로 둠
    if (remainingQty > 0 || remainingAmt > 0) {
      out.shortQty += remainingQty;
      out.shortAmt += remainingAmt;
    }

    return out;
  }

  function consumeFromInventory(invArr, reqQty, reqAmt) {
    let remainingQty = toNumber(reqQty);
    let remainingAmt = toNumber(reqAmt);

    let usedQty = 0;
    let usedAmt = 0;

    if ((remainingQty <= 0 && remainingAmt <= 0) || !Array.isArray(invArr) || invArr.length === 0) {
      return {
        usedQty,
        usedAmt,
        remainingQty,
        remainingAmt,
      };
    }

    for (const lot of invArr) {
      if (remainingQty <= 0 && remainingAmt <= 0) break;
      if (lot.qtyRemaining <= 0 && lot.amtRemaining <= 0) continue;

      const qtyTake = Math.min(lot.qtyRemaining, remainingQty);
      const unitAmt = lot.qtyRemaining > 0 ? safeDiv(lot.amtRemaining, lot.qtyRemaining) : 0;
      let amtTakeByQty = round2(qtyTake * unitAmt);

      // 남은 요청금액보다 과도하면 조정
      if (remainingAmt > 0) {
        amtTakeByQty = Math.min(amtTakeByQty, lot.amtRemaining, remainingAmt);
      } else {
        amtTakeByQty = 0;
      }

      // 수량 먼저, 그 수량 비례금액만 차감
      if (qtyTake > 0) {
        lot.qtyRemaining = round2(lot.qtyRemaining - qtyTake);
        lot.amtRemaining = round2(lot.amtRemaining - amtTakeByQty);

        usedQty = round2(usedQty + qtyTake);
        usedAmt = round2(usedAmt + amtTakeByQty);

        remainingQty = round2(remainingQty - qtyTake);
        remainingAmt = round2(Math.max(0, remainingAmt - amtTakeByQty));
      }

      // 수량은 다 찼지만 금액만 남아 있는 비정상 케이스에 대해
      // 같은 lot 에서 추가 금액을 무리하게 당기지 않음
      // -> 남는 금액은 이후 전년/당해/부족으로 처리
    }

    return {
      usedQty,
      usedAmt,
      remainingQty,
      remainingAmt,
    };
  }

  // =========================
  // 결과 렌더링
  // =========================
  function updateBranchFilter(rows) {
    const filter = qs("branchFilter");
    if (!filter) return;

    const current = filter.value;
    const branches = Array.from(new Set(rows.map((r) => r["지점명"]))).sort((a, b) =>
      a.localeCompare(b, "ko")
    );

    filter.innerHTML = `<option value="">전체</option>` +
      branches.map((b) => `<option value="${escapeHtml(b)}">${escapeHtml(b)}</option>`).join("");

    if (branches.includes(current)) {
      filter.value = current;
    }
  }

  function getFilteredRows() {
    const filter = qs("branchFilter");
    const branch = filter ? filter.value : "";
    if (!branch) return [...finalRows];
    return finalRows.filter((r) => r["지점명"] === branch);
  }

  function renderPreview() {
    const tbody = qs("previewBody");
    const info = qs("previewInfo");
    if (!tbody || !info) return;

    const filtered = getFilteredRows();
    previewRows = filtered;

    const rowsToShow = filtered.slice(0, currentPreviewLimit);

    let html = "";
    for (const row of rowsToShow) {
      html += "<tr>";
      for (const h of OUTPUT_HEADERS) {
        html += `<td style="border:1px solid #ddd; padding:4px; white-space:nowrap;">${escapeHtml(
          row[h]
        )}</td>`;
      }
      html += "</tr>";
    }

    tbody.innerHTML = html;

    info.textContent = `미리보기 ${rowsToShow.length.toLocaleString()} / 전체 ${filtered.length.toLocaleString()} 행 표시`;
  }

  // =========================
  // 다운로드
  // =========================
  function downloadExcel() {
    if (!downloadReady || finalRows.length === 0) {
      alert("먼저 파일을 처리해주세요.");
      return;
    }

    const wsResult = XLSX.utils.json_to_sheet(finalRows, { header: OUTPUT_HEADERS });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, wsResult, "자동소진결과");

    if (validationErrors.length > 0) {
      const errRows = validationErrors.map((msg, idx) => ({
        번호: idx + 1,
        오류내용: msg,
      }));
      const wsErr = XLSX.utils.json_to_sheet(errRows);
      XLSX.utils.book_append_sheet(wb, wsErr, "검증오류");
    }

    XLSX.writeFile(wb, currentFileName);
  }

  // =========================
  // 메인 처리
  // =========================
  async function processWorkbook() {
    validationErrors = [];
    finalRows = [];
    previewRows = [];
    currentPreviewLimit = 300;
    downloadReady = false;
    qs("downloadBtn").disabled = true;

    const fileInput = qs("fileInput");
    if (!fileInput || !fileInput.files || !fileInput.files[0]) {
      alert("엑셀 파일을 먼저 선택해주세요.");
      return;
    }

    const file = fileInput.files[0];
    currentFileName = file.name.replace(/\.(xlsx|xls)$/i, "") + "_자동소진결과.xlsx";

    try {
      setStatus("파일 읽는 중...");
      setSummary("");
      setErrors([]);

      const data = await file.arrayBuffer();
      workbook = XLSX.read(data, {
        type: "array",
        cellDates: false,
        raw: true,
      });

      const ws2025 = workbook.Sheets["2025"];
      const wsPrev = workbook.Sheets["전년재고_DB"];
      const wsPrev2 = workbook.Sheets["전전년재고_DB"];

      if (!ws2025) throw new Error("2025 시트를 찾을 수 없습니다.");
      if (!wsPrev) throw new Error("전년재고_DB 시트를 찾을 수 없습니다.");
      if (!wsPrev2) throw new Error("전전년재고_DB 시트를 찾을 수 없습니다.");

      setStatus("2025 시트 파싱 중...");
      const lines2025 = parse2025Sheet(ws2025);

      setStatus("재고 시트 파싱 중...");
      const prevRows = parseInventorySheet(wsPrev, "전년");
      const prev2Rows = parseInventorySheet(wsPrev2, "전전년");

      setStatus("재고 버킷 구성 중...");
      const inventoryBuckets = buildInventoryBuckets(prev2Rows, prevRows);

      setStatus("FIFO 자동소진 계산 중...");
      finalRows = allocateFromBuckets(lines2025, inventoryBuckets);
      sortRows(finalRows);

      updateBranchFilter(finalRows);
      renderPreview();

      const totalSalesQty = finalRows.reduce((s, r) => s + toNumber(r["판매수량"]), 0);
      const totalSalesAmt = finalRows.reduce((s, r) => s + toNumber(r["판매금액"]), 0);
      const totalDiscardQty = finalRows.reduce((s, r) => s + toNumber(r["폐기수량"]), 0);
      const totalDiscardAmt = finalRows.reduce((s, r) => s + toNumber(r["폐기금액"]), 0);
      const totalShortQty = finalRows.reduce((s, r) => s + toNumber(r["부족수량"]), 0);
      const totalShortAmt = finalRows.reduce((s, r) => s + toNumber(r["부족금액"]), 0);

      setSummary(
        [
          `처리 완료`,
          `- 결과 행 수: ${finalRows.length.toLocaleString()}건`,
          `- 판매수량 합계: ${round2(totalSalesQty).toLocaleString()}`,
          `- 판매금액 합계: ${round2(totalSalesAmt).toLocaleString()}`,
          `- 폐기수량 합계: ${round2(totalDiscardQty).toLocaleString()}`,
          `- 폐기금액 합계: ${round2(totalDiscardAmt).toLocaleString()}`,
          `- 부족수량 합계: ${round2(totalShortQty).toLocaleString()}`,
          `- 부족금액 합계: ${round2(totalShortAmt).toLocaleString()}`,
          `- 검증 오류 수: ${validationErrors.length.toLocaleString()}건`,
        ].join("\n")
      );

      setErrors(validationErrors);

      downloadReady = true;
      qs("downloadBtn").disabled = false;
      setStatus("완료");
    } catch (err) {
      console.error(err);
      setStatus(`오류 발생: ${err.message}`);
      setSummary("");
      setErrors(validationErrors);
      alert(`처리 중 오류가 발생했습니다.\n${err.message}`);
    }
  }

  // =========================
  // 이벤트 바인딩
  // =========================
  function bindEvents() {
    qs("processBtn").addEventListener("click", processWorkbook);

    qs("downloadBtn").addEventListener("click", downloadExcel);

    qs("branchFilter").addEventListener("change", () => {
      currentPreviewLimit = 300;
      renderPreview();
    });

    qs("moreBtn").addEventListener("click", () => {
      currentPreviewLimit += 300;
      renderPreview();
    });
  }

  // =========================
  // 시작
  // =========================
  initUI();
  bindEvents();
  setStatus("엑셀 파일을 선택한 뒤 '처리 시작'을 눌러주세요.");
  setSummary("");
  setErrors([]);
})();
