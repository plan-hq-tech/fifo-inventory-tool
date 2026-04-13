(() => {
  "use strict";

  if (typeof XLSX === "undefined") {
    alert("XLSX 라이브러리(sheetjs)가 먼저 로드되어야 합니다.");
    return;
  }

  const ITEM_ORDER = ["의류", "잡화", "생활", "문화", "건강미용", "식품", "기증파트너"];
  const ITEM_ORDER_MAP = new Map(ITEM_ORDER.map((v, i) => [v, i]));
  const REQUIRED_2025_HEADERS = ["판매수량", "판매금액(사용명세서)", "최종폐기", "금액"];

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

  let finalRows = [];
  let validationErrors = [];
  let previewLimit = 300;
  let currentFileName = "자동소진_결과.xlsx";
  let currentWorkbook = null;
  let lastSelectedFile = null;
  let autoBound = false;

  function log(...args) {
    console.log("[FIFO]", ...args);
  }

  function q(id) {
    return document.getElementById(id);
  }

  function qq(selector) {
    return document.querySelector(selector);
  }

  function toText(v) {
    return String(v ?? "").trim();
  }

  function toNumber(v) {
    if (v === null || v === undefined || v === "") return 0;
    if (typeof v === "number") return Number.isFinite(v) ? v : 0;
    const n = Number(String(v).replace(/,/g, "").trim());
    return Number.isFinite(n) ? n : 0;
  }

  function round2(n) {
    return Math.round((n + Number.EPSILON) * 100) / 100;
  }

  function escapeHtml(str) {
    return String(str ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function excelDateToDate(value) {
    if (value instanceof Date) return value;
    if (typeof value === "number") {
      const parsed = XLSX.SSF.parse_date_code(value);
      if (!parsed) return null;
      return new Date(parsed.y, parsed.m - 1, parsed.d);
    }
    if (typeof value === "string" && value.trim()) {
      const s = value.trim().replace(/\./g, "-").replace(/\//g, "-");
      const d = new Date(s);
      if (!Number.isNaN(d.getTime())) return d;
    }
    return null;
  }

  function formatDate(value) {
    const d = excelDateToDate(value);
    if (!d) return toText(value);
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  }

  function dateSortValue(value) {
    const d = excelDateToDate(value);
    return d ? d.getTime() : Number.MAX_SAFE_INTEGER;
  }

  function rowKey(branch, date, item) {
    return `${branch}|||${date}|||${item}`;
  }

  function inventoryKey(branch, item) {
    return `${branch}|||${item}`;
  }

  function sortRows(rows) {
    rows.sort((a, b) => {
      const c = String(a["지점명"]).localeCompare(String(b["지점명"]), "ko");
      if (c !== 0) return c;

      const d1 = dateSortValue(a["일자"]);
      const d2 = dateSortValue(b["일자"]);
      if (d1 !== d2) return d1 - d2;

      const i1 = ITEM_ORDER_MAP.has(a["품목군"]) ? ITEM_ORDER_MAP.get(a["품목군"]) : 999;
      const i2 = ITEM_ORDER_MAP.has(b["품목군"]) ? ITEM_ORDER_MAP.get(b["품목군"]) : 999;
      return i1 - i2;
    });
  }

  function ensureUI() {
    let app = q("fifoDebugApp");
    if (!app) {
      app = document.createElement("div");
      app.id = "fifoDebugApp";
      app.style.maxWidth = "1400px";
      app.style.margin = "20px auto";
      app.style.padding = "16px";
      app.style.fontFamily = "Arial, sans-serif";
      document.body.appendChild(app);
    }

    if (!q("fifoStatusBox")) {
      const html = `
        <div id="fifoStatusBox" style="margin-bottom:12px; padding:12px; border:1px solid #dbe2ea; border-radius:12px; background:#f8fafc; white-space:pre-wrap;"></div>
        <div style="display:flex; gap:8px; flex-wrap:wrap; margin-bottom:12px;">
          <button id="fifoDownloadBtn" type="button" style="padding:10px 14px; border-radius:10px; border:1px solid #cbd5e1; background:#fff; cursor:pointer;">결과 엑셀 다운로드</button>
          <button id="fifoReprocessBtn" type="button" style="padding:10px 14px; border-radius:10px; border:1px solid #cbd5e1; background:#fff; cursor:pointer;">마지막 파일 다시 처리</button>
          <button id="fifoMoreBtn" type="button" style="padding:10px 14px; border-radius:10px; border:1px solid #cbd5e1; background:#fff; cursor:pointer;">더 보기</button>
          <select id="fifoBranchFilter" style="padding:10px; border-radius:10px; border:1px solid #cbd5e1;">
            <option value="">전체 지점</option>
          </select>
        </div>
        <div id="fifoSummaryBox" style="margin-bottom:12px; padding:12px; border:1px solid #e5e7eb; border-radius:12px; background:#fff; white-space:pre-wrap;"></div>
        <div id="fifoErrorBox" style="margin-bottom:12px; padding:12px; border:1px solid #fecaca; border-radius:12px; background:#fff7f7; max-height:240px; overflow:auto; white-space:pre-wrap;"></div>
        <div id="fifoPreviewInfo" style="margin-bottom:8px;"></div>
        <div style="overflow:auto; max-height:600px; border:1px solid #e5e7eb; border-radius:12px; background:#fff;">
          <table id="fifoPreviewTable" style="width:100%; border-collapse:collapse; font-size:12px;">
            <thead><tr id="fifoPreviewHead"></tr></thead>
            <tbody id="fifoPreviewBody"></tbody>
          </table>
        </div>
      `;
      app.innerHTML = html;

      q("fifoPreviewHead").innerHTML = OUTPUT_HEADERS.map(
        (h) =>
          `<th style="position:sticky; top:0; background:#f8fafc; border-bottom:1px solid #cbd5e1; padding:8px; text-align:left; white-space:nowrap;">${escapeHtml(h)}</th>`
      ).join("");

      q("fifoDownloadBtn").addEventListener("click", downloadExcel);
      q("fifoReprocessBtn").addEventListener("click", () => {
        if (lastSelectedFile) processWorkbook(lastSelectedFile);
        else alert("아직 선택된 파일이 없습니다.");
      });
      q("fifoMoreBtn").addEventListener("click", () => {
        previewLimit += 300;
        renderPreview();
      });
      q("fifoBranchFilter").addEventListener("change", () => {
        previewLimit = 300;
        renderPreview();
      });
    }
  }

  function setStatus(msg) {
    ensureUI();
    q("fifoStatusBox").textContent = msg;
  }

  function setSummary(msg) {
    ensureUI();
    q("fifoSummaryBox").textContent = msg;
  }

  function setErrors(errors) {
    ensureUI();
    if (!errors.length) {
      q("fifoErrorBox").textContent = "오류 없음";
      return;
    }
    const show = errors.slice(0, 300);
    q("fifoErrorBox").textContent =
      `오류 ${errors.length}건\n\n` +
      show.map((x, i) => `${i + 1}. ${x}`).join("\n") +
      (errors.length > 300 ? `\n\n... 외 ${errors.length - 300}건` : "");
  }

  function updateBranchFilter() {
    ensureUI();
    const select = q("fifoBranchFilter");
    const current = select.value;
    const branches = [...new Set(finalRows.map((r) => r["지점명"]))].sort((a, b) =>
      String(a).localeCompare(String(b), "ko")
    );

    select.innerHTML =
      `<option value="">전체 지점</option>` +
      branches.map((b) => `<option value="${escapeHtml(b)}">${escapeHtml(b)}</option>`).join("");

    if (branches.includes(current)) select.value = current;
  }

  function getFilteredRows() {
    const branch = q("fifoBranchFilter") ? q("fifoBranchFilter").value : "";
    if (!branch) return [...finalRows];
    return finalRows.filter((r) => r["지점명"] === branch);
  }

  function renderPreview() {
    ensureUI();
    const rows = getFilteredRows();
    const sliced = rows.slice(0, previewLimit);

    q("fifoPreviewInfo").textContent =
      `미리보기 ${sliced.length.toLocaleString()} / 전체 ${rows.length.toLocaleString()} 행`;

    q("fifoPreviewBody").innerHTML = sliced
      .map((row) => {
        return `<tr>${OUTPUT_HEADERS.map((h) => {
          return `<td style="border-bottom:1px solid #eef2f7; padding:6px 8px; white-space:nowrap;">${escapeHtml(row[h])}</td>`;
        }).join("")}</tr>`;
      })
      .join("");
  }

  function downloadExcel() {
    if (!finalRows.length) {
      alert("처리된 결과가 없습니다.");
      return;
    }

    const wb = XLSX.utils.book_new();
    const wsResult = XLSX.utils.json_to_sheet(finalRows, { header: OUTPUT_HEADERS });
    XLSX.utils.book_append_sheet(wb, wsResult, "자동소진결과");

    const wsErrors = XLSX.utils.json_to_sheet(
      validationErrors.map((msg, idx) => ({ 번호: idx + 1, 오류내용: msg }))
    );
    XLSX.utils.book_append_sheet(wb, wsErrors, "오류이상치");

    XLSX.writeFile(wb, currentFileName);
  }

  function parse2025Sheet(ws) {
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const BRANCH_ROW = 8;
    const HEADER_ROW = 9;
    const DATA_START_ROW = 10;
    const COL_DATE = 0;
    const COL_ITEM = 1;

    const branchColumnMap = new Map();
    const resultMap = new Map();

    for (let c = 2; c <= range.e.c; c++) {
      const branchName = toText((ws[XLSX.utils.encode_cell({ r: BRANCH_ROW, c })] || {}).v);
      const headerName = toText((ws[XLSX.utils.encode_cell({ r: HEADER_ROW, c })] || {}).v);

      if (!branchName) continue;
      if (!branchColumnMap.has(branchName)) branchColumnMap.set(branchName, []);
      branchColumnMap.get(branchName).push({ col: c, header: headerName });
    }

    log("지점 블록 개수:", branchColumnMap.size);

    const branchHeaderMap = new Map();

    for (const [branchName, cols] of branchColumnMap.entries()) {
      const map = {};
      for (const req of REQUIRED_2025_HEADERS) {
        const found = cols.find((x) => x.header === req);
        if (found) map[req] = found.col;
      }

      const missing = REQUIRED_2025_HEADERS.filter((x) => map[x] === undefined);
      if (missing.length) {
        validationErrors.push(`[2025 시트] ${branchName} 지점 헤더 누락: ${missing.join(", ")}`);
      }

      branchHeaderMap.set(branchName, map);
    }

    for (let r = DATA_START_ROW; r <= range.e.r; r++) {
      const date = formatDate((ws[XLSX.utils.encode_cell({ r, c: COL_DATE })] || {}).v);
      const item = toText((ws[XLSX.utils.encode_cell({ r, c: COL_ITEM })] || {}).v);

      if (!date && !item) continue;
      if (!date || !item) continue;

      for (const [branchName, map] of branchHeaderMap.entries()) {
        const saleQtyCol = map["판매수량"];
        const saleAmtCol = map["판매금액(사용명세서)"];
        const discardQtyCol = map["최종폐기"];
        const discardAmtCol = map["금액"];

        if (
          saleQtyCol === undefined ||
          saleAmtCol === undefined ||
          discardQtyCol === undefined ||
          discardAmtCol === undefined
        ) {
          continue;
        }

        const saleQty = toNumber((ws[XLSX.utils.encode_cell({ r, c: saleQtyCol })] || {}).v);
        const saleAmt = toNumber((ws[XLSX.utils.encode_cell({ r, c: saleAmtCol })] || {}).v);
        const discardQty = toNumber((ws[XLSX.utils.encode_cell({ r, c: discardQtyCol })] || {}).v);
        const discardAmt = toNumber((ws[XLSX.utils.encode_cell({ r, c: discardAmtCol })] || {}).v);

        if (saleQty === 0 && saleAmt === 0 && discardQty === 0 && discardAmt === 0) continue;

        if (saleQty > 0 && saleAmt === 0) {
          validationErrors.push(`[검증오류] ${branchName} / ${date} / ${item} : 판매수량 > 0 인데 판매금액 = 0`);
        }
        if (saleQty === 0 && saleAmt > 0) {
          validationErrors.push(`[검증오류] ${branchName} / ${date} / ${item} : 판매수량 = 0 인데 판매금액 > 0`);
        }
        if (discardQty > 0 && discardAmt === 0) {
          validationErrors.push(`[검증오류] ${branchName} / ${date} / ${item} : 폐기수량 > 0 인데 폐기금액 = 0`);
        }
        if (discardQty === 0 && discardAmt > 0) {
          validationErrors.push(`[검증오류] ${branchName} / ${date} / ${item} : 폐기수량 = 0 인데 폐기금액 > 0`);
        }

        const key = rowKey(branchName, date, item);
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

    const rows = Array.from(resultMap.values());
    log("2025 파싱 결과 행수:", rows.length);
    return rows;
  }

  function parseInventorySheet(ws, label) {
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
    const out = [];

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i] || [];
      const branch = toText(row[0]);
      const item = toText(row[1]);
      const date = formatDate(row[2]);
      const qty = toNumber(row[3]);
      const amt = toNumber(row[4]);

      const joined = row.map((x) => toText(x)).join("|");
      if (!branch && !item && !date && qty === 0 && amt === 0) continue;
      if (joined.includes("지점명") && joined.includes("품목")) continue;
      if (!branch || !item) continue;

      out.push({
        yearLabel: label,
        지점명: branch,
        품목군: item,
        일자: date,
        수량: qty,
        금액: amt,
      });
    }

    log(label, "재고 행수:", out.length);
    return out;
  }

  function buildInventoryBuckets(prev2Rows, prevRows) {
    const buckets = new Map();

    function add(kind, row) {
      const key = inventoryKey(row["지점명"], row["품목군"]);
      if (!buckets.has(key)) buckets.set(key, { prev2: [], prev1: [] });
      buckets.get(key)[kind].push({
        date: row["일자"],
        sortDate: dateSortValue(row["일자"]),
        qtyRemaining: toNumber(row["수량"]),
        amtRemaining: toNumber(row["금액"]),
      });
    }

    prev2Rows.forEach((r) => add("prev2", r));
    prevRows.forEach((r) => add("prev1", r));

    for (const bucket of buckets.values()) {
      bucket.prev2.sort((a, b) => a.sortDate - b.sortDate);
      bucket.prev1.sort((a, b) => a.sortDate - b.sortDate);
    }

    return buckets;
  }

  function consumeLots(lots, reqQty, reqAmt) {
    let remainingQty = toNumber(reqQty);
    let remainingAmt = toNumber(reqAmt);
    let usedQty = 0;
    let usedAmt = 0;

    if ((!remainingQty && !remainingAmt) || !lots || !lots.length) {
      return { usedQty, usedAmt, remainingQty, remainingAmt };
    }

    for (const lot of lots) {
      if (remainingQty <= 0 && remainingAmt <= 0) break;
      if (lot.qtyRemaining <= 0 && lot.amtRemaining <= 0) continue;

      const takeQty = Math.min(lot.qtyRemaining, Math.max(0, remainingQty));
      if (takeQty <= 0) continue;

      const unitAmt = lot.qtyRemaining > 0 ? lot.amtRemaining / lot.qtyRemaining : 0;
      let takeAmt = round2(takeQty * unitAmt);

      if (remainingAmt > 0) takeAmt = Math.min(takeAmt, lot.amtRemaining, remainingAmt);
      else takeAmt = 0;

      lot.qtyRemaining = round2(lot.qtyRemaining - takeQty);
      lot.amtRemaining = round2(lot.amtRemaining - takeAmt);

      usedQty = round2(usedQty + takeQty);
      usedAmt = round2(usedAmt + takeAmt);
      remainingQty = round2(remainingQty - takeQty);
      remainingAmt = round2(Math.max(0, remainingAmt - takeAmt));
    }

    return { usedQty, usedAmt, remainingQty, remainingAmt };
  }

  function allocateUsage(reqQty, reqAmt, prev2Lots, prev1Lots) {
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

    const a1 = consumeLots(prev2Lots, remainingQty, remainingAmt);
    out.prev2Qty = round2(a1.usedQty);
    out.prev2Amt = round2(a1.usedAmt);
    remainingQty = round2(a1.remainingQty);
    remainingAmt = round2(a1.remainingAmt);

    const a2 = consumeLots(prev1Lots, remainingQty, remainingAmt);
    out.prev1Qty = round2(a2.usedQty);
    out.prev1Amt = round2(a2.usedAmt);
    remainingQty = round2(a2.remainingQty);
    remainingAmt = round2(a2.remainingAmt);

    if (remainingQty > 0 && remainingAmt > 0) {
      out.currentQty = round2(remainingQty);
      out.currentAmt = round2(remainingAmt);
      remainingQty = 0;
      remainingAmt = 0;
    }

    if (remainingQty > 0 || remainingAmt > 0) {
      out.shortQty = round2(remainingQty);
      out.shortAmt = round2(remainingAmt);
    }

    return out;
  }

  function allocateAll(lines2025, inventoryBuckets) {
    const rows = [...lines2025];
    sortRows(rows);

    return rows.map((line) => {
      const branch = line["지점명"];
      const item = line["품목군"];
      const saleQty = toNumber(line["판매수량"]);
      const saleAmt = toNumber(line["판매금액"]);
      const discardQty = toNumber(line["폐기수량"]);
      const discardAmt = toNumber(line["폐기금액"]);
      const totalQty = round2(saleQty + discardQty);
      const totalAmt = round2(saleAmt + discardAmt);

      const key = inventoryKey(branch, item);
      const bucket = inventoryBuckets.get(key) || { prev2: [], prev1: [] };

      const saleAlloc = allocateUsage(saleQty, saleAmt, bucket.prev2, bucket.prev1);
      const discardAlloc = allocateUsage(discardQty, discardAmt, bucket.prev2, bucket.prev1);

      const result = {
        지점명: branch,
        일자: line["일자"],
        품목군: item,
        판매수량: round2(saleQty),
        판매금액: round2(saleAmt),
        폐기수량: round2(discardQty),
        폐기금액: round2(discardAmt),
        총사용수량: totalQty,
        총사용금액: totalAmt,
        전전년_판매수량: saleAlloc.prev2Qty,
        전전년_판매금액: saleAlloc.prev2Amt,
        전전년_폐기수량: discardAlloc.prev2Qty,
        전전년_폐기금액: discardAlloc.prev2Amt,
        전년_판매수량: saleAlloc.prev1Qty,
        전년_판매금액: saleAlloc.prev1Amt,
        전년_폐기수량: discardAlloc.prev1Qty,
        전년_폐기금액: discardAlloc.prev1Amt,
        당해_판매수량: saleAlloc.currentQty,
        당해_판매금액: saleAlloc.currentAmt,
        당해_폐기수량: discardAlloc.currentQty,
        당해_폐기금액: discardAlloc.currentAmt,
        부족수량: round2(saleAlloc.shortQty + discardAlloc.shortQty),
        부족금액: round2(saleAlloc.shortAmt + discardAlloc.shortAmt),
      };

      return result;
    });
  }

  async function processWorkbook(file) {
    if (!file) return;

    try {
      ensureUI();
      setStatus("파일 읽는 중...");
      validationErrors = [];
      finalRows = [];
      previewLimit = 300;
      lastSelectedFile = file;
      currentFileName = file.name.replace(/\.(xlsx|xls)$/i, "") + "_자동소진결과.xlsx";

      const buf = await file.arrayBuffer();
      currentWorkbook = XLSX.read(buf, { type: "array", raw: true, cellDates: false });

      log("시트 목록:", currentWorkbook.SheetNames);

      const ws2025 = currentWorkbook.Sheets["2025"];
      const wsPrev = currentWorkbook.Sheets["전년재고_DB"];
      const wsPrev2 = currentWorkbook.Sheets["전전년재고_DB"];

      if (!ws2025) throw new Error("2025 시트를 찾을 수 없습니다.");
      if (!wsPrev) throw new Error("전년재고_DB 시트를 찾을 수 없습니다.");
      if (!wsPrev2) throw new Error("전전년재고_DB 시트를 찾을 수 없습니다.");

      setStatus("2025 시트 파싱 중...");
      const lines2025 = parse2025Sheet(ws2025);

      setStatus("전년/전전년 재고 파싱 중...");
      const prevRows = parseInventorySheet(wsPrev, "전년");
      const prev2Rows = parseInventorySheet(wsPrev2, "전전년");

      setStatus("FIFO 계산 중...");
      const inventoryBuckets = buildInventoryBuckets(prev2Rows, prevRows);
      finalRows = allocateAll(lines2025, inventoryBuckets);
      sortRows(finalRows);

      const totalSales = finalRows.reduce((s, r) => s + (toNumber(r["판매수량"]) > 0 ? 1 : 0), 0);
      const totalDiscard = finalRows.reduce((s, r) => s + (toNumber(r["폐기수량"]) > 0 ? 1 : 0), 0);
      const totalShort = finalRows.reduce(
        (s, r) => s + (toNumber(r["부족수량"]) > 0 || toNumber(r["부족금액"]) > 0 ? 1 : 0),
        0
      );

      setSummary(
        [
          `처리 완료`,
          `- 결과 행수: ${finalRows.length.toLocaleString()}`,
          `- 판매 처리건수: ${totalSales.toLocaleString()}`,
          `- 폐기 처리건수: ${totalDiscard.toLocaleString()}`,
          `- 부족 발생건: ${totalShort.toLocaleString()}`,
          `- 오류 수: ${validationErrors.length.toLocaleString()}`
        ].join("\n")
      );

      setErrors(validationErrors);
      updateBranchFilter();
      renderPreview();
      setStatus(`완료: ${file.name}`);

      log("최종 결과 행수:", finalRows.length);
      if (!finalRows.length) {
        validationErrors.push("결과 행이 0건입니다. 2025 시트의 9행/10행 헤더 구조 또는 실제 헤더명이 다를 가능성이 큽니다.");
        setErrors(validationErrors);
      }
    } catch (err) {
      console.error(err);
      setStatus(`오류 발생: ${err.message}`);
      validationErrors.unshift(err.message);
      setErrors(validationErrors);
      alert(`처리 중 오류가 발생했습니다.\n${err.message}`);
    }
  }

  function findFileInput() {
    return (
      q("fileInput") ||
      q("excelFile") ||
      qq('input[type="file"]')
    );
  }

  function bindFileInput() {
    const fileInput = findFileInput();
    if (!fileInput) {
      log("파일 input 아직 없음");
      return false;
    }

    if (fileInput.dataset.fifoBound === "1") {
      return true;
    }

    fileInput.addEventListener("change", async (e) => {
      const file = e.target.files && e.target.files[0] ? e.target.files[0] : null;
      log("파일 선택 change 이벤트 감지:", file ? file.name : "없음");
      if (!file) return;
      await processWorkbook(file);
    });

    fileInput.dataset.fifoBound = "1";
    log("파일 input 바인딩 완료");
    setStatus("엑셀 파일을 선택하면 자동으로 처리됩니다.");
    return true;
  }

  function startBindingWatcher() {
    ensureUI();

    bindFileInput();

    const observer = new MutationObserver(() => {
      if (!autoBound) {
        autoBound = bindFileInput();
      } else {
        bindFileInput();
      }
    });

    observer.observe(document.documentElement, {
      childList: true,
      subtree: true,
    });

    let tries = 0;
    const timer = setInterval(() => {
      tries += 1;
      const ok = bindFileInput();
      if (ok || tries >= 20) {
        clearInterval(timer);
      }
    }, 500);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", startBindingWatcher);
  } else {
    startBindingWatcher();
  }
})();
