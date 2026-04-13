(() => {
  "use strict";

  /******************************************************************
   * 사용명세서 자동소진 프로그램 - script.js 전체 교체본
   *
   * 동작 방식
   * - 파일 선택 즉시 자동 처리
   * - 필수 시트:
   *   1) 2025
   *   2) 전년재고_DB
   *   3) 전전년재고_DB
   *
   * 2025 시트 구조
   * - A열: 날짜
   * - B열: 품목
   * - 9행: 지점명
   * - 10행: 항목명
   * - 11행부터 데이터
   *
   * 지점 블록 필수 헤더명
   * - 판매수량
   * - 판매금액(사용명세서)
   * - 최종폐기
   * - 금액
   *
   * 재고 시트 구조 (A:E 고정)
   * - A열: 지점명
   * - B열: 품목
   * - C열: 일자
   * - D열: 수량
   * - E열: 금액
   *
   * FIFO
   * - 전전년 -> 전년 -> 당해
   * - 세 연차를 모두 배분하고도 남을 때만 부족
   ******************************************************************/

  if (typeof XLSX === "undefined") {
    alert("XLSX 라이브러리가 먼저 로드되어야 합니다.");
    return;
  }

  const ITEM_ORDER = ["의류", "잡화", "생활", "문화", "건강미용", "식품", "기증파트너"];
  const ITEM_ORDER_MAP = new Map(ITEM_ORDER.map((v, i) => [v, i]));

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

  let workbook = null;
  let finalRows = [];
  let validationErrors = [];
  let currentFileName = "자동소진_결과.xlsx";
  let previewLimit = 300;

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
    const cleaned = String(v).replace(/,/g, "").trim();
    if (!cleaned) return 0;
    const n = Number(cleaned);
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

  function sortResultRows(rows) {
    rows.sort((a, b) => {
      const c1 = String(a["지점명"]).localeCompare(String(b["지점명"]), "ko");
      if (c1 !== 0) return c1;

      const d1 = dateSortValue(a["일자"]);
      const d2 = dateSortValue(b["일자"]);
      if (d1 !== d2) return d1 - d2;

      const i1 = ITEM_ORDER_MAP.has(a["품목군"]) ? ITEM_ORDER_MAP.get(a["품목군"]) : 999;
      const i2 = ITEM_ORDER_MAP.has(b["품목군"]) ? ITEM_ORDER_MAP.get(b["품목군"]) : 999;
      if (i1 !== i2) return i1 - i2;

      return 0;
    });
  }

  function ensurePreviewArea() {
    let wrap = q("resultPreviewWrap");
    if (!wrap) {
      wrap = document.createElement("div");
      wrap.id = "resultPreviewWrap";
      wrap.style.marginTop = "20px";
      wrap.innerHTML = `
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px; gap:8px; flex-wrap:wrap;">
          <div id="previewInfo" style="font-size:13px; color:#334155;">미리보기 준비중</div>
          <button id="moreRowsBtn" type="button" style="padding:8px 12px; border-radius:8px; border:1px solid #cbd5e1; background:#fff; cursor:pointer;">더 보기</button>
        </div>
        <div style="overflow:auto; max-height:520px; border:1px solid #e5e7eb; border-radius:12px; background:#fff;">
          <table id="resultPreviewTable" style="width:100%; border-collapse:collapse; font-size:12px;">
            <thead>
              <tr id="resultPreviewHead"></tr>
            </thead>
            <tbody id="resultPreviewBody"></tbody>
          </table>
        </div>
      `;
      document.body.appendChild(wrap);
    }

    const head = q("resultPreviewHead");
    if (head && !head.dataset.ready) {
      head.innerHTML = OUTPUT_HEADERS.map(
        (h) =>
          `<th style="position:sticky; top:0; z-index:1; background:#f8fafc; border-bottom:1px solid #cbd5e1; padding:8px; text-align:left; white-space:nowrap;">${escapeHtml(h)}</th>`
      ).join("");
      head.dataset.ready = "1";
    }

    const moreBtn = q("moreRowsBtn");
    if (moreBtn && !moreBtn.dataset.bound) {
      moreBtn.addEventListener("click", () => {
        previewLimit += 300;
        renderPreview();
      });
      moreBtn.dataset.bound = "1";
    }
  }

  function setTextIfExists(ids, text) {
    for (const id of ids) {
      const el = q(id);
      if (el) {
        el.textContent = text;
        return true;
      }
    }
    return false;
  }

  function setHtmlIfExists(ids, html) {
    for (const id of ids) {
      const el = q(id);
      if (el) {
        el.innerHTML = html;
        return true;
      }
    }
    return false;
  }

  function setStatus(text) {
    if (setTextIfExists(["status", "statusText", "processStatus"], text)) return;

    let el = q("autoStatusBox");
    if (!el) {
      el = document.createElement("div");
      el.id = "autoStatusBox";
      el.style.margin = "12px 0";
      el.style.padding = "12px";
      el.style.border = "1px solid #e5e7eb";
      el.style.borderRadius = "12px";
      el.style.background = "#f8fafc";
      document.body.appendChild(el);
    }
    el.textContent = text;
  }

  function setErrorList(errors) {
    const html =
      errors.length === 0
        ? `<div style="color:#334155;">오류 없음</div>`
        : errors
            .slice(0, 300)
            .map((e) => `<div style="margin-bottom:6px;">• ${escapeHtml(e)}</div>`)
            .join("") +
          (errors.length > 300
            ? `<div style="margin-top:8px; color:#64748b;">외 ${errors.length - 300}건</div>`
            : "");

    if (setHtmlIfExists(["errorList", "errorBox", "anomalyList"], html)) return;

    let box = q("autoErrorBox");
    if (!box) {
      box = document.createElement("div");
      box.id = "autoErrorBox";
      box.style.marginTop = "16px";
      box.style.padding = "12px";
      box.style.border = "1px solid #fecaca";
      box.style.borderRadius = "12px";
      box.style.background = "#fff7f7";
      box.style.maxHeight = "260px";
      box.style.overflow = "auto";
      document.body.appendChild(box);
    }
    box.innerHTML = `<div style="font-weight:700; margin-bottom:8px;">오류 / 이상치</div>${html}`;
  }

  function setSummaryCards(rows, errors) {
    const salesCount = rows.reduce((sum, r) => sum + (toNumber(r["판매수량"]) > 0 ? 1 : 0), 0);
    const discardCount = rows.reduce((sum, r) => sum + (toNumber(r["폐기수량"]) > 0 ? 1 : 0), 0);
    const shortCount = rows.reduce(
      (sum, r) => sum + (toNumber(r["부족수량"]) > 0 || toNumber(r["부족금액"]) > 0 ? 1 : 0),
      0
    );

    setTextIfExists(["salesCount", "saleCount", "판매처리건수"], String(salesCount));
    setTextIfExists(["discardCount", "disposeCount", "폐기처리건수"], String(discardCount));
    setTextIfExists(["errorCount", "anomalyCount", "오류건수"], String(errors.length));
    setTextIfExists(["shortageCount", "lackCount", "부족건수"], String(shortCount));

    const statCards = Array.from(document.querySelectorAll("[data-stat]"));
    if (statCards.length > 0) {
      statCards.forEach((el) => {
        const key = el.getAttribute("data-stat");
        if (key === "sales") el.textContent = String(salesCount);
        if (key === "discard") el.textContent = String(discardCount);
        if (key === "error") el.textContent = String(errors.length);
        if (key === "shortage") el.textContent = String(shortCount);
      });
    }
  }

  function updateBranchFilter(rows) {
    let select =
      q("branchFilter") ||
      qq("select[data-role='branch-filter']") ||
      qq("select");

    if (!select) return;

    const current = select.value;
    const branches = [...new Set(rows.map((r) => r["지점명"]))].sort((a, b) =>
      String(a).localeCompare(String(b), "ko")
    );

    select.innerHTML =
      `<option value="">전체</option>` +
      branches.map((b) => `<option value="${escapeHtml(b)}">${escapeHtml(b)}</option>`).join("");

    if (branches.includes(current)) {
      select.value = current;
    }

    if (!select.dataset.boundFifo) {
      select.addEventListener("change", () => {
        previewLimit = 300;
        renderPreview();
      });
      select.dataset.boundFifo = "1";
    }
  }

  function getSelectedBranch() {
    const select =
      q("branchFilter") ||
      qq("select[data-role='branch-filter']") ||
      qq("select");
    return select ? select.value : "";
  }

  function getFilteredRows() {
    const branch = getSelectedBranch();
    if (!branch) return [...finalRows];
    return finalRows.filter((r) => r["지점명"] === branch);
  }

  function renderPreview() {
    ensurePreviewArea();
    const body = q("resultPreviewBody");
    const info = q("previewInfo");
    if (!body || !info) return;

    const rows = getFilteredRows();
    const sliced = rows.slice(0, previewLimit);

    body.innerHTML = sliced
      .map((row) => {
        return `<tr>${OUTPUT_HEADERS.map((h) => {
          const v = row[h];
          return `<td style="border-bottom:1px solid #eef2f7; padding:6px 8px; white-space:nowrap;">${escapeHtml(v)}</td>`;
        }).join("")}</tr>`;
      })
      .join("");

    info.textContent = `미리보기 ${sliced.length.toLocaleString()} / 전체 ${rows.length.toLocaleString()} 행`;
  }

  function downloadExcel() {
    if (!finalRows.length) {
      alert("먼저 파일을 업로드해서 처리 결과를 만든 뒤 다운로드 해주세요.");
      return;
    }

    const wb = XLSX.utils.book_new();
    const wsResult = XLSX.utils.json_to_sheet(finalRows, { header: OUTPUT_HEADERS });
    XLSX.utils.book_append_sheet(wb, wsResult, "자동소진결과");

    const wsErrors = XLSX.utils.json_to_sheet(
      validationErrors.map((msg, i) => ({ 번호: i + 1, 오류내용: msg }))
    );
    XLSX.utils.book_append_sheet(wb, wsErrors, "오류이상치");

    XLSX.writeFile(wb, currentFileName);
  }

  function bindDownloadButton() {
    const btn =
      q("downloadBtn") ||
      q("downloadButton") ||
      qq("button[id*='download']") ||
      qq("button");

    if (!btn) return;

    if (!btn.dataset.boundDownloadFifo) {
      btn.addEventListener("click", (e) => {
        e.preventDefault();
        downloadExcel();
      });
      btn.dataset.boundDownloadFifo = "1";
    }

    btn.disabled = finalRows.length === 0;
  }

  function parse2025Sheet(ws) {
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const BRANCH_ROW = 8; // 9행
    const HEADER_ROW = 9; // 10행
    const DATA_START_ROW = 10; // 11행
    const COL_DATE = 0; // A
    const COL_ITEM = 1; // B

    const REQUIRED_HEADERS = ["판매수량", "판매금액(사용명세서)", "최종폐기", "금액"];
    const branchColumnCandidates = new Map();
    const resultMap = new Map();

    for (let c = 2; c <= range.e.c; c++) {
      const branchCell = ws[XLSX.utils.encode_cell({ r: BRANCH_ROW, c })];
      const headerCell = ws[XLSX.utils.encode_cell({ r: HEADER_ROW, c })];
      const branchName = toText(branchCell ? branchCell.v : "");
      const headerName = toText(headerCell ? headerCell.v : "");

      if (!branchName) continue;
      if (!branchColumnCandidates.has(branchName)) {
        branchColumnCandidates.set(branchName, []);
      }
      branchColumnCandidates.get(branchName).push({ col: c, header: headerName });
    }

    const branchHeaderMap = new Map();

    for (const [branchName, columns] of branchColumnCandidates.entries()) {
      const map = {};
      for (const need of REQUIRED_HEADERS) {
        const found = columns.find((x) => x.header === need);
        if (found) map[need] = found.col;
      }

      const missing = REQUIRED_HEADERS.filter((h) => map[h] === undefined);
      if (missing.length > 0) {
        validationErrors.push(`[2025 시트] ${branchName} 지점 블록에서 필수 헤더 누락: ${missing.join(", ")}`);
      }

      branchHeaderMap.set(branchName, map);
    }

    for (let r = DATA_START_ROW; r <= range.e.r; r++) {
      const dateValue = ws[XLSX.utils.encode_cell({ r, c: COL_DATE })];
      const itemValue = ws[XLSX.utils.encode_cell({ r, c: COL_ITEM })];

      const date = formatDate(dateValue ? dateValue.v : "");
      const item = toText(itemValue ? itemValue.v : "");

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

        const saleQty = toNumber(
          (ws[XLSX.utils.encode_cell({ r, c: saleQtyCol })] || {}).v
        );
        const saleAmt = toNumber(
          (ws[XLSX.utils.encode_cell({ r, c: saleAmtCol })] || {}).v
        );
        const discardQty = toNumber(
          (ws[XLSX.utils.encode_cell({ r, c: discardQtyCol })] || {}).v
        );
        const discardAmt = toNumber(
          (ws[XLSX.utils.encode_cell({ r, c: discardAmtCol })] || {}).v
        );

        if (saleQty === 0 && saleAmt === 0 && discardQty === 0 && discardAmt === 0) {
          continue;
        }

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

    return Array.from(resultMap.values());
  }

  function parseInventorySheet(ws, label) {
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
    const result = [];

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i] || [];
      const branch = toText(row[0]);
      const item = toText(row[1]);
      const date = formatDate(row[2]);
      const qty = toNumber(row[3]);
      const amt = toNumber(row[4]);

      const rowJoined = row.map((v) => toText(v)).join("|");
      if (!branch && !item && !date && qty === 0 && amt === 0) continue;
      if (rowJoined.includes("지점명") && rowJoined.includes("품목")) continue;
      if (!branch || !item) continue;

      if (qty < 0 || amt < 0) {
        validationErrors.push(`[재고오류] ${label} / ${branch} / ${item} / ${date} : 수량 또는 금액이 음수`);
      }

      result.push({
        yearLabel: label,
        지점명: branch,
        품목군: item,
        일자: date,
        수량: qty,
        금액: amt,
      });
    }

    return result;
  }

  function buildInventoryBuckets(prev2Rows, prevRows) {
    const buckets = new Map();

    function addLot(kind, row) {
      const key = inventoryKey(row["지점명"], row["품목군"]);
      if (!buckets.has(key)) {
        buckets.set(key, { prev2: [], prev1: [] });
      }
      buckets.get(key)[kind].push({
        date: row["일자"],
        sortDate: dateSortValue(row["일자"]),
        qtyRemaining: toNumber(row["수량"]),
        amtRemaining: toNumber(row["금액"]),
      });
    }

    prev2Rows.forEach((r) => addLot("prev2", r));
    prevRows.forEach((r) => addLot("prev1", r));

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

    if ((!remainingQty && !remainingAmt) || !lots || lots.length === 0) {
      return { usedQty, usedAmt, remainingQty, remainingAmt };
    }

    for (const lot of lots) {
      if (remainingQty <= 0 && remainingAmt <= 0) break;
      if (lot.qtyRemaining <= 0 && lot.amtRemaining <= 0) continue;

      const takeQty = Math.min(lot.qtyRemaining, Math.max(0, remainingQty));
      if (takeQty <= 0) continue;

      const unitAmt = lot.qtyRemaining > 0 ? lot.amtRemaining / lot.qtyRemaining : 0;
      let takeAmt = round2(takeQty * unitAmt);

      if (remainingAmt > 0) {
        takeAmt = Math.min(takeAmt, lot.amtRemaining, remainingAmt);
      } else {
        takeAmt = 0;
      }

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

    const step1 = consumeLots(prev2Lots, remainingQty, remainingAmt);
    out.prev2Qty = round2(step1.usedQty);
    out.prev2Amt = round2(step1.usedAmt);
    remainingQty = round2(step1.remainingQty);
    remainingAmt = round2(step1.remainingAmt);

    const step2 = consumeLots(prev1Lots, remainingQty, remainingAmt);
    out.prev1Qty = round2(step2.usedQty);
    out.prev1Amt = round2(step2.usedAmt);
    remainingQty = round2(step2.remainingQty);
    remainingAmt = round2(step2.remainingAmt);

    // 당해 흡수
    if (remainingQty > 0 && remainingAmt > 0) {
      out.currentQty = round2(remainingQty);
      out.currentAmt = round2(remainingAmt);
      remainingQty = 0;
      remainingAmt = 0;
    }

    // 세 연차를 모두 배분하고도 남는 경우만 부족
    if (remainingQty > 0 || remainingAmt > 0) {
      out.shortQty = round2(remainingQty);
      out.shortAmt = round2(remainingAmt);
    }

    return out;
  }

  function allocateAll(lines2025, inventoryBuckets) {
    const rows = [...lines2025];
    sortResultRows(rows);

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

      const qtyCheck =
        result["전전년_판매수량"] +
        result["전전년_폐기수량"] +
        result["전년_판매수량"] +
        result["전년_폐기수량"] +
        result["당해_판매수량"] +
        result["당해_폐기수량"] +
        result["부족수량"];

      const amtCheck =
        result["전전년_판매금액"] +
        result["전전년_폐기금액"] +
        result["전년_판매금액"] +
        result["전년_폐기금액"] +
        result["당해_판매금액"] +
        result["당해_폐기금액"] +
        result["부족금액"];

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

      return result;
    });
  }

  async function processWorkbook(file) {
    if (!file) return;

    validationErrors = [];
    finalRows = [];
    previewLimit = 300;
    bindDownloadButton();

    try {
      setStatus("파일 읽는 중...");
      const buf = await file.arrayBuffer();
      workbook = XLSX.read(buf, { type: "array", raw: true, cellDates: false });
      currentFileName = file.name.replace(/\.(xlsx|xls)$/i, "") + "_자동소진결과.xlsx";

      const ws2025 = workbook.Sheets["2025"];
      const wsPrev = workbook.Sheets["전년재고_DB"];
      const wsPrev2 = workbook.Sheets["전전년재고_DB"];

      if (!ws2025) throw new Error("2025 시트를 찾을 수 없습니다.");
      if (!wsPrev) throw new Error("전년재고_DB 시트를 찾을 수 없습니다.");
      if (!wsPrev2) throw new Error("전전년재고_DB 시트를 찾을 수 없습니다.");

      setStatus("2025 시트 읽는 중...");
      const lines2025 = parse2025Sheet(ws2025);

      setStatus("재고 시트 읽는 중...");
      const prevRows = parseInventorySheet(wsPrev, "전년");
      const prev2Rows = parseInventorySheet(wsPrev2, "전전년");

      setStatus("FIFO 계산 중...");
      const inventoryBuckets = buildInventoryBuckets(prev2Rows, prevRows);
      finalRows = allocateAll(lines2025, inventoryBuckets);
      sortResultRows(finalRows);

      updateBranchFilter(finalRows);
      setSummaryCards(finalRows, validationErrors);
      setErrorList(validationErrors);
      renderPreview();
      bindDownloadButton();

      setStatus(
        `처리 완료 · 결과 ${finalRows.length.toLocaleString()}건 · 오류 ${validationErrors.length.toLocaleString()}건`
      );
    } catch (err) {
      console.error(err);
      setStatus(`오류 발생: ${err.message}`);
      setErrorList([err.message, ...validationErrors]);
      bindDownloadButton();
      alert(`처리 중 오류가 발생했습니다.\n${err.message}`);
    }
  }

  function bindFileInputAutoProcess() {
    const fileInput =
      q("fileInput") ||
      qq('input[type="file"]') ||
      q("excelFile");

    if (!fileInput) {
      setStatus("파일 업로드 input을 찾지 못했습니다.");
      return;
    }

    if (!fileInput.dataset.boundAutoFifo) {
      fileInput.addEventListener("change", async (e) => {
        const file = e.target.files && e.target.files[0] ? e.target.files[0] : null;
        if (!file) return;
        await processWorkbook(file);
      });
      fileInput.dataset.boundAutoFifo = "1";
    }
  }

  function init() {
    ensurePreviewArea();
    bindFileInputAutoProcess();
    bindDownloadButton();
    setErrorList([]);
    setStatus("엑셀 파일을 선택하면 자동으로 처리됩니다.");
  }

  init();
})();
