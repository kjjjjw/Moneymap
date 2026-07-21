const CATEGORY_TREE = {
  "고정비": {
    "주거비": ["주담대 원리금", "아파트 관리비"],
    "통신비": ["준우", "소희"],
    "보험비": ["준우", "소희"]
  },
  "생활비": {
    "식비": ["개인식비 (준우)", "개인식비 (소희)", "장보기", "외식비"],
    "취미": ["준우 탁구장 이용료", "소희 네일, 요가"],
    "용돈": ["준우", "소희"],
    "교통비": ["준우", "소희"],
    "여행": ["여행"],
    "육아비": ["육아 제반"],
    "미용": ["미용"],
    "생활": ["생필품"],
    "건강": ["병원/약"],
    "쇼핑": ["쇼핑"],
    "구독": ["구독"],
    "기타": ["기타"],
    "계모임": ["맛집탐방·외식", "미술사 회비"]
  },
  "비정기 지출": {
    "명절": ["명절 현금"],
    "가족": ["가족 생일"],
    "경조사": ["경조사"],
    "세금": ["세금"]
  }
};

const PAGE_SIZE = 20;

const msalInstance = new msal.PublicClientApplication({
  auth: {
    clientId: APP_CONFIG.clientId,
    authority: APP_CONFIG.authority,
    redirectUri: window.location.href.split("#")[0].split("?")[0]
  },
  cache: { cacheLocation: "localStorage" }
});

const GRAPH_SCOPES = ["Files.ReadWrite"];

const el = (id) => document.getElementById(id);

// ── 목록 상태 ──────────────────────────────────────────────
let loadedRows = [];        // { index, values } — 최신 항목이 배열 앞쪽
let totalRows = 0;
let isLoading = false;

// ── 수정 상태 ──────────────────────────────────────────────
let editingIndex = null;    // 표 안에서의 행 위치
let editingOriginal = null; // 불러왔을 때의 값 (덮어쓰기 전 대조용)

function todayStr() {
  const d = new Date();
  const tzOffsetMs = d.getTimezoneOffset() * 60000;
  return new Date(d - tzOffsetMs).toISOString().slice(0, 10);
}

// 엑셀 날짜는 문자열로 올 수도, 일련번호(숫자)로 올 수도 있습니다.
function toDateInput(v) {
  if (typeof v === "number" && isFinite(v)) {
    return new Date(Math.round((v - 25569) * 86400000)).toISOString().slice(0, 10);
  }
  const s = String(v ?? "").trim();
  const m = s.match(/(\d{4})[-./\s]+(\d{1,2})[-./\s]+(\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;
  return s;
}

function norm(v) {
  return String(v ?? "").trim();
}

function sameRow(a, b) {
  if (!a || !b || a.length !== b.length) return false;
  return a.every((v, i) => norm(v) === norm(b[i]));
}

// ── 분류 셀렉트 ────────────────────────────────────────────
function fillSelect(sel, items, selected) {
  sel.innerHTML = "";
  const list = items.slice();
  if (selected && !list.includes(selected)) list.push(selected); // 표에만 있는 옛 항목 보존
  for (const item of list) {
    const opt = document.createElement("option");
    opt.value = item;
    opt.textContent = item;
    sel.appendChild(opt);
  }
  if (selected) sel.value = selected;
}

function populateMajor(major, minor, detail) {
  fillSelect(el("major"), Object.keys(CATEGORY_TREE), major);
  populateMinor(minor, detail);
}

function populateMinor(minor, detail) {
  const major = el("major").value;
  fillSelect(el("minor"), Object.keys(CATEGORY_TREE[major] || {}), minor);
  populateDetail(detail);
}

function populateDetail(detail) {
  const major = el("major").value;
  const minor = el("minor").value;
  fillSelect(el("detail"), (CATEGORY_TREE[major] || {})[minor] || [], detail);
}

function showStatus(msg, isError) {
  const s = el("statusMsg");
  s.textContent = msg;
  s.className = "status" + (isError ? " error" : msg ? " success" : "");
}

// ── Graph 호출 ─────────────────────────────────────────────
async function getAccessToken() {
  const account = msalInstance.getAllAccounts()[0];
  if (!account) throw new Error("로그인이 필요합니다.");
  try {
    const result = await msalInstance.acquireTokenSilent({ scopes: GRAPH_SCOPES, account });
    return result.accessToken;
  } catch (e) {
    await msalInstance.acquireTokenRedirect({ scopes: GRAPH_SCOPES, account });
    throw new Error("다시 로그인 중입니다...");
  }
}

async function graphFetch(path, options = {}) {
  const token = await getAccessToken();
  const res = await fetch("https://graph.microsoft.com/v1.0" + path, {
    ...options,
    headers: {
      Authorization: "Bearer " + token,
      "Content-Type": "application/json",
      ...(options.headers || {})
    }
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Graph API 오류 (${res.status}): ${text}`);
  }
  if (res.status === 204) return null;
  return res.json();
}

async function getFileId() {
  const cached = localStorage.getItem("gagyebu_fileId");
  if (cached) return cached;
  const q = encodeURIComponent(APP_CONFIG.fileName);
  const candidates = [
    `/me/drive/special/documents:/가계부/${q}`,
    `/me/drive/root:/문서/가계부/${q}`,
    `/me/drive/root:/Documents/가계부/${q}`,
    `/me/drive/root:/${q}`
  ];
  let match = null;
  for (const path of candidates) {
    try {
      match = await graphFetch(path);
      if (match && match.id) break;
    } catch (e) {
      match = null;
    }
  }
  if (!match) {
    const result = await graphFetch(`/me/drive/root/search(q='${q}')`);
    match = (result.value || []).find((f) => f.name === APP_CONFIG.fileName);
  }
  if (!match) throw new Error(`OneDrive에서 '${APP_CONFIG.fileName}' 파일을 찾을 수 없습니다.`);
  localStorage.setItem("gagyebu_fileId", match.id);
  return match.id;
}

function tablePath(fileId) {
  return `/me/drive/items/${fileId}/workbook/tables/${encodeURIComponent(APP_CONFIG.tableName)}`;
}

function rowPath(fileId, index) {
  // /rows/{index} 형식은 ApiNotFound가 나는 경우가 있어 ItemAt을 씁니다.
  return `${tablePath(fileId)}/rows/$/ItemAt(index=${index})`;
}

async function recalc(fileId) {
  try {
    await graphFetch(`/me/drive/items/${fileId}/workbook/application/calculate`, {
      method: "POST",
      body: JSON.stringify({ calculationType: "Recalculate" })
    });
  } catch (e) {
    // 재계산 실패는 무시 - 다음에 파일을 열면 자동으로 재계산됩니다.
  }
}

async function getRowCount(fileId) {
  try {
    const range = await graphFetch(`${tablePath(fileId)}/dataBodyRange?$select=rowCount`);
    if (range && typeof range.rowCount === "number") return range.rowCount;
  } catch (e) {
    // 아래 폴백으로 진행
  }
  const all = await graphFetch(`${tablePath(fileId)}/rows?$select=index`);
  return (all.value || []).length;
}

// ── 쓰기 ───────────────────────────────────────────────────
async function addRow(values) {
  const fileId = await getFileId();
  await graphFetch(`${tablePath(fileId)}/rows/add`, {
    method: "POST",
    body: JSON.stringify({ values: [values] })
  });
  await recalc(fileId);
}

// 목록을 불러온 뒤 다른 곳에서 행이 지워지면 index가 밀립니다.
// 쓰기 직전에 그 행만 다시 읽어 값이 그대로인지 확인합니다.
async function assertRowUnchanged(fileId, index, expected) {
  const row = await graphFetch(rowPath(fileId, index));
  const actual = (row && row.values && row.values[0]) || null;
  if (!sameRow(expected, actual)) {
    throw new Error("이 내역이 다른 곳에서 바뀌었습니다. 새로고침한 뒤 다시 시도하세요.");
  }
}

async function updateRow(index, original, values) {
  const fileId = await getFileId();
  await assertRowUnchanged(fileId, index, original);
  await graphFetch(rowPath(fileId, index), {
    method: "PATCH",
    body: JSON.stringify({ values: [values] })
  });
  await recalc(fileId);
}

async function deleteRow(index, original) {
  const fileId = await getFileId();
  await assertRowUnchanged(fileId, index, original);
  await graphFetch(rowPath(fileId, index), { method: "DELETE" });
  await recalc(fileId);
}

// ── 목록 읽기 (뒤에서부터 20개씩) ──────────────────────────
async function loadRows(reset) {
  if (isLoading) return;
  isLoading = true;
  const list = el("recentList");
  const moreBtn = el("moreBtn");
  moreBtn.disabled = true;

  try {
    const fileId = await getFileId();
    if (reset) {
      loadedRows = [];
      list.innerHTML = "<li class='muted'>불러오는 중...</li>";
      el("listMeta").textContent = "";
      totalRows = await getRowCount(fileId);
    } else {
      moreBtn.textContent = "불러오는 중...";
    }

    const remaining = totalRows - loadedRows.length;
    if (remaining > 0) {
      const top = Math.min(PAGE_SIZE, remaining);
      const skip = remaining - top;
      const data = await graphFetch(`${tablePath(fileId)}/rows?$top=${top}&$skip=${skip}`);
      const batch = (data.value || []).map((r, i) => ({
        index: typeof r.index === "number" ? r.index : skip + i,
        values: (r.values && r.values[0]) || []
      }));
      loadedRows = loadedRows.concat(batch.reverse()); // 최신이 위로
    }
    renderRows();
  } catch (e) {
    if (reset) el("recentList").innerHTML = "";
    el("listMeta").textContent = "내역을 불러오지 못했습니다: " + e.message;
  } finally {
    isLoading = false;
    moreBtn.textContent = "더 보기";
    moreBtn.disabled = false;
  }
}

function renderRows() {
  const list = el("recentList");
  list.innerHTML = "";

  if (loadedRows.length === 0) {
    list.innerHTML = "<li class='muted'>아직 입력한 내역이 없습니다. 위에서 첫 지출을 남겨보세요.</li>";
  }

  for (const row of loadedRows) {
    const [date, major, minor, detail, memo, amount] = row.values;
    const li = document.createElement("li");
    if (row.index === editingIndex) li.classList.add("is-editing");

    const dateEl = document.createElement("span");
    dateEl.className = "rdate";
    dateEl.textContent = toDateInput(date);

    const catEl = document.createElement("span");
    catEl.className = "rcat";
    catEl.textContent = [major, minor, detail].filter(Boolean).join(" · ");

    const amtEl = document.createElement("span");
    amtEl.className = "ramt";
    amtEl.textContent = Number(amount || 0).toLocaleString("ko-KR") + "원";

    li.append(dateEl, catEl, amtEl);

    if (memo) {
      const memoEl = document.createElement("span");
      memoEl.className = "rmemo";
      memoEl.textContent = memo;
      li.append(memoEl);
    }

    const actions = document.createElement("span");
    actions.className = "ractions";

    const editBtn = document.createElement("button");
    editBtn.type = "button";
    editBtn.className = "linklike";
    editBtn.textContent = "수정";
    editBtn.addEventListener("click", () => startEdit(row));

    const delBtn = document.createElement("button");
    delBtn.type = "button";
    delBtn.className = "linklike danger";
    delBtn.textContent = "삭제";
    delBtn.addEventListener("click", () => handleDelete(row));

    actions.append(editBtn, delBtn);
    li.append(actions);
    list.appendChild(li);
  }

  const more = totalRows - loadedRows.length;
  el("moreBtn").hidden = more <= 0;
  el("listMeta").textContent = totalRows
    ? `전체 ${totalRows.toLocaleString("ko-KR")}건 중 ${loadedRows.length.toLocaleString("ko-KR")}건`
    : "";
}

// ── 수정 모드 ──────────────────────────────────────────────
function startEdit(row) {
  const [date, major, minor, detail, memo, amount] = row.values;
  editingIndex = row.index;
  editingOriginal = row.values.slice();

  el("date").value = toDateInput(date);
  populateMajor(norm(major), norm(minor), norm(detail));
  el("memo").value = norm(memo);
  el("amount").value = Number(amount || 0);

  el("editBanner").hidden = false;
  el("editBanner").textContent = `${toDateInput(date)} · ${norm(detail)} 내역을 수정하는 중`;
  el("submitBtn").textContent = "수정 저장";
  el("cancelEditBtn").hidden = false;
  showStatus("", false);
  renderRows();

  el("entryForm").scrollIntoView({ behavior: "smooth", block: "start" });
  el("amount").focus();
}

function cancelEdit() {
  editingIndex = null;
  editingOriginal = null;
  el("editBanner").hidden = true;
  el("submitBtn").textContent = "입력";
  el("cancelEditBtn").hidden = true;
  el("date").value = todayStr();
  el("memo").value = "";
  el("amount").value = "";
  populateMajor();
  showStatus("", false);
  renderRows();
}

function formValues() {
  return [
    el("date").value,
    el("major").value,
    el("minor").value,
    el("detail").value,
    el("memo").value,
    Number(el("amount").value)
  ];
}

async function handleSubmit(e) {
  e.preventDefault();
  const submitBtn = el("submitBtn");
  submitBtn.disabled = true;
  const editing = editingIndex !== null;
  showStatus(editing ? "수정하는 중..." : "저장하는 중...", false);

  try {
    const values = formValues();
    if (editing) {
      await updateRow(editingIndex, editingOriginal, values);
      cancelEdit();
      showStatus("수정했습니다.", false);
    } else {
      await addRow(values);
      showStatus("저장했습니다.", false);
      el("amount").value = "";
      el("memo").value = "";
    }
    await loadRows(true);
  } catch (err) {
    showStatus(err.message, true);
  } finally {
    submitBtn.disabled = false;
  }
}

async function handleDelete(row) {
  const [date, , , detail, , amount] = row.values;
  const label = `${toDateInput(date)} · ${norm(detail)} · ${Number(amount || 0).toLocaleString("ko-KR")}원`;
  if (!window.confirm(`이 내역을 삭제할까요?\n\n${label}\n\n엑셀에서도 함께 지워지며 되돌릴 수 없습니다.`)) return;

  showStatus("삭제하는 중...", false);
  try {
    await deleteRow(row.index, row.values);
    if (editingIndex === row.index) cancelEdit();
    showStatus("삭제했습니다.", false);
    await loadRows(true); // 삭제하면 뒤쪽 행 index가 밀리므로 전체를 다시 읽습니다
  } catch (err) {
    showStatus(err.message, true);
  }
}

// ── 계획 대비 실적 ─────────────────────────────────────────
let summaryYear = null;
let summaryMonth = null;
let summaryMode = "month";
let expandedMajors = new Set();

function summarySheetPath() {
  return `/me/drive/items/${cachedFileId()}/workbook/worksheets/${encodeURIComponent(APP_CONFIG.summarySheet)}`;
}

function cachedFileId() {
  return localStorage.getItem("gagyebu_fileId");
}

function monthKey(year, month) {
  return `${year}-${String(month).padStart(2, "0")}`;
}

function columnsFor(mode, year, month) {
  if (mode === "year") return YEAR_COLUMNS[String(year)] || null;
  return MONTH_COLUMNS[monthKey(year, month)] || null;
}

function colIndex(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function colOffset(fromCol, toCol) {
  return colIndex(toCol) - colIndex(fromCol);
}

// 시트에서 필요한 셀만 하나의 range 주소로 모아 한 번에 읽습니다.
async function fetchSummaryRange(cols, rows) {
  const [planCol, actualCol] = cols;
  const minRow = Math.min(...rows);
  const maxRow = Math.max(...rows);
  const address = `${planCol}${minRow}:${actualCol}${maxRow}`;
  const data = await graphFetch(
    `${summarySheetPath()}/range(address='${address}')?$select=values`
  );
  return { values: data.values, minRow, planCol };
}

function cellFromRange(range, targetCol, targetRow) {
  const rowOffset = targetRow - range.minRow;
  const colOff = colOffset(range.planCol, targetCol);
  const line = range.values[rowOffset];
  return line ? line[colOff] : null;
}

function fmtWon(n) {
  const v = Number(n || 0);
  return v.toLocaleString("ko-KR") + "원";
}

// isIncome=true(소득/저축투자): 실적이 계획보다 많을수록 좋음(초록)
// isIncome=false(지출): 실적이 계획보다 적을수록 좋음(초록) - 시트의 "차이" 정의와 동일
function fmtDiff(planVal, actualVal, isIncome) {
  const plan = Number(planVal || 0);
  const actual = Number(actualVal || 0);
  const diff = isIncome ? (actual - plan) : (plan - actual);
  if (diff === 0) return { text: "±0", cls: "" };
  const sign = diff > 0 ? "+" : "";
  return { text: `${sign}${diff.toLocaleString("ko-KR")}원`, cls: diff > 0 ? "under" : "over" };
}

function populateSummarySelectors() {
  const yearSel = el("summaryYear");
  const monthSel = el("summaryMonth");
  const years = Object.keys(YEAR_COLUMNS);
  yearSel.innerHTML = "";
  for (const y of years) {
    const opt = document.createElement("option");
    opt.value = y;
    opt.textContent = `${y}년`;
    yearSel.appendChild(opt);
  }

  monthSel.innerHTML = "";
  for (let m = 1; m <= 12; m++) {
    const opt = document.createElement("option");
    opt.value = m;
    opt.textContent = `${m}월`;
    monthSel.appendChild(opt);
  }

  const now = new Date();
  const curYear = String(now.getFullYear());
  const curMonth = now.getMonth() + 1;
  summaryYear = years.includes(curYear) ? curYear : years[0];
  summaryMonth = MONTH_COLUMNS[monthKey(summaryYear, curMonth)] ? curMonth : 1;
  yearSel.value = summaryYear;
  monthSel.value = summaryMonth;
  el("summaryMode").value = summaryMode;
  syncSummaryModeUI();
}

function syncSummaryModeUI() {
  el("summaryMonth").hidden = summaryMode === "year";
}

async function loadSummary() {
  const body = el("summaryBody");
  const netEl = el("summaryNet");
  body.innerHTML = "<p class=\'muted\'>불러오는 중...</p>";
  netEl.textContent = "";

  const cols = columnsFor(summaryMode, summaryYear, summaryMonth);
  if (!cols) {
    body.innerHTML = "<p class=\'muted\'>이 시트에 없는 기간입니다 (2026.07~2028.12 범위만 있습니다).</p>";
    return;
  }

  try {
    await getFileId(); // fileId 캐시 보장
    const allRows = SUMMARY_GROUPS.flatMap((g) => g.rows.map((r) => r.row).concat([g.totalRow]))
      .concat([GRAND_TOTAL_ROW, NET_ROW, CUMULATIVE_ROW]);
    const range = await fetchSummaryRange(cols, allRows);
    renderSummary(range, cols);
  } catch (e) {
    body.innerHTML = `<p class=\'muted\'>불러오지 못했습니다: ${e.message}</p>`;
  }
}

function renderSummary(range, cols) {
  const [planCol, actualCol] = cols;
  const body = el("summaryBody");
  body.innerHTML = "";

  // 막대는 "계획" 기준 트랙 위에 실적을 채웁니다.
  // 실적이 계획을 넘으면 트랙 밖으로 오버플로우 세그먼트가 삐져나옵니다.
  function barHTML(plan, actual, isIncome) {
    const p = Number(plan || 0);
    const a = Number(actual || 0);
    const base = Math.max(p, 1); // 0으로 나누기 방지
    const fillPct = Math.min(100, (a / base) * 100);
    const overPct = a > p ? Math.min(100, ((a - p) / base) * 100) : 0;
    const isOver = a > p;
    // 지출: 넘치면 나쁨(빨강). 소득/저축: 못 채우면 나쁨(옅게), 넘치면 좋음(그대로 채움)
    const fillClass = isIncome ? "sbar-fill-income" : (isOver ? "sbar-fill-over" : "sbar-fill");
    return `
      <div class="sbar" role="img" aria-label="계획 ${fmtWon(p)} 중 실적 ${fmtWon(a)}">
        <div class="sbar-track">
          <div class="${fillClass}" style="width:${fillPct}%"></div>
          ${overPct > 0 ? `<div class="sbar-over" style="width:${overPct}%"></div>` : ""}
        </div>
      </div>
    `;
  }

  for (const group of SUMMARY_GROUPS) {
    const isIncome = group.major === "소득" || group.major === "저축·투자";
    const plan = cellFromRange(range, planCol, group.totalRow);
    const actual = cellFromRange(range, actualCol, group.totalRow);
    const diff = fmtDiff(plan, actual, isIncome);

    const wrap = document.createElement("div");
    wrap.className = "sgroup";

    const head = document.createElement("button");
    head.type = "button";
    head.className = "sgroup-head";
    const isOpen = expandedMajors.has(group.major);
    head.setAttribute("aria-expanded", String(isOpen));
    head.innerHTML = `
      <div class="sgroup-toprow">
        <span class="sgroup-name">${group.major}</span>
        <span class="sgroup-caret">${isOpen ? "접기 ▲" : "펼치기 ▼"}</span>
      </div>
      ${barHTML(plan, actual, isIncome)}
      <div class="sgroup-nums">
        <span class="snum-block"><em>계획</em>${fmtWon(plan)}</span>
        <span class="snum-block"><em>실적</em>${fmtWon(actual)}</span>
        <span class="sdiff ${diff.cls}">${diff.text}</span>
      </div>
    `;
    head.addEventListener("click", () => {
      if (expandedMajors.has(group.major)) expandedMajors.delete(group.major);
      else expandedMajors.add(group.major);
      renderSummary(range, cols);
    });

    wrap.appendChild(head);

    if (isOpen) {
      const detailList = document.createElement("div");
      detailList.className = "sgroup-detail";
      for (const r of group.rows) {
        const p = cellFromRange(range, planCol, r.row);
        const a = cellFromRange(range, actualCol, r.row);
        const d = fmtDiff(p, a, isIncome);
        const line = document.createElement("div");
        line.className = "sline";
        const label = r.minor ? `${r.minor} · ${r.detail}` : r.detail;
        line.innerHTML = `
          <div class="sline-toprow">
            <span class="sline-label">${label}</span>
            <span class="sdiff ${d.cls}">${d.text}</span>
          </div>
          ${barHTML(p, a, isIncome)}
          <div class="sline-nums">
            <span>계획 ${fmtWon(p)}</span>
            <span>실적 ${fmtWon(a)}</span>
          </div>
        `;
        detailList.appendChild(line);
      }
      wrap.appendChild(detailList);
    }

    body.appendChild(wrap);
  }

  const grandPlan = cellFromRange(range, planCol, GRAND_TOTAL_ROW);
  const grandActual = cellFromRange(range, actualCol, GRAND_TOTAL_ROW);
  const grandDiff = fmtDiff(grandPlan, grandActual, false);
  const netPlan = cellFromRange(range, planCol, NET_ROW);
  const netActual = cellFromRange(range, actualCol, NET_ROW);

  const totalWrap = document.createElement("div");
  totalWrap.className = "sgroup stotal";
  totalWrap.innerHTML = `
    <div class="sgroup-head is-static">
      <div class="sgroup-toprow">
        <span class="sgroup-name">지출 합계</span>
      </div>
      ${barHTML(grandPlan, grandActual, false)}
      <div class="sgroup-nums">
        <span class="snum-block"><em>계획</em>${fmtWon(grandPlan)}</span>
        <span class="snum-block"><em>실적</em>${fmtWon(grandActual)}</span>
        <span class="sdiff ${grandDiff.cls}">${grandDiff.text}</span>
      </div>
    </div>
  `;
  body.appendChild(totalWrap);

  const netEl = el("summaryNet");
  const netActualNum = Number(netActual || 0);
  netEl.className = "summary-net " + (netActualNum >= 0 ? "positive" : "negative");
  netEl.innerHTML = `
    <span class="net-label">월 수지 (실적)</span>
    <span class="net-value">${fmtWon(netActual)}</span>
    <span class="net-plan">계획 ${fmtWon(netPlan)}</span>
  `;
}

function switchTab(name) {
  const isEntry = name === "entry";
  el("tab-entry").hidden = !isEntry;
  el("tab-summary").hidden = isEntry;
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.classList.toggle("is-active", btn.dataset.tab === name);
  });
  if (name === "summary") {
    if (!summaryYear) populateSummarySelectors();
    loadSummary();
  }
}

// ── 로그인 상태 ────────────────────────────────────────────
function showApp() {
  el("loginArea").hidden = true;
  el("mainArea").hidden = false;
  el("authArea").innerHTML = `<span class="who"></span> <button type="button" id="logoutBtn" class="linklike">로그아웃</button>`;
  el("authArea").querySelector(".who").textContent = msalInstance.getAllAccounts()[0].username;
  el("logoutBtn").addEventListener("click", () => {
    localStorage.removeItem("gagyebu_fileId");
    msalInstance.logoutRedirect();
  });
  loadRows(true);
}

function showLogin() {
  el("loginArea").hidden = false;
  el("mainArea").hidden = true;
  el("authArea").innerHTML = "";
}

async function init() {
  el("date").value = todayStr();
  populateMajor();
  el("major").addEventListener("change", () => populateMinor());
  el("minor").addEventListener("change", () => populateDetail());
  el("entryForm").addEventListener("submit", handleSubmit);
  el("cancelEditBtn").addEventListener("click", cancelEdit);
  el("refreshBtn").addEventListener("click", () => loadRows(true));
  el("moreBtn").addEventListener("click", () => loadRows(false));
  el("loginBtn").addEventListener("click", () => msalInstance.loginRedirect({ scopes: GRAPH_SCOPES }));

  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => switchTab(btn.dataset.tab));
  });
  el("summaryMode").addEventListener("change", (e) => {
    summaryMode = e.target.value;
    syncSummaryModeUI();
    loadSummary();
  });
  el("summaryYear").addEventListener("change", (e) => {
    summaryYear = e.target.value;
    loadSummary();
  });
  el("summaryMonth").addEventListener("change", (e) => {
    summaryMonth = Number(e.target.value);
    loadSummary();
  });
  el("summaryRefreshBtn").addEventListener("click", () => loadSummary());

  try {
    await msalInstance.initialize();
    await msalInstance.handleRedirectPromise();
  } catch (e) {
    el("loginError").textContent = "로그인 처리 중 오류: " + e.message;
  }

  if (msalInstance.getAllAccounts().length > 0) {
    showApp();
  } else {
    showLogin();
  }
}

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => navigator.serviceWorker.register("sw.js").catch(() => {}));
}

init();
