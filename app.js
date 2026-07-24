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

const FIRST_PAGE_SIZE = 10;  // 처음 보여줄 건수
const PAGE_SIZE = 20;        // "더 보기" 한 번에 추가할 건수

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
let recentMonthFilter = "";   // "" = 전체 기간, "2026-07" 형태
let allRowsCache = null;      // 월 필터용 전체 행 캐시
let shownCount = 0;           // 월 필터 상태에서 현재 보여준 건수

// ── 수정 상태 ──────────────────────────────────────────────
let editingIndex = null;    // 표 안에서의 행 위치
let editingOriginal = null; // 불러왔을 때의 값 (덮어쓰기 전 대조용)

// ── 금액 입력 포맷 ─────────────────────────────────────────
// 입력창에는 천 단위 콤마를 넣어 보여주고, 저장할 때는 숫자만 꺼냅니다.
function amountToNumber(str) {
  const digits = String(str ?? "").replace(/[^0-9]/g, "");
  return digits ? Number(digits) : 0;
}

function formatAmountInput(str) {
  const digits = String(str ?? "").replace(/[^0-9]/g, "");
  if (!digits) return "";
  return Number(digits).toLocaleString("ko-KR");
}

function setAmountValue(n) {
  const num = Number(n || 0);
  el("amount").value = num ? num.toLocaleString("ko-KR") : "";
}

function getAmountValue() {
  return amountToNumber(el("amount").value);
}

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

// ── 입력 모드: 지출 / 소득 ──────────────────────────────────
// 지출은 기존 CATEGORY_TREE(하드코딩)를 그대로 쓰고,
// 소득은 가계부_월별 시트에서 자동 인식한 구조(getStructure)를 그대로 재사용합니다.
// 그래서 시트에 소득/저축 항목이 추가돼도 카테고리 목록이 저절로 따라옵니다.
let entryMode = "expense"; // "expense" | "income"
let incomeCategoryTree = null; // { 대분류: { 소분류(""면 무분류): [세부항목...] } }
const INCOME_MAJORS = ["소득", "저축·투자"];
const NO_MINOR = ""; // 소분류가 없는 항목(예: 소득)의 자리표시 키

function buildIncomeTree(st) {
  const tree = {};
  for (const g of st.groups) {
    if (!INCOME_MAJORS.includes(g.major)) continue;
    const minors = {};
    for (const r of g.rows) {
      const key = r.minor || NO_MINOR;
      if (!minors[key]) minors[key] = [];
      minors[key].push(r.detail);
    }
    tree[g.major] = minors;
  }
  return tree;
}

function currentCategoryTree() {
  return entryMode === "income" ? (incomeCategoryTree || {}) : CATEGORY_TREE;
}

function currentTableName() {
  return entryMode === "income" ? APP_CONFIG.incomeTableName : APP_CONFIG.tableName;
}

// 소분류가 전부 무분류(NO_MINOR)인 대분류는 소분류 선택을 건너뜁니다 (예: 소득).
function minorFieldNeeded(tree, major) {
  const minors = Object.keys(tree[major] || {});
  return !(minors.length === 1 && minors[0] === NO_MINOR);
}

function populateMajor(major, minor, detail) {
  const tree = currentCategoryTree();
  fillSelect(el("major"), Object.keys(tree), major);
  populateMinor(minor, detail);
}

function populateMinor(minor, detail) {
  const tree = currentCategoryTree();
  const major = el("major").value;
  const needsMinor = minorFieldNeeded(tree, major);
  el("minorField").hidden = !needsMinor;
  // 숨긴 소분류는 required를 풀어야 폼 검증(required)에 걸리지 않습니다.
  el("minor").required = needsMinor;
  if (needsMinor) {
    fillSelect(el("minor"), Object.keys(tree[major] || {}), minor);
  } else {
    fillSelect(el("minor"), [NO_MINOR], NO_MINOR);
  }
  populateDetail(detail);
}

function populateDetail(detail) {
  const tree = currentCategoryTree();
  const major = el("major").value;
  const minor = el("minor").value;
  fillSelect(el("detail"), (tree[major] || {})[minor] || [], detail);
}

async function ensureIncomeTree() {
  if (incomeCategoryTree) return incomeCategoryTree;
  const st = await getStructure();
  incomeCategoryTree = buildIncomeTree(st);
  return incomeCategoryTree;
}

async function setEntryMode(mode) {
  if (mode === entryMode) return;
  if (mode === "income") {
    showStatus("소득 분류를 불러오는 중...", false);
    try {
      await ensureIncomeTree();
    } catch (e) {
      showStatus("소득 분류를 불러오지 못했습니다: " + e.message, true);
      return;
    }
  }
  entryMode = mode;
  cancelEdit(); // 모드를 바꾸면 진행 중이던 수정은 취소합니다.
  document.querySelectorAll(".mode-btn").forEach((btn) => {
    btn.classList.toggle("is-active", btn.dataset.mode === mode);
  });
  el("amountLabel").textContent = mode === "income" ? "입금액 (원)" : "이용금액 (원)";
  el("submitBtn").textContent = "입력";
  el("recentTitle").textContent = mode === "income" ? "소득 · 저축 내역" : "지출 내역";
  populateMajor();
  showStatus("", false);
  recentMonthFilter = "";
  el("recentMonth").value = "";
  allRowsCache = null;
  await loadRows(true);
  populateRecentMonths();
}

// ── 캐릭터 토스트 ──────────────────────────────────────────
let toastTimer = null;

const TOAST_CHARS = {
  saved:   { img: "icons/char-saved.png",   alt: "" },
  deleted: { img: "icons/char-deleted.png", alt: "" },
  edited:  { img: "icons/char-edited.png",  alt: "" }
};

function showToast(kind, message) {
  const box = el("toast");
  const conf = TOAST_CHARS[kind];
  if (!box || !conf) return;

  el("toastImg").src = conf.img;
  el("toastText").textContent = message;
  box.hidden = false;
  // 재생 중이던 애니메이션을 초기화합니다.
  box.classList.remove("is-on");
  void box.offsetWidth;
  box.classList.add("is-on");

  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    box.classList.remove("is-on");
    setTimeout(() => { box.hidden = true; }, 260);
  }, 2200);
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

function tablePath(fileId, tableName) {
  return `/me/drive/items/${fileId}/workbook/tables/${encodeURIComponent(tableName)}`;
}

function rowPath(fileId, tableName, index) {
  // /rows/{index} 형식은 ApiNotFound가 나는 경우가 있어 ItemAt을 씁니다.
  return `${tablePath(fileId, tableName)}/rows/$/ItemAt(index=${index})`;
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

async function getRowCount(fileId, tableName) {
  try {
    const range = await graphFetch(`${tablePath(fileId, tableName)}/dataBodyRange?$select=rowCount`);
    if (range && typeof range.rowCount === "number") return range.rowCount;
  } catch (e) {
    // 아래 폴백으로 진행
  }
  const all = await graphFetch(`${tablePath(fileId, tableName)}/rows?$select=index`);
  return (all.value || []).length;
}

// ── 쓰기 ───────────────────────────────────────────────────
async function addRow(tableName, values) {
  const fileId = await getFileId();
  await graphFetch(`${tablePath(fileId, tableName)}/rows/add`, {
    method: "POST",
    body: JSON.stringify({ values: [values] })
  });
  await recalc(fileId);
}

// 목록을 불러온 뒤 다른 곳에서 행이 지워지면 index가 밀립니다.
// 쓰기 직전에 그 행만 다시 읽어 값이 그대로인지 확인합니다.
async function assertRowUnchanged(fileId, tableName, index, expected) {
  const row = await graphFetch(rowPath(fileId, tableName, index));
  const actual = (row && row.values && row.values[0]) || null;
  if (!sameRow(expected, actual)) {
    throw new Error("이 내역이 다른 곳에서 바뀌었습니다. 새로고침한 뒤 다시 시도하세요.");
  }
}

async function updateRow(tableName, index, original, values) {
  const fileId = await getFileId();
  await assertRowUnchanged(fileId, tableName, index, original);
  await graphFetch(rowPath(fileId, tableName, index), {
    method: "PATCH",
    body: JSON.stringify({ values: [values] })
  });
  await recalc(fileId);
}

async function deleteRow(tableName, index, original) {
  const fileId = await getFileId();
  await assertRowUnchanged(fileId, tableName, index, original);
  await graphFetch(rowPath(fileId, tableName, index), { method: "DELETE" });
  await recalc(fileId);
}

// ── 목록 읽기 (뒤에서부터 20개씩) ──────────────────────────
async function loadRows(reset) {
  if (isLoading) return;
  isLoading = true;
  const list = el("recentList");
  const moreBtn = el("moreBtn");
  moreBtn.disabled = true;

  const tableName = currentTableName();
  try {
    const fileId = await getFileId();

    // 월 필터가 걸린 경우: 그 달 데이터가 표 어디에 있을지 알 수 없으므로
    // 전체를 한 번 받아 클라이언트에서 거릅니다.
    if (recentMonthFilter) {
      if (reset || allRowsCache === null) {
        list.innerHTML = "<li class='muted'>불러오는 중...</li>";
        el("listMeta").textContent = "";
        const data = await graphFetch(`${tablePath(fileId, tableName)}/rows`);
        allRowsCache = (data.value || []).map((r, i) => ({
          index: typeof r.index === "number" ? r.index : i,
          values: (r.values && r.values[0]) || []
        }));
      }
      const filtered = allRowsCache
        .filter((r) => toDateInput(r.values[0]).startsWith(recentMonthFilter))
        .reverse();
      totalRows = filtered.length;
      if (reset) shownCount = FIRST_PAGE_SIZE;
      else shownCount += PAGE_SIZE;
      loadedRows = filtered.slice(0, shownCount);
      renderRows();
      return;
    }

    // 전체 기간: 표 뒤쪽부터 필요한 만큼만 받아옵니다.
    if (reset) {
      loadedRows = [];
      allRowsCache = null;
      list.innerHTML = "<li class='muted'>불러오는 중...</li>";
      el("listMeta").textContent = "";
      totalRows = await getRowCount(fileId, tableName);
    } else {
      moreBtn.textContent = "불러오는 중...";
    }

    const remaining = totalRows - loadedRows.length;
    if (remaining > 0) {
      const pageSize = loadedRows.length === 0 ? FIRST_PAGE_SIZE : PAGE_SIZE;
      const top = Math.min(pageSize, remaining);
      const skip = remaining - top;
      const data = await graphFetch(`${tablePath(fileId, tableName)}/rows?$top=${top}&$skip=${skip}`);
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

// 표 전체를 훑어 데이터가 있는 월 목록을 만듭니다.
async function populateRecentMonths() {
  const sel = el("recentMonth");
  const keep = sel.value;
  try {
    const fileId = await getFileId();
    const tableName = currentTableName();
    const data = await graphFetch(`${tablePath(fileId, tableName)}/rows`);
    allRowsCache = (data.value || []).map((r, i) => ({
      index: typeof r.index === "number" ? r.index : i,
      values: (r.values && r.values[0]) || []
    }));

    const months = new Set();
    for (const r of allRowsCache) {
      const d = toDateInput(r.values[0]);
      if (/^\d{4}-\d{2}/.test(d)) months.add(d.slice(0, 7));
    }

    sel.innerHTML = '<option value="">전체 기간</option>';
    for (const m of Array.from(months).sort().reverse()) {
      const [y, mo] = m.split("-");
      const opt = document.createElement("option");
      opt.value = m;
      opt.textContent = `${y}년 ${Number(mo)}월`;
      sel.appendChild(opt);
    }
    if (keep && months.has(keep)) sel.value = keep;
  } catch (e) {
    // 월 목록을 못 만들어도 전체 기간 보기는 동작합니다.
  }
}

function renderRows() {
  const list = el("recentList");
  list.innerHTML = "";

  if (loadedRows.length === 0) {
    const msg = recentMonthFilter
      ? "이 달에는 내역이 없어요."
      : (entryMode === "income"
          ? "아직 소득·저축 내역이 없어요."
          : "아직 지출 내역이 없어요.");
    const sub = recentMonthFilter ? "다른 달을 골라보세요." : "위에서 첫 항목을 남겨보세요.";
    list.innerHTML = `
      <li class="empty-state">
        <img class="empty-img" src="icons/char-empty.png" alt="">
        <p class="empty-msg">${msg}</p>
        <p class="empty-sub">${sub}</p>
      </li>`;
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
  const scope = recentMonthFilter ? "이 달" : "전체";
  el("listMeta").textContent = totalRows
    ? `${scope} ${totalRows.toLocaleString("ko-KR")}건 중 ${loadedRows.length.toLocaleString("ko-KR")}건`
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
  setAmountValue(amount);

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
    getAmountValue()
  ];
}

async function handleSubmit(e) {
  e.preventDefault();

  // type=text로 바꾸면서 브라우저 기본 검증이 빠지므로 직접 확인합니다.
  if (getAmountValue() <= 0) {
    showStatus("금액을 입력하세요.", true);
    el("amount").focus();
    return;
  }

  const submitBtn = el("submitBtn");
  submitBtn.disabled = true;
  const editing = editingIndex !== null;
  showStatus(editing ? "수정하는 중..." : "저장하는 중...", false);

  const tableName = currentTableName();
  try {
    const values = formValues();
    if (editing) {
      await updateRow(tableName, editingIndex, editingOriginal, values);
      cancelEdit();
      showStatus("수정했습니다.", false);
      showToast("edited", "수정했어요!");
    } else {
      await addRow(tableName, values);
      showStatus("저장했습니다.", false);
      showToast("saved", "저장 완료!");
      el("amount").value = "";
      el("memo").value = "";
    }
    allRowsCache = null;
    await loadRows(true);
    populateRecentMonths();
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
    await deleteRow(currentTableName(), row.index, row.values);
    if (editingIndex === row.index) cancelEdit();
    showStatus("삭제했습니다.", false);
    showToast("deleted", "삭제했어요");
    allRowsCache = null;
    await loadRows(true); // 삭제하면 뒤쪽 행 index가 밀리므로 전체를 다시 읽습니다
    populateRecentMonths();
  } catch (err) {
    showStatus(err.message, true);
  }
}

// ── 계획 대비 실적 ─────────────────────────────────────────
let summaryYear = null;
let summaryMonth = null;
let summaryMode = "month";
let expandedMajors = new Set();
let expandedMinors = new Set(); // "대분류::소분류" 형태로 저장

function summarySheetPath() {
  return `/me/drive/items/${cachedFileId()}/workbook/worksheets/${encodeURIComponent(APP_CONFIG.summarySheet)}`;
}

function cachedFileId() {
  return localStorage.getItem("gagyebu_fileId");
}

function monthKey(year, month) {
  return `${year}-${String(month).padStart(2, "0")}`;
}

// ── 시트 구조 자동 인식 ────────────────────────────────────
// 행/열 위치를 하드코딩하지 않고 시트에서 직접 읽습니다.
// 시트에 행이나 월이 추가돼도 앱이 알아서 따라갑니다.
let sheetStructure = null;

function numToCol(n) {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function cellText(v) {
  return String(v ?? "").trim();
}

// A~C열을 훑어 대분류/소분류/세부항목과 소계 행 위치를 알아냅니다.
// 규칙: C열(세부항목)이 있으면 항목 행, A열만 있고 C열이 비면 소계/합계 행.
function parseRowStructure(values) {
  const groups = [];
  let current = null;
  let lastMinor = "";
  let grandTotalRow = null;
  let netRow = null;
  let cumulativeRow = null;

  for (let i = 0; i < values.length; i++) {
    const rowNum = i + 1;
    const a = cellText(values[i][0]);
    const b = cellText(values[i][1]);
    const c = cellText(values[i][2]);

    if (a === "대분류" || (a === "" && b === "" && c === "")) continue;

    if (c) {
      if (a) {
        current = { major: a, totalRow: null, rows: [] };
        groups.push(current);
        lastMinor = "";
      }
      if (!current) continue;
      if (b) lastMinor = b;
      current.rows.push({ row: rowNum, minor: lastMinor, detail: c });
      continue;
    }

    if (a) {
      if (current && current.totalRow === null) {
        current.totalRow = rowNum;   // 방금 끝난 그룹의 소계
        current = null;
        continue;
      }
      // 그룹 밖의 총계 행들
      if (a.includes("누적")) cumulativeRow = rowNum;
      else if (a.includes("수지")) netRow = rowNum;
      else if (a.includes("합계")) grandTotalRow = rowNum;
    }
  }

  return {
    groups: groups.filter((g) => g.totalRow && g.rows.length),
    grandTotalRow,
    netRow,
    cumulativeRow
  };
}

// 5행(월 라벨)과 6행(계획/실적/차이)을 읽어 연-월 → 열 매핑을 만듭니다.
function parseColumnStructure(headerValues, startColNum) {
  const monthCols = {};
  const yearCols = {};
  const labelRow = headerValues[0] || [];
  const kindRow = headerValues[1] || [];
  let pending = [];

  for (let i = 0; i < kindRow.length; i++) {
    const kind = cellText(kindRow[i]);
    const label = cellText(labelRow[i]);
    const colNum = startColNum + i;

    if (kind !== "계획") continue;

    const cols = [numToCol(colNum), numToCol(colNum + 1), numToCol(colNum + 2)];
    const yearMatch = label.match(/(\d{4})\s*합계/);

    if (yearMatch) {
      const year = yearMatch[1];
      yearCols[year] = cols;
      for (const p of pending) monthCols[monthKey(year, p.month)] = p.cols;
      pending = [];
      continue;
    }

    const monthMatch = label.match(/(\d{1,2})\s*월/);
    if (monthMatch) pending.push({ month: Number(monthMatch[1]), cols });
  }

  return { monthCols, yearCols };
}

async function getStructure() {
  if (sheetStructure) return sheetStructure;
  await getFileId();

  const [labelRes, headerRes] = await Promise.all([
    graphFetch(`${summarySheetPath()}/range(address='A1:C120')?$select=values`),
    graphFetch(`${summarySheetPath()}/range(address='D5:EZ6')?$select=values`)
  ]);

  const rowPart = parseRowStructure(labelRes.values || []);
  const colPart = parseColumnStructure(headerRes.values || [], 4); // D = 4번째 열

  const ok = rowPart.groups.length > 0 && Object.keys(colPart.monthCols).length > 0;
  if (!ok) {
    // 시트 형식이 예상과 다르면 config.js의 값으로 되돌아갑니다.
    sheetStructure = {
      groups: SUMMARY_GROUPS,
      grandTotalRow: GRAND_TOTAL_ROW,
      netRow: NET_ROW,
      cumulativeRow: CUMULATIVE_ROW,
      monthCols: MONTH_COLUMNS,
      yearCols: YEAR_COLUMNS,
      fallback: true
    };
    return sheetStructure;
  }

  sheetStructure = { ...rowPart, ...colPart, fallback: false };
  return sheetStructure;
}

function columnsFor(st, mode, year, month) {
  if (mode === "year") return st.yearCols[String(year)] || null;
  return st.monthCols[monthKey(year, month)] || null;
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

function populateSummarySelectors(st) {
  const yearSel = el("summaryYear");
  const monthSel = el("summaryMonth");
  const years = Object.keys(st.yearCols).sort();
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
  summaryMonth = st.monthCols[monthKey(summaryYear, curMonth)] ? curMonth : firstMonthOf(st, summaryYear);
  yearSel.value = summaryYear;
  monthSel.value = summaryMonth;
  el("summaryMode").value = summaryMode;
  syncSummaryModeUI();
}

function firstMonthOf(st, year) {
  for (let m = 1; m <= 12; m++) {
    if (st.monthCols[monthKey(year, m)]) return m;
  }
  return 1;
}

function syncSummaryModeUI() {
  el("summaryMonth").hidden = summaryMode === "year";
}

async function loadSummary() {
  const body = el("summaryBody");
  const netEl = el("summaryNet");
  body.innerHTML = "<p class='muted'>불러오는 중...</p>";
  netEl.textContent = "";

  try {
    const st = await getStructure();

    const cols = columnsFor(st, summaryMode, summaryYear, summaryMonth);
    if (!cols) {
      body.innerHTML = `
        <div class="empty-state">
          <img class="empty-img" src="icons/char-chart.png" alt="">
          <p class="empty-msg">이 기간은 아직 없어요.</p>
          <p class="empty-sub">다른 월이나 연도를 골라보세요.</p>
        </div>`;
      return;
    }

    const allRows = st.groups
      .flatMap((g) => g.rows.map((r) => r.row).concat([g.totalRow]))
      .concat([st.grandTotalRow, st.netRow, st.cumulativeRow].filter((r) => r));
    const range = await fetchSummaryRange(cols, allRows);
    renderSummary(st, range, cols);

    if (st.fallback) {
      const warn = document.createElement("p");
      warn.className = "sfallback";
      warn.textContent = "시트 구조를 자동으로 읽지 못해 기본 설정값으로 표시하고 있습니다. 숫자가 어긋나면 시트 형식을 확인해주세요.";
      body.prepend(warn);
    }
  } catch (e) {
    body.innerHTML = `<p class='muted'>불러오지 못했습니다: ${e.message}</p>`;
  }
}

function renderSummary(st, range, cols) {
  const [planCol, actualCol] = cols;
  const body = el("summaryBody");
  body.innerHTML = "";

  // 달성률(%) — 계획 대비 실적이 몇 %인지
  function pctInfo(plan, actual, isIncome) {
    const p = Number(plan || 0);
    const a = Number(actual || 0);
    if (p === 0 && a === 0) return { text: "—", cls: "neutral", pct: 0 };
    if (p === 0) return { text: "계획 없음", cls: isIncome ? "good" : "bad", pct: 100 };
    const pct = Math.round((a / p) * 100);
    let cls = "neutral";
    if (isIncome) cls = pct >= 100 ? "good" : (pct >= 80 ? "neutral" : "bad");
    else cls = pct > 100 ? "bad" : (pct >= 85 ? "warn" : "good");
    return { text: pct + "%", cls, pct };
  }

  // 계획 트랙 위에 실적을 채우고, 넘친 만큼은 트랙 밖으로 빗금 처리
  function barHTML(plan, actual, isIncome) {
    const p = Number(plan || 0);
    const a = Number(actual || 0);
    const base = Math.max(p, 1);
    const fillPct = Math.min(100, (a / base) * 100);
    const overPct = a > p ? Math.min(60, ((a - p) / base) * 100) : 0;
    const isOver = a > p;
    const fillClass = isIncome ? "sbar-fill income" : (isOver ? "sbar-fill over" : "sbar-fill");
    return `
      <div class="sbar-track">
        <div class="${fillClass}" style="width:${fillPct}%"></div>
        ${overPct > 0 ? `<div class="sbar-over" style="width:${overPct}%"></div>` : ""}
      </div>
    `;
  }

  for (const group of st.groups) {
    const isIncome = group.major === "소득" || group.major === "저축·투자";
    const plan = cellFromRange(range, planCol, group.totalRow);
    const actual = cellFromRange(range, actualCol, group.totalRow);
    const diff = fmtDiff(plan, actual, isIncome);
    const pct = pctInfo(plan, actual, isIncome);

    const wrap = document.createElement("div");
    wrap.className = "sgroup";

    const head = document.createElement("button");
    head.type = "button";
    head.className = "sgroup-head";
    const isOpen = expandedMajors.has(group.major);
    head.setAttribute("aria-expanded", String(isOpen));
    head.innerHTML = `
      <div class="sgroup-top">
        <span class="sgroup-name">${group.major}</span>
        <span class="spct ${pct.cls}">${pct.text}</span>
      </div>
      ${barHTML(plan, actual, isIncome)}
      <div class="sgroup-foot">
        <span class="sfig"><b>${fmtWon(actual)}</b> <i>/ ${fmtWon(plan)}</i></span>
        <span class="sdiff ${diff.cls}">${diff.text}</span>
      </div>
      <span class="sgroup-toggle">${isOpen ? "닫기" : "자세히"}</span>
    `;
    head.addEventListener("click", () => {
      if (expandedMajors.has(group.major)) {
        expandedMajors.delete(group.major);
        // 대분류를 접으면 그 아래 소분류 펼침 상태도 정리합니다.
        for (const key of Array.from(expandedMinors)) {
          if (key.startsWith(group.major + "::")) expandedMinors.delete(key);
        }
      } else {
        expandedMajors.add(group.major);
      }
      renderSummary(st, range, cols);
    });

    wrap.appendChild(head);

    if (isOpen) {
      // 1단계: 소분류. 시트에 소분류 소계 행이 없으므로 세부항목 값을 합산합니다.
      const buckets = [];
      for (const r of group.rows) {
        const key = r.minor || "";
        let b = buckets.find((x) => x.minor === key);
        if (!b) { b = { minor: key, items: [] }; buckets.push(b); }
        b.items.push(r);
      }

      const detail = document.createElement("div");
      detail.className = "sgroup-detail";

      for (const bucket of buckets) {
        let bPlan = 0;
        let bActual = 0;
        for (const r of bucket.items) {
          bPlan += Number(cellFromRange(range, planCol, r.row) || 0);
          bActual += Number(cellFromRange(range, actualCol, r.row) || 0);
        }
        const bDiff = fmtDiff(bPlan, bActual, isIncome);
        const bPct = pctInfo(bPlan, bActual, isIncome);
        const minorKey = `${group.major}::${bucket.minor}`;
        const minorOpen = !bucket.minor || expandedMinors.has(minorKey);

        const sub = document.createElement("div");
        sub.className = "ssub" + (minorOpen ? " is-open" : "");

        // 소분류가 없는 그룹(소득 등)은 소분류 행 없이 세부항목만 보여줍니다.
        if (bucket.minor) {
          const subHead = document.createElement("button");
          subHead.type = "button";
          subHead.className = "ssub-head";
          subHead.setAttribute("aria-expanded", String(minorOpen));
          subHead.innerHTML = `
            <div class="ssub-top">
              <span class="ssub-name">${bucket.minor}</span>
              <span class="ssub-right">
                <span class="spct sm ${bPct.cls}">${bPct.text}</span>
                <span class="ssub-caret">${minorOpen ? "▲" : "▼"}</span>
              </span>
            </div>
            ${barHTML(bPlan, bActual, isIncome)}
            <div class="ssub-foot">
              <span class="sfig"><b>${fmtWon(bActual)}</b> <i>/ ${fmtWon(bPlan)}</i></span>
              <span class="sdiff sm ${bDiff.cls}">${bDiff.text}</span>
            </div>
          `;
          subHead.addEventListener("click", () => {
            if (expandedMinors.has(minorKey)) expandedMinors.delete(minorKey);
            else expandedMinors.add(minorKey);
            renderSummary(st, range, cols);
          });
          sub.appendChild(subHead);
        }

        // 2단계: 세부항목
        if (minorOpen) {
          const items = document.createElement("div");
          items.className = "ssub-items";
          let html = "";
          for (const r of bucket.items) {
            const p = cellFromRange(range, planCol, r.row);
            const a = cellFromRange(range, actualCol, r.row);
            const d = fmtDiff(p, a, isIncome);
            const ip = pctInfo(p, a, isIncome);
            html += `
              <div class="sitem">
                <div class="sitem-top">
                  <span class="sitem-name">${r.detail}</span>
                  <span class="spct sm ${ip.cls}">${ip.text}</span>
                </div>
                ${barHTML(p, a, isIncome)}
                <div class="sitem-foot">
                  <span class="sfig"><b>${fmtWon(a)}</b> <i>/ ${fmtWon(p)}</i></span>
                  <span class="sdiff sm ${d.cls}">${d.text}</span>
                </div>
              </div>
            `;
          }
          items.innerHTML = html;
          sub.appendChild(items);
        }

        detail.appendChild(sub);
      }

      wrap.appendChild(detail);
    }

    body.appendChild(wrap);
  }

  const grandPlan = cellFromRange(range, planCol, st.grandTotalRow);
  const grandActual = cellFromRange(range, actualCol, st.grandTotalRow);
  const grandDiff = fmtDiff(grandPlan, grandActual, false);
  const grandPct = pctInfo(grandPlan, grandActual, false);
  const netPlan = cellFromRange(range, planCol, st.netRow);
  const netActual = cellFromRange(range, actualCol, st.netRow);

  const totalWrap = document.createElement("div");
  totalWrap.className = "sgroup stotal";
  totalWrap.innerHTML = `
    <div class="sgroup-head is-static">
      <div class="sgroup-top">
        <span class="sgroup-name">지출 합계</span>
        <span class="spct ${grandPct.cls}">${grandPct.text}</span>
      </div>
      ${barHTML(grandPlan, grandActual, false)}
      <div class="sgroup-foot">
        <span class="sfig"><b>${fmtWon(grandActual)}</b> <i>/ ${fmtWon(grandPlan)}</i></span>
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
    if (!summaryYear) {
      getStructure()
        .then((st) => {
          populateSummarySelectors(st);
          loadSummary();
        })
        .catch(() => loadSummary());
    } else {
      loadSummary();
    }
  }
}

// ── 로그인 상태 ────────────────────────────────────────────
function showApp() {
  el("loginArea").hidden = true;
  el("mainArea").hidden = false;
  el("authArea").innerHTML = `<button type="button" id="logoutBtn" class="linklike">로그아웃</button>`;
  el("logoutBtn").addEventListener("click", () => {
    localStorage.removeItem("gagyebu_fileId");
    msalInstance.logoutRedirect();
  });
  loadRows(true);
  populateRecentMonths();
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

  // 금액 입력 중 천 단위 콤마를 실시간으로 적용합니다.
  el("amount").addEventListener("input", (e) => {
    const input = e.target;
    const before = input.value;
    const caretFromEnd = before.length - (input.selectionStart ?? before.length);
    const formatted = formatAmountInput(before);
    if (formatted === before) return;
    input.value = formatted;
    // 콤마가 늘거나 줄어도 커서가 같은 자리에 남도록 뒤에서부터 위치를 잡습니다.
    const pos = Math.max(0, formatted.length - caretFromEnd);
    input.setSelectionRange(pos, pos);
  });
  el("cancelEditBtn").addEventListener("click", cancelEdit);
  el("refreshBtn").addEventListener("click", () => {
    allRowsCache = null;
    loadRows(true);
    populateRecentMonths();
  });

  el("recentMonth").addEventListener("change", (e) => {
    recentMonthFilter = e.target.value;
    loadRows(true);
  });
  el("moreBtn").addEventListener("click", () => loadRows(false));
  el("loginBtn").addEventListener("click", () => msalInstance.loginRedirect({ scopes: GRAPH_SCOPES }));

  document.querySelectorAll(".mode-btn").forEach((btn) => {
    btn.addEventListener("click", () => setEntryMode(btn.dataset.mode));
  });

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
  el("summaryRefreshBtn").addEventListener("click", () => {
    sheetStructure = null; // 시트 구조도 다시 읽습니다
    getStructure()
      .then((st) => {
        populateSummarySelectors(st);
        loadSummary();
      })
      .catch(() => loadSummary());
  });

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
