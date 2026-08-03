// ── 홈 화면 설치 이벤트 ────────────────────────────────────
// 이 리스너는 페이지 초기에 등록돼야 이벤트를 놓치지 않습니다.
let deferredInstall = null;

window.addEventListener("beforeinstallprompt", (e) => {
  e.preventDefault();
  deferredInstall = e;
  const btn = document.getElementById("installBtn");
  if (btn) btn.hidden = false;
});

window.addEventListener("appinstalled", () => {
  deferredInstall = null;
  const btn = document.getElementById("installBtn");
  if (btn) btn.hidden = true;
});

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

// 작성자 (엑셀 표의 7번째 열)
const WRITERS = ["준우", "소히"];
let currentWriter = localStorage.getItem("gagyebu_writer") || WRITERS[0];

function setWriter(name) {
  if (!WRITERS.includes(name)) return;
  currentWriter = name;
  localStorage.setItem("gagyebu_writer", name);
  document.querySelectorAll(".who-btn").forEach((b) => {
    const on = b.dataset.who === name;
    b.classList.toggle("is-active", on);
    b.setAttribute("aria-checked", String(on));
  });
}

// 작성자를 CSS 클래스로 (준우 → who-junwoo)
function writerClass(name) {
  const w = norm(name);
  if (w === "준우") return "who-junwoo";
  if (w === "소히" || w === "소희") return "who-sohee";
  return "";
}

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
let recentSearchTerm = "";    // 내역 검색어 (분류·항목·메모)
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
  const num = Math.abs(Number(n || 0));   // 부호는 입금/인출 토글로 다룹니다
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
  // 표에만 있는 옛 항목도 고를 수 있게 남겨두되, 목록에 없는 값임을 표시합니다.
  const isForeign = selected && !list.includes(selected);
  if (isForeign) list.push(selected);
  for (const item of list) {
    const opt = document.createElement("option");
    opt.value = item;
    opt.textContent = (isForeign && item === selected) ? `${item} (목록에 없음)` : item;
    sel.appendChild(opt);
  }
  if (selected) sel.value = selected;
  sel.classList.toggle("is-foreign", Boolean(isForeign));
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
  syncFlowPick();
}

function populateDetail(detail) {
  const tree = currentCategoryTree();
  const major = el("major").value;
  const minor = el("minor").value;
  fillSelect(el("detail"), (tree[major] || {})[minor] || [], detail);
  updateItemStatus();
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
  setFlow("in");
  recentSearchTerm = "";
  el("recentSearch").value = "";
  el("searchClear").hidden = true;
  allRowsCache = null;
  await loadRows(true);
  populateRecentMonths();
  loadQuickItems(false);
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
// 사용한 날짜(내역의 날짜 칸) 기준 최신순으로 정렬합니다.
// 같은 날이면 나중에 입력한 것이 위로 옵니다.
function sortByUsedDate(rows) {
  return rows.slice().sort((a, b) => {
    const da = toDateInput(a.values[0]);
    const db = toDateInput(b.values[0]);
    if (da !== db) return db < da ? -1 : 1;   // 날짜 내림차순
    return b.index - a.index;                  // 같은 날 → 입력 순서 역순
  });
}

// 표 전체를 받아 캐시합니다. 날짜순 정렬이 필요하므로 전체가 있어야 합니다.
async function fetchAllRows(fileId, tableName) {
  const data = await graphFetch(`${tablePath(fileId, tableName)}/rows`);
  return (data.value || []).map((r, i) => ({
    index: typeof r.index === "number" ? r.index : i,
    values: (r.values && r.values[0]) || []
  }));
}

async function loadRows(reset) {
  if (isLoading) return;
  isLoading = true;
  const list = el("recentList");
  const moreBtn = el("moreBtn");
  moreBtn.disabled = true;

  const tableName = currentTableName();
  try {
    const fileId = await getFileId();

    if (reset || allRowsCache === null) {
      list.innerHTML = "<li class='muted'>불러오는 중...</li>";
      el("listMeta").textContent = "";
      allRowsCache = await fetchAllRows(fileId, tableName);
    } else {
      moreBtn.textContent = "불러오는 중...";
    }

    const filtered = sortByUsedDate(
      allRowsCache
        .filter((r) => !recentMonthFilter || toDateInput(r.values[0]).startsWith(recentMonthFilter))
        .filter((r) => rowMatchesSearch(r, recentSearchTerm))
    );

    totalRows = filtered.length;
    if (reset) shownCount = FIRST_PAGE_SIZE;
    else shownCount += PAGE_SIZE;
    loadedRows = filtered.slice(0, shownCount);
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

// 데이터가 있는 월 목록을 만듭니다. loadRows가 받아둔 캐시를 재사용합니다.
async function populateRecentMonths() {
  const sel = el("recentMonth");
  const keep = sel.value;
  try {
    if (allRowsCache === null) {
      const fileId = await getFileId();
      allRowsCache = await fetchAllRows(fileId, currentTableName());
    }

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

function rowMatchesSearch(row, term) {
  if (!term) return true;
  const [date, major, minor, detail, memo, amount, writer] = row.values;
  const hay = [
    toDateInput(date),
    norm(major), norm(minor), norm(detail), norm(memo), norm(writer),
    String(Number(amount || 0)),                       // 15000 으로 검색
    Number(amount || 0).toLocaleString("ko-KR")        // 15,000 으로도 검색
  ].join(" ").toLowerCase();
  // 공백으로 나눈 여러 단어가 모두 포함되어야 합니다.
  return term.toLowerCase().split(/\s+/).filter(Boolean).every((w) => hay.includes(w));
}

function renderRows() {
  const list = el("recentList");
  list.innerHTML = "";

  if (loadedRows.length === 0) {
    let msg, sub;
    if (recentSearchTerm) {
      msg = "찾는 내역이 없어요.";
      sub = "다른 말로 검색해보세요.";
    } else if (recentMonthFilter) {
      msg = "이 달에는 내역이 없어요.";
      sub = "다른 달을 골라보세요.";
    } else {
      msg = entryMode === "income" ? "아직 소득·저축 내역이 없어요." : "아직 지출 내역이 없어요.";
      sub = "위에서 첫 항목을 남겨보세요.";
    }
    list.innerHTML = `
      <li class="empty-state">
        <img class="empty-img" src="icons/char-empty.png" alt="">
        <p class="empty-msg">${msg}</p>
        <p class="empty-sub">${sub}</p>
      </li>`;
  }

  for (const row of loadedRows) {
    const [date, major, minor, detail, memo, amount, writer] = row.values;
    const li = document.createElement("li");
    const wc = writerClass(writer);
    if (wc) li.classList.add(wc);
    if (row.index === editingIndex) li.classList.add("is-editing");

    const dateEl = document.createElement("span");
    dateEl.className = "rdate";
    dateEl.textContent = toDateInput(date);
    if (norm(writer)) {
      const wb = document.createElement("span");
      wb.className = "rwho " + wc;
      wb.textContent = norm(writer);
      dateEl.appendChild(wb);
    }

    const catEl = document.createElement("span");
    catEl.className = "rcat";
    catEl.textContent = [major, minor, detail].filter(Boolean).join(" · ");

    const amtEl = document.createElement("span");
    amtEl.className = "ramt";
    const amtNum = Number(amount || 0);
    amtEl.textContent = amtNum.toLocaleString("ko-KR") + "원";
    if (amtNum < 0) amtEl.classList.add("is-out");   // 인출은 눈에 띄게

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
  const scope = recentSearchTerm ? "검색 결과" : (recentMonthFilter ? "이 달" : "전체");
  el("listMeta").textContent = totalRows
    ? `${scope} ${totalRows.toLocaleString("ko-KR")}건 중 ${loadedRows.length.toLocaleString("ko-KR")}건`
    : "";
}

// ── 수정 모드 ──────────────────────────────────────────────
function startEdit(row) {
  const [date, major, minor, detail, memo, amount, writer] = row.values;
  editingIndex = row.index;
  editingOriginal = row.values.slice();
  if (norm(writer)) setWriter(norm(writer) === "소희" ? "소히" : norm(writer));

  el("date").value = toDateInput(date);
  populateMajor(norm(major), norm(minor), norm(detail));
  el("memo").value = norm(memo);
  const amt = Number(amount || 0);
  setAmountValue(Math.abs(amt));
  syncFlowPick();
  setFlow(amt < 0 ? "out" : "in");

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

// 분류 조합이 실제로 존재하는지 확인합니다.
// (브라우저 폼 복원이나 옛 데이터 때문에 어긋난 조합이 저장되는 것을 막습니다)
function validateCategory() {
  const tree = currentCategoryTree();
  const major = el("major").value;
  const minor = el("minor").value;
  const detail = el("detail").value;

  if (!tree[major]) return `대분류 "${major}"를 찾을 수 없습니다.`;

  const needsMinor = minorFieldNeeded(tree, major);
  if (needsMinor) {
    if (!tree[major][minor]) {
      return `"${major}" 아래에 "${minor}" 소분류가 없습니다. 분류를 다시 선택해주세요.`;
    }
    if (!tree[major][minor].includes(detail)) {
      return `"${major} · ${minor}" 아래에 "${detail}" 항목이 없습니다. 분류를 다시 선택해주세요.`;
    }
  } else {
    const details = tree[major][NO_MINOR] || [];
    if (!details.includes(detail)) {
      return `"${major}" 아래에 "${detail}" 항목이 없습니다. 분류를 다시 선택해주세요.`;
    }
  }
  return null;
}

function formValues() {
  return [
    el("date").value,
    el("major").value,
    el("minor").value,
    el("detail").value,
    el("memo").value,
    (flowPickAvailable() && amountFlow === "out") ? -getAmountValue() : getAmountValue(),
    currentWriter
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

  // 분류 조합이 어긋나면 계획 시트의 SUMIFS가 집계하지 못하므로 미리 막습니다.
  const catError = validateCategory();
  if (catError) {
    showStatus(catError, true);
    el("major").focus();
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
    calLoadedKey = null;   // 달력도 다시 읽도록
    clearStatusCache();    // 실적이 바뀌었으니 항목 현황도 새로
    await loadRows(true);
    populateRecentMonths();
    loadQuickItems(true);
    updateItemStatus();
    loadMonthCard();
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
    calLoadedKey = null;   // 달력도 다시 읽도록
    clearStatusCache();
    await loadRows(true); // 삭제하면 뒤쪽 행 index가 밀리므로 전체를 다시 읽습니다
    populateRecentMonths();
    updateItemStatus();
    loadMonthCard();
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

// ═══════════════════════════════════════════════════════════
// 이번 달 지출 요약 카드
// ═══════════════════════════════════════════════════════════
// 저축·투자를 뺀 순수 지출(고정비+생활비+비정기)만 봅니다.
const SPEND_MAJORS = ["고정비", "생활비", "비정기 지출"];

// 이번 달 남은 날짜 (오늘 포함)
function daysLeftInMonth() {
  const now = new Date();
  const last = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  return last - now.getDate() + 1;
}

async function loadMonthCard() {
  const card = el("monthCard");
  try {
    const st = await getStructure();
    const now = new Date();
    const y = now.getFullYear(), m = now.getMonth() + 1;
    const ym = `${y}-${String(m).padStart(2, "0")}`;
    const cols = columnsFor(st, "month", y, m);
    if (!cols) { card.hidden = true; return; }

    if (!statusCache[ym]) {
      const allRows = st.groups.flatMap((g) => g.rows).map((r) => r.row);
      statusCache[ym] = await fetchSummaryRange(cols, allRows);
    }
    const range = statusCache[ym];

    let plan = 0, actual = 0;
    for (const g of st.groups) {
      if (!SPEND_MAJORS.includes(g.major)) continue;
      for (const r of g.rows) {
        plan += Number(cellFromRange(range, cols[0], r.row) || 0);
        actual += Number(cellFromRange(range, cols[1], r.row) || 0);
      }
    }
    renderMonthCard(m, plan, actual);
  } catch (e) {
    card.hidden = true;   // 부가 정보이므로 조용히 숨깁니다
  }
}

function renderMonthCard(month, plan, actual) {
  const card = el("monthCard");
  if (plan <= 0 && actual <= 0) { card.hidden = true; return; }

  const pct = plan > 0 ? Math.round((actual / plan) * 100) : 0;
  const left = plan - actual;
  const days = daysLeftInMonth();

  // 지출은 적게 쓸수록 좋습니다.
  const tone = plan <= 0 ? "none" : (pct > 100 ? "bad" : (pct >= 80 ? "warn" : "good"));

  const fill = Math.min(100, pct);
  const over = pct > 100 ? Math.min(100, pct - 100) : 0;

  const leftLabel = plan <= 0
    ? "계획 없음"
    : (left >= 0 ? `${shortWon(left)}원 남음` : `${shortWon(-left)}원 초과`);

  card.hidden = false;
  card.className = `month-card tone-${tone}`;
  card.innerHTML = `
    <div class="mc-top">
      <span class="mc-title">${month}월 지출</span>
      <span class="mc-days">${days}일 남음</span>
    </div>
    <div class="mc-main">
      <span class="mc-amt">${shortWon(actual)}</span>
      <span class="mc-plan">/ 계획 ${shortWon(plan)}</span>
    </div>
    <div class="mc-bar">
      <span class="mc-fill" style="width:${fill}%"></span>
      ${over ? `<span class="mc-over" style="width:${over}%"></span>` : ""}
    </div>
    <div class="mc-foot">
      <span class="mc-left">${leftLabel}</span>
      ${plan > 0 ? `<span class="mc-pct">${pct}%</span>` : ""}
    </div>
  `;
  card.title = `실적 ${fmtWon(actual)} · 계획 ${fmtWon(plan)}`;
}

// ═══════════════════════════════════════════════════════════
// 선택한 분류의 이번 달 계획 대비 실적 (라벨 옆 간략 표시)
// ═══════════════════════════════════════════════════════════
let statusCache = {};
let statusToken = 0;

function statusMonthKey() {
  const d = el("date").value;
  return /^\d{4}-\d{2}/.test(d) ? d.slice(0, 7) : null;
}

// 좁은 자리에 넣기 위해 만 단위로 줄입니다 (164,340 → 16.4만)
function shortWon(n) {
  const v = Math.round(Number(n || 0));
  const sign = v < 0 ? "-" : "";
  const a = Math.abs(v);
  if (a === 0) return "0";
  if (a >= 100000000) return sign + (a / 100000000).toFixed(1).replace(/\.0$/, "") + "억";
  if (a >= 10000) {
    const man = a / 10000;
    return sign + (man >= 100 ? Math.round(man) : man.toFixed(1).replace(/\.0$/, "")) + "만";
  }
  return sign + a.toLocaleString("ko-KR");
}

// 진행률 도넛 (작은 원형 게이지)
function donutSVG(pct, tone) {
  const p = Math.max(0, Math.min(100, pct));
  const r = 7, C = 2 * Math.PI * r;
  const dash = (p / 100) * C;
  return `<svg class="fs-donut ${tone}" viewBox="0 0 20 20" aria-hidden="true">
    <circle cx="10" cy="10" r="${r}" class="fs-track"/>
    <circle cx="10" cy="10" r="${r}" class="fs-arc"
      stroke-dasharray="${dash.toFixed(2)} ${(C - dash).toFixed(2)}"
      transform="rotate(-90 10 10)"/>
  </svg>`;
}

function toneOf(plan, actual, isIncome) {
  if (plan <= 0) return actual > 0 ? "bad" : "none";
  const pct = (actual / plan) * 100;
  if (isIncome) return pct >= 100 ? "good" : (pct >= 70 ? "warn" : "bad");
  return pct > 100 ? "bad" : (pct >= 80 ? "warn" : "good");
}

function paintStatus(elm, plan, actual, isIncome) {
  if (!elm) return;
  if (plan <= 0 && actual <= 0) { elm.hidden = true; return; }

  const pct = plan > 0 ? Math.round((actual / plan) * 100) : 0;
  const left = plan - actual;
  const tone = toneOf(plan, actual, isIncome);

  const leftLabel = plan <= 0
    ? "계획없음"
    : (left >= 0 ? `잔액 ${shortWon(left)}` : `초과 ${shortWon(-left)}`);

  elm.hidden = false;
  elm.className = `fs tone-${tone}`;
  elm.innerHTML = `${donutSVG(pct, tone)}<span class="fs-fig">${shortWon(actual)}/${shortWon(plan)}</span><span class="fs-left">${leftLabel}</span>`;
  elm.title = `계획 ${fmtWon(plan)} · 실적 ${fmtWon(actual)}${plan > 0 ? ` (${pct}%)` : ""}`;
}

// ── 저축·투자의 입금 / 인출 ────────────────────────────────
// 비상금에서 꺼내 쓰는 것처럼, 쌓아둔 돈을 헐어 쓰는 경우를 음수로 기록합니다.
let amountFlow = "in";   // "in" = 입금(+), "out" = 인출(−)

// 저축·투자 항목에서만 인출을 고를 수 있게 합니다.
function flowPickAvailable() {
  return entryMode === "income" && el("major").value === "저축·투자";
}

function setFlow(dir) {
  amountFlow = (dir === "out") ? "out" : "in";
  document.querySelectorAll(".flow-btn").forEach((b) => {
    b.classList.toggle("is-active", b.dataset.flow === amountFlow);
  });
  el("amount").classList.toggle("is-out", amountFlow === "out");
}

function syncFlowPick() {
  const box = el("flowPick");
  if (!box) return;
  const avail = flowPickAvailable();
  box.hidden = !avail;
  if (!avail) setFlow("in");   // 대상이 아니면 항상 입금으로
}

// 계획 금액을 자동으로 채워줄 항목
// 매달 금액이 달라져 시트를 확인해야 하는 ISA 두 계좌만 대상으로 합니다.
const AUTOFILL_DETAILS = ["소희 ISA", "준우 ISA"];

function isAutofillTarget(major, detail) {
  return INCOME_MAJORS.includes(major) && AUTOFILL_DETAILS.includes(norm(detail));
}

// 계획 금액을 금액 칸에 채웁니다.
// 사용자가 이미 입력한 값이 있으면 건드리지 않습니다.
function autofillPlanAmount(plan) {
  if (!plan || plan <= 0) return;
  if (editingIndex !== null) return;        // 수정 중에는 원래 금액을 지키기
  if (amountFlow === "out") return;         // 인출 중에는 계획 금액을 넣지 않습니다
  if (getAmountValue() > 0) return;         // 이미 입력한 값이 있으면 그대로

  setAmountValue(plan);
  const box = el("amount");
  box.classList.add("is-autofilled");
  setTimeout(() => box.classList.remove("is-autofilled"), 1200);
  showStatus(`계획 금액 ${fmtWon(plan)}을 넣었어요. 실제 금액이 다르면 고쳐주세요.`, false);
}

async function updateItemStatus() {
  const major = el("major").value;
  const minor = el("minor").value;
  const detail = el("detail").value;
  const ym = statusMonthKey();

  const boxes = [el("stMajor"), el("stMinor"), el("stDetail")];
  if (!major || !ym) { boxes.forEach((b) => b && (b.hidden = true)); return; }

  const my = ++statusToken;
  try {
    const st = await getStructure();
    const group = st.groups.find((g) => g.major === major);
    if (!group) { boxes.forEach((b) => b && (b.hidden = true)); return; }

    const [y, m] = ym.split("-").map(Number);
    const cols = columnsFor(st, "month", y, m);
    if (!cols) { boxes.forEach((b) => b && (b.hidden = true)); return; }

    if (!statusCache[ym]) {
      const allRows = st.groups.flatMap((g) => g.rows).map((r) => r.row);
      statusCache[ym] = await fetchSummaryRange(cols, allRows);
    }
    if (my !== statusToken) return;

    const range = statusCache[ym];
    const isIncome = INCOME_MAJORS.includes(major);
    const val = (row, ci) => Number(cellFromRange(range, cols[ci], row) || 0);

    // 대분류: 그룹 전체 합
    let gPlan = 0, gActual = 0;
    for (const r of group.rows) { gPlan += val(r.row, 0); gActual += val(r.row, 1); }
    paintStatus(el("stMajor"), gPlan, gActual, isIncome);

    // 소분류: 시트에 소계 행이 없으므로 해당 소분류의 세부항목을 합산합니다.
    const needsMinor = minorFieldNeeded(currentCategoryTree(), major);
    if (needsMinor && minor) {
      let mPlan = 0, mActual = 0;
      for (const r of group.rows) {
        if (r.minor !== minor) continue;
        mPlan += val(r.row, 0); mActual += val(r.row, 1);
      }
      paintStatus(el("stMinor"), mPlan, mActual, isIncome);
    } else if (el("stMinor")) {
      el("stMinor").hidden = true;
    }

    // 세부항목: 해당 행 하나
    const target = group.rows.find(
      (r) => r.detail === detail && (!needsMinor || r.minor === minor)
    );
    if (target) {
      const tPlan = val(target.row, 0);
      paintStatus(el("stDetail"), tPlan, val(target.row, 1), isIncome);
      if (isAutofillTarget(major, detail)) autofillPlanAmount(tPlan);
    } else if (el("stDetail")) {
      el("stDetail").hidden = true;
    }
  } catch (e) {
    boxes.forEach((b) => b && (b.hidden = true));
  }
}

function clearStatusCache() {
  statusCache = {};
}

// ═══════════════════════════════════════════════════════════
// 자주 쓰는 항목 (빠른 입력)
// ═══════════════════════════════════════════════════════════
// 최근 기록에서 자주 등장한 분류 조합을 뽑아 한 번에 채웁니다.
const QUICK_SCAN = 120;   // 최근 몇 건을 살펴볼지
const QUICK_MAX = 6;      // 칩 최대 개수

let quickCache = { expense: null, income: null };

function quickSignature(v) {
  return [norm(v[1]), norm(v[2]), norm(v[3])].join("|");
}

// 최근 것일수록 점수를 더 줍니다 (오래된 습관보다 요즘 습관이 중요)
function buildQuickItems(rows) {
  const recent = rows.slice(-QUICK_SCAN);
  const score = new Map();

  recent.forEach((r, i) => {
    const v = r.values || [];
    const detail = norm(v[3]);
    if (!detail) return;
    const sig = quickSignature(v);
    // 뒤쪽(최근)일수록 1.0 → 앞쪽(과거)은 0.4까지 낮춥니다.
    const w = 0.4 + 0.6 * (i / Math.max(1, recent.length - 1));
    const cur = score.get(sig) || {
      major: norm(v[1]), minor: norm(v[2]), detail,
      count: 0, weight: 0, amounts: []
    };
    cur.count += 1;
    cur.weight += w;
    const amt = Number(v[5] || 0);
    if (amt > 0) cur.amounts.push(amt);
    score.set(sig, cur);
  });

  return Array.from(score.values())
    .filter((x) => x.count >= 2)          // 한 번뿐인 건 제외
    .sort((a, b) => b.weight - a.weight)
    .slice(0, QUICK_MAX);
}

async function loadQuickItems(force) {
  const box = el("quickBox");
  const mode = entryMode;

  if (!force && quickCache[mode]) {
    renderQuickChips(quickCache[mode]);
    return;
  }

  try {
    const fileId = await getFileId();
    const data = await graphFetch(`${tablePath(fileId, currentTableName())}/rows`);
    const rows = (data.value || []).map((r, i) => ({
      index: typeof r.index === "number" ? r.index : i,
      values: (r.values && r.values[0]) || []
    }));
    const items = buildQuickItems(rows);
    quickCache[mode] = items;
    if (mode === entryMode) renderQuickChips(items);
  } catch (e) {
    box.hidden = true;   // 실패해도 조용히 숨깁니다 (부가 기능이므로)
  }
}

function renderQuickChips(items) {
  const box = el("quickBox");
  const wrap = el("quickChips");

  if (!items || !items.length) { box.hidden = true; return; }
  box.hidden = false;
  wrap.innerHTML = "";

  for (const it of items) {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "quick-chip";

    // "준우"처럼 세부항목만으로 뜻이 안 통하면 소분류를 앞에 세웁니다.
    const vague = it.minor && (it.detail === "준우" || it.detail === "소희" || it.detail === "소히");
    const name = document.createElement("span");
    name.className = "qc-name";
    name.textContent = vague ? `${it.minor} · ${it.detail}` : it.detail;

    const sub = document.createElement("span");
    sub.className = "qc-sub";
    sub.textContent = vague ? it.major : (it.minor || it.major);

    chip.append(name, sub);
    chip.title = [it.major, it.minor, it.detail].filter(Boolean).join(" · ");

    chip.addEventListener("click", () => applyQuickItem(it, chip));
    wrap.appendChild(chip);
  }
}

function applyQuickItem(it, chip) {
  // 수정 중이면 헷갈리니 먼저 빠져나옵니다.
  if (editingIndex !== null) cancelEdit();

  populateMajor(it.major, it.minor, it.detail);

  // 눌린 표시를 잠깐 보여줍니다.
  document.querySelectorAll(".quick-chip").forEach((c) => c.classList.remove("is-picked"));
  if (chip) {
    chip.classList.add("is-picked");
    setTimeout(() => chip.classList.remove("is-picked"), 900);
  }

  el("amount").focus();
}

// ═══════════════════════════════════════════════════════════
// 고정비 한번에 입력
// ═══════════════════════════════════════════════════════════
// 매달 비슷하게 나가는 고정비를 계획 시트에서 불러와 한 번에 넣습니다.
const FIXED_MAJORS = ["고정비"];   // 대상 대분류

let fixedItems = [];        // { minor, detail, plan, checked, already }
let fixedWriter = null;

function fixedMonthKey() {
  return el("fixedMonth").value;   // "2026-07"
}

async function openFixedSheet() {
  el("fixedSheet").hidden = false;
  document.body.classList.add("sheet-open");

  fixedWriter = currentWriter;
  document.querySelectorAll("#fixedWho .who-btn").forEach((b) => {
    const on = b.dataset.who === fixedWriter;
    b.classList.toggle("is-active", on);
    b.setAttribute("aria-checked", String(on));
  });

  const list = el("fixedList");
  list.innerHTML = "<p class='fixed-loading'>불러오는 중...</p>";
  el("fixedSummary").textContent = "";

  try {
    const st = await getStructure();

    // 월 선택지 채우기 (시트에 있는 달만)
    const sel = el("fixedMonth");
    if (!sel.options.length) {
      const months = Object.keys(st.monthCols).sort();
      for (const m of months) {
        const [y, mo] = m.split("-");
        const opt = document.createElement("option");
        opt.value = m;
        opt.textContent = `${y}년 ${Number(mo)}월`;
        sel.appendChild(opt);
      }
      const now = new Date();
      const cur = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}`;
      sel.value = st.monthCols[cur] ? cur : months[months.length - 1];
    }

    await loadFixedItems(st);
  } catch (e) {
    list.innerHTML = `<p class='fixed-loading'>불러오지 못했습니다: ${e.message}</p>`;
  }
}

function closeFixedSheet() {
  const sheet = el("fixedSheet");
  if (!sheet) return;
  sheet.hidden = true;
  document.body.classList.remove("sheet-open");
}

async function loadFixedItems(st) {
  const list = el("fixedList");
  list.innerHTML = "<p class='fixed-loading'>불러오는 중...</p>";

  const key = fixedMonthKey();
  const cols = st.monthCols[key];
  if (!cols) {
    list.innerHTML = "<p class='fixed-loading'>이 달은 계획 시트에 없습니다.</p>";
    fixedItems = [];
    updateFixedSummary();
    return;
  }

  const groups = st.groups.filter((g) => FIXED_MAJORS.includes(g.major));
  const rows = groups.flatMap((g) => g.rows);
  if (!rows.length) {
    list.innerHTML = "<p class='fixed-loading'>고정비 항목을 찾지 못했습니다.</p>";
    fixedItems = [];
    updateFixedSummary();
    return;
  }

  const fileId = await getFileId();

  // 계획 금액 읽기 + 이미 입력된 내역 확인을 동시에
  const [planRange, existing] = await Promise.all([
    fetchSummaryRange(cols, rows.map((r) => r.row)),
    graphFetch(`${tablePath(fileId, APP_CONFIG.tableName)}/rows`)
  ]);

  const already = new Set();
  for (const r of (existing.value || [])) {
    const v = (r.values && r.values[0]) || [];
    if (!toDateInput(v[0]).startsWith(key)) continue;
    already.add([norm(v[1]), norm(v[2]), norm(v[3])].join("|"));
  }

  fixedItems = rows.map((r) => {
    const plan = Number(cellFromRange(planRange, cols[0], r.row) || 0);
    const sig = ["고정비", r.minor, r.detail].join("|");
    const dup = already.has(sig);
    return {
      minor: r.minor,
      detail: r.detail,
      plan,
      amount: plan,
      checked: !dup && plan > 0,   // 이미 넣었거나 계획이 0이면 기본 해제
      already: dup
    };
  });

  renderFixedList();
}

function renderFixedList() {
  const list = el("fixedList");
  list.innerHTML = "";

  fixedItems.forEach((item, i) => {
    const row = document.createElement("div");
    row.className = "fixed-item" + (item.already ? " is-done" : "");

    const label = document.createElement("label");
    label.className = "fixed-check";

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = item.checked;
    cb.addEventListener("change", () => {
      fixedItems[i].checked = cb.checked;
      row.classList.toggle("is-on", cb.checked);
      updateFixedSummary();
    });

    const nameBox = document.createElement("span");
    nameBox.className = "fixed-name";
    nameBox.innerHTML = `
      <b>${item.detail}</b>
      <em>${item.minor}</em>
      ${item.already ? '<span class="fixed-done-tag">입력됨</span>' : ""}
    `;

    label.append(cb, nameBox);

    const amtWrap = document.createElement("div");
    amtWrap.className = "fixed-amt";
    const amt = document.createElement("input");
    amt.type = "text";
    amt.inputMode = "numeric";
    amt.value = item.amount ? item.amount.toLocaleString("ko-KR") : "";
    amt.placeholder = "0";
    amt.addEventListener("input", (e) => {
      const digits = String(e.target.value).replace(/[^0-9]/g, "");
      e.target.value = digits ? Number(digits).toLocaleString("ko-KR") : "";
      fixedItems[i].amount = digits ? Number(digits) : 0;
      updateFixedSummary();
    });
    const won = document.createElement("span");
    won.textContent = "원";
    amtWrap.append(amt, won);

    row.append(label, amtWrap);
    if (item.checked) row.classList.add("is-on");
    list.appendChild(row);
  });

  updateFixedSummary();
}

function updateFixedSummary() {
  const picked = fixedItems.filter((i) => i.checked && i.amount > 0);
  const sum = picked.reduce((s, i) => s + i.amount, 0);
  el("fixedSummary").innerHTML = picked.length
    ? `<b>${picked.length}건</b> · ${fmtWon(sum)}`
    : "선택한 항목이 없어요";
  el("fixedSubmit").disabled = picked.length === 0;
}

async function submitFixed() {
  const picked = fixedItems.filter((i) => i.checked && i.amount > 0);
  if (!picked.length) return;

  const btn = el("fixedSubmit");
  btn.disabled = true;
  btn.textContent = "저장하는 중...";

  const key = fixedMonthKey();
  const [y, m] = key.split("-");
  // 그 달 1일로 기록합니다.
  const dateStr = `${y}-${m}-01`;

  try {
    const fileId = await getFileId();
    const values = picked.map((i) => [
      dateStr, "고정비", i.minor, i.detail, "고정비 자동입력", i.amount, fixedWriter
    ]);

    // 표에 여러 행을 한 번에 넣습니다.
    await graphFetch(`${tablePath(fileId, APP_CONFIG.tableName)}/rows/add`, {
      method: "POST",
      body: JSON.stringify({ values })
    });
    await recalc(fileId);

    closeFixedSheet();
    showToast("saved", `${picked.length}건 저장 완료!`);

    allRowsCache = null;
    calLoadedKey = null;
    clearStatusCache();
    await loadRows(true);
    populateRecentMonths();
    loadQuickItems(true);
    updateItemStatus();
    loadMonthCard();
  } catch (e) {
    alert("저장하지 못했습니다: " + e.message);
  } finally {
    btn.disabled = false;
    btn.textContent = "선택 항목 저장";
  }
}

// ═══════════════════════════════════════════════════════════
// 달력 보기
// ═══════════════════════════════════════════════════════════
let calYear = null;
let calMonth = null;          // 1~12
let calMode = "expense";      // "expense" | "income"
let calRows = [];             // 그 달의 행들
let calSelectedDay = null;    // 선택한 날짜(1~31)
let calLoadedKey = null;      // "expense-2026-07" 형태, 중복 로딩 방지

function calTableName() {
  return calMode === "income" ? APP_CONFIG.incomeTableName : APP_CONFIG.tableName;
}

function calKey() {
  return `${calMode}-${calYear}-${String(calMonth).padStart(2, "0")}`;
}

function initCalendarDate() {
  if (calYear) return;
  const now = new Date();
  calYear = now.getFullYear();
  calMonth = now.getMonth() + 1;
}

function calShiftMonth(delta) {
  calMonth += delta;
  if (calMonth < 1) { calMonth = 12; calYear -= 1; }
  else if (calMonth > 12) { calMonth = 1; calYear += 1; }
  calSelectedDay = null;
  loadCalendar();
}

async function loadCalendar() {
  initCalendarDate();
  const grid = el("calGrid");
  const key = calKey();

  el("calTitle").textContent = `${calYear}년 ${calMonth}월`;
  document.querySelectorAll(".cal-mode-btn").forEach((b) => {
    b.classList.toggle("is-active", b.dataset.calmode === calMode);
  });

  if (calLoadedKey === key && calRows.length >= 0) {
    renderCalendar();
    return;
  }

  grid.innerHTML = "<p class='cal-loading'>불러오는 중...</p>";
  el("calTotal").textContent = "";
  el("calDetail").hidden = true;

  try {
    const fileId = await getFileId();
    const data = await graphFetch(`${tablePath(fileId, calTableName())}/rows`);
    const prefix = `${calYear}-${String(calMonth).padStart(2, "0")}`;
    calRows = (data.value || [])
      .map((r, i) => ({
        index: typeof r.index === "number" ? r.index : i,
        values: (r.values && r.values[0]) || []
      }))
      .filter((r) => toDateInput(r.values[0]).startsWith(prefix));
    calLoadedKey = key;
    renderCalendar();
  } catch (e) {
    grid.innerHTML = `<p class='cal-loading'>불러오지 못했습니다: ${e.message}</p>`;
  }
}

// 날짜별로 금액을 모읍니다.
function calDayTotals() {
  const map = {};
  for (const r of calRows) {
    const d = toDateInput(r.values[0]);
    const day = Number(d.slice(8, 10));
    if (!day) continue;
    if (!map[day]) map[day] = { sum: 0, items: [] };
    map[day].sum += Number(r.values[5] || 0);
    map[day].items.push(r);
  }
  return map;
}

// 공휴일이면 이름을 돌려줍니다.
function holidayName(dateStr) {
  return (typeof KR_HOLIDAYS !== "undefined" && KR_HOLIDAYS[dateStr]) || null;
}

function renderCalendar() {
  const grid = el("calGrid");
  const totals = calDayTotals();
  const first = new Date(calYear, calMonth - 1, 1);
  const startDow = first.getDay();                     // 0=일
  const daysInMonth = new Date(calYear, calMonth, 0).getDate();
  const today = todayStr();

  // 월 합계
  const monthSum = calRows.reduce((s, r) => s + Number(r.values[5] || 0), 0);
  const label = calMode === "income" ? "이 달 수입·저축" : "이 달 지출";
  el("calTotal").innerHTML = `
    <span class="cal-total-label">${label}</span>
    <span class="cal-total-value">${fmtWon(monthSum)}</span>
    <span class="cal-total-count">${calRows.length}건</span>
  `;

  grid.innerHTML = "";

  // 앞쪽 빈 칸
  for (let i = 0; i < startDow; i++) {
    const cell = document.createElement("div");
    cell.className = "cal-cell is-blank";
    grid.appendChild(cell);
  }

  for (let d = 1; d <= daysInMonth; d++) {
    const dateStr = `${calYear}-${String(calMonth).padStart(2,"0")}-${String(d).padStart(2,"0")}`;
    const info = totals[d];
    const dow = (startDow + d - 1) % 7;

    const holiday = holidayName(dateStr);

    const cell = document.createElement("button");
    cell.type = "button";
    cell.className = "cal-cell";
    if (dow === 0) cell.classList.add("is-sun");
    if (dow === 6) cell.classList.add("is-sat");
    if (holiday) cell.classList.add("is-holiday");
    if (dateStr === today) cell.classList.add("is-today");
    if (calSelectedDay === d) cell.classList.add("is-selected");
    if (info) cell.classList.add("has-data");
    if (holiday) cell.title = holiday;

    const num = document.createElement("span");
    num.className = "cal-day";
    num.textContent = d;
    cell.appendChild(num);

    if (holiday) {
      const hd = document.createElement("span");
      hd.className = "cal-holi";
      // 칸이 좁아 짧게 줄입니다 (설날 연휴 → 설날)
      hd.textContent = holiday.replace(/\s*(연휴|대체)$/, "");
      cell.appendChild(hd);
    }

    if (info) {
      const amt = document.createElement("span");
      amt.className = "cal-amt";
      // 만 원 단위로 줄여 표시 (칸이 좁아서)
      amt.textContent = compactWon(info.sum);
      cell.appendChild(amt);
    }

    cell.addEventListener("click", () => {
      calSelectedDay = (calSelectedDay === d) ? null : d;
      renderCalendar();
    });

    grid.appendChild(cell);
  }

  renderCalDetail(totals);
}

// 좁은 칸에 넣기 위해 금액을 짧게 (12,345 → 1.2만)
function compactWon(n) {
  const v = Number(n || 0);
  if (v === 0) return "";
  if (v >= 100000000) return (v / 100000000).toFixed(1).replace(/\.0$/, "") + "억";
  if (v >= 10000) {
    const man = v / 10000;
    return (man >= 100 ? Math.round(man) : man.toFixed(1).replace(/\.0$/, "")) + "만";
  }
  return v.toLocaleString("ko-KR");
}

function renderCalDetail(totals) {
  const box = el("calDetail");
  if (!calSelectedDay) { box.hidden = true; return; }

  const info = totals[calSelectedDay];
  const dateStr = `${calYear}-${String(calMonth).padStart(2,"0")}-${String(calSelectedDay).padStart(2,"0")}`;
  const dowNames = ["일","월","화","수","목","금","토"];
  const dow = dowNames[new Date(calYear, calMonth-1, calSelectedDay).getDay()];
  const holi = holidayName(dateStr);
  const holiTag = holi ? ` <span class="cd-holi">${holi}</span>` : "";

  box.hidden = false;

  if (!info) {
    box.innerHTML = `
      <div class="cal-detail-head">
        <span class="cal-detail-date">${calMonth}월 ${calSelectedDay}일 (${dow})${holiTag}</span>
      </div>
      <p class="cal-detail-empty">이 날은 내역이 없어요.</p>`;
    return;
  }

  let html = `
    <div class="cal-detail-head">
      <span class="cal-detail-date">${calMonth}월 ${calSelectedDay}일 (${dow})${holiTag}</span>
      <span class="cal-detail-sum">${fmtWon(info.sum)}</span>
    </div>
    <ul class="cal-detail-list">`;

  for (const r of info.items) {
    const [, major, minor, detail, memo, amount, writer] = r.values;
    const cat = [major, minor, detail].filter(Boolean).join(" · ");
    const wc = writerClass(writer);
    const wTag = norm(writer) ? `<span class="rwho ${wc}">${norm(writer)}</span>` : "";
    html += `
      <li class="${wc}">
        <span class="cd-cat">${cat}${wTag}</span>
        <span class="cd-amt">${fmtWon(amount)}</span>
        ${memo ? `<span class="cd-memo">${norm(memo)}</span>` : ""}
      </li>`;
  }
  html += `</ul>`;
  box.innerHTML = html;
}

function switchTab(name) {
  el("tab-entry").hidden = name !== "entry";
  el("tab-calendar").hidden = name !== "calendar";
  el("tab-summary").hidden = name !== "summary";
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.classList.toggle("is-active", btn.dataset.tab === name);
  });
  if (name === "calendar") {
    loadCalendar();
  }
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
  loadQuickItems(false);
  updateItemStatus();
  loadMonthCard();
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
  el("detail").addEventListener("change", () => updateItemStatus());
  document.querySelectorAll(".flow-btn").forEach((btn) => {
    btn.addEventListener("click", () => setFlow(btn.dataset.flow));
  });
  el("date").addEventListener("change", () => updateItemStatus());
  el("entryForm").addEventListener("submit", handleSubmit);
  setupInstallButton();

  document.querySelectorAll("#whoPick .who-btn").forEach((btn) => {
    btn.addEventListener("click", () => setWriter(btn.dataset.who));
  });
  setWriter(currentWriter);   // 마지막에 고른 사람을 기억합니다

  // 고정비 한번에 입력
  el("fixedBtn").addEventListener("click", openFixedSheet);
  el("fixedClose").addEventListener("click", closeFixedSheet);
  el("fixedDim").addEventListener("click", closeFixedSheet);
  // 시트 본문 클릭은 배경 클릭으로 취급하지 않습니다.
  document.querySelector("#fixedSheet .sheet-body")
    .addEventListener("click", (e) => e.stopPropagation());
  el("fixedSubmit").addEventListener("click", submitFixed);
  el("fixedMonth").addEventListener("change", async () => {
    try {
      const st = await getStructure();
      await loadFixedItems(st);
    } catch (e) {
      el("fixedList").innerHTML = `<p class='fixed-loading'>${e.message}</p>`;
    }
  });
  document.querySelectorAll("#fixedWho .who-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      fixedWriter = btn.dataset.who;
      document.querySelectorAll("#fixedWho .who-btn").forEach((b) => {
        const on = b === btn;
        b.classList.toggle("is-active", on);
        b.setAttribute("aria-checked", String(on));
      });
    });
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !el("fixedSheet").hidden) closeFixedSheet();
  });

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

  // 검색: 타이핑이 멈춘 뒤에 한 번만 걸러 화면 깜빡임을 줄입니다.
  let searchTimer = null;
  el("recentSearch").addEventListener("input", (e) => {
    const v = e.target.value;
    el("searchClear").hidden = !v;
    if (searchTimer) clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      recentSearchTerm = v.trim();
      loadRows(true);
    }, 250);
  });

  el("searchClear").addEventListener("click", () => {
    el("recentSearch").value = "";
    el("searchClear").hidden = true;
    recentSearchTerm = "";
    loadRows(true);
    el("recentSearch").focus();
  });
  el("moreBtn").addEventListener("click", () => loadRows(false));
  el("loginBtn").addEventListener("click", () => msalInstance.loginRedirect({ scopes: GRAPH_SCOPES }));

  document.querySelectorAll(".mode-btn").forEach((btn) => {
    btn.addEventListener("click", () => setEntryMode(btn.dataset.mode));
  });

  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => switchTab(btn.dataset.tab));
  });

  el("calPrev").addEventListener("click", () => calShiftMonth(-1));
  el("calNext").addEventListener("click", () => calShiftMonth(1));
  document.querySelectorAll(".cal-mode-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      if (calMode === btn.dataset.calmode) return;
      calMode = btn.dataset.calmode;
      calSelectedDay = null;
      calLoadedKey = null;
      loadCalendar();
    });
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

// ── 홈 화면에 앱 설치 ──────────────────────────────────────
// 안드로이드/데스크톱 크롬은 조건이 맞으면 beforeinstallprompt를 보냅니다.
// 그 이벤트를 잡아뒀다가 버튼을 눌렀을 때 설치창을 띄웁니다.
function setupInstallButton() {
  const btn = document.getElementById("installBtn");
  if (!btn) return;
  // 이미 설치되어 실행 중이면 버튼을 숨깁니다.
  const standalone = window.matchMedia("(display-mode: standalone)").matches
    || window.navigator.standalone === true;
  if (standalone) { btn.hidden = true; return; }
  if (deferredInstall) btn.hidden = false;

  btn.addEventListener("click", async () => {
    if (!deferredInstall) return;
    deferredInstall.prompt();
    try { await deferredInstall.userChoice; } catch (e) {}
    deferredInstall = null;
    btn.hidden = true;
  });
}

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => navigator.serviceWorker.register("sw.js").catch(() => {}));
}

init();
