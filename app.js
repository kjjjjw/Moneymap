const CATEGORY_TREE = {
  "고정비": {
    "주거비": ["주담대 원리금", "아파트 관리비"],
    "통신비": ["준우", "소희"],
    "보험비": ["준우", "소희"]
  },
  "생활비": {
    "식비": ["개인식비 (준우)", "개인식비 (소희)", "장보기", "외식비"],
    "취미": ["준우 탁구장 이용료", "소히 네일, 요가"],
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

function todayStr() {
  const d = new Date();
  const tzOffsetMs = d.getTimezoneOffset() * 60000;
  return new Date(d - tzOffsetMs).toISOString().slice(0, 10);
}

function populateMajor() {
  const majorSel = el("major");
  majorSel.innerHTML = "";
  for (const major of Object.keys(CATEGORY_TREE)) {
    const opt = document.createElement("option");
    opt.value = major;
    opt.textContent = major;
    majorSel.appendChild(opt);
  }
  populateMinor();
}

function populateMinor() {
  const major = el("major").value;
  const minorSel = el("minor");
  minorSel.innerHTML = "";
  for (const minor of Object.keys(CATEGORY_TREE[major] || {})) {
    const opt = document.createElement("option");
    opt.value = minor;
    opt.textContent = minor;
    minorSel.appendChild(opt);
  }
  populateDetail();
}

function populateDetail() {
  const major = el("major").value;
  const minor = el("minor").value;
  const detailSel = el("detail");
  detailSel.innerHTML = "";
  const details = (CATEGORY_TREE[major] || {})[minor] || [];
  for (const detail of details) {
    const opt = document.createElement("option");
    opt.value = detail;
    opt.textContent = detail;
    detailSel.appendChild(opt);
  }
}

function showStatus(msg, isError) {
  const s = el("statusMsg");
  s.textContent = msg;
  s.className = "status" + (isError ? " error" : " success");
}

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
  const result = await graphFetch(`/me/drive/root/search(q='${q}')`);
  const match = (result.value || []).find((f) => f.name === APP_CONFIG.fileName);
  if (!match) throw new Error(`OneDrive에서 '${APP_CONFIG.fileName}' 파일을 찾을 수 없습니다.`);
  localStorage.setItem("gagyebu_fileId", match.id);
  return match.id;
}

function tablePath(fileId) {
  return `/me/drive/items/${fileId}/workbook/tables/${encodeURIComponent(APP_CONFIG.tableName)}`;
}

async function addRow(values) {
  const fileId = await getFileId();
  await graphFetch(`${tablePath(fileId)}/rows/add`, {
    method: "POST",
    body: JSON.stringify({ values: [values] })
  });
  try {
    await graphFetch(`/me/drive/items/${fileId}/workbook/application/calculate`, {
      method: "POST",
      body: JSON.stringify({ calculationType: "Recalculate" })
    });
  } catch (e) {
    // 재계산 실패는 무시 - 다음에 파일을 열면 자동으로 재계산됩니다.
  }
}

async function loadRecent() {
  const list = el("recentList");
  list.innerHTML = "<li class='muted'>불러오는 중...</li>";
  try {
    const fileId = await getFileId();
    const data = await graphFetch(`${tablePath(fileId)}/rows`);
    const rows = (data.value || []).slice(-10).reverse();
    if (rows.length === 0) {
      list.innerHTML = "<li class='muted'>입력된 내역이 없습니다.</li>";
      return;
    }
    list.innerHTML = "";
    for (const row of rows) {
      const [date, major, minor, detail, memo, amount] = row.values[0];
      const li = document.createElement("li");
      const amountStr = Number(amount).toLocaleString("ko-KR");
      li.innerHTML = `<span class="rdate">${date}</span><span class="rcat">${major} · ${minor} · ${detail}</span><span class="ramt">${amountStr}원</span>${memo ? `<span class="rmemo">${memo}</span>` : ""}`;
      list.appendChild(li);
    }
  } catch (e) {
    list.innerHTML = `<li class='muted'>불러오기 실패: ${e.message}</li>`;
  }
}

async function handleSubmit(e) {
  e.preventDefault();
  const submitBtn = el("submitBtn");
  submitBtn.disabled = true;
  showStatus("저장 중...", false);
  try {
    const values = [
      el("date").value,
      el("major").value,
      el("minor").value,
      el("detail").value,
      el("memo").value,
      Number(el("amount").value)
    ];
    await addRow(values);
    showStatus("저장되었습니다.", false);
    el("amount").value = "";
    el("memo").value = "";
    loadRecent();
  } catch (e) {
    showStatus(e.message, true);
  } finally {
    submitBtn.disabled = false;
  }
}

function showApp() {
  el("loginArea").hidden = true;
  el("mainArea").hidden = false;
  el("authArea").innerHTML = `<span class="who">${msalInstance.getAllAccounts()[0].username}</span> <button type="button" id="logoutBtn" class="linklike">로그아웃</button>`;
  el("logoutBtn").addEventListener("click", () => {
    localStorage.removeItem("gagyebu_fileId");
    msalInstance.logoutRedirect();
  });
  loadRecent();
}

function showLogin() {
  el("loginArea").hidden = false;
  el("mainArea").hidden = true;
  el("authArea").innerHTML = "";
}

async function init() {
  el("date").value = todayStr();
  populateMajor();
  el("major").addEventListener("change", populateMinor);
  el("minor").addEventListener("change", populateDetail);
  el("entryForm").addEventListener("submit", handleSubmit);
  el("refreshBtn").addEventListener("click", loadRecent);
  el("loginBtn").addEventListener("click", () => msalInstance.loginRedirect({ scopes: GRAPH_SCOPES }));

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
