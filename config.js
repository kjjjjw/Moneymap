// Azure AD 앱 등록 후 아래 값을 채워주세요.
// portal.azure.com > App registrations > 등록한 앱 > Overview > "Application (client) ID"
const APP_CONFIG = {
clientId: "5d6de376-b2c3-40c7-a395-1b750b7c9925",  // 개인 Microsoft 계정만 사용하므로 authority는 고정값입니다. 변경하지 마세요.
  authority: "https://login.microsoftonline.com/consumers",
  // OneDrive에 올린 엑셀 파일명 (경로는 자동 검색합니다)
  fileName: "가계부_R2.xlsx",
  // 이용내역 상세 시트의 표(Table) 이름
  tableName: "이용내역상세",
  // 계획 대비 실적을 읽어오는 시트 이름
  summarySheet: "가계부_월별"
};

// ── 가계부_월별 시트 구조 ──────────────────────────────────
// A~C열: 대분류/소분류/세부항목, 행은 시트에 고정된 위치입니다.
// 그룹 순서와 소계행은 시트의 실제 배치를 그대로 따릅니다.
const SUMMARY_GROUPS = [
  {
    major: "소득",
    totalRow: 13,
    rows: [
      { row: 7, minor: "", detail: "준우 월급" },
      { row: 8, minor: "", detail: "준우 상여·성과급" },
      { row: 9, minor: "", detail: "소희 월급/휴직급여" },
      { row: 10, minor: "", detail: "기타" },
      { row: 11, minor: "", detail: "출산전후휴가급여" },
      { row: 12, minor: "", detail: "부모급여·아동수당·첫만남" }
    ]
  },
  {
    major: "저축·투자",
    totalRow: 21,
    rows: [
      { row: 14, minor: "비상금", detail: "현금 (다올)" },
      { row: 15, minor: "ISA", detail: "소희 ISA" },
      { row: 16, minor: "ISA", detail: "준우 ISA" },
      { row: 17, minor: "투자", detail: "소희" },
      { row: 18, minor: "투자", detail: "준우" },
      { row: 19, minor: "계", detail: "가족계 (소희)" },
      { row: 20, minor: "계", detail: "가족계 (준우)" }
    ]
  },
  {
    major: "고정비",
    totalRow: 28,
    rows: [
      { row: 22, minor: "주거비", detail: "주담대 원리금" },
      { row: 23, minor: "주거비", detail: "아파트 관리비" },
      { row: 24, minor: "통신비", detail: "준우" },
      { row: 25, minor: "통신비", detail: "소희" },
      { row: 26, minor: "보험비", detail: "준우" },
      { row: 27, minor: "보험비", detail: "소희" }
    ]
  },
  {
    major: "생활비",
    totalRow: 49,
    rows: [
      { row: 29, minor: "식비", detail: "개인식비 (준우)" },
      { row: 30, minor: "식비", detail: "개인식비 (소희)" },
      { row: 31, minor: "식비", detail: "장보기" },
      { row: 32, minor: "식비", detail: "외식비" },
      { row: 33, minor: "취미", detail: "준우 탁구장 이용료" },
      { row: 34, minor: "취미", detail: "소희 네일, 요가" },
      { row: 35, minor: "용돈", detail: "준우" },
      { row: 36, minor: "용돈", detail: "소희" },
      { row: 37, minor: "교통비", detail: "준우" },
      { row: 38, minor: "교통비", detail: "소희" },
      { row: 39, minor: "여행", detail: "여행" },
      { row: 40, minor: "육아비", detail: "육아 제반" },
      { row: 41, minor: "미용", detail: "미용" },
      { row: 42, minor: "생활", detail: "생필품" },
      { row: 43, minor: "건강", detail: "병원/약" },
      { row: 44, minor: "쇼핑", detail: "쇼핑" },
      { row: 45, minor: "구독", detail: "구독" },
      { row: 46, minor: "기타", detail: "기타" },
      { row: 47, minor: "계모임", detail: "맛집탐방·외식" },
      { row: 48, minor: "계모임", detail: "미술사 회비" }
    ]
  },
  {
    major: "비정기 지출",
    totalRow: 54,
    rows: [
      { row: 50, minor: "명절", detail: "명절 현금" },
      { row: 51, minor: "가족", detail: "가족 생일" },
      { row: 52, minor: "경조사", detail: "경조사" },
      { row: 53, minor: "세금", detail: "세금" }
    ]
  }
];

const GRAND_TOTAL_ROW = 55;   // 지출 합계 (시트 원본 그대로 - 소득·저축투자 제외)
const NET_ROW = 56;           // 월 수지
const CUMULATIVE_ROW = 57;    // 누적 수지

// 연-월 → [계획열, 실적열, 차이열] 매핑. 2026.07 ~ 2028.12, 시트에 실제 존재하는 범위만.
const MONTH_COLUMNS = {
  "2026-07": ["E", "F", "G"], "2026-08": ["H", "I", "J"], "2026-09": ["K", "L", "M"],
  "2026-10": ["N", "O", "P"], "2026-11": ["Q", "R", "S"], "2026-12": ["T", "U", "V"],
  "2027-01": ["Z", "AA", "AB"], "2027-02": ["AC", "AD", "AE"], "2027-03": ["AF", "AG", "AH"],
  "2027-04": ["AI", "AJ", "AK"], "2027-05": ["AL", "AM", "AN"], "2027-06": ["AO", "AP", "AQ"],
  "2027-07": ["AR", "AS", "AT"], "2027-08": ["AU", "AV", "AW"], "2027-09": ["AX", "AY", "AZ"],
  "2027-10": ["BA", "BB", "BC"], "2027-11": ["BD", "BE", "BF"], "2027-12": ["BG", "BH", "BI"],
  "2028-01": ["BM", "BN", "BO"], "2028-02": ["BP", "BQ", "BR"], "2028-03": ["BS", "BT", "BU"],
  "2028-04": ["BV", "BW", "BX"], "2028-05": ["BY", "BZ", "CA"], "2028-06": ["CB", "CC", "CD"],
  "2028-07": ["CE", "CF", "CG"], "2028-08": ["CH", "CI", "CJ"], "2028-09": ["CK", "CL", "CM"],
  "2028-10": ["CN", "CO", "CP"], "2028-11": ["CQ", "CR", "CS"], "2028-12": ["CT", "CU", "CV"]
};

// 연도 선택(전체 합계) 시 사용하는 열
const YEAR_COLUMNS = {
  "2026": ["W", "X", "Y"],
  "2027": ["BJ", "BK", "BL"],
  "2028": ["CW", "CX", "CY"]
};
