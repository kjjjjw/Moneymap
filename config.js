// Azure AD 앱 등록 후 아래 값을 채워주세요.
// portal.azure.com > App registrations > 등록한 앱 > Overview > "Application (client) ID"
const APP_CONFIG = {
clientId: "5d6de376-b2c3-40c7-a395-1b750b7c9925",  // 개인 Microsoft 계정만 사용하므로 authority는 고정값입니다. 변경하지 마세요.
  authority: "https://login.microsoftonline.com/consumers",
  // OneDrive에 올린 엑셀 파일명 (경로는 자동 검색합니다)
  fileName: "가계부_R2.xlsx",
  // 이용내역 상세 시트의 표(Table) 이름
  tableName: "이용내역상세"
};
