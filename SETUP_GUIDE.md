# 가계부 지출 입력 앱 — 설정 가이드

`app` 폴더에 있는 정적 파일들을 배포하고 OneDrive와 연동하기 위한 단계별 안내입니다. 아래 순서대로 진행하세요.

## 1. 가계부_R2.xlsx를 OneDrive에 업로드

1. `C:\Users\kim6w\Desktop\Moneymap\가계부_R2.xlsx` (이미 앱 연동용으로 구조가 변경된 파일)를 OneDrive 폴더로 이동 또는 업로드합니다.
   - OneDrive 데스크톱 앱을 쓰신다면 OneDrive 동기화 폴더로 파일을 옮기면 됩니다.
   - 아니라면 [onedrive.com](https://onedrive.com) 에 로그인해서 웹으로 직접 업로드하세요.
2. 파일명은 반드시 `가계부_R2.xlsx` 그대로 유지하세요 (앱이 이 이름으로 파일을 찾습니다).
3. 원본 백업(`가계부_R2_backup_20260719.xlsx`)은 OneDrive에 올리지 않아도 됩니다 — 로컬 백업용입니다.

## 2. Azure AD 앱 등록 (Microsoft Graph API 접근 권한)

1. [portal.azure.com](https://portal.azure.com) 에 로그인 (OneDrive와 같은 개인 Microsoft 계정 사용).
2. 상단 검색창에 "App registrations" 입력 후 이동 → **New registration** 클릭.
3. 입력값:
   - **Name**: `가계부입력앱` (원하는 이름으로 자유롭게)
   - **Supported account types**: **Personal Microsoft accounts only** 선택
   - **Redirect URI**: 지금은 비워두고 3단계(GitHub Pages 배포) 완료 후 다시 와서 등록합니다.
4. **Register** 클릭 후, Overview 화면에서 **Application (client) ID** 값을 복사해둡니다. (예: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`)
5. 왼쪽 메뉴 **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions** → `Files.ReadWrite` 검색해서 체크 → **Add permissions**.
6. 왼쪽 메뉴 **Authentication** → **Add a platform** → **Single-page application (SPA)** 선택.
   - Redirect URI에 3단계에서 확정될 GitHub Pages 주소를 입력합니다 (예: `https://내계정.github.io/gagyebu-app/`). 지금은 임시로 `http://localhost:5500/` 를 추가해도 되고, 나중에 추가/수정도 가능합니다.
   - **Configure** 저장.

## 3. GitHub Pages 배포

1. [github.com](https://github.com) 에서 새 저장소(Repository) 생성 (예: `gagyebu-app`), Public으로 설정.
2. 저장소 페이지에서 **Add file → Upload files** 클릭.
3. `app` 폴더 안의 모든 파일을 통째로 드래그 앤 드롭으로 업로드합니다:
   - `index.html`, `app.js`, `config.js`, `styles.css`, `manifest.json`, `sw.js`, `icons/icon-192.png`, `icons/icon-512.png`
   - **주의**: `app` 폴더 자체가 아니라 그 **안의 파일들**이 저장소 루트에 오도록 업로드하세요.
4. **Commit changes** 로 저장.
5. 저장소 **Settings → Pages** 로 이동 → Source를 `Deploy from a branch`, Branch를 `main` / `/(root)` 로 설정 → **Save**.
6. 몇 분 후 `https://내계정.github.io/gagyebu-app/` 형태의 주소가 발급됩니다. 이 주소를 2단계의 Azure AD **Redirect URI**에 정확히 등록/수정하세요 (끝의 `/` 까지 동일하게).

## 4. 앱 설정값 채우기

1. `app/config.js` 파일을 열어 `clientId` 값을 2단계에서 복사한 Application (client) ID로 교체합니다.
2. GitHub 저장소에서 `config.js` 파일을 다시 업로드(덮어쓰기)해서 반영합니다.
   - Git을 쓸 수 있는 환경이라면 `git add / commit / push`로도 가능합니다.

## 5. 최종 테스트

1. 휴대폰 또는 PC 브라우저(Chrome/Edge 권장)에서 GitHub Pages 주소로 접속합니다.
2. "Microsoft 계정으로 로그인" 클릭 → OneDrive 계정으로 로그인 및 권한 동의.
3. 로그인 후 입력 폼이 보이면: 날짜(기본 오늘) · 대분류 · 소분류 · 세부항목을 선택하고, 금액을 입력 후 **입력** 버튼 클릭.
4. "저장되었습니다" 메시지가 뜨고 "최근 입력 내역"에 방금 넣은 항목이 나타나는지 확인합니다.
5. OneDrive에서 `가계부_R2.xlsx`를 열어:
   - `이용내역 상세` 시트에 새 행이 추가되었는지
   - `가계부_월별` 시트의 해당 월 '실적' 값이 갱신되었는지
   확인합니다.
6. (선택) 휴대폰 브라우저에서 "홈 화면에 추가"를 하면 앱처럼 아이콘으로 설치되어 바로 실행할 수 있습니다.

## 문제가 생기면

- **로그인이 안 됨 / 리디렉션 오류**: Azure AD의 Redirect URI가 실제 접속 주소와 정확히 일치하는지 확인 (끝 슬래시 포함).
- **"파일을 찾을 수 없습니다"**: OneDrive에 업로드한 파일명이 정확히 `가계부_R2.xlsx`인지 확인.
- **저장은 되는데 가계부_월별에 반영이 안 보임**: 파일을 껐다 켜면(또는 Excel에서 수동으로 Ctrl+Alt+F9 전체 재계산) 반영됩니다. 워크북에 열 때 자동 재계산 설정을 이미 켜두었습니다.
- 그 외 오류 메시지는 화면의 상태 메시지에 그대로 표시되니, 캡처해서 알려주시면 원인을 확인해드릴 수 있습니다.
