# 설치·실행 절차서 — A트랙 Outlook 애드인 + 백엔드 (NFR-07)

이 문서만 보고 제3자가 개발 PC에서 LEONA 자동견적 애드인을 실행할 수 있도록 작성했다.
2026-09-28 Windows 11 + Chrome + Outlook 웹(개인 학교 Microsoft 365 계정)에서 실제로 수행한 순서다.

---

## 0. 구성 한눈에 보기

```
Outlook 웹 ──(패널 화면)── https://localhost:3000   애드인 개발 서버 (addin/, Node.js)
                                │  /api 요청을 넘김(프록시)
                                ▼
                          http://localhost:8000   A트랙 백엔드 API (server/, Python FastAPI)
                                │
                                ├─ server/data/leona.db   개발용 DB (SQLite)
                                └─ server/storage/        메일 원문(.eml)·첨부·견적서 파일
```

- 메일 본문은 내 PC의 로컬 서버로만 가며 외부 서비스로 보내지 않는다 (NFR-04).
- 애드인은 **자동 발송·사서함 주기 조회(폴링)를 하지 않는다.** 모든 동작은 담당자가 버튼을 눌러야 실행된다.
- 두 서버는 **내 PC에서만** 돈다. 서버 창을 닫거나 PC를 끄면 패널도 동작하지 않는다.

## 1. 준비물 (최초 1회)

| 항목 | 버전(검증) | 설치 |
|---|---|---|
| Node.js | 24.x | `winget install OpenJS.NodeJS.LTS` |
| Python | 3.12.x | `winget install Python.Python.3.12 --scope user` |
| Git (또는 GitHub Desktop) | 2.x | `winget install Git.Git` |
| Outlook 계정 | 본인 Microsoft 계정 | **각자 준비.** 발주 측은 계정을 제공하지 않으며 개인 Microsoft 계정 사용을 허락했다(`docs/client_messages.md` 회신 1). 학교 Microsoft 365 계정(`@office.deu.ac.kr`) 또는 개인 Outlook.com 계정 |

설치 후 **새 터미널**을 열어야 `node`, `python` 명령이 잡힌다.
(열려 있던 창에서는 `python`이 Microsoft Store 바로가기로 연결될 수 있다)

## 2. 소스 받기와 의존성 설치 (최초 1회)

1. 팀 저장소를 `C:\dev\leona-addin`에 받는다.
   - GitHub Desktop: **File → Clone repository** → 팀 비공개 저장소 선택 → Local path `C:\dev\leona-addin`
   - 또는 전달받은 zip을 `C:\dev\leona-addin`에 압축 해제
2. 의존성을 설치한다.

```powershell
cd C:\dev\leona-addin\addin
npm install

cd C:\dev\leona-addin\server
python -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
copy .env.example .env      # 기본값(SQLite)이면 수정 없이 사용
```

> 저장소에는 발주 측 자료(정의서·샘플 메일, 외부 재배포 금지)가 들어 있으므로 **비공개로만** 공유한다.

## 3. 개발용 HTTPS 인증서 신뢰 (최초 1회, 직접 클릭 필요)

```powershell
cd C:\dev\leona-addin\addin
npx office-addin-dev-certs install
```

Windows **"보안 경고 — 인증서를 설치하시겠습니까?"** 창이 뜨면 **[예]**. 이 인증서는 이 PC의 `localhost` 전용이며
개인키(`%USERPROFILE%\.office-addin-dev-certs`)는 git에 올라가지 않는다.

## 4. 서버 실행 (작업할 때마다)

**방법 A — 더블클릭 (권장):** 프로젝트 폴더의 **`start-dev.bat`** 을 더블클릭한다.
"LEONA 백엔드 API (8000)", "LEONA 애드인 개발 서버 (3000)" 창 2개가 열린다. **끄려면 두 창을 닫는다.**

**방법 B — 직접 실행 (터미널 2개):**

```powershell
# 터미널 1 — 백엔드 API
cd C:\dev\leona-addin\server
.\.venv\Scripts\python -m uvicorn app.main:app --port 8000 --reload
```

```powershell
# 터미널 2 — 애드인 개발 서버
cd C:\dev\leona-addin\addin
npm run dev-server
```

확인: 10~20초 뒤 브라우저에서 `https://localhost:3000/api/health` → `{"status":"ok"}`
API 문서(자동 생성, 직접 호출해 볼 수 있음): `http://localhost:8000/docs`

> `npm start`는 데스크톱 Outlook에 자동 설치를 시도한다. Outlook 웹을 쓸 때는 `npm run dev-server`(또는 start-dev.bat)를 쓴다.

## 5. Outlook 웹에 애드인 설치 (사이드로드, 계정마다 1회)

1. Outlook 웹에 로그인한다.
   - 학교·회사 계정: `https://outlook.cloud.microsoft/mail/` (검증함)
   - 개인 Outlook.com 계정: `https://outlook.live.com/mail/` (**미검증** — 아래 2의 주소를 `https://outlook.live.com/mail/0/inclientstore`로 바꿔 시도)
2. 주소창에 **`https://outlook.cloud.microsoft/mail/inclientstore`** 입력 → "Outlook용 추가 기능" 창이 열린다.
   (새 "앱" 스토어 화면에는 파일 업로드 버튼이 없어서 이 주소를 쓴다)
3. 왼쪽 **[내 추가 기능]** → 맨 아래 **"추가 기능 사용자 지정"** → **[+ 사용자 지정 추가 기능 추가 ∨] → [파일에서 추가…]**
4. `C:\dev\leona-addin\addin\manifest.xml` 선택 → 경고 창에서 **[설치]**
5. 창을 닫고 **F5**(새로고침)

### 매니페스트를 바꾼 뒤 다시 설치하는 법
메일의 [앱](격자 아이콘) → 왼쪽 아래 **[앱 관리]** → "LEONA 자동견적" **[⋯] → 제거** → 위 5-2부터 다시.
패널 화면(html·ts·css)이나 백엔드만 바꾼 경우에는 재설치가 필요 없다.

## 6. 사용 확인 (담당자 흐름)

패널은 메일 위쪽의 **[앱](격자 아이콘) → LEONA 자동견적**으로 연다.
패널 오른쪽 위 **📌(고정)** 을 누르면 다른 메일을 선택해도 패널이 닫히지 않고 새 메일을 자동으로 읽는다.

| 단계 | 담당자 동작 | 결과 | 요구사항 |
|---|---|---|---|
| 1 | 견적 요청 메일 열기 → LEONA 패널 | 메일 정보, "아직 케이스로 등록되지 않은 메일" | FR-101 |
| 2 | **[케이스 생성]** | `LQ-YYYY-MMDD-NNN` 발급, 원문·첨부 저장. 같은 메일은 다시 만들어지지 않음 | FR-101·102·103·105 |
| 3 | 보완 문항 입력 → **[보완 요청 초안 열기]** | 원래 메일에 대한 회신 초안이 열림. **문항은 C트랙 검증 연결 전까지 담당자가 직접 입력** | FR-304 |
| 4 | 그 초안(작성 화면)에서 [앱] → **LEONA 자동견적** | 제목에 `[케이스ID]`, 초안함 저장, "초안함에 저장했습니다" | FR-304·305·510 |
| 5 | 내용 확인 후 **담당자가 직접 [보내기]** | 시스템은 보내지 않는다 | 12장 설계 고정 |
| 6 | 보낸 편지함에서 그 메일 → LEONA → **[발송 기록]** | 메일 원문의 실제 발송 시각·수행자 기록, 상태 `보완대기` | FR-505 |
| 7 | 화주 회신 열기 → LEONA | 근거(제목 케이스 ID·회신 헤더·같은 대화)와 함께 대상 케이스 제시 | FR-104 |
| 8 | **[LQ-…에 병합]** (못 찾으면 [케이스 직접 선택]) | 회신 저장, 추출·검증 재실행 호출 | FR-104·207 |
| 9 | 견적서 등록 → 케이스 메일에서 **[견적서 송부 초안 열기]** → 새 메일 창의 [앱] → **LEONA** | 견적서 PDF·XLSX 자동 첨부, 초안함 저장 | FR-505 |

**9단계 견적서 등록 (C트랙 연결 전 테스트 방법):** `http://localhost:8000/docs` → `POST /api/cases/{case_id}/quotes` →
**Try it out** → case_id 입력, PDF 또는 XLSX 파일 선택 → **Execute**. C트랙이 연결되면 견적서 생성 결과가 자동으로 여기에 들어온다.

## 7. 문제 해결

| 증상 | 원인·조치 |
|---|---|
| 패널이 비어 있음 / "추가 기능을 로드할 수 없음" | 애드인 개발 서버(3000)가 꺼져 있음. start-dev.bat 실행 |
| "백엔드 서버에 연결할 수 없습니다" | 백엔드(8000)가 꺼져 있음. start-dev.bat 실행 |
| 서버 창에 "address already in use" / "포트 사용 중" | 이미 서버가 켜져 있음(start-dev.bat 두 번 실행 등). 기존 서버 창을 닫고 다시 실행 |
| 인증서 오류 | 3단계 다시 실행 |
| 메뉴에 LEONA가 안 보임 | 설치 직후에는 F5 필요. 그래도 없으면 5단계 재설치 |
| 버튼을 눌러도 반응 없음 | 패널을 처음 클릭하면 활성화만 되는 경우가 있음. 한 번 더 누른다 |
| 작성 창 LEONA에 "패널에서 연 초안이 아닙니다" | 받은 메일의 패널에서 [초안 열기]를 먼저 눌러야 한다(30분 안에) |
| 보낼 때 "추가 기능이 예상보다 길어지고 있습니다" | 예전 매니페스트(보내기 이벤트 포함)가 설치된 상태. 5단계로 재설치 |

## 8. 알려진 제한 (2026-09-28 실제 테스트 결과)

- **추가 클릭 2회:** 사이드로드한 Outlook 웹에서는 백그라운드 실행 환경(이벤트 기반 실행)이 동작하지 않았다.
  그래서 초안 저장은 작성 창에서 LEONA 패널을 한 번 더 여는 방식(6-4, 6-9), 발송 기록은 [발송 기록] 버튼(6-6)으로 한다.
  - 줄이기 시험 예정: 보완 요청 인라인 회신에서 고정 패널이 제목을 자동 수정, 보낸 메일·대화를 열면 발송 자동 기록
  - 완전 해결 후보: 사내 IT의 **중앙 배포**(이벤트 기반 실행 지원). 처리기 코드는 `addin/src/commands`에 보존
- **Graph 미사용:** 초안 생성은 Microsoft Graph 대신 Office.js로 한다(열린 메일 1통 권한 `ReadWriteItem`만 사용).
  승인된 변경 요청서의 "선택 메일 권한만" 방향과 맞추기 위함이며, 정의서 7장 "Graph API로 회신·발송" 표현은 Rev 1.3 반영을 요청한다.
- **원문 보관 조건:** 메일 원문(.eml)·첨부는 Outlook 요구사항 세트 1.14(`getAsFileAsync`)가 되는 환경에서만 받는다.
  안 되는 구버전 Outlook에서는 패널이 읽은 제목·본문 등만 저장된다.
- **앱 목록 아이콘:** Microsoft 서버가 설치 시 가져가므로 공개 HTTPS 주소여야 한다(`localhost` 불가).
  현재는 발주 측 홈페이지 로고 주소를 쓰는데 정사각형이 아니라 가로로 눌려 보인다.
  → 정사각형 아이콘(`addin/assets/icon-128.png`)을 공개 주소에 올리고 매니페스트 `IconUrl`·`HighResolutionIconUrl`을 바꾼다(로고 사용 허락 확인 후).
- **견적서 파일:** 현재 테스트용 임시 PDF. 실제 견적서는 C트랙 생성 결과(발주 측 양식 파일 수령 필요).
- **단일 계정 테스트:** 한 계정으로 화주·담당자를 모두 테스트하면 자기 메일에 병합 경고가 뜨고 [발송 기록]도 함께 보인다.
  가능하면 화주 역할 계정을 따로 쓴다.
- **검증 범위:** Outlook 웹(학교 계정)만 검증했다. 데스크톱 Outlook·개인 Outlook.com 계정은 미검증.

## 9. 테스트 실행

```powershell
cd C:\dev\leona-addin\server
.\.venv\Scripts\python -m pytest -q          # 백엔드 수용 기준 테스트 (34건)
.\.venv\Scripts\python -m app.eml_reader samples\eml   # .eml → 메일 표준 JSON 출력
```
