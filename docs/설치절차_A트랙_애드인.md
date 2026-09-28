# 설치·실행 절차서 — A트랙 Outlook 애드인 + 백엔드 (NFR-07)

이 문서만 보고 제3자가 개발 PC에서 LEONA 자동견적 애드인을 실행할 수 있도록 작성했다.
2026-09-28 Windows 11 + Chrome + Outlook 웹(학교 Microsoft 365 계정)에서 실제로 수행한 순서다.

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

## 1. 준비물 (최초 1회)

| 항목 | 버전(검증) | 설치 |
|---|---|---|
| Node.js | 24.x | `winget install OpenJS.NodeJS.LTS` |
| Python | 3.12.x | `winget install Python.Python.3.12 --scope user` |
| Git | 2.x | `winget install Git.Git` |
| Outlook 계정 | Microsoft 365(회사·학교) 또는 Outlook.com | 계정 담당자에게 받음 |

설치 후 **새 터미널**을 열어야 `node`, `python` 명령이 잡힌다.

## 2. 소스 받기와 의존성 설치 (최초 1회)

```powershell
cd C:\dev\leona-addin\addin
npm install

cd C:\dev\leona-addin\server
python -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
copy .env.example .env      # 기본값(SQLite)이면 수정 없이 사용
```

## 3. 개발용 HTTPS 인증서 신뢰 (최초 1회, 직접 클릭 필요)

```powershell
cd C:\dev\leona-addin\addin
npx office-addin-dev-certs install
```

Windows **"보안 경고 — 인증서를 설치하시겠습니까?"** 창이 뜨면 **[예]**. 이 인증서는 이 PC의 `localhost` 전용이며
개인키(`%USERPROFILE%\.office-addin-dev-certs`)는 git에 올라가지 않는다.

## 4. 서버 실행 (작업할 때마다, 터미널 2개)

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

확인: 브라우저에서 `https://localhost:3000/api/health` → `{"status":"ok"}`
API 문서(자동 생성): `http://localhost:8000/docs`

> `npm start`는 데스크톱 Outlook에 자동 설치를 시도한다. Outlook 웹을 쓸 때는 `npm run dev-server`를 쓴다.

## 5. Outlook 웹에 애드인 설치 (사이드로드)

1. Outlook 웹에 로그인한다. 회사·학교 계정은 `https://outlook.cloud.microsoft/mail/`
2. 주소창에 **`https://outlook.cloud.microsoft/mail/inclientstore`** 입력 → "Outlook용 추가 기능" 창이 열린다.
   (새 "앱" 스토어 화면에는 파일 업로드 버튼이 없어서 이 주소를 쓴다)
3. 왼쪽 **[내 추가 기능]** → 맨 아래 **"추가 기능 사용자 지정"** → **[+ 사용자 지정 추가 기능 추가 ∨] → [파일에서 추가…]**
4. `C:\dev\leona-addin\addin\manifest.xml` 선택 → 경고 창에서 **[설치]**
5. 창을 닫고 **F5**(새로고침)

### 매니페스트를 바꾼 뒤 다시 설치하는 법
앱 목록(격자 아이콘) → 왼쪽 아래 **[앱 관리]** → "LEONA 자동견적" **[⋯] → 제거** → 위 5-2부터 다시.
패널 화면(html·ts·css)만 바꾼 경우에는 재설치가 필요 없다.

## 6. 사용 확인 (담당자 흐름)

| 단계 | 담당자 동작 | 결과 | 요구사항 |
|---|---|---|---|
| 1 | 견적 요청 메일 열기 → 메일 위 [앱] → **LEONA 자동견적** | 패널에 메일 정보, "아직 케이스로 등록되지 않은 메일" | FR-101 |
| 2 | **[케이스 생성]** | `LQ-YYYY-MMDD-NNN` 발급, 원문·첨부 저장. 같은 메일은 다시 만들어지지 않음 | FR-101·102·103·105 |
| 3 | 보완 문항 입력 → **[보완 요청 초안 열기]** | 원래 메일에 대한 회신 초안이 열림 | FR-304 |
| 4 | 초안(작성 화면)에서 [앱] → **LEONA 자동견적** | 제목에 `[케이스ID]`, 초안함 저장, "초안함에 저장했습니다" | FR-304·305·510 |
| 5 | 내용 확인 후 **담당자가 직접 [보내기]** | 시스템은 보내지 않는다 | 12장 설계 고정 |
| 6 | 보낸 편지함에서 그 메일 → LEONA → **[발송 기록]** | 실제 발송 시각·수행자 기록, 상태 `보완대기` | FR-505 |
| 7 | 화주 회신 열기 → LEONA | 근거(제목 케이스 ID·회신 헤더·같은 대화)와 함께 대상 케이스 제시 | FR-104 |
| 8 | **[LQ-…에 병합]** | 회신 저장, 추출·검증 재실행 호출 | FR-104·207 |
| 9 | 견적서 등록 후 **[견적서 송부 초안 열기]** → 작성 창 [앱] → LEONA | 견적서 PDF·XLSX 첨부, 초안함 저장 | FR-505 |

## 7. 문제 해결

| 증상 | 원인·조치 |
|---|---|
| 패널이 비어 있음 / "추가 기능을 로드할 수 없음" | 애드인 개발 서버(터미널 2)가 꺼져 있음. `npm run dev-server` |
| "백엔드 서버에 연결할 수 없습니다" | 백엔드(터미널 1)가 꺼져 있음. uvicorn 실행 |
| 인증서 오류 | 3단계 다시 실행 |
| 메뉴에 LEONA가 안 보임 | 설치 직후에는 F5 필요. 그래도 없으면 5단계 재설치 |
| 버튼을 눌러도 반응 없음 | 패널을 처음 클릭하면 활성화만 되는 경우가 있음. 한 번 더 누른다 |
| 보낼 때 "추가 기능이 예상보다 길어지고 있습니다" | 예전 매니페스트(보내기 이벤트 포함)가 설치된 상태. 5단계로 재설치 |

## 8. 알려진 제한 (2026-09-28 실제 테스트 결과)

- 사이드로드한 Outlook 웹에서는 **백그라운드 실행 환경(이벤트 기반 실행·함수 버튼)이 동작하지 않았다.**
  그래서 초안 마무리는 작성 창에서 LEONA 패널을 여는 방식, 발송 기록은 보낸 편지함의 [발송 기록] 버튼으로 한다.
  사내 관리자 배포에서는 보내기 이벤트로 자동 기록이 가능할 수 있으며 코드(`addin/src/commands`)는 남겨 두었다.
- 초안 생성은 Microsoft Graph 대신 Office.js로 한다(열린 메일 1통 권한 `ReadWriteItem`만 사용).
  정의서 7장의 "Graph API로 회신·발송" 표현과 다르므로 Rev 1.3 반영을 발주 측에 요청해야 한다.
- 한 계정으로 화주·담당자를 모두 테스트하면 자기 메일에 병합 경고가 뜬다. 가능하면 화주용 계정을 따로 쓴다.

## 9. 테스트 실행

```powershell
cd C:\dev\leona-addin\server
.\.venv\Scripts\python -m pytest -q          # 백엔드 수용 기준 테스트
.\.venv\Scripts\python -m app.eml_reader samples\eml   # .eml → 메일 표준 JSON 출력
```
