@echo off
chcp 65001 > nul
rem ============================================================
rem  LEONA 자동견적 개발 서버 한 번에 켜기 (더블클릭)
rem  - 창 2개가 열린다: 백엔드 API(8000), 애드인 개발 서버(3000)
rem  - 끄려면 두 창을 닫는다
rem ============================================================
cd /d "%~dp0"

if not exist "server\.venv\Scripts\python.exe" (
  echo [오류] server\.venv 가 없습니다. docs\설치절차_A트랙_애드인.md 2단계를 먼저 하세요.
  pause
  exit /b 1
)
if not exist "addin\node_modules" (
  echo [오류] addin\node_modules 가 없습니다. addin 폴더에서 npm install 을 먼저 하세요.
  pause
  exit /b 1
)

start "LEONA 백엔드 API (8000)" cmd /k "cd /d "%~dp0server" && .venv\Scripts\python -m uvicorn app.main:app --port 8000 --reload"
start "LEONA 애드인 개발 서버 (3000)" cmd /k "cd /d "%~dp0addin" && npm run dev-server"

echo.
echo 서버 2개를 켰습니다. 약 10~20초 뒤 Outlook에서 LEONA 패널을 열어 주세요.
echo 확인 주소: https://localhost:3000/api/health  (status ok 가 나오면 정상)
echo.
timeout /t 5 > nul
