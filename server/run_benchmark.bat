@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo  LEONA 로컬 LLM 모델 비교 측정
echo  사용법: run_benchmark.bat            (기본: qwen3.5:9b qwen3:14b)
echo          run_benchmark.bat qwen3:8b   (모델을 직접 지정)
echo ============================================================

if not exist ".venv\Scripts\python.exe" (
  echo [오류] server\.venv 가 없습니다. 설치 절차서대로 가상환경을 먼저 만드세요.
  pause
  exit /b 1
)
where ollama >nul 2>nul
if errorlevel 1 (
  echo [오류] ollama 명령을 찾을 수 없습니다. https://ollama.com 에서 설치하세요.
  pause
  exit /b 1
)

set "MODELS=%*"
if "%MODELS%"=="" set "MODELS=qwen3.5:9b qwen3:14b"

echo.
echo [1/3] 설정 점검
.venv\Scripts\python -m app.diagnostics

echo.
echo [2/3] 모델 받기 (이미 있으면 바로 넘어감)
for %%m in (%MODELS%) do (
  echo   - %%m
  ollama pull %%m
)

echo.
echo [3/3] 측정 (모델마다 몇 분 걸릴 수 있음)
.venv\Scripts\python -m eval.compare --models %MODELS%

echo.
echo 끝났습니다. 보고서: server\eval\reports\compare_*.md  ^<- 이 파일을 공유해 주세요
pause
