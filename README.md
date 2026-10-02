# LogiFlow — LEONA AI 자동견적 (동의대 학생팀)

Outlook에서 견적 요청 메일을 선택 → 패널이 케이스를 만들고 → 누락 항목 보완 요청 초안 → 회신 병합 → 견적서 송부 초안까지 만든다.
**발송은 항상 담당자가 직접 누른다.** 자동 발송·사서함 자동 조회는 하지 않는다.

> 비공개 저장소. `docs/`에 발주 측 자료(요구사항 정의서·샘플 메일·견적서 양식)가 있다 — **외부 공유·재배포 금지.**

## 폴더
| 폴더 | 내용 | 담당 |
|---|---|---|
| `addin/` | Outlook 추가 기능(패널, TypeScript) | A트랙 |
| `server/` | 백엔드 API(Python FastAPI + SQLAlchemy), 테스트 `server/tests/`, 샘플 .eml `server/samples/eml/` | A트랙 (B·C 연결 지점 `server/app/pipeline.py`) |
| `docs/` | 정의서·발주 측 회신·샘플·견적서 양식·절차서·트랙 간 전달 사항 | 공통 |

## 처음 실행
**`docs/설치절차_A트랙_애드인.md`** 를 순서대로 따라 한다 (Node.js·Python 설치 → 의존성 → 인증서 → `start-dev.bat` → Outlook 웹에 매니페스트 추가).

## 다른 트랙이 먼저 볼 문서
- `docs/트랙간_전달사항.md` — 메일 표준 JSON, B 추출·C 검증 연결 규격, 견적서 파일 두는 곳, 실제 검증에서 발견한 입력 특성
- `CLAUDE.md` — 확정된 결정 사항, 수용 기준 대조표, 진행 상황

## 테스트
```powershell
cd server
.\.venv\Scripts\python -m pytest -q
```
