"""설정값. 코드에 비밀값을 쓰지 않고 .env 파일(또는 환경 변수)에서 읽는다 (NFR-04)."""

from datetime import timedelta, timezone
from pathlib import Path

from pydantic_settings import BaseSettings, SettingsConfigDict

SERVER_DIR = Path(__file__).resolve().parent.parent

# 케이스 ID 날짜 기준 시간대. 한국은 서머타임이 없어 고정 +09:00으로 충분하다.
KST = timezone(timedelta(hours=9))


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=SERVER_DIR / ".env", extra="ignore")

    database_url: str = f"sqlite:///{(SERVER_DIR / 'data' / 'leona.db').as_posix()}"
    storage_dir: Path = SERVER_DIR / "storage"

    # 로컬 LLM (Ollama). 메일 본문이 외부로 나가지 않도록 로컬·사내 주소만 허용한다 (NFR-04).
    # off: 추출 미연결(기본) / rules: 키워드 사전·정규식만 / hybrid: 룰 + 로컬 LLM
    extraction_mode: str = "off"
    llm_url: str = "http://localhost:11434"
    llm_model: str = "qwen3:8b"
    llm_think: bool = False
    # all: LLM이 모든 항목을 읽고 규칙은 교차 확인 (기본, 유연) / missing: 규칙이 못 찾은 빈 칸만 LLM (느린 PC)
    llm_scope: str = "all"  # 생각(thinking) 모드. 켜면 긴 추론이 먼저 나와 60초를 넘기기 쉽다 → 기본 끔
    llm_timeout_s: float = 45.0  # 추출+검증+견적 전체가 60초 안에 끝나야 한다 (FR-101). 넘으면 룰 결과로 진행
    llm_num_ctx: int = 8192  # Ollama 기본 컨텍스트는 짧아 긴 메일이 '조용히' 잘린다 → 명시
    llm_keep_alive: str = "30m"  # 모델을 메모리에 유지 (매 요청마다 로딩하면 10초 이상 추가)

    # 견적서 PDF (MUST-SHIP ④: PDF·XLSX). LibreOffice로 양식 XLSX를 변환한다. 없으면 XLSX만 만든다.
    quote_pdf: bool = True
    soffice_path: str | None = None  # 비우면 PATH·기본 설치 경로에서 찾는다
    pdf_timeout_s: float = 30.0

    # 견적서 양식 (발주 측 제공 XLSX)
    quote_template: Path = SERVER_DIR.parent / "docs" / "양식" / "견적서_양식.xlsx"


settings = Settings()
