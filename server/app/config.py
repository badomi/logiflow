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


settings = Settings()
