"""DB 연결. 연결 주소(DATABASE_URL)만 바꾸면 SQLite ↔ Oracle ↔ 발주 측 DB로 옮길 수 있다."""

from collections.abc import Iterator
from pathlib import Path

from sqlalchemy import create_engine
from sqlalchemy.engine import make_url
from sqlalchemy.orm import DeclarativeBase, Session, sessionmaker

from .config import settings


class Base(DeclarativeBase):
    pass


def _make_engine(url: str):
    if url.startswith("sqlite"):
        db_path = make_url(url).database
        if db_path and db_path != ":memory:":
            Path(db_path).parent.mkdir(parents=True, exist_ok=True)
        return create_engine(url, connect_args={"check_same_thread": False})
    return create_engine(url, pool_pre_ping=True)


engine = _make_engine(settings.database_url)
SessionLocal = sessionmaker(bind=engine, expire_on_commit=False)


def init_db() -> None:
    """테이블이 없으면 만든다. (스키마 확정 후에는 마이그레이션 도구로 교체 예정)"""
    from . import models  # noqa: F401  테이블 정의를 등록하기 위해 불러온다

    Base.metadata.create_all(engine)


def get_session() -> Iterator[Session]:
    """API 요청마다 DB 세션을 하나 열고, 끝나면 닫는다."""
    with SessionLocal() as session:
        yield session
