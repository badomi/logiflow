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
    """테이블이 없으면 만들고, 있는 테이블에 새로 생긴 칸이 없으면 추가한다.

    create_all은 이미 있는 테이블에 칸을 더하지 못한다 → 팀원 PC의 기존 DB가 깨지지 않게 빠진 칸만 ALTER TABLE로 추가.
    빈 칸(NULL 허용)만 추가하므로 기존 데이터는 그대로다. 칸 이름 변경·삭제는 하지 않는다
    (스키마 확정 후에는 마이그레이션 도구로 교체 예정).
    """
    from . import models  # noqa: F401  테이블 정의를 등록하기 위해 불러온다

    Base.metadata.create_all(engine)
    add_missing_columns()


def add_missing_columns() -> list[str]:
    from sqlalchemy import inspect, text

    inspector = inspect(engine)
    added = []
    with engine.begin() as conn:
        for table in Base.metadata.sorted_tables:
            if not inspector.has_table(table.name):
                continue
            have = {c["name"] for c in inspector.get_columns(table.name)}
            for column in table.columns:
                if column.name in have or not column.nullable or column.primary_key:
                    continue
                col_type = column.type.compile(dialect=engine.dialect)
                name = engine.dialect.identifier_preparer.quote(column.name)
                conn.execute(text(f"ALTER TABLE {table.name} ADD {'COLUMN ' if engine.dialect.name == 'sqlite' else ''}"
                                  f"{name} {col_type}"))
                added.append(f"{table.name}.{column.name}")
    return added


def get_session() -> Iterator[Session]:
    """API 요청마다 DB 세션을 하나 열고, 끝나면 닫는다."""
    with SessionLocal() as session:
        yield session
