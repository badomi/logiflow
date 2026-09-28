"""테스트 공통 설정: 실제 DB·저장소를 건드리지 않도록 임시 폴더를 쓴다."""

import os
import shutil
import tempfile
from pathlib import Path

# app 모듈을 불러오기 전에 설정해야 한다 (설정은 import 시점에 한 번 읽힌다)
_TMP = Path(tempfile.mkdtemp(prefix="leona-test-"))
os.environ["DATABASE_URL"] = f"sqlite:///{(_TMP / 'test.db').as_posix()}"
os.environ["STORAGE_DIR"] = str(_TMP / "storage")

import pytest  # noqa: E402
from fastapi.testclient import TestClient  # noqa: E402

from app.config import settings  # noqa: E402
from app.db import Base, SessionLocal, engine, init_db  # noqa: E402
from app.main import app  # noqa: E402

SAMPLES = Path(__file__).resolve().parent.parent / "samples" / "eml"


@pytest.fixture(autouse=True)
def clean_db():
    """테스트마다 빈 DB·빈 저장소에서 시작한다."""
    init_db()
    yield
    Base.metadata.drop_all(engine)
    shutil.rmtree(settings.storage_dir, ignore_errors=True)


@pytest.fixture
def session():
    with SessionLocal() as s:
        yield s


@pytest.fixture
def client():
    with TestClient(app) as c:
        yield c


def sample(name: str) -> bytes:
    return (SAMPLES / name).read_bytes()
