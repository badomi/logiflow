"""메일 원문·첨부파일 파일 저장소 (FR-103). DB에는 여기서 돌려준 상대 경로만 저장한다."""

import re
from pathlib import Path

from .config import settings

_UNSAFE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def safe_filename(name: str, fallback: str = "file") -> str:
    """경로 조작(../)이나 Windows에서 못 쓰는 글자를 막고 길이를 제한한다."""
    cleaned = _UNSAFE.sub("_", Path(name).name).strip(" .")
    return (cleaned or fallback)[:150]


def save_file(relative_dir: str, filename: str, data: bytes) -> str:
    """storage_dir/relative_dir/filename 에 저장하고 상대 경로(슬래시 구분)를 돌려준다.

    같은 이름이 이미 있으면 "이름 (2).확장자"처럼 번호를 붙인다.
    """
    folder = settings.storage_dir / relative_dir
    folder.mkdir(parents=True, exist_ok=True)

    target = folder / safe_filename(filename)
    counter = 2
    while target.exists():
        target = folder / f"{target.stem.split(' (')[0]} ({counter}){target.suffix}"
        counter += 1

    target.write_bytes(data)
    return target.relative_to(settings.storage_dir).as_posix()


def read_file(relative_path: str) -> bytes:
    return full_path(relative_path).read_bytes()


def full_path(relative_path: str) -> Path:
    """DB에 저장된 상대 경로 → 실제 파일 경로. 저장소 폴더 밖을 가리키면 거부한다."""
    root = settings.storage_dir.resolve()
    path = (root / relative_path).resolve()
    if root not in path.parents:
        raise ValueError("저장소 밖의 경로입니다.")
    return path
