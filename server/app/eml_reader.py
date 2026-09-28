""".eml 리더 — .eml 파일을 메일 표준 JSON(MailSnapshot)으로 바꾼다 (FR-101 우회 경로).

Outlook 없이도 B트랙(추출)·D트랙(정밀도 측정)이 같은 형식의 입력을 대량으로 얻기 위한 도구.

사용법 (server 폴더에서):
  .venv\\Scripts\\python -m app.eml_reader samples\\eml                 # 화면에 JSON 출력
  .venv\\Scripts\\python -m app.eml_reader samples\\eml --out out\\json  # 파일마다 .json 저장
  .venv\\Scripts\\python -m app.eml_reader samples\\eml --import        # DB에 케이스로 등록
"""

import argparse
import json
import sys
from pathlib import Path

from .mail_parser import parse_eml


def _eml_files(target: Path) -> list[Path]:
    if target.is_dir():
        return sorted(target.glob("*.eml"))
    return [target]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=".eml → 메일 표준 JSON 변환")
    parser.add_argument("path", type=Path, help=".eml 파일 또는 .eml이 들어 있는 폴더")
    parser.add_argument("--out", type=Path, help="JSON 파일을 저장할 폴더 (없으면 화면 출력)")
    parser.add_argument("--import", dest="do_import", action="store_true", help="DB에 케이스로 등록")
    parser.add_argument("--actor", default="eml-reader", help="이력에 남길 수행자 이름")
    args = parser.parse_args(argv)
    # Windows 콘솔 기본 인코딩(cp949)에서 한글·기호가 깨지지 않게 UTF-8로 출력한다
    sys.stdout.reconfigure(encoding="utf-8")

    files = _eml_files(args.path)
    if not files:
        print(f".eml 파일이 없습니다: {args.path}", file=sys.stderr)
        return 1

    if args.do_import:
        return _import(files, args.actor)

    if args.out:
        args.out.mkdir(parents=True, exist_ok=True)
    outputs = []
    for file in files:
        snapshot = parse_eml(file.read_bytes()).snapshot.model_dump(mode="json", by_alias=True)
        if args.out:
            (args.out / f"{file.stem}.json").write_text(
                json.dumps(snapshot, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            print(f"{file.name} → {args.out / (file.stem + '.json')}")
        else:
            outputs.append({"file": file.name, "mail": snapshot})
    if not args.out:
        print(json.dumps(outputs, ensure_ascii=False, indent=2))
    return 0


def _import(files: list[Path], actor: str) -> int:
    from . import services
    from .db import SessionLocal, init_db

    init_db()
    with SessionLocal() as session:
        for file in files:
            raw = file.read_bytes()
            case, duplicate = services.create_case(session, parse_eml(raw), raw, "EML", actor)
            print(f"{file.name}: {case.case_id}{' (이미 등록됨)' if duplicate else ''}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
