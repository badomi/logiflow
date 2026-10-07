"""설정 점검 — 추출·LLM·요율·PDF·.env 상태를 한 번에 확인하고, 문제가 있으면 무엇을 하라고 알려준다.

    python -m app.diagnostics          서버를 켜지 않고 점검 (server 폴더에서)
    GET /api/health/detail             서버가 켜져 있을 때
"""

import importlib.util
import json
import urllib.request

from sqlalchemy import func, select
from sqlalchemy.orm import Session

from .config import SERVER_DIR, settings
from .models import ContainerType, Rate


def _ollama(url: str, model: str) -> dict:
    from .llm import LlmError, ensure_local

    info = {"url": url, "model": model, "reachable": False, "modelInstalled": False, "installed": []}
    try:
        ensure_local(url)
        with urllib.request.urlopen(f"{url.rstrip('/')}/api/tags", timeout=3) as res:
            names = [m.get("name", "") for m in json.loads(res.read().decode()).get("models", [])]
        info.update(reachable=True, installed=names,
                    modelInstalled=model in names or f"{model}:latest" in names)
    except LlmError as error:
        info["error"] = str(error)
    except Exception as error:  # 연결 실패 등
        info["error"] = f"{error.__class__.__name__}: {error}"
    return info


def collect(session: Session) -> dict:
    from . import quotation, rates
    from .extraction.base import dictionary
    from .validation import rules

    warnings: list[str] = []
    env_file = SERVER_DIR / ".env"
    if not env_file.exists():
        warnings.append("server\\.env 파일이 없습니다 — .env.example을 복사해 만드세요 (모든 설정이 기본값)")
    mode = settings.extraction_mode
    if mode not in ("off", "rules", "hybrid"):
        warnings.append(f"EXTRACTION_MODE 값이 이상합니다: {mode} (off / rules / hybrid)")
    elif mode == "off":
        warnings.append("EXTRACTION_MODE=off — 항목 추출·검증·견적서가 꺼져 있습니다 (.env에서 rules 또는 hybrid)")

    llm = None
    if mode == "hybrid":
        llm = _ollama(settings.llm_url, settings.llm_model)
        if not llm["reachable"]:
            warnings.append(f"LLM 서버({settings.llm_url})에 연결되지 않습니다 — Ollama가 켜져 있는지 확인 "
                            f"(지금은 키워드 규칙만으로 추출)")
        elif not llm["modelInstalled"]:
            warnings.append(f"모델 {settings.llm_model}이 설치돼 있지 않습니다 — `ollama pull {settings.llm_model}` "
                            f"(설치된 모델: {', '.join(llm['installed']) or '없음'})")

    rate_count = session.scalar(select(func.count()).select_from(Rate)) or 0
    if rate_count == 0:
        warnings.append("요율이 0개입니다 — 모든 견적이 '요율 미등록'으로 보류됩니다 "
                        "(요율 엑셀 올리기, 또는 테스트용: python -m app.seed --sample-rates)")
    sources = rates.summary(session)
    for s in sources:
        if s["expired"]:
            warnings.append(f"적용 기간이 끝난 요율이 있습니다: {s['source']} (~{s['validUntil']}) — 새 운임표를 올리세요")

    soffice = quotation.find_soffice() if settings.quote_pdf else None
    builtin = importlib.util.find_spec("reportlab") is not None  # 내장 엔진(quote_pdf.py)
    if settings.quote_pdf and soffice is None and not builtin:
        warnings.append("LibreOffice도 reportlab도 없어 견적서 PDF를 만들지 못합니다 (XLSX만). "
                        "pip install -r requirements.txt 로 reportlab을 설치하세요")
    if not settings.quote_template.exists():
        warnings.append(f"견적서 양식 파일이 없습니다: {settings.quote_template}")

    return {
        "ok": not warnings,
        "warnings": warnings,
        "envFile": str(env_file) if env_file.exists() else None,
        "extractionMode": mode,
        "llm": llm,
        "rates": {"count": rate_count, "sources": sources},
        "containers": list(session.scalars(select(ContainerType.code))),
        "pdf": {"enabled": settings.quote_pdf, "soffice": soffice, "builtin": builtin},
        "quoteTemplate": str(settings.quote_template),
        "database": settings.database_url.split("@")[-1],  # 비밀번호가 있으면 가린다
        "dictionaryVersion": dictionary().get("version"),
        "rulesVersion": rules().get("version"),
    }


def main() -> None:
    from .db import SessionLocal, init_db
    from .seed import seed_containers

    init_db()
    with SessionLocal() as session:
        seed_containers(session)
        session.commit()
        info = collect(session)
    print("LEONA 자동견적 설정 점검")
    print(f"  .env 파일      : {info['envFile'] or '없음'}")
    print(f"  추출 모드      : {info['extractionMode']}")
    if info["llm"]:
        l = info["llm"]
        print(f"  LLM            : {l['model']} @ {l['url']} — 연결 {'O' if l['reachable'] else 'X'}, "
              f"모델 설치 {'O' if l['modelInstalled'] else 'X'}")
    print(f"  요율           : {info['rates']['count']}개 ({len(info['rates']['sources'])}개 출처)")
    print(f"  견적서 PDF     : {info['pdf']['soffice'] or ('내장 엔진 (reportlab)' if info['pdf']['builtin'] else '만들지 않음')}")
    print(f"  DB             : {info['database']}")
    print()
    if info["ok"]:
        print("문제 없음")
    else:
        print("확인할 것:")
        for w in info["warnings"]:
            print("  -", w)


if __name__ == "__main__":
    main()
