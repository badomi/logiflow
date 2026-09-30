"""추출 정밀도 측정 (NFR-01) — 평가 세트로 엔진을 돌려 필드별 정밀도·채움률·처리 시간을 보고서로 남긴다.

    cd server
    python -m eval.run_eval                         룰만
    python -m eval.run_eval --mode hybrid           룰 + 로컬 LLM (Ollama, .env의 LLM_URL·LLM_MODEL)
    python -m eval.run_eval --mode hybrid --model qwen2.5:14b-instruct   모델 비교

지표 (정의서 NFR-01: '추출한 값의 정밀도(오추출률)', 미추출은 오류 아님)
  정밀도 = 맞게 채운 값 / 채운 값        ← 목표 90% 이상
  채움률 = 맞게 채운 값 / 정답이 있는 값   ← 높을수록 보완 요청이 줄어듦
  처리 시간: 메일 1건 60초 이내 (NFR-02)
보고서: eval/reports/<시각>_<모드>.md — 오추출 목록은 W3–4 '오추출 사례 피드백' 자료로 쓴다.
"""

import argparse
import io
import statistics
import sys
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

from openpyxl import Workbook

SERVER = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SERVER))

from app.extraction import attachments, engine  # noqa: E402
from app.extraction.base import FIELDS, MailText  # noqa: E402
from app.mail_parser import parse_eml  # noqa: E402

from eval.dataset import all_cases  # noqa: E402

SAMPLES = SERVER / "samples" / "eml"
REPORTS = Path(__file__).resolve().parent / "reports"


def xlsx_bytes(rows: list[list[str]]) -> bytes:
    wb = Workbook()
    for row in rows:
        wb.active.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_eml(mail: dict) -> bytes:
    """변형 메일 정의 → 실제 .eml 바이트 (A트랙과 같은 파서를 거치게 하려고)."""
    msg = EmailMessage()
    msg["Subject"] = mail["subject"]
    msg["From"] = f'{mail["from"][0]} <{mail["from"][1]}>'
    msg["To"] = "LEONA 견적 담당자 <quote@leona.example.com>"
    msg.set_content(mail["body"])
    for filename, rows in mail.get("attachments", []):
        msg.add_attachment(xlsx_bytes(rows), maintype="application",
                           subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=filename)
    return msg.as_bytes()


def load_mails(case: dict) -> list[MailText]:
    out = []
    for mail in case["mails"]:
        if "path" in mail:
            raw = Path(mail["path"]).read_bytes()
        elif "eml" in mail:
            raw = (SAMPLES / mail["eml"]).read_bytes()
        else:
            raw = build_eml(mail)
        parsed = parse_eml(raw)
        snap = parsed.snapshot
        files = [(a.filename, attachments.extract_text(a.filename, a.data)) for a in parsed.attachments]
        out.append(MailText(snap.subject, snap.from_.name, snap.from_.email, snap.body_text,
                            mail.get("role", "ORIGINAL"), files))
    return out


def same(got, expected) -> bool:
    if isinstance(expected, (int, float)) and isinstance(got, (int, float)):
        return abs(got - expected) <= max(0.01, abs(expected) * 0.005)
    return engine._same(got, expected)


def evaluate(llm=None, model: str | None = None, cases: list[dict] | None = None, scope: str = "all") -> dict:
    cases_in = cases if cases is not None else all_cases()
    per_field = {f: {"filled": 0, "correct": 0, "expected": 0} for f in FIELDS}
    errors, cases, times, multi_ok, unparsed = [], [], [], 0, []
    for case in cases_in:
        result = engine.run(load_mails(case), llm=llm, model_name=model, scope=scope, asked=case.get("asked"))
        times.append(result["elapsedMs"])
        expected = case["expected"]
        c_filled = c_correct = 0
        for f in FIELDS:
            want = expected.get(f)
            got = result["fields"][f]
            if got.get("reason") == "UNPARSED":
                unparsed.append({"case": case["id"], "field": f, "evidence": got["evidence"], "expected": want})
            if want is not None:
                per_field[f]["expected"] += 1
            if got["status"] != "filled":
                continue
            per_field[f]["filled"] += 1
            c_filled += 1
            if want is not None and same(got["value"], want):
                per_field[f]["correct"] += 1
                c_correct += 1
            else:
                errors.append({"case": case["id"], "field": f, "got": got["value"], "expected": want,
                               "evidence": got["evidence"], "method": got["method"], "score": got["score"]})
        multi_ok += result["multipleItems"] == case.get("multipleItems", False)
        cases.append({"id": case["id"], "desc": case["desc"], "filled": c_filled, "correct": c_correct,
                      "expected": sum(v is not None for v in expected.values()), "ms": result["elapsedMs"],
                      "llm": result["llm"], "multi": result["multipleItems"]})

    filled = sum(v["filled"] for v in per_field.values())
    correct = sum(v["correct"] for v in per_field.values())
    expected = sum(v["expected"] for v in per_field.values())
    return {
        "precision": correct / filled if filled else 0.0, "coverage": correct / expected if expected else 0.0,
        "filled": filled, "correct": correct, "expectedTotal": expected, "perField": per_field, "cases": cases,
        "errors": errors, "unparsed": unparsed, "multiAccuracy": multi_ok / len(cases_in), "caseCount": len(cases_in),
        "msMedian": statistics.median(times), "msMax": max(times),
    }


def report(r: dict, mode: str, model: str | None) -> str:
    pct = lambda x: f"{x * 100:.1f}%"  # noqa: E731
    lines = [
        f"# 추출 정밀도 측정 — {mode}{f' ({model})' if model else ''}",
        f"측정 시각: {datetime.now():%Y-%m-%d %H:%M} · 평가 세트 {r['caseCount']}건 (eval/dataset.py + eval/cases/) · 사전 {engine.dictionary()['version']}",
        "",
        "| 지표 | 값 | 기준 |", "|---|---|---|",
        f"| **정밀도** (맞게 채움 / 채움) | **{pct(r['precision'])}** ({r['correct']}/{r['filled']}) | NFR-01 90% 이상 |",
        f"| 채움률 (맞게 채움 / 정답 있음) | {pct(r['coverage'])} ({r['correct']}/{r['expectedTotal']}) | 높을수록 보완 요청 감소 |",
        f"| 다품목 판정 정확도 | {pct(r['multiAccuracy'])} | FR-210 |",
        f"| 처리 시간 (중앙값 / 최대) | {r['msMedian'] / 1000:.2f}s / {r['msMax'] / 1000:.2f}s | NFR-02 60초 |",
        "", "## 필드별", "| 필드 | 정밀도 | 채움률 | 채움 | 정답 있음 |", "|---|---|---|---|---|",
    ]
    for f, v in r["perField"].items():
        p = pct(v["correct"] / v["filled"]) if v["filled"] else "-"
        c = pct(v["correct"] / v["expected"]) if v["expected"] else "-"
        lines.append(f"| {f} | {p} | {c} | {v['filled']} | {v['expected']} |")
    lines += ["", "## 케이스별", "| 케이스 | 설명 | 맞음/채움 | 정답 있음 | 시간 | LLM |", "|---|---|---|---|---|---|"]
    for c in r["cases"]:
        llm = "사용" if c["llm"].get("used") else ("오류: " + c["llm"]["error"][:40] if c["llm"].get("error") else "-")
        lines.append(f"| {c['id']} | {c['desc']} | {c['correct']}/{c['filled']} | {c['expected']} | {c['ms'] / 1000:.2f}s | {llm} |")
    lines += ["", "## 오추출 목록 (정밀도를 깎은 값)"]
    if not r["errors"]:
        lines.append("없음")
    else:
        lines += ["| 케이스 | 필드 | 추출값 | 정답 | 방법·점수 | 근거 |", "|---|---|---|---|---|---|"]
        for e in r["errors"]:
            lines.append(f"| {e['case']} | {e['field']} | {e['got']} | {e['expected']} | {e['method']} {e['score']} | "
                         f"{(e['evidence'] or '').replace('|', '/')[:60]} |")
    lines += ["", "## 찾았지만 읽지 못한 값 (저장 안 함 → 원문을 짚어 확인 질문, 사전 보강 후보)"]
    if not r["unparsed"]:
        lines.append("없음")
    else:
        lines += ["| 케이스 | 필드 | 원문 | 정답 |", "|---|---|---|---|"]
        for u in r["unparsed"]:
            lines.append(f"| {u['case']} | {u['field']} | {(u['evidence'] or '').replace('|', '/')[:60]} | {u['expected']} |")
    return "\n".join(lines) + "\n"


def main() -> None:
    parser = argparse.ArgumentParser(description="추출 정밀도 측정 (NFR-01)")
    parser.add_argument("--mode", choices=["rules", "hybrid"], default="rules")
    parser.add_argument("--model", help="Ollama 모델 이름 (기본: .env LLM_MODEL)")
    parser.add_argument("--url", help="Ollama 주소 (기본: .env LLM_URL)")
    parser.add_argument("--scope", choices=["all", "missing"], help="LLM이 읽을 범위 (기본: .env LLM_SCOPE)")
    args = parser.parse_args()

    llm, model = None, None
    if args.mode == "hybrid":
        from app.config import settings
        from app.llm import OllamaClient

        model = args.model or settings.llm_model
        llm = OllamaClient(url=args.url, model=model)

    from app.config import settings as _s

    scope = args.scope or _s.llm_scope
    result = evaluate(llm, model, scope=scope)
    text = report(result, args.mode + (f"·{scope}" if llm else ""), model)
    REPORTS.mkdir(exist_ok=True)
    out = REPORTS / f"{datetime.now():%Y%m%d-%H%M}_{args.mode}{'_' + model.replace(':', '-') if model else ''}.md"
    out.write_text(text, encoding="utf-8")
    print(text)
    print(f"보고서: {out}")


if __name__ == "__main__":
    main()
