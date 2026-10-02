"""모델 비교 측정 — 여러 로컬 LLM을 같은 평가 세트로 돌려 한 표로 비교하고, 쓸 모델을 추천한다.

    cd server
    .venv\\Scripts\\python -m eval.compare --models qwen3.5:9b qwen3:14b
    (또는 server\\run_benchmark.bat 더블클릭)

추천 기준 (정의서):
  1) 정밀도 90% 이상 (NFR-01)          2) 메일 1건 최대 45초 이내 (NFR-02, LLM 제한 시간)
  3) LLM 오류(시간 초과 등) 10% 이하    → 통과한 모델 중 채움률이 가장 높은 것, 같으면 빠른 것
보고서: eval/reports/compare_<시각>.md
"""

import argparse
import sys
import time
from datetime import datetime
from pathlib import Path

SERVER = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SERVER))

from eval.dataset import all_cases  # noqa: E402
from eval.run_eval import REPORTS, evaluate  # noqa: E402

PRECISION_MIN, MAX_MS, ERROR_RATE_MAX = 0.90, 45_000, 0.10


def warm_up(client) -> float:
    """첫 호출은 모델을 메모리에 올리느라 느리다 → 측정 전에 한 번 불러 둔다 (걸린 초를 돌려준다)."""
    started = time.monotonic()
    try:
        client.complete_json("JSON으로만 답하라.", "ok 라고만 답하라", {"type": "object", "properties": {"ok": {"type": "string"}}})
    except Exception:  # 로딩 실패는 본 측정에서 오류로 드러난다
        pass
    return time.monotonic() - started


def summarize(name: str, r: dict, warm: float | None) -> dict:
    errors = [c for c in r["cases"] if c["llm"].get("error")]
    used = [c for c in r["cases"] if c["llm"].get("used")]
    attempted = [c for c in r["cases"] if c["llm"].get("asked")]  # 다품목 등 LLM을 부르지 않은 케이스는 오류율에서 뺀다
    llm_ms = [c["llm"].get("elapsedMs", 0) for c in used]
    return {
        "name": name, "precision": r["precision"], "coverage": r["coverage"], "filled": r["filled"],
        "correct": r["correct"], "expected": r["expectedTotal"], "msMax": r["msMax"], "msMedian": r["msMedian"],
        "llmMsMax": max(llm_ms) if llm_ms else 0, "errorRate": len(errors) / max(1, len(attempted)),
        "errors": errors, "wrong": r["errors"], "unparsed": r["unparsed"], "cases": r["cases"], "warm": warm,
    }


def verdict(s: dict) -> list[str]:
    problems = []
    if s["precision"] < PRECISION_MIN:
        problems.append(f"정밀도 {s['precision'] * 100:.1f}% < 90%")
    if s["msMax"] > MAX_MS:
        problems.append(f"최대 {s['msMax'] / 1000:.0f}초 > 45초")
    if s["errorRate"] > ERROR_RATE_MAX:
        problems.append(f"LLM 오류 {s['errorRate'] * 100:.0f}% (시간 초과 등)")
    return problems


def recommend(rows: list[dict]) -> tuple[dict | None, str]:
    llm_rows = [r for r in rows if r["name"] != "규칙만"]
    ok = [r for r in llm_rows if not verdict(r)]
    if ok:
        best = max(ok, key=lambda r: (round(r["coverage"], 3), -r["msMax"]))
        base = next((r for r in rows if r["name"] == "규칙만"), None)
        gain = f" (규칙만 대비 채움률 +{(best['coverage'] - base['coverage']) * 100:.1f}%p)" if base else ""
        return best, f"**{best['name']}** 추천 — 기준 통과 모델 중 채움률이 가장 높음{gain}. `.env`의 `LLM_MODEL={best['name']}`"
    if llm_rows:
        why = "; ".join(f"{r['name']}: {', '.join(verdict(r))}" for r in llm_rows)
        return None, (f"기준을 통과한 모델이 없습니다 ({why}). 시간 초과가 원인이면 더 작은 모델 또는 `LLM_SCOPE=missing`, "
                      f"정밀도가 원인이면 오추출 목록을 보내 주세요.")
    return None, "측정한 LLM이 없습니다."


def report(rows: list[dict], scope: str, case_count: int) -> str:
    pct = lambda x: f"{x * 100:.1f}%"  # noqa: E731
    best, advice = recommend(rows)
    lines = [
        f"# 모델 비교 측정 — LLM_SCOPE={scope}",
        f"측정 시각: {datetime.now():%Y-%m-%d %H:%M} · 평가 세트 {case_count}건 · 기준: 정밀도 ≥90%, 최대 ≤45초, LLM 오류 ≤10%",
        "", "## 결론", advice, "",
        "## 비교", "| 모델 | 정밀도 | 채움률 | 처리 시간 중앙값 / 최대 | LLM 오류 | 첫 로딩 | 판정 |",
        "|---|---|---|---|---|---|---|",
    ]
    for r in rows:
        v = verdict(r) if r["name"] != "규칙만" else []
        mark = "기준" if r["name"] == "규칙만" else ("✅ 통과" if not v else "❌ " + ", ".join(v))
        warm = f"{r['warm']:.0f}초" if r.get("warm") is not None else "-"
        lines.append(f"| {r['name']}{' ⭐' if best is r else ''} | {pct(r['precision'])} ({r['correct']}/{r['filled']}) | "
                     f"{pct(r['coverage'])} ({r['correct']}/{r['expected']}) | "
                     f"{r['msMedian'] / 1000:.1f}s / {r['msMax'] / 1000:.1f}s | {pct(r['errorRate'])} | {warm} | {mark} |")

    lines += ["", "## 케이스별 맞음/채움", "| 케이스 | 설명 | " + " | ".join(r["name"] for r in rows) + " |",
              "|---|---|" + "---|" * len(rows)]
    for i, case in enumerate(rows[0]["cases"]):
        cells = []
        for r in rows:
            c = r["cases"][i]
            mark = " ⚠" if c["llm"].get("error") else ""
            cells.append(f"{c['correct']}/{c['filled']} ({c['ms'] / 1000:.1f}s){mark}")
        lines.append(f"| {case['id']} | {case['desc'][:40]} | " + " | ".join(cells) + " |")
    lines.append("")
    lines.append("⚠ = LLM 오류(시간 초과 등)로 그 케이스는 규칙 결과만 들어감")

    for r in rows:
        if r["name"] == "규칙만":
            continue
        lines += ["", f"## {r['name']} — 오추출 (정밀도를 깎은 값)"]
        if not r["wrong"]:
            lines.append("없음")
        else:
            lines += ["| 케이스 | 필드 | 추출값 | 정답 | 방법 | 근거 |", "|---|---|---|---|---|---|"]
            for e in r["wrong"]:
                lines.append(f"| {e['case']} | {e['field']} | {e['got']} | {e['expected']} | {e['method']} | "
                             f"{(e['evidence'] or '').replace('|', '/')[:50]} |")
        if r["errors"]:
            lines.append("")
            lines.append("LLM 오류: " + "; ".join(f"{c['id']}: {c['llm']['error'][:60]}" for c in r["errors"][:5]))
    return "\n".join(lines) + "\n"


def compare(clients: dict, scope: str = "all", cases: list[dict] | None = None, warm: bool = True) -> tuple[list[dict], str]:
    """clients = {이름: LLM 클라이언트}. 규칙만 결과를 기준으로 함께 잰다."""
    cases = cases if cases is not None else all_cases()
    rows = [summarize("규칙만", evaluate(None, None, cases=cases, scope=scope), None)]
    for name, client in clients.items():
        print(f"  {name} 측정 중… (케이스 {len(cases)}건)", flush=True)
        warm_s = warm_up(client) if warm else None
        rows.append(summarize(name, evaluate(client, name, cases=cases, scope=scope), warm_s))
    return rows, report(rows, scope, len(cases))


def main() -> None:
    from app.config import settings
    from app.diagnostics import _ollama
    from app.llm import OllamaClient

    parser = argparse.ArgumentParser(description="로컬 LLM 모델 비교 측정")
    parser.add_argument("--models", nargs="+", default=[settings.llm_model], help="비교할 Ollama 모델 이름들")
    parser.add_argument("--scope", choices=["all", "missing"], default=settings.llm_scope)
    parser.add_argument("--url", default=settings.llm_url)
    args = parser.parse_args()

    clients = {}
    for model in args.models:
        info = _ollama(args.url, model)
        if not info["reachable"]:
            print(f"Ollama({args.url})에 연결할 수 없습니다: {info.get('error')} — Ollama를 켜고 다시 실행하세요.")
            sys.exit(1)
        if not info["modelInstalled"]:
            print(f"  {model}: 설치되지 않아 건너뜁니다 (ollama pull {model})")
            continue
        clients[model] = OllamaClient(url=args.url, model=model)
    if not clients:
        print("측정할 모델이 없습니다.")
        sys.exit(1)

    print(f"모델 비교 측정 시작 — {', '.join(clients)} (LLM_SCOPE={args.scope})")
    _, text = compare(clients, args.scope)
    REPORTS.mkdir(exist_ok=True)
    out = REPORTS / f"compare_{datetime.now():%Y%m%d-%H%M}.md"
    out.write_text(text, encoding="utf-8")
    print()
    print(text)
    print(f"보고서: {out}")


if __name__ == "__main__":
    main()
