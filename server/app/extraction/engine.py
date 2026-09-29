"""추출 엔진: 룰 + 서명 + 로컬 LLM 후보를 합쳐 필드마다 값·근거·점수·상태를 정한다 (FR-201·202·209).

결합 규칙 (정밀도 우선, NFR-01)
  - 룰과 LLM이 같은 값        → 점수 +0.1 (두 방법이 독립적으로 동의)
  - 룰과 LLM이 다른 값        → 점수 높은 쪽을 쓰되 -0.25, 상태 conflict (대부분 임계치 아래 → 검토 필요)
  - 한쪽만 있음               → 그 점수
  - 최신 메일(회신)의 후보가 임계치를 넘으면 앞 메일 값보다 우선 (FR-207 부족분 병합)
상태: filled(임계치 이상 → 저장) / review(값은 있으나 임계치 미만 → 저장 안 함, 검토 필요) / missing

DB 없이 동작한다 — 파이프라인(extractor.py)과 정밀도 측정(eval/run_eval.py)이 같은 엔진을 쓴다.
"""

import time

from . import grounding, llm_stage, parsers, rules, signature
from .base import EXTRA_FIELDS, FIELDS, Candidate, MailText, Source, dictionary, threshold

SCHEMA_VERSION = "extract-v1"
CARGO_FIELDS = tuple(f for f in FIELDS if f not in ("customerName", "contactName", "contactEmail"))


def build_sources(mails: list[MailText]) -> tuple[list[Source], list[Candidate]]:
    """메일들 → 추출 대상 텍스트(출처) + 서명·헤더 후보."""
    from .preprocess import split_mail

    sources: list[Source] = []
    people: list[Candidate] = []
    for i, mail in enumerate(mails, start=1):
        role = "최초 요청" if mail.role == "ORIGINAL" else "회신"
        base = i * 100
        parts = split_mail(mail.body)
        sources.append(Source(base, f"메일 {i} · 제목", "subject", mail.subject or ""))
        sources.append(Source(base + 1, f"메일 {i} · {role} 본문", "body", parts.body))
        for j, (filename, text) in enumerate(mail.attachments, start=2):
            if text.strip():
                sources.append(Source(base + j, f"메일 {i} · 첨부 {filename}", "attachment", text))
        if parts.signature:
            sources.append(Source(base + 50, f"메일 {i} · 서명", "signature", parts.signature))
        people += signature.extract(parts.signature, mail.from_name, mail.from_email, f"메일 {i}", base + 50)
    return sources, people


def _same(a, b) -> bool:
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= max(abs(a), abs(b)) * 0.01
    x, y = str(a).strip().lower(), str(b).strip().lower()
    if x == y:
        return True
    # 글자 값은 한쪽이 다른 쪽을 포함하면 같은 값으로 본다 ('스프링노트' ⊂ '스프링노트(일반 문구류)')
    short, long_ = sorted((x, y), key=len)
    return len(short) >= 2 and short in long_


def _pick(cands: list[Candidate], field: str) -> Candidate | None:
    """최신 출처부터, 임계치를 넘는 첫 후보. 없으면 점수가 가장 높은 후보."""
    if not cands:
        return None
    t = threshold(field)
    ordered = sorted(cands, key=lambda c: (-c.order, -c.score))
    return next((c for c in ordered if c.score >= t), max(cands, key=lambda c: c.score))


def combine(field: str, rule_c: Candidate | None, llm_c: Candidate | None) -> dict:
    t = threshold(field)
    if rule_c is None and llm_c is None:
        return {"value": None, "score": 0.0, "evidence": None, "origin": None, "method": None, "status": "missing"}
    if rule_c and llm_c:
        if _same(rule_c.value, llm_c.value):
            chosen = rule_c if rule_c.order >= llm_c.order else llm_c
            if isinstance(rule_c.value, str) and isinstance(llm_c.value, str) and rule_c.order == llm_c.order:
                chosen = max((rule_c, llm_c), key=lambda c: len(c.value))  # 더 온전한 표기
            score, method, conflict = min(1.0, max(rule_c.score, llm_c.score) + 0.1), "rule+llm", None
        elif abs(rule_c.order - llm_c.order) >= 100 and max(rule_c.score, llm_c.score) >= t:
            chosen = rule_c if rule_c.order > llm_c.order else llm_c  # 다른 메일이면 최신 메일 값
            score, method, conflict = chosen.score, chosen.method, None
        else:
            chosen = rule_c if rule_c.score >= llm_c.score else llm_c
            score, method = max(0.0, chosen.score - 0.25), chosen.method
            conflict = {"rule": rule_c.value, "llm": llm_c.value}
    else:
        chosen = rule_c or llm_c
        score, method, conflict = chosen.score, chosen.method, None
    status = "filled" if score >= t and not conflict else "review"  # 룰·LLM이 엇갈리면 저장하지 않는다
    result = {
        "value": chosen.value, "score": round(score, 3), "evidence": chosen.evidence, "origin": chosen.origin,
        "method": method, "status": status,
    }
    if conflict:
        result["conflict"] = conflict
    return result


def run(mails: list[MailText], llm=None, model_name: str | None = None) -> dict:
    """고정 스키마 JSON (FR-201). 각 필드 = {value, score, evidence, origin, method, status}."""
    began = time.monotonic()
    sources, people = build_sources(mails)
    rule_cands, notes, multi_rule = rules.extract(sources)
    rule_cands += people

    llm_info: dict = {"used": False, "model": model_name}
    llm_cands: list[Candidate] = []
    multi_llm = False
    if llm is not None:
        rule_only = {n: combine(n, _pick([c for c in rule_cands if c.field == n], n), None) for n in FIELDS}
        keys = llm_targets(rule_only, notes, multi_rule, sources, rule_cands)
        llm_info["asked"] = keys
        if not keys:
            llm_info["skipped"] = "다품목" if multi_rule else "규칙으로 모두 찾음"
    if llm is not None and keys:
        started = time.monotonic()
        try:
            raw = llm.complete_json(llm_stage.SYSTEM_PROMPT, llm_stage.build_prompt(sources, keys),
                                    llm_stage.schema_for(keys))
            llm_cands, llm_notes, multi_llm = llm_stage.to_candidates(raw, sources, grounding.ground, asked=keys)
            notes += [{**n, "origin": "LLM"} for n in llm_notes]
            llm_info.update(used=True, dropped=_count_dropped(raw, llm_cands))
        except Exception as error:  # LLM이 없거나 실패해도 룰 결과로 계속 (NFR-03)
            llm_info["error"] = str(error)
        llm_info["elapsedMs"] = int((time.monotonic() - started) * 1000)

    fields: dict[str, dict] = {}
    for name in (*FIELDS, *EXTRA_FIELDS):
        r = _pick([c for c in rule_cands if c.field == name], name)
        m = _pick([c for c in llm_cands if c.field == name], name)
        fields[name] = combine(name, r, m)

    _keep_box_together(fields)
    multi = multi_rule or multi_llm
    if multi:  # FR-210: 품목별 값이 섞이지 않도록 화물 필드는 저장하지 않는다 (고객 정보는 유지)
        for name in CARGO_FIELDS:
            if fields[name]["status"] == "filled":
                fields[name]["status"] = "review"
                fields[name]["reason"] = "MULTIPLE_ITEMS"
    return {
        "schemaVersion": SCHEMA_VERSION,
        "dictionaryVersion": dictionary().get("version"),
        "fields": fields,
        "multipleItems": multi,
        "multipleItemsBy": [k for k, v in (("rule", multi_rule), ("llm", multi_llm)) if v],
        "notes": notes,
        "llm": llm_info,
        "elapsedMs": int((time.monotonic() - began) * 1000),
    }


# LLM 항목 → 사전(fields.json)의 라벨 그룹 (회신에 그 항목 이야기가 나오는지 볼 때 쓴다)
_LABEL_GROUP = {"quantity": "qty", "boxDimensions": "box", "totalVolume": "totalCbm", "grossWeight": "grossWeightKg",
                "portOfLoading": "pol", "portOfDischarge": "pod", "container": "containerType"}


def llm_targets(rule_only: dict, notes: list[dict], multi_rule: bool, sources: list[Source],
                rule_cands: list[Candidate] | None = None) -> list[str]:
    """LLM에게 물어볼 항목: 룰이 못 찾은 것 + 회신이 문장으로 고쳤을 수 있는 것.

    CPU만 있는 PC에서도 60초(NFR-02) 안에 끝나도록 출력을 줄인다. 룰이 다 찾았으면 LLM을 부르지 않는다.
    """
    if multi_rule:
        return []  # 다품목은 화물 값을 저장하지 않으므로 물어볼 필요 없음 (FR-210)
    filled = {n for n, f in rule_only.items() if f["status"] == "filled"}
    has_volume = {"boxL", "boxW", "boxH"} <= filled or "totalCbm" in filled
    multiple_terms = any(n["code"] == "MULTIPLE_INCOTERMS" for n in notes)

    replies = [s for s in sources if s.kind == "body" and s.order >= 200]
    latest_reply = max(replies, key=lambda s: s.order) if replies else None
    # 회신에서 룰이 이미 '라벨: 값'으로 읽은 줄은 뺀다 ('화물 준비일: …'의 '화물'을 품목 이야기로 오해하지 않게)
    parsed_lines = {c.evidence.strip() for c in rule_cands or [] if latest_reply and c.origin == latest_reply.label}
    reply_text = parsers.clean("\n".join(
        ln for ln in (latest_reply.text.split("\n") if latest_reply else []) if ln.strip() not in parsed_lines
    )).lower()
    labels = dictionary()["labels"]

    keys: list[str] = []
    for key, names in llm_stage.KEY_FIELDS.items():
        if key == "contactEmail" and "contactEmail" in filled:
            continue
        if key in ("boxDimensions", "totalVolume") and has_volume:
            need = False  # 규격이나 부피 중 하나만 있어도 견적 가능
        elif key == "incoterms" and multiple_terms:
            need = False  # 조건이 여러 개인 건 LLM도 못 정한다 → 보완 요청
        else:
            need = not set(names) <= filled
        if not need and latest_reply is not None and key not in llm_stage.CONTACT_KEYS:
            older = all((rule_only[n]["origin"] or "").split(" · ")[0] != latest_reply.label.split(" · ")[0]
                        for n in names if n in filled)
            words = [w.lower() for w in labels.get(_LABEL_GROUP.get(key, key), []) if len(w) >= 3 or not w.isascii()]
            need = older and any(w in reply_text for w in words)  # 회신이 이 항목을 문장으로 고쳤을 수 있음
        if need:
            keys.append(key)
    # 고객사명·담당자명은 필수가 아니고 없으면 헤더 발신자로 대신한다 → 이것만을 위해 LLM을 부르지 않는다
    if all(k in llm_stage.CONTACT_KEYS for k in keys):
        return []
    return keys


def _keep_box_together(fields: dict) -> None:
    """박스 규격은 셋이 한 벌 — 하나라도 저장 못 하면 셋 다 검토로 (부분 규격으로 부피 계산 방지)."""
    box = [fields[k] for k in ("boxL", "boxW", "boxH")]
    if any(b["status"] != "filled" for b in box) and any(b["status"] == "filled" for b in box):
        for b in box:
            if b["status"] == "filled":
                b["status"] = "review"


def _count_dropped(raw: dict, kept: list[Candidate]) -> int:
    """LLM이 값을 냈지만 근거 검증에서 버려진 개수 (환각 의심) — 로그용."""
    given = sum(1 for v in (raw.get("fields") or {}).values() if isinstance(v, dict) and v.get("value"))
    kept_keys = {c.evidence for c in kept}
    return max(0, given - len(kept_keys))


def saved_values(result: dict) -> dict:
    """quote_inputs에 저장할 값 (filled만)."""
    return {k: v["value"] for k, v in result["fields"].items() if k in FIELDS and v["status"] == "filled"}
