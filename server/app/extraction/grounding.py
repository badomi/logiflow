"""근거 검증 (FR-202): LLM이 인용한 문장이 실제 메일에 있는지, 값이 그 문장에서 나왔는지 확인한다.

점수 = 인용 일치도(원문과 얼마나 같은가) × 값 뒷받침(값의 숫자·글자가 인용 안에 있는가).
인용이 원문 어디에도 없으면 0점 → 버린다 (LLM이 지어낸 값 차단, 정밀도 우선).
"""

import difflib
import re

from .base import Source
from .parsers import clean

MIN_QUOTE_MATCH = 0.75


def _norm(text: str) -> str:
    return re.sub(r"[^0-9a-z가-힣.]+", "", clean(text).lower())


def _digits(text: str) -> list[str]:
    return re.findall(r"\d+(?:\.\d+)?", clean(text).replace(",", ""))


def value_supported(value: str, quote: str) -> float:
    """값이 인용 안에 있는가: 숫자는 모두 들어 있어야 1.0, 글자는 겹치는 비율."""
    v, q = _norm(value), _norm(quote)
    if not v:
        return 0.0
    if v in q:
        return 1.0
    numbers = _digits(value)
    if numbers:
        quoted = set(_digits(quote))
        return 1.0 if all(n in quoted for n in numbers) else 0.0
    return difflib.SequenceMatcher(None, v, q).find_longest_match(0, len(v), 0, len(q)).size / len(v)


def locate(quote: str, sources: list[Source]) -> tuple[float, Source | None]:
    """인용이 가장 잘 맞는 출처와 일치도. 뒤 출처(최신)부터 본다."""
    target = _norm(quote)
    if not target:
        return 0.0, None
    best, best_src = 0.0, None
    for src in sorted(sources, key=lambda s: -s.order):
        whole = _norm(src.text)
        if target in whole:
            return 1.0, src
        lines = [_norm(ln) for ln in src.text.split("\n") if ln.strip()]
        windows = lines + [a + b for a, b in zip(lines, lines[1:])]
        for w in windows:
            ratio = difflib.SequenceMatcher(None, target, w).ratio()
            if ratio > best:
                best, best_src = ratio, src
    return best, best_src


def ground(quote: str | None, value: str | None, sources: list[Source]) -> tuple[float, str, int]:
    """→ (점수, 출처 라벨, 출처 순서). 근거 없으면 점수 0."""
    if not quote or not value:
        return 0.0, "", -1
    match, src = locate(quote, sources)
    if match < MIN_QUOTE_MATCH or src is None:
        return 0.0, "", -1
    # 비슷한 문장이어도 숫자가 원문과 다르면 탈락 ('500kg'을 '5,000kg'으로 옮긴 경우)
    if match < 1.0 and not set(_digits(quote)) <= set(_digits(src.text)):
        return 0.0, "", -1
    support = value_supported(value, quote)
    if support < 0.5:
        return 0.0, "", -1
    return round(match * support, 3), src.label, src.order


# 말로 쓴 수 ('오백', '반 톤', '한 대', 'five hundred') — 값의 숫자가 원문에 없어도 이런 말이 있으면 LLM이 정리한 것으로 본다
_NUMBER_WORDS = re.compile(
    r"(?:[일이삼사오육칠팔구]?[십백천만])+|반\s*(?:톤|t)|(?:한|두|세|네|다섯)\s*(?:대|개|박스|팔레트)|"
    r"\b(?:half|one|two|three|four|five|six|seven|eight|nine|ten|hundred|thousand)\b", re.I)


def support_lines(value: str, lines: list, index: list) -> tuple[float, object, str]:
    """LLM이 준 줄 번호로 근거를 찾고, 값이 그 줄에서 나왔는지 본다 → (지지도, 출처, 근거 원문).

    - 줄 번호가 없거나 범위 밖이면 0 (지어낸 근거)
    - 값의 숫자가 근거 줄에 모두 있으면 1.0, 말로 쓴 수가 있으면 0.8, 숫자가 다르면 0
    - 숫자 없는 값(품목·이름 등)은 글자가 겹치는 비율
    뜻이 같은지(하이큐브 = 40HQ, 싱가포르 = Singapore)는 llm_stage가 규칙 해석기로 한 번 더 본다.
    """
    valid = [int(n) for n in lines if isinstance(n, (int, float)) and 1 <= int(n) <= len(index)]
    if not valid:
        return 0.0, None, ""
    picked = [index[n - 1] for n in dict.fromkeys(valid)]
    evidence = " / ".join(text for _, text in picked)
    src = max((s for s, _ in picked), key=lambda s: s.order)
    numbers = _digits(value)
    if numbers:
        quoted = set(_digits(evidence))
        if all(n in quoted for n in numbers):
            return 1.0, src, evidence
        if _NUMBER_WORDS.search(clean(evidence)):
            return 0.8, src, evidence
        return 0.0, src, evidence
    return value_supported(value, evidence), src, evidence
