"""로컬 LLM 호출 (Ollama HTTP API).

- 메일 본문·고객정보를 외부로 보내지 않는다 (NFR-04): localhost·사설 IP 주소가 아니면 호출 자체를 거부한다.
- 표준 라이브러리(urllib)만 쓴다 — 새 의존성 없음.
- 응답은 JSON 스키마로 형식을 강제해서 받는다 (Ollama `format` 옵션). 그래도 형식이 깨지면 LlmError.
"""

import ipaddress
import json
import socket
import urllib.error
import urllib.request
from typing import Protocol
from urllib.parse import urlparse

from .config import settings


class LlmError(RuntimeError):
    """LLM 호출 실패 — 파이프라인이 잡아서 케이스 상태 '실패' + 이력으로 남긴다 (NFR-03)."""


class _ThinkUnsupported(Exception):
    """이 모델은 think 옵션을 모른다 → 옵션 없이 다시 요청"""


class LlmClient(Protocol):
    def complete_json(self, system: str, user: str, schema: dict) -> dict: ...


def ensure_local(url: str) -> None:
    """LLM 주소가 이 PC 또는 사내망(사설 IP)인지 확인한다. 아니면 실행 거부 (NFR-04)."""
    host = urlparse(url).hostname
    if not host:
        raise LlmError(f"LLM 주소가 올바르지 않습니다: {url}")
    try:
        addresses = {info[4][0] for info in socket.getaddrinfo(host, None)}
    except socket.gaierror as error:
        raise LlmError(f"LLM 주소를 찾을 수 없습니다: {host}") from error
    for address in addresses:
        ip = ipaddress.ip_address(address)
        if not (ip.is_loopback or ip.is_private):
            raise LlmError(f"외부 주소({host})의 LLM은 쓸 수 없습니다. 로컬·사내망 주소만 허용됩니다 (NFR-04).")


class OllamaClient:
    def __init__(self, url: str | None = None, model: str | None = None, timeout_s: float | None = None):
        self.url = (url or settings.llm_url).rstrip("/")
        self.model = model or settings.llm_model
        self.timeout_s = timeout_s or settings.llm_timeout_s
        ensure_local(self.url)

    def complete_json(self, system: str, user: str, schema: dict) -> dict:
        try:
            return self._chat(system, user, schema, send_think=True)
        except _ThinkUnsupported:
            return self._chat(system, user, schema, send_think=False)  # 생각 모드가 없는 모델(qwen2.5 등)

    def _timeout_message(self) -> str:
        return (f"LLM 응답이 {self.timeout_s:.0f}초 안에 오지 않았습니다 ({self.model}이 이 PC에서 느림 — "
                f"더 작은 모델 또는 LLM_TIMEOUT_S 확인)")

    def _chat(self, system: str, user: str, schema: dict, send_think: bool) -> dict:
        payload = {
                "model": self.model,
                "stream": False,
                "format": schema,
                "keep_alive": settings.llm_keep_alive,
                "options": {"temperature": 0, "num_ctx": settings.llm_num_ctx, "num_predict": 1500},
                "messages": [{"role": "system", "content": system}, {"role": "user", "content": user}],
        }
        if send_think:
            payload["think"] = settings.llm_think  # Qwen3 등은 기본이 생각 모드 켬 → 명시적으로 끈다
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        request = urllib.request.Request(
            f"{self.url}/api/chat", data=body, headers={"Content-Type": "application/json"}, method="POST"
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout_s) as response:
                payload = json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as error:
            detail = error.read().decode("utf-8", errors="replace") if error.fp else ""
            if send_think and error.code == 400 and "think" in detail.lower():
                raise _ThinkUnsupported from error
            raise LlmError(f"LLM 서버 오류 {error.code} — 모델({self.model}) 설치 여부를 확인하세요.") from error
        except (TimeoutError, socket.timeout) as error:
            raise LlmError(self._timeout_message()) from error
        except urllib.error.URLError as error:
            if isinstance(error.reason, (TimeoutError, socket.timeout)):
                raise LlmError(self._timeout_message()) from error
            raise LlmError(f"LLM 서버({self.url})에 연결하지 못했습니다 — Ollama가 켜져 있는지 확인: {error.reason}") from error

        content = (payload.get("message") or {}).get("content", "")
        if "</think>" in content:  # think 옵션을 무시하는 구버전 Ollama: 추론 부분을 떼고 JSON만
            content = content.split("</think>", 1)[1]
        try:
            result = json.loads(content)
        except json.JSONDecodeError as error:
            raise LlmError("LLM 응답이 JSON 형식이 아닙니다.") from error
        if not isinstance(result, dict):
            raise LlmError("LLM 응답이 JSON 객체가 아닙니다.")
        return result
