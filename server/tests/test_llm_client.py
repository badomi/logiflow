"""Ollama 요청 형식 테스트 (실제 서버 없이 urlopen을 가로챈다)."""

import io
import json
import urllib.error

import pytest

from app import llm


class FakeResponse(io.BytesIO):
    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


def fake_urlopen(replies, sent):
    def _open(request, timeout):
        sent.append(json.loads(request.data))
        reply = replies.pop(0)
        if isinstance(reply, Exception):
            raise reply
        return FakeResponse(json.dumps({"message": {"content": reply}}).encode())
    return _open


def test_thinking_is_turned_off_and_context_set(monkeypatch):
    sent = []
    monkeypatch.setattr(llm.urllib.request, "urlopen", fake_urlopen(['{"fields": {}}'], sent))
    result = llm.OllamaClient(model="qwen3:8b").complete_json("sys", "user", {"type": "object"})
    assert result == {"fields": {}}
    assert sent[0]["think"] is False and sent[0]["options"]["num_ctx"] >= 8192 and sent[0]["format"]


def test_model_without_think_option_is_retried(monkeypatch):
    sent = []
    err = urllib.error.HTTPError("u", 400, "bad", {}, io.BytesIO(b'{"error":"model does not support thinking"}'))
    monkeypatch.setattr(llm.urllib.request, "urlopen", fake_urlopen([err, '{"ok": true}'], sent))
    assert llm.OllamaClient(model="qwen2.5:7b-instruct").complete_json("s", "u", {}) == {"ok": True}
    assert "think" in sent[0] and "think" not in sent[1]


def test_think_block_is_stripped(monkeypatch):
    monkeypatch.setattr(llm.urllib.request, "urlopen", fake_urlopen(['<think>음…</think>{"a": 1}'], []))
    assert llm.OllamaClient().complete_json("s", "u", {}) == {"a": 1}


def test_other_http_errors_raise(monkeypatch):
    err = urllib.error.HTTPError("u", 404, "nf", {}, io.BytesIO(b'{"error":"model not found"}'))
    monkeypatch.setattr(llm.urllib.request, "urlopen", fake_urlopen([err], []))
    with pytest.raises(llm.LlmError, match="모델"):
        llm.OllamaClient(model="없는모델").complete_json("s", "u", {})



def test_timeout_message_says_model_is_slow(monkeypatch):
    monkeypatch.setattr(llm.urllib.request, "urlopen", fake_urlopen([TimeoutError("timed out")], []))
    with pytest.raises(llm.LlmError, match="초 안에 오지 않았습니다"):
        llm.OllamaClient(model="qwen3:4b", timeout_s=45).complete_json("s", "u", {})
