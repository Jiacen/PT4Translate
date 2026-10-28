"""Small shared Qwen-MT relay for portable teacher clients.

Only the relay host knows the upstream API key.  Its public route accepts the
single-text translation payload already emitted by the local Qwen-MT adapter.
Put this app behind an HTTPS reverse proxy; do not expose its HTTP port.
"""
from __future__ import annotations

import hmac
import json
import os
import threading
import time
from pathlib import Path
from urllib.parse import urlparse

import requests
from fastapi import FastAPI, Header, HTTPException
from pydantic import BaseModel, Field


class Message(BaseModel):
    role: str
    content: str = Field(min_length=1, max_length=12000)


class TermPair(BaseModel):
    source: str = Field(min_length=1, max_length=300)
    target: str = Field(min_length=1, max_length=300)


class TranslationOptions(BaseModel):
    source_lang: str = Field(max_length=40)
    target_lang: str = Field(max_length=40)
    terms: list[TermPair] = Field(default_factory=list, max_length=50)
    tm_list: list[TermPair] = Field(default_factory=list, max_length=50)
    domains: list[str] = Field(default_factory=list, max_length=10)


class TranslationPayload(BaseModel):
    model: str = Field(max_length=100)
    messages: list[Message] = Field(min_length=1, max_length=1)
    translation_options: TranslationOptions
    temperature: float = Field(default=0.1, ge=0, le=1)


app = FastAPI(title="SITRAIN 统一翻译入口", docs_url=None, redoc_url=None, openapi_url=None)
_gate = threading.Semaphore(8)
_clock_lock = threading.Lock()
_next_start = 0.0


def _teacher_name(bearer: str) -> str:
    path = os.getenv("SITRAIN_RELAY_TOKENS_FILE", "")
    if not path or not Path(path).is_file():
        raise HTTPException(503, "翻译入口尚未配置老师凭证")
    try:
        tokens = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise HTTPException(503, "翻译入口凭证不可用") from exc
    if not isinstance(tokens, dict) or len(tokens) > 100:
        raise HTTPException(503, "翻译入口凭证不可用")
    presented = bearer.removeprefix("Bearer ") if bearer.startswith("Bearer ") else ""
    result = ""
    for teacher, token in tokens.items():
        if isinstance(token, str) and len(token) >= 32 and hmac.compare_digest(presented, token):
            result = str(teacher)
    if not result:
        raise HTTPException(401, "访问凭证无效")
    return result


def _upstream_url() -> str:
    url = os.getenv("SITRAIN_RELAY_UPSTREAM_URL", "https://dashscope.aliyuncs.com/compatible-mode/v1").rstrip("/")
    parsed = urlparse(url)
    if parsed.scheme != "https" or not parsed.netloc or parsed.username or parsed.password:
        raise HTTPException(503, "上游翻译地址配置无效")
    return url + "/chat/completions"


def _wait_for_slot() -> None:
    global _next_start
    rpm = max(1, min(600, int(os.getenv("SITRAIN_RELAY_RPM", "60"))))
    with _clock_lock:
        now = time.monotonic()
        scheduled = max(now, _next_start)
        _next_start = scheduled + 60.0 / rpm
    delay = scheduled - now
    if delay > 0:
        time.sleep(delay)


@app.get("/health")
def health():
    return {"ok": True}


@app.post("/v1/chat/completions")
def translate(payload: TranslationPayload, authorization: str = Header(default="")):
    _teacher_name(authorization)
    if payload.messages[0].role != "user":
        raise HTTPException(422, "仅接受翻译文字")
    key = os.getenv("SITRAIN_RELAY_UPSTREAM_KEY", "")
    if not key:
        raise HTTPException(503, "统一翻译服务尚未配置上游密钥")
    body = payload.model_dump(exclude_none=True)
    body["model"] = os.getenv("SITRAIN_RELAY_MODEL", "qwen-mt-plus")
    if not _gate.acquire(timeout=180):
        raise HTTPException(503, "翻译请求较多，请稍后重试")
    try:
        _wait_for_slot()
        try:
            response = requests.post(_upstream_url(), headers={"Authorization": f"Bearer {key}",
                                    "Content-Type": "application/json"}, json=body, timeout=130)
        except requests.RequestException as exc:
            raise HTTPException(502, "上游翻译服务暂时不可用") from exc
        if response.status_code >= 400:
            raise HTTPException(502, f"上游翻译服务返回 HTTP {response.status_code}")
        try:
            result = response.json()
            content = result["choices"][0]["message"]["content"]
            if not isinstance(content, str):
                raise ValueError("invalid content")
        except (ValueError, KeyError, IndexError, TypeError) as exc:
            raise HTTPException(502, "上游翻译结果格式无效") from exc
        return {"choices": [{"message": {"content": content}}]}
    finally:
        _gate.release()
