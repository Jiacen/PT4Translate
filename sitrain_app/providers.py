from __future__ import annotations

import json
import os
import random
import re
import time
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from typing import Callable

import requests


@dataclass
class TranslationRequest:
    units: list[tuple[str, str]]
    source_lang: str
    target_lang: str
    terms: list[dict]
    memory: list[dict]


class ProviderError(RuntimeError):
    pass


def _key(config: dict) -> str:
    env_name = config.get("config", {}).get("api_key_env")
    return os.getenv(env_name, "") if env_name else config.get("api_key", "")


class _RateLimiter:
    """Small, provider-local start-rate limiter shared by concurrent Qwen calls."""

    def __init__(self, requests_per_minute: int):
        self.interval = 60.0 / requests_per_minute if requests_per_minute > 0 else 0.0
        self.next_start = 0.0
        self.lock = threading.Lock()

    def wait(self) -> None:
        if not self.interval:
            return
        with self.lock:
            now = time.monotonic()
            scheduled = max(now, self.next_start)
            self.next_start = scheduled + self.interval
        delay = scheduled - now
        if delay > 0:
            time.sleep(delay)


def _post(config: dict, payload: dict) -> dict:
    key = _key(config)
    if not key:
        raise ProviderError(f"Provider {config['name']} 未配置 API Key")
    settings = config.get("config", {})
    headers = {"Authorization":f"Bearer {key}","Content-Type":"application/json",**settings.get("headers",{})}
    url = config["base_url"].rstrip("/") + "/chat/completions"
    retries = max(0,int(settings.get("retries",3)))
    for attempt in range(retries + 1):
        try:
            response = requests.post(url,headers=headers,json=payload,timeout=float(settings.get("timeout",120)))
            if response.status_code >= 400:
                retry_after = response.headers.get("Retry-After") if response.status_code == 429 else None
                if retry_after:
                    try:
                        time.sleep(min(60.0, max(0.0, float(retry_after))))
                    except ValueError:
                        pass
                raise ProviderError(f"HTTP {response.status_code}: {response.text[:300]}")
            return response.json()
        except (requests.RequestException, ValueError, ProviderError) as exc:
            if attempt >= retries:
                raise ProviderError(str(exc)) from exc
            time.sleep(float(settings.get("backoff",2)) * (2 ** attempt) + random.random())
    raise ProviderError("unreachable")


def _choice_text(response: dict) -> str:
    try:
        content = response["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as exc:
        raise ProviderError("Provider 返回结构无 choices[0].message.content") from exc
    if not isinstance(content,str) or not content.strip():
        raise ProviderError("Provider 返回空译文")
    return content.strip()


def translate_qwen(config: dict, request: TranslationRequest, progress: Callable[[int,int],None]) -> dict[str,str]:
    settings = config.get("config",{})
    # Qwen-MT accepts one source text per request.  Parallelise only the network
    # calls, so each stable unit ID still maps to exactly one returned translation.
    concurrency = min(8, max(1, int(settings.get("concurrency", 4))))
    limiter = _RateLimiter(max(0, int(settings.get("rpm", 0))))

    def translate_one(key: str, text: str) -> tuple[str, str]:
        options = {"source_lang":request.source_lang,"target_lang":request.target_lang}
        if request.terms:
            options["terms"] = [{"source":t["source"],"target":t["target"]} for t in request.terms[:50]]
        if request.memory:
            options["tm_list"] = [{"source":m["source"],"target":m["target"]} for m in request.memory[:50]]
        if settings.get("domains"):
            options["domains"] = settings["domains"]
        payload = {"model":config["model"],"messages":[{"role":"user","content":text}],
                   "temperature":settings.get("temperature",0.1),"translation_options":options}
        limiter.wait()
        return key, _choice_text(_post(config,payload))

    output: dict[str, str] = {}
    with ThreadPoolExecutor(max_workers=concurrency, thread_name_prefix="qwen-mt") as executor:
        futures = [executor.submit(translate_one, key, text) for key, text in request.units]
        for completed, future in enumerate(as_completed(futures), 1):
            key, translated = future.result()
            output[key] = translated
            progress(completed, len(request.units))
    return output


def _json_object(text: str) -> dict:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.I)
    try:
        value = json.loads(text)
    except json.JSONDecodeError as exc:
        raise ProviderError(f"Provider 未返回有效 JSON：{exc}") from exc
    if not isinstance(value,dict):
        raise ProviderError("Provider JSON 顶层必须是对象")
    return value


def translate_openai(config: dict, request: TranslationRequest, progress: Callable[[int,int],None]) -> dict[str,str]:
    settings = config.get("config",{})
    batch_size = max(1,int(settings.get("batch_size",20)))
    output = {}
    for offset in range(0,len(request.units),batch_size):
        batch = request.units[offset:offset+batch_size]
        body = {key:text for key,text in batch}
        term_text = "\n".join(f"{t['source']} => {t['target']}" for t in request.terms)
        prompt = settings.get("prompt") or "Translate each item faithfully from {source_lang} into {target_lang}. Return only a JSON object using the same keys."
        source_name = "the automatically detected source language" if request.source_lang == "auto" else request.source_lang
        prompt = prompt.format(source_lang=source_name,target_lang=request.target_lang)
        if term_text:
            prompt += "\nMandatory terminology:\n" + term_text
        payload = {"model":config["model"],"messages":[{"role":"system","content":prompt},
                   {"role":"user","content":json.dumps(body,ensure_ascii=False)}],"temperature":settings.get("temperature",0.1)}
        if settings.get("structured_output",True):
            payload["response_format"] = {"type":"json_object"}
        value = _json_object(_choice_text(_post(config,payload)))
        expected = {key for key,_ in batch}
        if set(value) != expected or not all(isinstance(value[k],str) for k in expected if k in value):
            raise ProviderError(f"Provider JSON 键集合不匹配：expected={sorted(expected)}, actual={sorted(value)}")
        output.update(value)
        progress(min(offset+len(batch),len(request.units)),len(request.units))
    return output


def translate(config: dict, request: TranslationRequest, progress=lambda a,b:None) -> dict[str,str]:
    if config["protocol"] == "qwen_mt":
        return translate_qwen(config,request,progress)
    if config["protocol"] == "mock":
        result = {key:f"[机翻]{text}" for key,text in request.units}
        progress(len(result),len(result))
        return result
    return translate_openai(config,request,progress)
