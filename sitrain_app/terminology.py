from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass


def normalize(value: str) -> str:
    value = unicodedata.normalize("NFKC", value or "").casefold()
    value = value.replace("‐", "-").replace("‑", "-").replace("–", "-").replace("—", "-")
    value = re.sub(r"\s+", " ", value)
    return value.strip()


@dataclass(frozen=True)
class Term:
    id: str
    source: str
    target: str
    aliases: tuple[str, ...] = ()
    mode: str = "required"
    case_sensitive: bool = False


class TerminologyEngine:
    def __init__(self, terms: list[Term]):
        self.terms = sorted(terms, key=lambda t: max([len(t.source), *map(len, t.aliases)]), reverse=True)

    def protect(self, text: str) -> tuple[str, dict[str, tuple[Term, str]]]:
        protected, mapping = text, {}
        for term in self.terms:
            variants = sorted({term.source, *term.aliases}, key=len, reverse=True)
            for variant in variants:
                flags = 0 if term.case_sensitive else re.IGNORECASE
                pattern = re.compile(r"(?<!\w)" + re.escape(variant) + r"(?!\w)", flags)
                if pattern.search(protected):
                    token = f"\ue000TERM_{len(mapping):04d}\ue001"
                    protected = pattern.sub(token, protected)
                    mapping[token] = (term, variant)
                    break
        return protected, mapping

    @staticmethod
    def restore(translated: str, mapping: dict[str, tuple[Term, str]]) -> str:
        for token, (term, source_variant) in mapping.items():
            replacement = source_variant if term.mode in {"keep", "forbidden"} else term.target
            translated = translated.replace(token, replacement)
        return translated

    def validate(self, source: str, translated: str) -> list[dict]:
        issues = []
        src_norm, dst_norm = normalize(source), normalize(translated)
        for term in self.terms:
            present = any(normalize(v) in src_norm for v in (term.source, *term.aliases))
            if not present:
                continue
            required = term.source if term.mode in {"keep", "forbidden"} else term.target
            if normalize(required) not in dst_norm:
                issues.append({"code":"TERM_MISSING","severity":"blocker" if term.mode != "recommended" else "warning",
                               "term_id":term.id,"message":f"必须使用术语：{term.source} → {required}"})
        return issues


_LATIN_WORD = re.compile(r"[A-Za-z][A-Za-z0-9_-]{2,}")


def residual_source_issues(source: str, translated: str) -> list[dict]:
    issues = []
    source_words = {w.casefold() for w in _LATIN_WORD.findall(source)}
    target_words = {w.casefold() for w in _LATIN_WORD.findall(translated)}
    residual = sorted(source_words & target_words)
    if residual and len(residual) >= max(2, len(source_words) // 2):
        issues.append({"code":"SOURCE_RESIDUAL","severity":"warning","message":"译文可能残留较多原文","tokens":residual[:12]})
    for number in re.findall(r"\b\d+(?:[.,]\d+)?\b", source):
        if number not in translated:
            issues.append({"code":"NUMBER_CHANGED","severity":"blocker","message":f"数字可能丢失：{number}"})
    return issues
