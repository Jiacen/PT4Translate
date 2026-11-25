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
                # Short all-caps terms are acronyms.  Treating TO as case
                # insensitive would incorrectly match the English preposition
                # "to" in ordinary sentences.
                acronym = variant.isupper() and len(variant) <= 3
                flags = 0 if term.case_sensitive or acronym else re.IGNORECASE
                pattern = re.compile(r"(?<![A-Za-z0-9])" + re.escape(variant) + r"(?![A-Za-z0-9])", flags)
                if pattern.search(protected):
                    # Plain ASCII is far more robust across translation APIs
                    # than private-use Unicode markers. Some providers turn
                    # ``TERM`` into “术语” or discard private-use glyphs,
                    # which made a detected term impossible to restore.
                    token = f"[[SITRAIN_TERM_{len(mapping):04d}]]"
                    protected = pattern.sub(token, protected)
                    mapping[token] = (term, variant)
                    break
        return protected, mapping

    @staticmethod
    def restore(translated: str, mapping: dict[str, tuple[Term, str]]) -> str:
        for token, (term, source_variant) in mapping.items():
            replacement = source_variant if term.mode in {"keep", "forbidden"} else term.target
            translated = translated.replace(token, replacement)
            token_no = re.search(r"_(\d{4})\]\]$", token)
            if token_no:
                number = token_no.group(1)
                # Accept legacy private-use markers as well as the Chinese
                # “术语” spelling emitted by earlier Qwen retries.
                translated = re.sub(r"\ue000(?:TERM|变量|术语)_" + number + r"\ue001", replacement, translated)
                translated = re.sub(r"(?:\[\[|【)?SITRAIN_(?:TERM|术语)_" + number + r"(?:\]\]|】)?", replacement, translated)
        return translated

    def validate(self, source: str, translated: str) -> list[dict]:
        issues = []
        dst_norm = normalize(translated)
        matches: list[tuple[int, int, Term]] = []
        for term in self.terms:
            for variant in (term.source, *term.aliases):
                acronym = variant.isupper() and len(variant) <= 3
                flags = 0 if term.case_sensitive or acronym else re.IGNORECASE
                pattern = re.compile(r"(?<![A-Za-z0-9])" + re.escape(variant) + r"(?![A-Za-z0-9])", flags)
                matches.extend((match.start(), match.end(), term) for match in pattern.finditer(source))
                if pattern.search(source):
                    break
        # A specific glossary entry must take precedence over a shorter entry
        # nested inside it. For example, if "Comfort Panel → 精智面板" is
        # configured, the general "Comfort → Comfort" rule must not demand a
        # literal "Comfort" in that same span.
        selected: list[tuple[int, int, Term]] = []
        for start, end, term in sorted(matches, key=lambda item: (-(item[1] - item[0]), item[0])):
            if any(start >= chosen_start and end <= chosen_end
                   for chosen_start, chosen_end, _chosen_term in selected):
                continue
            selected.append((start, end, term))
        for _start, _end, term in selected:
            required = term.source if term.mode in {"keep", "forbidden"} else term.target
            if normalize(required) not in dst_norm:
                issues.append({"code":"TERM_MISSING","severity":"blocker" if term.mode != "recommended" else "warning",
                               "term_id":term.id,"message":f"必须使用术语：{term.source} → {required}"})
        return issues


_LATIN_WORD = re.compile(r"[A-Za-z][A-Za-z0-9_-]{2,}")
_SINGLE_DIGIT_CN = {
    "0": ("零", "〇"), "1": ("一", "壹"), "2": ("二", "两", "贰"),
    "3": ("三", "叁"), "4": ("四", "肆"), "5": ("五", "伍"),
    "6": ("六", "陆"), "7": ("七", "柒"), "8": ("八", "捌"), "9": ("九", "玖"),
}
_PHASE_CN = {"1": ("单相",), "2": ("双相", "两相")}


def residual_source_issues(source: str, translated: str) -> list[dict]:
    issues = []
    source_words = {w.casefold() for w in _LATIN_WORD.findall(source)}
    target_words = {w.casefold() for w in _LATIN_WORD.findall(translated)}
    residual = sorted(source_words & target_words)
    if residual and len(residual) >= max(2, len(source_words) // 2):
        issues.append({"code":"SOURCE_RESIDUAL","severity":"warning","message":"译文可能残留较多原文","tokens":residual[:12]})
    for number in re.findall(r"\b\d+(?:[.,]\d+)?\b", source):
        if number in _PHASE_CN and re.search(r"(?<![\w.])" + number + r"\s*-?\s*(?:phasig|phases?)\b", normalize(source)):
            if any(name in translated for name in _PHASE_CN[number]):
                continue
        # Natural Chinese expressions such as “第三级”“两相” retain the
        # meaning of a single digit.  Keep strict checks for multi-digit values,
        # decimals, units and model identifiers, where changing the notation is
        # materially riskier.
        if len(number) == 1 and any(char in translated for char in _SINGLE_DIGIT_CN[number]):
            continue
        if number not in translated:
            issues.append({"code":"NUMBER_CHANGED","severity":"blocker","message":f"数字可能丢失：{number}"})
    return issues
