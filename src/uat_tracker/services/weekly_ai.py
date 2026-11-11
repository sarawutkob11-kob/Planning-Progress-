"""Weekly report - the AI part (Part 6, approved 10-Nov-2026). Claude writes ONLY the executive summary,
top risks and follow-ups, in Thai and English, from the data the tracker computed.

- API key: Windows Credential Manager (per Windows user), never in config.toml / the DB / logs / Excel.
- Model: Settings (default claude-opus-5-5, effort medium). Structured output (JSON schema). Server-side
  refusal fallback "default" on the models that support it.
- Timeout 60 s, 2 retries (the SDK's own retry of 408/409/429/5xx and connection errors).
- Every number in the AI text must appear in the input data; otherwise the AI part is not used.
- Tokens + cost are returned for the report log.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field

DEFAULT_MODEL = "claude-opus-5-5"
S_MODEL = "settings.ai_model"
KEY_TARGET = "PlanningProgress/anthropic"
LEGACY_KEY_TARGET = "UATTracker/anthropic"        # before the rename; moved on first read
TIMEOUT_S = 60.0
MAX_RETRIES = 2
FALLBACK_MODELS = {"claude-opus-5-5", "claude-opus-5", "claude-fable-5-1", "claude-sonnet-5-5"}
PRICES = {                       # USD per 1M tokens (input, output), Anthropic first-party rates
    "claude-fable-5-1": (10.0, 50.0), "claude-fable-5": (10.0, 50.0),
    "claude-opus-5-5": (4.0, 20.0), "claude-opus-5": (5.0, 25.0), "claude-opus-4-8": (5.0, 25.0),
    "claude-sonnet-5-5": (2.0, 10.0), "claude-sonnet-5": (2.0, 10.0),
    "claude-haiku-5-5": (0.10, 0.50),
}

SYSTEM = (
    "You write the executive summary of a weekly UAT progress report for a bank IT project (the project name is in the data).\n"
    "Use ONLY the JSON data in the user message. Do not calculate, estimate or round anything: every number "
    "you write must appear exactly as given in the data (copy \"91.5%\" as \"91.5%\", dates as written).\n"
    "Do not mention people or email addresses. If the data shows no risk in an area, do not invent one.\n"
    "Write each field twice: Thai (_th) and English (_en). Plain sentences, no tables, no markdown.\n"
    "In `related`, use only application, task, defect or item names exactly as they appear in the data."
)

_ITEM = lambda a, b: {
    "type": "object", "additionalProperties": False, "required": [f"{a}_th", f"{a}_en", f"{b}_th", f"{b}_en", "related"],
    "properties": {f"{a}_th": {"type": "string"}, f"{a}_en": {"type": "string"},
                   f"{b}_th": {"type": "string"}, f"{b}_en": {"type": "string"},
                   "related": {"type": "array", "items": {"type": "string"}}}}
SCHEMA = {
    "type": "object", "additionalProperties": False,
    "required": ["summary_th", "summary_en", "risks", "follow_ups"],
    "properties": {
        "summary_th": {"type": "string"}, "summary_en": {"type": "string"},
        "risks": {"type": "array", "items": _ITEM("title", "why")},
        "follow_ups": {"type": "array", "items": _ITEM("action", "why")},
    },
}


# --------------------------------------------------------------------------- key store

class WinCredKeyStore:
    def get(self) -> str | None:
        import pywintypes
        import win32cred
        for target in (KEY_TARGET, LEGACY_KEY_TARGET):
            try:
                c = win32cred.CredRead(target, win32cred.CRED_TYPE_GENERIC)
            except pywintypes.error:
                continue
            blob = c["CredentialBlob"]
            key = (blob.decode("utf-16-le") if isinstance(blob, bytes) else str(blob)) or None
            if key and target == LEGACY_KEY_TARGET:      # rename: keep the key under the new name
                self.set(key)
                try:
                    win32cred.CredDelete(LEGACY_KEY_TARGET, win32cred.CRED_TYPE_GENERIC)
                except pywintypes.error:
                    pass
            return key
        return None

    def set(self, key: str) -> None:
        import win32cred
        win32cred.CredWrite({"Type": win32cred.CRED_TYPE_GENERIC, "TargetName": KEY_TARGET,
                             "UserName": "anthropic-api-key", "CredentialBlob": key,
                             "Comment": "Planning Progress - Claude API key (weekly report)",
                             "Persist": win32cred.CRED_PERSIST_LOCAL_MACHINE}, 0)

    def clear(self) -> None:
        import pywintypes
        import win32cred
        try:
            win32cred.CredDelete(KEY_TARGET, win32cred.CRED_TYPE_GENERIC)
        except pywintypes.error:
            pass


KEYS = WinCredKeyStore()


def model_of(repo) -> str:
    return (repo.get_setting(S_MODEL) or DEFAULT_MODEL).strip() or DEFAULT_MODEL


# --------------------------------------------------------------------------- number check

_NUM = re.compile(r"\d+(?:[.,]\d+)*")


def _norm(tok: str) -> str:
    tok = tok.replace(",", "")
    if "." in tok:
        tok = tok.rstrip("0").rstrip(".") or "0"
    return tok.lstrip("0") or "0"


def numbers_in(text: str) -> set[str]:
    return {_norm(m) for m in _NUM.findall(text or "")}


def ai_texts(ai: dict) -> list[str]:
    out = [ai.get("summary_th", ""), ai.get("summary_en", "")]
    for it in ai.get("risks", []) + ai.get("follow_ups", []):
        out += [v for k, v in it.items() if k != "related" and isinstance(v, str)]
        out += [r for r in it.get("related", []) if isinstance(r, str)]
    return out


def unknown_numbers(ai: dict, data: dict) -> list[str]:
    """Numbers the AI wrote that are not in the input data (acceptance 5)."""
    allowed = numbers_in(json.dumps(data, ensure_ascii=False))
    seen = set()
    for t in ai_texts(ai):
        seen |= numbers_in(t)
    return sorted(seen - allowed, key=lambda s: (len(s), s))


def keep_known_related(ai: dict, data: dict) -> dict:
    """`related` must name things that are in the data; anything else is dropped (not an error)."""
    blob = json.dumps(data, ensure_ascii=False)
    for it in ai.get("risks", []) + ai.get("follow_ups", []):
        it["related"] = [r for r in it.get("related", []) if isinstance(r, str) and r and r in blob]
    return ai


# --------------------------------------------------------------------------- call

@dataclass
class AiResult:
    ok: bool
    content: dict | None = None
    error: str = ""
    model: str = ""
    input_tokens: int = 0
    output_tokens: int = 0
    cost_usd: float | None = None
    raw_stop: str = ""
    extra: dict = field(default_factory=dict)


def cost_of(model: str, input_tokens: int, output_tokens: int) -> float | None:
    p = PRICES.get(model)
    if p is None:
        return None
    return round(input_tokens * p[0] / 1_000_000 + output_tokens * p[1] / 1_000_000, 6)


def _scrub(msg: str, key: str | None) -> str:
    return msg.replace(key, "***") if key else msg


def summarize(data: dict, *, model: str = DEFAULT_MODEL, key: str | None = None, client=None) -> AiResult:
    """Ask Claude for the summary. Never raises: problems come back as AiResult(ok=False, error=...)."""
    import anthropic
    key = key if key is not None else KEYS.get()
    if not key and client is None:
        return AiResult(False, error="No Claude API key set (Settings > Claude API key)", model=model)
    try:
        client = client or anthropic.Anthropic(api_key=key, timeout=TIMEOUT_S, max_retries=MAX_RETRIES)
        user = ("Report data (computed by the tracker; treat as facts):\n"
                + json.dumps(data, ensure_ascii=False, sort_keys=True) +
                "\n\nReturn: summary_th / summary_en (3-5 sentences on overall status), risks (up to 5: "
                "title, why), follow_ups (up to 5: action, why). Each with Thai and English text.")
        kwargs = dict(model=model, max_tokens=16000, system=SYSTEM,
                      messages=[{"role": "user", "content": user}],
                      output_config={"effort": "medium", "format": {"type": "json_schema", "schema": SCHEMA}})
        if model in FALLBACK_MODELS:
            kwargs.update(betas=["server-side-fallback-2026-07-01"], fallbacks="default")
        resp = client.beta.messages.create(**kwargs)
    except anthropic.AuthenticationError:
        return AiResult(False, error="Claude API key was refused (check it in Settings)", model=model)
    except anthropic.APITimeoutError:
        return AiResult(False, error=f"Claude API timed out ({TIMEOUT_S:.0f} s, {MAX_RETRIES} retries)", model=model)
    except anthropic.APIConnectionError as exc:
        return AiResult(False, error=_scrub(f"Cannot reach the Claude API: {exc}", key), model=model)
    except anthropic.APIStatusError as exc:
        return AiResult(False, error=_scrub(f"Claude API error {exc.status_code}: {exc.message}", key), model=model)
    except Exception as exc:                                    # never let the report fail because of the AI
        return AiResult(False, error=_scrub(f"Claude API call failed: {exc}", key), model=model)

    used = getattr(resp, "model", None) or model
    usage = getattr(resp, "usage", None)
    tin = int(getattr(usage, "input_tokens", 0) or 0) + int(getattr(usage, "cache_read_input_tokens", 0) or 0) \
        + int(getattr(usage, "cache_creation_input_tokens", 0) or 0)
    tout = int(getattr(usage, "output_tokens", 0) or 0)
    res = AiResult(False, model=used, input_tokens=tin, output_tokens=tout, cost_usd=cost_of(used, tin, tout),
                   raw_stop=str(getattr(resp, "stop_reason", "")))
    if res.raw_stop == "refusal":
        res.error = "Claude declined the request (refusal)"
        return res
    if res.raw_stop == "max_tokens":
        res.error = "Claude's answer was cut off (max_tokens)"
        return res
    text = next((b.text for b in resp.content if getattr(b, "type", "") == "text"), "")
    try:
        content = json.loads(text)
    except ValueError:
        res.error = "Claude did not return valid JSON"
        return res
    bad = unknown_numbers(content, data)
    if bad:
        res.error = "AI text contains numbers that are not in the report data: " + ", ".join(bad[:10])
        return res
    res.ok, res.content = True, keep_known_related(content, data)
    return res
