# uvicorn python_form_filler:app --reload --host 127.0.0.1 --port 8000
#
# Environment variables (set before running):
#   LLM_API_KEY   – key for an OpenAI-compatible service  (empty = demo mode)
#   LLM_BASE_URL  – base URL        (default: https://api.openai.com/v1)
#   LLM_MODEL     – model to use    (default: gpt-4o)

import os
import re
import json
import logging
from typing import List
from fastapi import FastAPI
from pydantic import BaseModel
from fastapi.responses import PlainTextResponse
from urllib.parse import quote
from urllib.request import Request, urlopen

# ── Logging ─────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("forminator")

# ── Config ──────────────────────────────────────────────
LLM_API_KEY = os.environ.get("LLM_API_KEY", "")
LLM_BASE_URL = os.environ.get("LLM_BASE_URL", "https://api.openai.com/v1")
LLM_MODEL = os.environ.get("LLM_MODEL", "gpt-4o")

app = FastAPI(title="Forminator")


@app.get("/")
def root():
    return {
        "status": "running",
        "llm_configured": bool(LLM_API_KEY),
        "llm_model": LLM_MODEL if LLM_API_KEY else "demo-mode",
    }


# ── Models ──────────────────────────────────────────────
class ContentControlInfo(BaseModel):
    id: int
    tag: str = ""
    title: str = ""
    cc_type: str = ""
    context: str = ""


class PlaceholderInfo(BaseModel):
    token: str
    key: str
    occurrence: int
    context: str = ""


class UnderscoreInfo(BaseModel):
    occurrence: int
    context: str = ""


class CheckboxBox(BaseModel):
    index: int
    offset: int
    label: str = ""


class CheckboxGroup(BaseModel):
    occurrence: int
    text: str
    boxes: List[CheckboxBox] = []
    context: str = ""


class FillRequest(BaseModel):
    doc_name: str = ""
    doc_folder: str = ""
    temp_text_path: str = ""
    deep_context: bool = False
    content_controls: List[ContentControlInfo] = []
    placeholders: List[PlaceholderInfo] = []
    underscore_runs: List[UnderscoreInfo] = []
    checkbox_groups: List[CheckboxGroup] = []


class ReviseRequest(BaseModel):
    instruction: str
    selected_text: str
    doc_name: str = ""
    doc_folder: str = ""


# ── Helpers ─────────────────────────────────────────────
def enc(s: str) -> str:
    return quote(s or "", safe="")


# Characters invisible in Word documents but present in the extracted text:
# control chars, BOM, soft-hyphens, zero-width joiners, object-replacement, etc.
_INVISIBLE_RE = re.compile(
    r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f"        # C0 control chars (keep \t \n \r)
    r"\xad"                                       # soft-hyphen
    r"\u200b-\u200f"                              # zero-width / direction marks
    r"\u2028\u2029"                               # line/paragraph separators
    r"\u202a-\u202e"                              # bidi embedding marks
    r"\u2060-\u2064"                              # word-joiner, invisible chars
    r"\ufeff"                                     # BOM / zero-width no-break space
    r"\ufffc\ufffd"                               # object replacement / replacement char
    r"\u0001-\u0003"                              # Word field markers
    r"]"
)


def clean_context(text: str, max_chars: int = 0) -> str:
    """Sanitise a context string before it reaches the LLM.

    1. Strip characters that are invisible in the Word document but leak
       into the extracted range (field markers, BOM, direction marks …).
    2. Collapse runs of whitespace into single spaces.
    3. If *max_chars* > 0, trim to that length but snap outward to the
       nearest word boundary so the LLM never sees a clipped word.
    """
    # Remove invisible characters
    t = _INVISIBLE_RE.sub("", text)
    # Normalise all whitespace (tabs, stray newlines inside context) to spaces
    t = re.sub(r"\s+", " ", t).strip()

    if max_chars > 0 and len(t) > max_chars:
        # Trim from each end equally, snapping to word boundaries
        half = max_chars // 2

        # ── left side: keep from start, cut at last space before half ──
        left = t[:half]
        sp = left.rfind(" ")
        if sp > 0:
            left = left[:sp]

        # ── right side: keep to end, cut at first space after len-half ──
        right = t[len(t) - half:]
        sp = right.find(" ")
        if sp != -1 and sp < len(right) - 1:
            right = right[sp + 1:]

        t = left + " … " + right

    return t


# ── LLM caller ─────────────────────────────────────────
def llm_chat(system_prompt: str, user_prompt: str, temperature: float = 0.2) -> str | None:
    """Call an OpenAI-compatible chat/completions endpoint.
    Returns the assistant message content, or None on any failure."""
    if not LLM_API_KEY:
        return None
    try:
        url = f"{LLM_BASE_URL.rstrip('/')}/chat/completions"
        body = json.dumps({
            "model": LLM_MODEL,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "temperature": temperature,
        }).encode()
        req = Request(url, data=body, headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {LLM_API_KEY}",
        })
        with urlopen(req, timeout=120) as resp:
            data = json.loads(resp.read())
        return data["choices"][0]["message"]["content"]
    except Exception as e:
        log.warning("LLM call failed: %s", e)
        return None


def strip_code_fences(text: str) -> str:
    """Remove ```json ... ``` wrappers that LLMs sometimes add."""
    t = text.strip()
    if t.startswith("```"):
        first_nl = t.find("\n")
        t = t[first_nl + 1:] if first_nl != -1 else t[3:]
        if t.endswith("```"):
            t = t[:-3]
    return t.strip()


# ── Context readers ─────────────────────────────────────
def read_text_snapshot(path: str, max_chars: int = 50_000) -> str:
    if not path or not os.path.exists(path):
        return ""
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read(max_chars)
    except Exception:
        return ""


def read_context_files(doc_folder: str, max_chars: int = 30_000) -> str:
    """Read .txt files from the document's folder as reference context."""
    if not doc_folder or not os.path.isdir(doc_folder):
        return ""
    parts: list[str] = []
    total = 0
    for name in sorted(os.listdir(doc_folder)):
        if not name.lower().endswith(".txt"):
            continue
        fpath = os.path.join(doc_folder, name)
        if not os.path.isfile(fpath):
            continue
        try:
            with open(fpath, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read(max_chars - total)
            if content.strip():
                parts.append(f"--- {name} ---\n{content}")
                total += len(content)
                if total >= max_chars:
                    break
        except Exception:
            continue
    return "\n\n".join(parts)


# ── Fill: system prompt ─────────────────────────────────
FILL_SYSTEM = """\
You are a smart form-filling assistant. You receive a document, optional \
reference files, and a list of form fields. Fill every field with the most \
appropriate value drawn from the available context.

Respond with ONLY valid JSON — no markdown fences, no extra text:
{
  "content_controls": [{"id": <int>, "value": "<text>", "reason": "<brief>"}],
  "placeholders":     [{"token": "<original token>", "value": "<text>", "reason": "<brief>"}],
  "underscores":      [{"occurrence": <int>, "value": "<text>", "reason": "<brief>"}],
  "checkboxes":       [{"occurrence": <int>, "selected": [<1-based box indices>], "reason": "<brief>"}]
}

Rules:
- Use information from reference files when available.
- Match the expected format (dates, names, etc.) implied by surrounding text.
- For checkboxes, return 1-based indices of the boxes to mark.
- Keep each reason under 30 words.
- Include an entry for EVERY field, even if you must make a reasonable guess.
- Omit a section key entirely if there are no fields of that type.\
"""


def _build_fill_user(req: FillRequest, doc_text: str, ctx_files: str) -> str:
    sections: list[str] = []

    if doc_text:
        sections.append(f"## Document Text\n{doc_text[:15_000]}")
    if ctx_files:
        sections.append(f"## Reference Files\n{ctx_files}")

    if req.content_controls:
        lines = ["## Content Controls"]
        for cc in req.content_controls:
            lines.append(
                f'- ID={cc.id}  tag="{cc.tag}"  title="{cc.title}"  type={cc.cc_type}'
            )
            if cc.context:
                lines.append(f"  Context: {clean_context(cc.context, 300)}")
        sections.append("\n".join(lines))

    if req.placeholders:
        lines = ["## Placeholders"]
        for p in req.placeholders:
            lines.append(f'- Token="{p.token}"  key="{p.key}"')
            if p.context:
                lines.append(f"  Context: {clean_context(p.context, 300)}")
        sections.append("\n".join(lines))

    if req.underscore_runs:
        lines = ["## Underscore Blanks"]
        for u in req.underscore_runs:
            lines.append(f"- Occurrence={u.occurrence}")
            if u.context:
                lines.append(f"  Context: {clean_context(u.context, 300)}")
        sections.append("\n".join(lines))

    if req.checkbox_groups:
        lines = ["## Checkbox Groups"]
        for g in req.checkbox_groups:
            box_labels = ", ".join(
                f'Box{b.index}="{b.label}"' for b in g.boxes
            )
            lines.append(
                f'- Occ={g.occurrence}  text="{clean_context(g.text, 200)}"  boxes=[{box_labels}]'
            )
            if g.context:
                lines.append(f"  Context: {clean_context(g.context, 300)}")
        sections.append("\n".join(lines))

    return "\n\n".join(sections)


def _parse_fill_llm(raw: str) -> str | None:
    """LLM JSON -> pipe-delimited lines.  Returns None on failure."""
    try:
        data = json.loads(strip_code_fences(raw))
        out: list[str] = []
        for cc in data.get("content_controls", []):
            out.append(
                f"CC|{cc['id']}|{enc(str(cc['value']))}|{enc(str(cc.get('reason', '')))}"
            )
        for ph in data.get("placeholders", []):
            out.append(
                f"PH|{enc(str(ph['token']))}|{enc(str(ph['value']))}|"
                f"{enc(str(ph.get('reason', '')))}"
            )
        for us in data.get("underscores", []):
            out.append(
                f"US|{us['occurrence']}|{enc(str(us['value']))}|"
                f"{enc(str(us.get('reason', '')))}"
            )
        for cb in data.get("checkboxes", []):
            sel = ",".join(str(i) for i in cb.get("selected", []))
            out.append(
                f"CB|{cb['occurrence']}|{enc(sel)}|{enc(str(cb.get('reason', '')))}"
            )
        return "\n".join(out) + "\n" if out else None
    except Exception as e:
        log.warning("Fill LLM parse error: %s", e)
        return None


# ── Fill: demo fallback ─────────────────────────────────
def _demo_value(key: str) -> tuple[str, str]:
    k = key.upper().strip()
    if k in ("FULL_NAME", "NAME"):
        return "John Doe", "Demo constant"
    if k in ("DATE", "TODAY"):
        return "2026-02-06", "Demo date"
    if k in ("ADDRESS",):
        return "1 Example St, Example City", "Demo address"
    return f"[{k}]", f"Demo placeholder for {k}"


def _fill_demo(req: FillRequest) -> str:
    out: list[str] = []
    for cc in req.content_controls:
        key = (cc.tag or cc.title or f"CC_{cc.id}").strip()
        val, reason = _demo_value(key)
        out.append(f"CC|{cc.id}|{enc(val)}|{enc(reason)}")
    for p in req.placeholders:
        val, reason = _demo_value(p.key)
        out.append(f"PH|{enc(p.token)}|{enc(val)}|{enc(reason)}")
    for u in req.underscore_runs:
        out.append(
            f"US|{u.occurrence}|{enc(f'FILLED_{u.occurrence}')}|{enc('Demo fill')}"
        )
    for g in req.checkbox_groups:
        sel = "1" if g.boxes else ""
        out.append(
            f"CB|{g.occurrence}|{enc(sel)}|{enc('Demo: selected first option')}"
        )
    return "\n".join(out) + "\n"


# ── Revise: system prompt ───────────────────────────────
REVISE_SYSTEM = """\
You are a precise text editor. Apply the user's instruction to revise the \
provided text.

Respond with ONLY valid JSON — no markdown fences, no extra text:
{"revised": "<the full revised text>", "reason": "<brief explanation>"}\
"""


def _build_revise_user(
    instruction: str, selected_text: str, ctx_files: str = ""
) -> str:
    parts = [f"Instruction: {instruction}\n\nOriginal text:\n{clean_context(selected_text)}"]
    if ctx_files:
        parts.append(f"\nReference files (for context):\n{ctx_files[:5000]}")
    return "\n".join(parts)


def _parse_revise_llm(raw: str) -> tuple[str, str] | None:
    try:
        data = json.loads(strip_code_fences(raw))
        return data["revised"], data.get("reason", "Applied instruction")
    except Exception as e:
        log.warning("Revise LLM parse error: %s", e)
        return None


def _revise_demo(instruction: str, text: str) -> tuple[str, str]:
    il = instruction.lower()
    if "uppercase" in il:
        return text.upper(), "Converted to uppercase"
    if "lowercase" in il:
        return text.lower(), "Converted to lowercase"
    if "trim" in il or "clean" in il:
        return " ".join(text.split()), "Trimmed whitespace"
    return text, f"Demo mode — no LLM configured (instruction: {instruction})"


# ═══════════════════════════════════════════════════════
# ENDPOINTS
# ═══════════════════════════════════════════════════════

@app.post("/audit", response_class=PlainTextResponse)
def audit(req: FillRequest) -> str:
    issues: list[str] = []
    warnings: list[str] = []

    if req.deep_context and not read_text_snapshot(req.temp_text_path):
        warnings.append("deep_context=true but temp_text_path missing/unreadable")

    cc_n = len(req.content_controls)
    ph_n = len(req.placeholders)
    us_n = len(req.underscore_runs)
    cb_n = len(req.checkbox_groups)

    if us_n > 0:
        occs = sorted(u.occurrence for u in req.underscore_runs)
        if occs != list(range(1, us_n + 1)):
            warnings.append("underscore occurrences not contiguous 1..N")

    for g in req.checkbox_groups:
        if len(g.boxes) == 1:
            warnings.append(f"checkbox group {g.occurrence} has only 1 box")

    if cc_n == 0 and ph_n == 0 and us_n == 0 and cb_n == 0:
        issues.append("No fields detected")

    status = "OK" if not issues else "FAIL"

    lines = [f"Detected: CC={cc_n} PH={ph_n} US={us_n} CB={cb_n}"]
    if LLM_API_KEY:
        lines.append(f"Mode: LLM ({LLM_MODEL})")
    else:
        lines.append("Mode: DEMO (set LLM_API_KEY for AI filling)")
    if req.doc_folder and os.path.isdir(req.doc_folder):
        ctx = [n for n in sorted(os.listdir(req.doc_folder)) if n.lower().endswith(".txt")]
        if ctx:
            lines.append(f"Context files: {', '.join(ctx)}")
    for x in issues:
        lines.append(f"ISSUE: {x}")
    for x in warnings:
        lines.append(f"WARNING: {x}")

    return f"AUDIT|{status}|{enc(chr(10).join(lines))}\n"


@app.post("/fill", response_class=PlainTextResponse)
def fill(req: FillRequest) -> str:
    doc_text = read_text_snapshot(req.temp_text_path) if req.deep_context else ""
    ctx_files = read_context_files(req.doc_folder)

    if LLM_API_KEY:
        user_prompt = _build_fill_user(req, doc_text, ctx_files)
        raw = llm_chat(FILL_SYSTEM, user_prompt)
        if raw:
            result = _parse_fill_llm(raw)
            if result:
                log.info("Fill completed via LLM (%d lines)", result.count("\n"))
                return result
        log.warning("Fill: LLM failed or returned bad JSON, falling back to demo")

    return _fill_demo(req)


@app.post("/revise", response_class=PlainTextResponse)
def revise(req: ReviseRequest) -> str:
    instr = (req.instruction or "").strip()
    txt = req.selected_text or ""

    if LLM_API_KEY:
        ctx_files = read_context_files(req.doc_folder)
        raw = llm_chat(REVISE_SYSTEM, _build_revise_user(instr, txt, ctx_files))
        if raw:
            parsed = _parse_revise_llm(raw)
            if parsed:
                revised, reason = parsed
                return f"REV|{enc(revised)}|{enc(reason)}\n"
        log.warning("Revise: LLM failed, falling back to demo")

    revised, reason = _revise_demo(instr, txt)
    return f"REV|{enc(revised)}|{enc(reason)}\n"
