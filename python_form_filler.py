#uvicorn python_form_filler:app --reload --host 127.0.0.1 --port 800
from typing import List
from fastapi import FastAPI
from pydantic import BaseModel
from fastapi.responses import PlainTextResponse
from urllib.parse import quote
import os

app = FastAPI()

# -----------------------------
# Models (shared)
# -----------------------------
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

# -----------------------------
# Helpers
# -----------------------------
def enc(s: str) -> str:
    return quote(s or "", safe="")

def read_text_snapshot(path: str, max_chars: int = 20000) -> str:
    if not path or not os.path.exists(path):
        return ""
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read(max_chars)
    except Exception:
        return ""

# -----------------------------
# Decision logic (replace with LLM)
# -----------------------------
def decide_value(field_key: str, local_ctx: str, doc_ctx: str) -> tuple[str, str]:
    """
    Returns (value, reason)
    Replace with LLM.
    """
    k = (field_key or "").strip().upper()

    if k in ("FULL_NAME", "NAME"):
        return "John Doe", "Demo constant for name"
    if k in ("DATE", "TODAY"):
        return "2026-02-04", "Demo constant date in YYYY-MM-DD"
    if k in ("ADDRESS",):
        return "1 Example St, Example City", "Demo constant address"

    return f"AUTO_{k or 'FIELD'}", f"Default mapping for key={k or 'FIELD'}"

def decide_underscore_value(occ: int, local_ctx: str, doc_ctx: str) -> tuple[str, str]:
    return f"FILLED_{occ}", f"Filled underscore blank #{occ} (demo)"

def decide_checkboxes(group: CheckboxGroup, doc_ctx: str) -> tuple[List[int], str]:
    """
    Returns (selected_indices, reason)
    Replace with LLM. For now: choose first option.
    """
    if not group.boxes:
        return [], "No boxes detected"
    # Demo heuristic: if contains "Yes" and "No", pick Yes (index 1)
    return [1], "Selected first option (demo)"

# -----------------------------
# AUDIT endpoint (plain text)
# -----------------------------
@app.post("/audit", response_class=PlainTextResponse)
def audit(req: FillRequest) -> str:
    issues: List[str] = []
    warnings: List[str] = []

    # Snapshot presence if deep_context requested
    if req.deep_context and (not req.temp_text_path or not os.path.exists(req.temp_text_path)):
        warnings.append("deep_context=true but temp_text_path is missing/unreadable")

    # Basic counts
    cc_n = len(req.content_controls)
    ph_n = len(req.placeholders)
    us_n = len(req.underscore_runs)
    cb_n = len(req.checkbox_groups)

    # Sanity checks
    # Underscores should be strictly 1..N
    if us_n > 0:
        occs = [u.occurrence for u in req.underscore_runs]
        if sorted(occs) != list(range(1, us_n + 1)):
            warnings.append("underscore occurrences are not contiguous 1..N (grouping mismatch)")

    # Checkboxes: each group should have at least 2 boxes in typical Yes/No cases
    for g in req.checkbox_groups:
        if len(g.boxes) == 1:
            warnings.append(f"checkbox group {g.occurrence} has only 1 box (may be parse issue)")

    # Decide pass/fail policy:
    # Fail only if payload is clearly broken (empty everything)
    if cc_n == 0 and ph_n == 0 and us_n == 0 and cb_n == 0:
        issues.append("No fields detected (content controls/placeholders/underscores/checkboxes all empty)")

    status = "OK" if not issues else "FAIL"

    summary_lines = []
    summary_lines.append(f"Detected: content_controls={cc_n}, placeholders={ph_n}, underscores={us_n}, checkbox_groups={cb_n}")
    if issues:
        summary_lines.append("Issues:")
        summary_lines.extend([f"- {x}" for x in issues])
    if warnings:
        summary_lines.append("Warnings:")
        summary_lines.extend([f"- {x}" for x in warnings])

    summary = "\n".join(summary_lines)
    return f"AUDIT|{status}|{enc(summary)}\n"

# -----------------------------
# FILL endpoint (plain text with reasons)
# -----------------------------
@app.post("/fill", response_class=PlainTextResponse)
def fill(req: FillRequest) -> str:
    doc_ctx = read_text_snapshot(req.temp_text_path) if req.deep_context else ""
    out: List[str] = []

    # Content controls
    for cc in req.content_controls:
        key = (cc.tag or cc.title or f"CC_{cc.id}").strip()
        value, reason = decide_value(key, cc.context, doc_ctx)
        out.append(f"CC|{cc.id}|{enc(value)}|{enc(reason)}")

    # Placeholders
    for p in req.placeholders:
        value, reason = decide_value(p.key, p.context, doc_ctx)
        out.append(f"PH|{enc(p.token)}|{enc(value)}|{enc(reason)}")

    # Underscores (per occurrence)
    for u in req.underscore_runs:
        value, reason = decide_underscore_value(u.occurrence, u.context, doc_ctx)
        out.append(f"US|{u.occurrence}|{enc(value)}|{enc(reason)}")

    # Checkboxes
    for g in req.checkbox_groups:
        selected, reason = decide_checkboxes(g, doc_ctx)
        selected_csv = ",".join(str(i) for i in selected)
        out.append(f"CB|{g.occurrence}|{enc(selected_csv)}|{enc(reason)}")

    return "\n".join(out) + "\n"

# -----------------------------
# REVISE endpoint (plain text)
# -----------------------------
@app.post("/revise", response_class=PlainTextResponse)
def revise(req: ReviseRequest) -> str:
    instr = (req.instruction or "").strip()
    txt = req.selected_text or ""

    # Replace with LLM call. Demo:
    revised = txt
    reason = f"Applied instruction: {instr}"

    il = instr.lower()
    if "uppercase" in il:
        revised = txt.upper()
    elif "lowercase" in il:
        revised = txt.lower()
    elif "trim" in il:
        revised = " ".join(txt.split())

    return f"REV|{enc(revised)}|{enc(reason)}\n"
