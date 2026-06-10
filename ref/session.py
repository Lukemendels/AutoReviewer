"""Python twin of the session-binding gate (CheckSessionGate in
modReviewImport.bas).

Why this gate exists: bookmark ids are generic ordinals (AR_PARA_00012), so a
stale JSONL payload left in LLM_Changes!A8 from a previous document can apply
cleanly to the WRONG document. The serializer therefore must begin its output
with a meta line:

    {"meta": "autoreviewer", "session": "<token>", "count": N}

where <token> is the export fingerprint carried verbatim from the hand-off
prompt and N is the number of edit lines that follow. The importer verifies
both before opening Word. Any mismatch is default-deny: no partial apply.

Failure codes (stable across Python and VBA; vectors compare these):
    NO_EXPORT_TOKEN  - no export fingerprint recorded (run export first)
    NO_PAYLOAD       - no non-empty lines pasted
    META_MISSING     - first line is not a valid autoreviewer meta line
    TOKEN_MISMATCH   - meta session != the export fingerprint
    COUNT_MISMATCH   - meta count != number of edit lines that follow
"""

from ref.jsonl_contract import parse_top_level, trim_ws

NO_EXPORT_TOKEN = "NO_EXPORT_TOKEN"
NO_PAYLOAD = "NO_PAYLOAD"
META_MISSING = "META_MISSING"
TOKEN_MISMATCH = "TOKEN_MISMATCH"
COUNT_MISMATCH = "COUNT_MISMATCH"

def is_fence(trimmed: str) -> bool:
    """A code-fence line: any trimmed line beginning with three backticks (so a
    language tag like ```jsonl or ```JSONL is recognized)."""
    return trimmed.startswith("```")


def filter_payload_lines(raw_lines):
    """Collect the payload lines from a paste.

    The serializer emits its meta+edit lines inside a single ```jsonl fenced
    block, with refuse-don't-guess notes (prose) AFTER the closing fence. So:
    if any fence line is present, take only the lines strictly between the
    FIRST fence and the next fence, dropping blanks; everything outside the
    fenced block (including stray prose after it) is ignored. If no fence is
    present, the operator pasted raw JSONL: take all non-blank lines. Both
    paths gate identically."""
    trimmed = [trim_ws(ln) for ln in raw_lines]
    fence_idx = [i for i, t in enumerate(trimmed) if is_fence(t)]
    if fence_idx:
        start = fence_idx[0]
        end = next((i for i in fence_idx if i > start), None)
        block = trimmed[start + 1 : end] if end is not None else trimmed[start + 1 :]
        return [t for t in block if t and not is_fence(t)]
    return [t for t in trimmed if t]


def parse_meta_line(line: str):
    """Returns (session_token, count) or None if the line is not a valid
    autoreviewer meta line."""
    pairs = parse_top_level(line)
    if pairs is None:
        return None
    meta = pairs.get("meta")
    if meta is None or meta[0] != "s" or meta[1] != "autoreviewer":
        return None
    sess = pairs.get("session")
    if sess is None or sess[0] != "s":
        return None
    cnt = pairs.get("count")
    if cnt is None or cnt[0] != "n":
        return None
    # Plain integer only: optional leading minus, then 1-9 ASCII digits.
    # (Matches the VBA exactly; rejects "+1", "3.5", "1e2", and absurd widths.)
    raw = cnt[1]
    body = raw[1:] if raw.startswith("-") else raw
    if not body or len(body) > 9 or any(c not in "0123456789" for c in body):
        return None
    n = int(raw)
    if n < 0:
        return None
    return sess[1], n


def check_session(lines, expected_token: str):
    """lines: already filtered payload lines (see filter_payload_lines),
    meta line included as lines[0] if present. Returns (ok, code)."""
    if expected_token == "":
        return False, NO_EXPORT_TOKEN
    if not lines:
        return False, NO_PAYLOAD
    meta = parse_meta_line(lines[0])
    if meta is None:
        return False, META_MISSING
    session, count = meta
    if session != expected_token:
        return False, TOKEN_MISMATCH
    if count != len(lines) - 1:
        return False, COUNT_MISMATCH
    return True, ""
