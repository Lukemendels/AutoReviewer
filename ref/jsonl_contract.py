"""Python twin of the AutoReviewer JSONL line contract (modReviewImport.bas).

This is the debuggable source of truth for ParseJsonLine + the per-change_type
validation. The VBA is a transliteration of THIS algorithm; the golden vectors
prove they agree.

Two historical VBA bugs are fixed here (and in the VBA):
1. Escape parity: the old closing-quote scan used `prevCh <> "\\"`, which
   misreads an escaped backslash before a quote (`"a\\\\"` never closed). The
   string reader below consumes backslash+next as a unit, so a quote closes
   iff it is preceded by an EVEN run of backslashes.
2. Key-in-value collision: the old InStr-over-whole-line key lookup could find
   a key name inside another field's value. This parser is a single
   left-to-right tokenizer that walks string literals and recognizes keys only
   at object top level.

Deliberate semantics, mirrored exactly in VBA (do not "improve" one side):
- trim_ws() strips space/tab/CR/LF from both ends (VBA Trim$ strips spaces
  only, so the VBA defines its own TrimWs to match).
- Duplicate keys: FIRST occurrence wins (matches the old extractor's
  find-first behavior).
- apply_change is honored only as a bare true/false literal; a quoted "true"
  is a string and is ignored (apply_change stays unset).
- JSON unescape supports the escapes backslash, quote, slash, b, f, n, r, t.
  Any OTHER escape drops the backslash and keeps the character literally --
  including \\uXXXX, which is NOT decoded (current VBA behavior; on the
  Found-not-fixed list).
- bookmark_id and change_type must be present AS STRINGS for a parse to
  succeed (the old extractor failed the line otherwise).
"""

WS = " \t\r\n"

CHANGE_TYPES = (
    "replace_text",
    "delete_element",
    "add_comment_only",
    "reply_to_comment",
    "accept_revision",
    "reject_revision",
)

# Validation reason codes (stable across Python and VBA; vectors compare these)
OK = ""
MISSING_BOOKMARK = "MISSING_BOOKMARK"
MISSING_CHANGE_TYPE = "MISSING_CHANGE_TYPE"
UNKNOWN_CHANGE_TYPE = "UNKNOWN_CHANGE_TYPE"
REPLACE_REQUIRES_NEW_TEXT = "REPLACE_REQUIRES_NEW_TEXT"
COMMENT_REQUIRES_TEXT = "COMMENT_REQUIRES_TEXT"
REPLY_REQUIRES_COMMENT_TARGET = "REPLY_REQUIRES_COMMENT_TARGET"
REPLY_REQUIRES_TEXT = "REPLY_REQUIRES_TEXT"
REVISION_REQUIRES_RANGE_TARGET = "REVISION_REQUIRES_RANGE_TARGET"
PARSE = "PARSE"  # structural failure / missing required string keys


def trim_ws(s: str) -> str:
    return s.strip(WS)


def json_unescape(s: str) -> str:
    """Mirror of JsonUnescapeString in modReviewImport.bas (incl. its gaps)."""
    out = []
    i = 0
    n = len(s)
    while i < n:
        ch = s[i]
        if ch == "\\" and i + 1 < n:
            esc = s[i + 1]
            if esc == "\\":
                out.append("\\")
            elif esc == '"':
                out.append('"')
            elif esc == "/":
                out.append("/")
            elif esc == "b":
                out.append("\b")
            elif esc == "f":
                out.append("\f")
            elif esc == "n":
                out.append("\n")
            elif esc == "r":
                out.append("\r")
            elif esc == "t":
                out.append("\t")
            else:
                # Unknown escape: keep the escaped char, drop the backslash.
                out.append(esc)
            i += 2
        else:
            out.append(ch)
            i += 1
    return "".join(out)


def _read_string_raw(s: str, i: int):
    """s[i] is the opening quote. Return (raw_contents, index_after_close)
    or (None, i) if unterminated. Backslash consumes the next char, so a
    closing quote is recognized iff preceded by an even backslash run."""
    n = len(s)
    i += 1
    out = []
    while i < n:
        ch = s[i]
        if ch == "\\":
            if i + 1 >= n:
                return None, i  # dangling backslash at end: unterminated
            out.append(ch)
            out.append(s[i + 1])
            i += 2
        elif ch == '"':
            return "".join(out), i + 1
        else:
            out.append(ch)
            i += 1
    return None, i  # no closing quote


def _skip_ws(s: str, i: int) -> int:
    n = len(s)
    while i < n and s[i] in WS:
        i += 1
    return i


def _read_nested_raw(s: str, i: int):
    """s[i] is '{' or '['. Walk to the matching close (string-aware).
    Return (raw_including_brackets, index_after) or (None, i)."""
    n = len(s)
    start = i
    depth = 0
    while i < n:
        ch = s[i]
        if ch == '"':
            raw, j = _read_string_raw(s, i)
            if raw is None:
                return None, i
            i = j
            continue
        if ch in "{[":
            depth += 1
        elif ch in "}]":
            depth -= 1
            if depth == 0:
                return s[start : i + 1], i + 1
        i += 1
    return None, i


NUM_CHARS = set("-+.eE0123456789")


def parse_top_level(line: str):
    """Tokenize one JSON object line. Returns a dict key -> (type, value)
    with type in {'s','b','n','z','c'} (string/bool/number/null/complex),
    string values unescaped, others raw -- or None on structural failure.
    First occurrence of a key wins."""
    s = trim_ws(line)
    n = len(s)
    if n < 2 or s[0] != "{" or s[-1] != "}":
        return None
    pairs = {}
    i = 1
    i = _skip_ws(s, i)
    if i < n and s[i] == "}":
        return pairs if i == n - 1 else None  # empty object
    while True:
        i = _skip_ws(s, i)
        if i >= n or s[i] != '"':
            return None
        key_raw, i = _read_string_raw(s, i)
        if key_raw is None:
            return None
        key = json_unescape(key_raw)
        i = _skip_ws(s, i)
        if i >= n or s[i] != ":":
            return None
        i = _skip_ws(s, i + 1)
        if i >= n:
            return None
        ch = s[i]
        if ch == '"':
            raw, i = _read_string_raw(s, i)
            if raw is None:
                return None
            tv = ("s", json_unescape(raw))
        elif s.startswith("true", i):
            tv = ("b", "true")
            i += 4
        elif s.startswith("false", i):
            tv = ("b", "false")
            i += 5
        elif s.startswith("null", i):
            tv = ("z", "null")
            i += 4
        elif ch in "{[":
            raw, i = _read_nested_raw(s, i)
            if raw is None:
                return None
            tv = ("c", raw)
        elif ch in NUM_CHARS:
            j = i
            while j < n and s[j] in NUM_CHARS:
                j += 1
            tv = ("n", s[i:j])
            i = j
        else:
            return None
        if key not in pairs:  # first wins
            pairs[key] = tv
        i = _skip_ws(s, i)
        if i >= n:
            return None
        if s[i] == ",":
            i += 1
            continue
        if s[i] == "}":
            return pairs if i == n - 1 else None
        return None


def _get_string(pairs, key) -> str:
    tv = pairs.get(key)
    if tv is not None and tv[0] == "s":
        return tv[1]
    return ""


def parse_jsonl_line(line: str):
    """Twin of ParseJsonLine. Returns (ok, fields). fields keys:
    bookmark_id, change_type, old_text, new_text, add_comment,
    apply_change (True/False/None), confidence. On ok=False all fields are
    empty/None (matching the VBA, which resets its ByRefs)."""
    empty = {
        "bookmark_id": "",
        "change_type": "",
        "old_text": "",
        "new_text": "",
        "add_comment": "",
        "apply_change": None,
        "confidence": "",
    }
    pairs = parse_top_level(line)
    if pairs is None:
        return False, dict(empty)
    bid = pairs.get("bookmark_id")
    cht = pairs.get("change_type")
    if bid is None or bid[0] != "s" or cht is None or cht[0] != "s":
        return False, dict(empty)
    fields = dict(empty)
    fields["bookmark_id"] = bid[1]
    fields["change_type"] = cht[1]
    fields["old_text"] = _get_string(pairs, "old_text")
    fields["new_text"] = _get_string(pairs, "new_text")
    fields["add_comment"] = _get_string(pairs, "add_comment")
    fields["confidence"] = _get_string(pairs, "confidence")
    ac = pairs.get("apply_change")
    if ac is not None and ac[0] == "b":
        fields["apply_change"] = ac[1] == "true"
    return True, fields


def validate_change(fields) -> str:
    """Twin of ValidateParsedChange. Returns "" (OK) or a reason code.
    Check order is part of the contract -- mirror exactly in VBA."""
    b = trim_ws(fields["bookmark_id"])
    ct = trim_ws(fields["change_type"]).lower()
    if not b:
        return MISSING_BOOKMARK
    if not ct:
        return MISSING_CHANGE_TYPE
    if ct not in CHANGE_TYPES:
        return UNKNOWN_CHANGE_TYPE
    is_comment_target = fields["bookmark_id"].startswith("AR_COMMENT_")
    if ct == "replace_text" and not trim_ws(fields["new_text"]):
        return REPLACE_REQUIRES_NEW_TEXT
    if ct == "add_comment_only" and not trim_ws(fields["add_comment"]):
        return COMMENT_REQUIRES_TEXT
    if ct == "reply_to_comment":
        if not is_comment_target:
            return REPLY_REQUIRES_COMMENT_TARGET
        if not trim_ws(fields["add_comment"]):
            return REPLY_REQUIRES_TEXT
    if ct in ("accept_revision", "reject_revision") and is_comment_target:
        return REVISION_REQUIRES_RANGE_TARGET
    return OK


def evaluate_line(line: str):
    """Combined verdict used by the golden vectors:
    returns (verdict 'PASS'|'REJECT', reason_code, fields)."""
    ok, fields = parse_jsonl_line(line)
    if not ok:
        return "REJECT", PARSE, fields
    code = validate_change(fields)
    if code:
        return "REJECT", code, fields
    return "PASS", "", fields
