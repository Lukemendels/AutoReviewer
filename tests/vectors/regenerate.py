"""Regenerate the golden vector files from the Python reference twin.

Run from the repo root:  python3 tests/vectors/regenerate.py

The vector files are GENERATED ARTIFACTS -- never hand-edit them. The Python
twin (ref/) is the source of truth; the VBA self-test harness (modSelfTest.bas)
replays these vectors against the VBA transliteration to prove the two agree.

Escape rule (shared by all files; the VBA harness reverses it):
  \\t -> TAB        \\n -> LF         \\r -> CR        \\\\ -> backslash
  \\uXXXX -> the UTF-16 code unit 0xXXXX (uppercase hex; surrogate halves
             appear as two consecutive \\u escapes, matching VBA's ChrW)
  every other char is literal printable ASCII 0x20..0x7E.
Files are therefore pure ASCII, so VBA native line I/O reads them safely.
"""

import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))

from ref.fingerprint import fingerprint, utf16_code_units  # noqa: E402
from ref.jsonl_contract import evaluate_line  # noqa: E402
from ref.session import check_session, filter_payload_lines  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))


def vector_escape(s: str) -> str:
    out = []
    for u in utf16_code_units(s):
        if u == 0x5C:
            out.append("\\\\")
        elif u == 0x09:
            out.append("\\t")
        elif u == 0x0A:
            out.append("\\n")
        elif u == 0x0D:
            out.append("\\r")
        elif 0x20 <= u <= 0x7E:
            out.append(chr(u))
        else:
            out.append("\\u%04X" % u)
    return "".join(out)


def vector_unescape(s: str) -> str:
    """Reference implementation of the un-escape (mirrored in modSelfTest)."""
    out = []
    i = 0
    n = len(s)
    while i < n:
        ch = s[i]
        if ch == "\\" and i + 1 < n:
            esc = s[i + 1]
            if esc == "t":
                out.append("\t")
                i += 2
            elif esc == "n":
                out.append("\n")
                i += 2
            elif esc == "r":
                out.append("\r")
                i += 2
            elif esc == "\\":
                out.append("\\")
                i += 2
            elif esc == "u" and i + 6 <= n:
                out.append(chr(int(s[i + 2 : i + 6], 16)))
                i += 6
            else:
                out.append(esc)
                i += 2
        else:
            out.append(ch)
            i += 1
    # surrogate halves were appended as lone chr(code); join them back into
    # real code points so the string round-trips through utf-16 encoding
    res = "".join(out)
    return res.encode("utf-16", "surrogatepass").decode("utf-16")


# --------------------------------------------------------------------------
# Fingerprint cases
# --------------------------------------------------------------------------

FINGERPRINT_INPUTS = [
    "",
    "a",
    "A",
    "z",
    "0",
    " ",
    "\t",
    "\n",
    "\r",
    "\r\n",
    "ab",
    "ba",
    "abc",
    "abd",
    "hello world",
    "Hello World",
    "The quick brown fox jumps over the lazy dog",
    "the quick brown fox jumps over the lazy dog",
    '{"bookmark_id":"AR_PARA_00012","change_type":"replace_text"}',
    '{"meta": "autoreviewer", "session": "0123ABCD4567EF89", "count": 3}',
    "line one\nline two",
    "line one\r\nline two",
    "col1\tcol2\tcol3",
    "back\\slash",
    "double\\\\backslash",
    'quote " inside',
    "em—dash",
    "en–dash",
    "curly “quotes” and ‘apostrophes’",
    "non breaking space",
    "café résumé",
    "安全运输",  # CJK
    "\U0001F600",  # emoji: surrogate pair D83D DE00
    "ok \U0001F600 done",
    "\U0001F1FA\U0001F1F8",  # regional indicators (two pairs)
    "A" * 100,
    "A" * 251,
    "A" * 252,
    "x" * 251 + "Q",
    "Q" + "x" * 251,
    "A" * 1000,
    ("payload " * 200).strip(),
    "mixed — ascii 安 emoji \U0001F600 tail",
    " leading and trailing ",
]

# --------------------------------------------------------------------------
# Parser cases (every adversarial case from the pytest suite appears here)
# --------------------------------------------------------------------------

PARSER_LINES = [
    # --- happy paths ---
    '{"bookmark_id":"AR_PARA_00012","change_type":"replace_text","old_text":"will conclude","new_text":"is expected to conclude","apply_change":true,"confidence":"High"}',
    '{"bookmark_id":"AR_PARA_00015","change_type":"delete_element","apply_change":true,"confidence":"Medium"}',
    '{"bookmark_id":"AR_COMMENT_3","change_type":"reply_to_comment","add_comment":"Agreed."}',
    '{"bookmark_id":"AR_COMMENT_12","change_type":"reply_to_comment","add_comment":"No action; see rationale.","confidence":"Low"}',
    '{"bookmark_id":"AR_REV_00001","change_type":"accept_revision"}',
    '{"bookmark_id":"AR_REV_00002","change_type":"reject_revision","add_comment":"Keeping original term."}',
    '{"bookmark_id":"AR_CELL_1_2_3","change_type":"replace_text","new_text":"$97,652"}',
    '{"bookmark_id":"AR_FN_001","change_type":"add_comment_only","add_comment":"Verify citation."}',
    '{"bookmark_id":"AR_PARA_00001","change_type":"replace_text","old_text":"","new_text":"full paragraph replacement"}',
    # whitespace variants
    '  { "bookmark_id" : "AR_PARA_00002" ,\t"change_type":"delete_element" }  ',
    '{"bookmark_id":"AR_PARA_00002","change_type":"delete_element"}\r',
    '\t{"bookmark_id":"AR_PARA_00002","change_type":"delete_element"}',
    # case-insensitive change_type
    '{"bookmark_id":"AR_PARA_00030","change_type":"Replace_Text","new_text":"x"}',
    '{"bookmark_id":"AR_PARA_00031","change_type":"DELETE_ELEMENT"}',
    # unicode values
    '{"bookmark_id":"AR_PARA_00032","change_type":"replace_text","new_text":"em—dash and café"}',
    '{"bookmark_id":"AR_PARA_00033","change_type":"replace_text","new_text":"安全 \U0001F600"}',
    # --- bug 1: escaped backslash before closing quote ---
    '{"bookmark_id":"AR_PARA_00003","change_type":"replace_text","old_text":"a\\\\","new_text":"b"}',
    '{"bookmark_id":"AR_PARA_00004","change_type":"replace_text","new_text":"x\\\\\\""}',
    '{"bookmark_id":"AR_PARA_00040","change_type":"replace_text","new_text":"C:\\\\Users\\\\doc.docx"}',
    '{"bookmark_id":"AR_PARA_00041","change_type":"replace_text","new_text":"ends with backslash\\\\"}',
    # --- bug 2: key name inside another value ---
    '{"old_text":"the \\"change_type\\" column","bookmark_id":"AR_PARA_00005","change_type":"replace_text","new_text":"the type column"}',
    '{"bookmark_id":"AR_PARA_00006","change_type":"add_comment_only","add_comment":"see bookmark_id: AR_PARA_99999 above"}',
    '{"bookmark_id":"AR_PARA_00042","change_type":"replace_text","old_text":"mention of new_text in prose","new_text":"fixed"}',
    '{"bookmark_id":"AR_PARA_00043","change_type":"replace_text","new_text":"apply_change should not bind here","apply_change":false}',
    # --- escapes ---
    '{"bookmark_id":"AR_PARA_00007","change_type":"replace_text","new_text":"line1\\nline2\\twith tab"}',
    '{"bookmark_id":"AR_PARA_00044","change_type":"replace_text","new_text":"quote: \\" slash: \\/ cr: \\r"}',
    '{"bookmark_id":"AR_PARA_00045","change_type":"replace_text","new_text":"unknown \\q escape"}',
    # --- \\uXXXX decoding (item 1) ---
    '{"bookmark_id":"AR_PARA_00046","change_type":"replace_text","new_text":"em dash \\u2014 here"}',
    '{"bookmark_id":"AR_PARA_00047b","change_type":"replace_text","new_text":"\\u201Cquoted\\u201D and \\u00e9"}',
    '{"bookmark_id":"AR_PARA_00047c","change_type":"replace_text","new_text":"ascii \\u0041\\u0042\\u0043 BC"}',
    '{"bookmark_id":"AR_PARA_00047d","change_type":"replace_text","new_text":"emoji \\uD83D\\uDE00 tail"}',
    '{"bookmark_id":"AR_PARA_00047e","change_type":"replace_text","new_text":"two faces \\uD83D\\uDE00\\uD83D\\uDE42"}',
    '{"bookmark_id":"AR_PARA_00047f","change_type":"replace_text","new_text":"malformed \\u12 short"}',
    '{"bookmark_id":"AR_PARA_00047g","change_type":"replace_text","new_text":"malformed \\uZZZZ here"}',
    '{"bookmark_id":"AR_COMMENT_9","change_type":"reply_to_comment","add_comment":"unicode reply \\u2014 \\uD83D\\uDC4D"}',
    # --- duplicate keys: first wins ---
    '{"bookmark_id":"AR_PARA_00047","bookmark_id":"AR_PARA_99999","change_type":"delete_element"}',
    '{"bookmark_id":"AR_PARA_00048","change_type":"delete_element","change_type":"replace_text"}',
    # --- nested / number / null values skipped ---
    '{"bookmark_id":"AR_PARA_00008","change_type":"delete_element","extra":{"nested":"with \\"quotes\\" and }"},"confidence":"Low"}',
    '{"bookmark_id":"AR_PARA_00009","change_type":"delete_element","count":42,"old_text":null}',
    '{"bookmark_id":"AR_PARA_00049","change_type":"delete_element","tags":["a","b}","c"]}',
    '{"bookmark_id":"AR_PARA_00050","change_type":"delete_element","weight":-3.5e2}',
    # apply_change handling
    '{"bookmark_id":"AR_PARA_00051","change_type":"delete_element","apply_change":false}',
    '{"bookmark_id":"AR_PARA_00052","change_type":"delete_element","apply_change":"false"}',
    '{"bookmark_id":"AR_PARA_00053","change_type":"delete_element","apply_change":true}',
    # --- structural rejects ---
    "",
    "   ",
    "just text",
    '["a","b"]',
    "{}",
    "{",
    "}",
    '{"bookmark_id":"AR_PARA_1',
    '{"bookmark_id":"AR_PARA_1\\',
    '{"bookmark_id":"A","change_type":"delete_element"} extra',
    '{"bookmark_id":"A","change_type":"delete_element"}}',
    '{"bookmark_id" "A"}',
    '{"bookmark_id":"A" "change_type":"delete_element"}',
    '{"bookmark_id":"A",}',
    '{:"A"}',
    '{"bookmark_id":}',
    "```jsonl",
    # missing / non-string required keys
    '{"change_type":"delete_element"}',
    '{"bookmark_id":"AR_PARA_1"}',
    '{"bookmark_id":true,"change_type":"delete_element"}',
    '{"bookmark_id":"AR_PARA_1","change_type":42}',
    '{"bookmark_id":null,"change_type":"delete_element"}',
    # --- validation rejects (parse OK) ---
    '{"bookmark_id":"","change_type":"delete_element"}',
    '{"bookmark_id":"   ","change_type":"delete_element"}',
    '{"bookmark_id":"AR_PARA_00060","change_type":""}',
    '{"bookmark_id":"AR_PARA_00061","change_type":"rewrite_all"}',
    '{"bookmark_id":"AR_PARA_00062","change_type":"replace_text"}',
    '{"bookmark_id":"AR_PARA_00063","change_type":"replace_text","new_text":"   "}',
    '{"bookmark_id":"AR_PARA_00064","change_type":"add_comment_only"}',
    '{"bookmark_id":"AR_PARA_00065","change_type":"reply_to_comment","add_comment":"hi"}',
    '{"bookmark_id":"AR_COMMENT_4","change_type":"reply_to_comment"}',
    '{"bookmark_id":"AR_COMMENT_5","change_type":"accept_revision"}',
    '{"bookmark_id":"AR_COMMENT_6","change_type":"reject_revision"}',
]

# --------------------------------------------------------------------------
# Session cases
# --------------------------------------------------------------------------

TOKEN = "0123ABCD4567EF89"
EDIT = '{"bookmark_id":"AR_PARA_00001","change_type":"delete_element"}'


def _meta(count, token=TOKEN, app="autoreviewer"):
    return '{"meta": "%s", "session": "%s", "count": %s}' % (app, token, count)


SESSION_CASES = [
    # (name, expected_token, raw_lines)
    ("valid_two_edits", TOKEN, [_meta(2), EDIT, EDIT]),
    ("valid_zero_edits", TOKEN, [_meta(0)]),
    ("valid_with_fences_and_blanks", TOKEN, ["```jsonl", _meta(2), EDIT, "", EDIT, "```"]),
    ("wrong_token", TOKEN, [_meta(1, token="FFFFFFFFFFFFFFFF"), EDIT]),
    ("count_too_high", TOKEN, [_meta(3), EDIT]),
    ("count_too_low", TOKEN, [_meta(1), EDIT, EDIT]),
    ("missing_meta", TOKEN, [EDIT, EDIT]),
    ("meta_not_on_line_1", TOKEN, [EDIT, _meta(1)]),
    ("empty_payload", TOKEN, []),
    ("only_blank_lines", TOKEN, ["", "   ", "\t"]),
    ("no_export_token", "", [_meta(0)]),
    ("meta_wrong_app", TOKEN, [_meta(0, app="otherapp")]),
    ("meta_count_not_number", TOKEN, ['{"meta": "autoreviewer", "session": "%s", "count": "two"}' % TOKEN]),
    ("meta_negative_count", TOKEN, ['{"meta": "autoreviewer", "session": "%s", "count": -1}' % TOKEN]),
    ("meta_count_plus_sign", TOKEN, ['{"meta": "autoreviewer", "session": "%s", "count": +1}' % TOKEN, EDIT]),
    ("meta_count_decimal", TOKEN, ['{"meta": "autoreviewer", "session": "%s", "count": 1.0}' % TOKEN, EDIT]),
]


def build_fingerprint_vectors() -> str:
    lines = [
        "# AutoReviewer golden vectors: fingerprint suite. GENERATED by tests/vectors/regenerate.py -- do not hand-edit.",
        "# Format: escaped_input <TAB> expected_hex16",
        "# Escapes: \\t tab, \\n LF, \\r CR, \\\\ backslash, \\uXXXX UTF-16 code unit (hex, uppercase).",
        "# All other characters are literal printable ASCII. Lines starting with # and blank lines are skipped.",
    ]
    for s in FINGERPRINT_INPUTS:
        lines.append("%s\t%s" % (vector_escape(s), fingerprint(s)))
    return "\n".join(lines) + "\n"


def build_parser_vectors() -> str:
    lines = [
        "# AutoReviewer golden vectors: JSONL parser + validation suite. GENERATED by tests/vectors/regenerate.py -- do not hand-edit.",
        "# Format: escaped_line <TAB> PASS|REJECT <TAB> reason <TAB> bookmark_id <TAB> change_type <TAB> old_text <TAB> new_text <TAB> add_comment <TAB> apply_change <TAB> confidence",
        "# Field columns are escaped with the same rule as the input column (\\t \\n \\r \\\\ \\uXXXX).",
        "# reason is empty for PASS; PARSE for structural failures; otherwise a validation code. On PARSE rows all fields are empty.",
        "# apply_change column: true | false | empty (absent or non-boolean).",
    ]
    for raw in PARSER_LINES:
        verdict, reason, f = evaluate_line(raw)
        ac = f["apply_change"]
        ac_s = "" if ac is None else ("true" if ac else "false")
        cols = [
            vector_escape(raw),
            verdict,
            reason,
            vector_escape(f["bookmark_id"]),
            vector_escape(f["change_type"]),
            vector_escape(f["old_text"]),
            vector_escape(f["new_text"]),
            vector_escape(f["add_comment"]),
            ac_s,
            vector_escape(f["confidence"]),
        ]
        lines.append("\t".join(cols))
    return "\n".join(lines) + "\n"


def build_session_vectors() -> str:
    lines = [
        "# AutoReviewer golden vectors: session-binding gate suite. GENERATED by tests/vectors/regenerate.py -- do not hand-edit.",
        "# Format: case_name <TAB> expected_token <TAB> escaped_payload (raw lines joined by \\n) <TAB> PASS|FAIL <TAB> code",
        "# The harness un-escapes the payload, splits on LF, trims/filters lines (blank + ``` fences), then calls the gate.",
    ]
    for name, token, raw_lines in SESSION_CASES:
        payload = "\n".join(raw_lines)
        ok, code = check_session(filter_payload_lines(raw_lines), token)
        cols = [
            name,
            token,
            vector_escape(payload),
            "PASS" if ok else "FAIL",
            code,
        ]
        lines.append("\t".join(cols))
    return "\n".join(lines) + "\n"


def main():
    files = {
        "fingerprint_vectors.txt": build_fingerprint_vectors(),
        "parser_vectors.txt": build_parser_vectors(),
        "session_vectors.txt": build_session_vectors(),
    }
    for name, content in files.items():
        path = os.path.join(HERE, name)
        content.encode("ascii")  # hard guarantee: vector files are pure ASCII
        with open(path, "w", encoding="ascii", newline="\r\n") as fh:
            fh.write(content)
        print("wrote %s (%d lines)" % (name, content.count("\n")))


if __name__ == "__main__":
    main()
