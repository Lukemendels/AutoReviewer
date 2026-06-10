from ref.session import (
    check_session,
    filter_payload_lines,
    parse_meta_line,
    NO_EXPORT_TOKEN,
    NO_PAYLOAD,
    META_MISSING,
    TOKEN_MISMATCH,
    COUNT_MISMATCH,
)

TOKEN = "0123ABCD4567EF89"


def meta(count, token=TOKEN):
    return f'{{"meta": "autoreviewer", "session": "{token}", "count": {count}}}'


EDIT = '{"bookmark_id":"AR_PARA_00001","change_type":"delete_element"}'


def test_valid_session():
    ok, code = check_session([meta(2), EDIT, EDIT], TOKEN)
    assert ok and code == ""


def test_valid_zero_edits():
    ok, code = check_session([meta(0)], TOKEN)
    assert ok


def test_wrong_token():
    ok, code = check_session([meta(1, token="FFFFFFFFFFFFFFFF"), EDIT], TOKEN)
    assert not ok and code == TOKEN_MISMATCH


def test_wrong_count():
    ok, code = check_session([meta(3), EDIT], TOKEN)
    assert not ok and code == COUNT_MISMATCH


def test_missing_meta():
    ok, code = check_session([EDIT, EDIT], TOKEN)
    assert not ok and code == META_MISSING


def test_meta_not_on_line_1():
    ok, code = check_session([EDIT, meta(1)], TOKEN)
    assert not ok and code == META_MISSING


def test_no_payload():
    ok, code = check_session([], TOKEN)
    assert not ok and code == NO_PAYLOAD


def test_no_export_token():
    ok, code = check_session([meta(0)], "")
    assert not ok and code == NO_EXPORT_TOKEN


def test_meta_wrong_app():
    bad = '{"meta": "otherapp", "session": "%s", "count": 0}' % TOKEN
    ok, code = check_session([bad], TOKEN)
    assert not ok and code == META_MISSING


def test_meta_count_not_a_number():
    bad = '{"meta": "autoreviewer", "session": "%s", "count": "two"}' % TOKEN
    ok, code = check_session([bad], TOKEN)
    assert not ok and code == META_MISSING


def test_parse_meta_line_negative_count():
    assert parse_meta_line(
        '{"meta": "autoreviewer", "session": "x", "count": -1}'
    ) is None


def test_parse_meta_line_strict_integer_rule():
    # optional minus + digits only, max 9 digits (mirrors the VBA exactly)
    assert parse_meta_line('{"meta": "autoreviewer", "session": "x", "count": +1}') is None
    assert parse_meta_line('{"meta": "autoreviewer", "session": "x", "count": 1.0}') is None
    assert parse_meta_line('{"meta": "autoreviewer", "session": "x", "count": 1e2}') is None
    assert parse_meta_line('{"meta": "autoreviewer", "session": "x", "count": 1234567890}') is None
    assert parse_meta_line('{"meta": "autoreviewer", "session": "x", "count": 999999999}') == ("x", 999999999)


def test_filter_drops_blanks_and_fences():
    raw = ["```jsonl", meta(1), "", "  ", EDIT, "```"]
    assert filter_payload_lines(raw) == [meta(1), EDIT]


def test_filter_no_fence_takes_all_nonblank():
    raw = [meta(2), "", EDIT, EDIT]
    assert filter_payload_lines(raw) == [meta(2), EDIT, EDIT]


def test_filter_language_tag_fence():
    # uppercase / arbitrary language tag still recognized as a fence
    raw = ["```JSONL", meta(1), EDIT, "```"]
    assert filter_payload_lines(raw) == [meta(1), EDIT]


def test_filter_ignores_prose_after_closing_fence():
    raw = ["```jsonl", meta(1), EDIT, "```", "Omitted: AR_PARA_9 missing new_text."]
    assert filter_payload_lines(raw) == [meta(1), EDIT]


def test_filter_fence_only_is_empty():
    assert filter_payload_lines(["```jsonl", "```"]) == []
    ok, code = check_session(filter_payload_lines(["```jsonl", "```"]), TOKEN)
    assert not ok and code == NO_PAYLOAD


def test_filter_unclosed_fence_takes_rest():
    raw = ["```jsonl", meta(1), EDIT]
    assert filter_payload_lines(raw) == [meta(1), EDIT]


def test_filtered_fenced_paste_passes_end_to_end():
    raw = ["```jsonl", meta(2), EDIT, EDIT, "```"]
    ok, code = check_session(filter_payload_lines(raw), TOKEN)
    assert ok


def test_fenced_with_prose_after_passes_end_to_end():
    raw = ["```jsonl", meta(2), EDIT, EDIT, "```", "Note: nothing omitted."]
    ok, code = check_session(filter_payload_lines(raw), TOKEN)
    assert ok
