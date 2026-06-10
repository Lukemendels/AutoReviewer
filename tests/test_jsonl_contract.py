from ref.jsonl_contract import (
    evaluate_line,
    json_unescape,
    parse_jsonl_line,
    parse_top_level,
    trim_ws,
    validate_change,
    PARSE,
    MISSING_BOOKMARK,
    MISSING_CHANGE_TYPE,
    UNKNOWN_CHANGE_TYPE,
    REPLACE_REQUIRES_NEW_TEXT,
    COMMENT_REQUIRES_TEXT,
    REPLY_REQUIRES_COMMENT_TARGET,
    REPLY_REQUIRES_TEXT,
    REVISION_REQUIRES_RANGE_TARGET,
)


def line_of(**kv):
    import json

    return json.dumps(kv)


# ---------- happy paths ----------

def test_basic_replace_text():
    v, code, f = evaluate_line(
        '{"bookmark_id":"AR_PARA_00012","change_type":"replace_text",'
        '"old_text":"will conclude","new_text":"is expected to conclude",'
        '"apply_change":true,"confidence":"High"}'
    )
    assert v == "PASS" and code == ""
    assert f["bookmark_id"] == "AR_PARA_00012"
    assert f["old_text"] == "will conclude"
    assert f["apply_change"] is True
    assert f["confidence"] == "High"


def test_reply_to_comment_pass():
    v, code, f = evaluate_line(
        '{"bookmark_id":"AR_COMMENT_3","change_type":"reply_to_comment",'
        '"add_comment":"Agreed."}'
    )
    assert v == "PASS"


def test_whitespace_variants_and_crlf():
    v, _, f = evaluate_line(
        '  { "bookmark_id" : "AR_PARA_00001" ,\t"change_type":"delete_element" }\r'
    )
    assert v == "PASS"
    assert f["bookmark_id"] == "AR_PARA_00001"


def test_unicode_values_pass_through():
    v, _, f = evaluate_line(
        '{"bookmark_id":"AR_PARA_00002","change_type":"replace_text",'
        '"new_text":"em—dash 安全 \U0001F600"}'
    )
    assert v == "PASS"
    assert f["new_text"] == "em—dash 安全 \U0001F600"


# ---------- the two fixed bugs ----------

def test_escaped_backslash_before_closing_quote():
    # old_text value ends with a real backslash: "a\\" in JSON.
    # The old prevCh scan never closed this string.
    v, _, f = evaluate_line(
        '{"bookmark_id":"AR_PARA_00003","change_type":"replace_text",'
        '"old_text":"a\\\\","new_text":"b"}'
    )
    assert v == "PASS"
    assert f["old_text"] == "a\\"
    assert f["new_text"] == "b"


def test_double_escaped_backslash_then_quote_in_value():
    # value is:  x\"  (escaped backslash, then escaped quote)
    v, _, f = evaluate_line(
        '{"bookmark_id":"AR_PARA_00004","change_type":"replace_text",'
        '"new_text":"x\\\\\\""}'
    )
    assert v == "PASS"
    assert f["new_text"] == 'x\\"'


def test_key_name_inside_another_value():
    # "change_type" appears inside old_text's VALUE; the old InStr lookup
    # could bind to it. The tokenizer must read the real top-level key.
    v, _, f = evaluate_line(
        '{"old_text":"the \\"change_type\\" column","bookmark_id":"AR_PARA_00005",'
        '"change_type":"replace_text","new_text":"the type column"}'
    )
    assert v == "PASS"
    assert f["change_type"] == "replace_text"
    assert f["old_text"] == 'the "change_type" column'


def test_bookmark_id_text_inside_value_not_picked_up():
    v, _, f = evaluate_line(
        '{"bookmark_id":"AR_PARA_00006","change_type":"add_comment_only",'
        '"add_comment":"see bookmark_id: AR_PARA_99999 above"}'
    )
    assert v == "PASS"
    assert f["bookmark_id"] == "AR_PARA_00006"


# ---------- unescaping ----------

def test_standard_escapes():
    assert json_unescape(r"a\nb\tc\rd\\e\"f\/g") == 'a\nb\tc\rd\\e"f/g'


def test_unknown_escape_drops_backslash():
    assert json_unescape(r"\q") == "q"


def test_u_escape_bmp():
    assert json_unescape(r"—") == "—"  # em dash
    assert json_unescape(r"ABC") == "ABC"
    assert json_unescape(r"é") == "é"


def test_u_escape_surrogate_pair_combines():
    assert json_unescape(r"😀") == "\U0001F600"  # one code point
    assert json_unescape(r"x😀y") == "x\U0001F600y"


def test_u_escape_malformed_drops_backslash():
    # not 4 hex digits after u -> unknown escape: drop backslash, keep 'u'
    assert json_unescape(r"\u12") == "u12"
    assert json_unescape(r"\uZZZZ") == "uZZZZ"
    assert json_unescape("end\\u") == "endu"


def test_u_escape_high_surrogate_without_low_stays_as_unit():
    # high surrogate followed by a NON-low-surrogate \u: each decodes alone.
    # (chr() of a lone surrogate is fine in Python as a code point.)
    assert json_unescape(r"\uD83DA") == "\ud83d" + "A"


def test_escapes_in_fields():
    v, _, f = evaluate_line(
        '{"bookmark_id":"AR_PARA_00007","change_type":"replace_text",'
        '"new_text":"line1\\nline2\\twith tab"}'
    )
    assert v == "PASS"
    assert f["new_text"] == "line1\nline2\twith tab"


# ---------- structural rejects ----------

def test_empty_line():
    assert evaluate_line("")[0:2] == ("REJECT", PARSE)


def test_not_an_object():
    assert evaluate_line("just text")[0:2] == ("REJECT", PARSE)
    assert evaluate_line('["a","b"]')[0:2] == ("REJECT", PARSE)


def test_unterminated_string():
    assert evaluate_line('{"bookmark_id":"AR_PARA_1')[0:2] == ("REJECT", PARSE)


def test_trailing_garbage():
    assert (
        evaluate_line('{"bookmark_id":"A","change_type":"delete_element"} extra')[0:2]
        == ("REJECT", PARSE)
    )
    assert (
        evaluate_line('{"bookmark_id":"A","change_type":"delete_element"}}')[0:2]
        == ("REJECT", PARSE)
    )


def test_missing_colon_or_comma():
    assert evaluate_line('{"bookmark_id" "A"}')[0:2] == ("REJECT", PARSE)
    assert (
        evaluate_line('{"bookmark_id":"A" "change_type":"delete_element"}')[0:2]
        == ("REJECT", PARSE)
    )


def test_missing_required_string_keys():
    assert evaluate_line('{"change_type":"delete_element"}')[0:2] == ("REJECT", PARSE)
    assert evaluate_line('{"bookmark_id":"AR_PARA_1"}')[0:2] == ("REJECT", PARSE)
    # present but non-string -> same as missing
    assert (
        evaluate_line('{"bookmark_id":true,"change_type":"delete_element"}')[0:2]
        == ("REJECT", PARSE)
    )


def test_empty_object():
    assert evaluate_line("{}")[0:2] == ("REJECT", PARSE)


# ---------- tokenizer details ----------

def test_duplicate_keys_first_wins():
    pairs = parse_top_level('{"a":"first","a":"second"}')
    assert pairs["a"] == ("s", "first")


def test_nested_object_value_skipped():
    v, _, f = evaluate_line(
        '{"bookmark_id":"AR_PARA_00008","change_type":"delete_element",'
        '"extra":{"nested":"with \\"quotes\\" and }"},"confidence":"Low"}'
    )
    assert v == "PASS"
    assert f["confidence"] == "Low"


def test_number_and_null_values_skipped():
    v, _, f = evaluate_line(
        '{"bookmark_id":"AR_PARA_00009","change_type":"delete_element",'
        '"count":42,"old_text":null}'
    )
    assert v == "PASS"
    assert f["old_text"] == ""  # null is not a string


def test_apply_change_string_is_ignored():
    ok, f = parse_jsonl_line(
        '{"bookmark_id":"A","change_type":"delete_element","apply_change":"false"}'
    )
    assert ok and f["apply_change"] is None


def test_apply_change_false_literal():
    ok, f = parse_jsonl_line(
        '{"bookmark_id":"A","change_type":"delete_element","apply_change":false}'
    )
    assert ok and f["apply_change"] is False


# ---------- per-change_type validation ----------

def base(**over):
    f = {
        "bookmark_id": "AR_PARA_00010",
        "change_type": "replace_text",
        "old_text": "",
        "new_text": "x",
        "add_comment": "",
        "apply_change": None,
        "confidence": "",
    }
    f.update(over)
    return f


def test_validation_blank_bookmark():
    assert validate_change(base(bookmark_id="   ")) == MISSING_BOOKMARK


def test_validation_blank_change_type():
    assert validate_change(base(change_type=" ")) == MISSING_CHANGE_TYPE


def test_validation_unknown_change_type():
    assert validate_change(base(change_type="rewrite_all")) == UNKNOWN_CHANGE_TYPE


def test_validation_change_type_case_insensitive():
    assert validate_change(base(change_type="Replace_Text")) == ""


def test_validation_replace_needs_new_text():
    assert validate_change(base(new_text="  ")) == REPLACE_REQUIRES_NEW_TEXT


def test_validation_comment_needs_text():
    assert (
        validate_change(base(change_type="add_comment_only", add_comment=""))
        == COMMENT_REQUIRES_TEXT
    )


def test_validation_reply_needs_comment_target():
    assert (
        validate_change(
            base(change_type="reply_to_comment", add_comment="hi")
        )
        == REPLY_REQUIRES_COMMENT_TARGET
    )


def test_validation_reply_needs_text():
    assert (
        validate_change(
            base(
                bookmark_id="AR_COMMENT_2",
                change_type="reply_to_comment",
                add_comment=" ",
            )
        )
        == REPLY_REQUIRES_TEXT
    )


def test_validation_revision_on_comment_target():
    assert (
        validate_change(
            base(bookmark_id="AR_COMMENT_2", change_type="accept_revision")
        )
        == REVISION_REQUIRES_RANGE_TARGET
    )
    assert (
        validate_change(
            base(bookmark_id="AR_REV_00001", change_type="reject_revision")
        )
        == ""
    )


def test_trim_ws_strips_crlf():
    assert trim_ws("\t {\"a\":1} \r\n") == '{"a":1}'
