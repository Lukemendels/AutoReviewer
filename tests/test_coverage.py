from ref.coverage import (
    addressed_ids,
    unaddressed_comments,
    ALL_ADDRESSED,
    UNADDRESSED,
)


def test_all_addressed():
    ids = ["AR_COMMENT_1", "AR_COMMENT_2"]
    edits = [("reply_to_comment", "AR_COMMENT_1"), ("add_comment_only", "AR_COMMENT_2")]
    missing, status = unaddressed_comments(ids, edits)
    assert missing == [] and status == ALL_ADDRESSED


def test_one_missing():
    ids = ["AR_COMMENT_1", "AR_COMMENT_2", "AR_COMMENT_3"]
    edits = [("reply_to_comment", "AR_COMMENT_1"), ("reply_to_comment", "AR_COMMENT_3")]
    missing, status = unaddressed_comments(ids, edits)
    assert missing == ["AR_COMMENT_2"] and status == UNADDRESSED


def test_none_addressed():
    ids = ["AR_COMMENT_1", "AR_COMMENT_2"]
    missing, status = unaddressed_comments(ids, [])
    assert missing == ["AR_COMMENT_1", "AR_COMMENT_2"] and status == UNADDRESSED


def test_no_comments():
    missing, status = unaddressed_comments([], [("replace_text", "AR_PARA_1")])
    assert missing == [] and status == ALL_ADDRESSED


def test_addressing_noncomment_target_does_not_count():
    ids = ["AR_COMMENT_1"]
    edits = [("add_comment_only", "AR_PARA_5")]
    missing, _ = unaddressed_comments(ids, edits)
    assert missing == ["AR_COMMENT_1"]


def test_replace_text_on_comment_id_is_not_addressing():
    ids = ["AR_COMMENT_1"]
    edits = [("replace_text", "AR_COMMENT_1")]
    missing, _ = unaddressed_comments(ids, edits)
    assert missing == ["AR_COMMENT_1"]


def test_revision_on_comment_id_is_not_addressing():
    ids = ["AR_COMMENT_1"]
    edits = [("accept_revision", "AR_COMMENT_1")]
    missing, _ = unaddressed_comments(ids, edits)
    assert missing == ["AR_COMMENT_1"]


def test_duplicate_addressing_is_fine():
    ids = ["AR_COMMENT_1", "AR_COMMENT_2"]
    edits = [
        ("reply_to_comment", "AR_COMMENT_1"),
        ("reply_to_comment", "AR_COMMENT_1"),
        ("reply_to_comment", "AR_COMMENT_2"),
    ]
    missing, status = unaddressed_comments(ids, edits)
    assert missing == [] and status == ALL_ADDRESSED


def test_order_preserved():
    ids = ["AR_COMMENT_3", "AR_COMMENT_1", "AR_COMMENT_2"]
    missing, _ = unaddressed_comments(ids, [])
    assert missing == ["AR_COMMENT_3", "AR_COMMENT_1", "AR_COMMENT_2"]


def test_case_insensitive_change_type():
    ids = ["AR_COMMENT_1"]
    edits = [("Reply_To_Comment", "AR_COMMENT_1")]
    missing, _ = unaddressed_comments(ids, edits)
    assert missing == []


def test_addressed_ids_helper():
    edits = [("reply_to_comment", "AR_COMMENT_9"), ("delete_element", "AR_PARA_1")]
    assert addressed_ids(edits) == {"AR_COMMENT_9"}
