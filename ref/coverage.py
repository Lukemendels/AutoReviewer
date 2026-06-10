"""Comment-coverage check (twin of ComputeUnaddressed in modReviewImport.bas).

A comment-laden document is what the tool adjudicates; a comment that receives
NO edit line is the silent false negative the tool exists to prevent (MKS
Profile s9.4). Given the ordered list of AR_COMMENT_ ids enumerated at export
and the parsed edit lines, this computes which comments were addressed -- a
comment is addressed iff some reply_to_comment or add_comment_only line targets
its id -- and returns the ordered list of unaddressed ids.

The apply step warn-gates on a non-empty list (Proceed / Abort): a no-action
ruling is legitimate, but it must be a visible act, never an omission.

Notes mirrored in VBA (do not change one side alone):
- Only reply_to_comment and add_comment_only count as addressing; a
  replace_text/delete_element that happens to target an AR_COMMENT_ id does NOT.
- apply_change is NOT consulted: a line that targets a comment counts as having
  addressed it (the operator's intent is recorded), matching the spec wording
  "ids targeted by any reply_to_comment / add_comment_only line".
- The unaddressed list preserves the comment_ids input order.
"""

ADDRESSING_CHANGE_TYPES = ("reply_to_comment", "add_comment_only")

ALL_ADDRESSED = "ALL_ADDRESSED"
UNADDRESSED = "UNADDRESSED"


def addressed_ids(edits):
    """edits: iterable of (change_type, bookmark_id). Returns the set of
    AR_COMMENT_ ids addressed by an addressing change type."""
    out = set()
    for ct, bid in edits:
        if ct.strip().lower() in ADDRESSING_CHANGE_TYPES:
            b = bid.strip()
            if b.startswith("AR_COMMENT_"):
                out.add(b)
    return out


def unaddressed_comments(comment_ids, edits):
    """Returns (ordered_unaddressed_list, status)."""
    addressed = addressed_ids(edits)
    missing = [c.strip() for c in comment_ids if c.strip() and c.strip() not in addressed]
    status = ALL_ADDRESSED if not missing else UNADDRESSED
    return missing, status
