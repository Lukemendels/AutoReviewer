# Expected behavior — 04 skipped comment (coverage)

Three comments. AR_COMMENT_1 and AR_COMMENT_2 are easy concrete edits;
AR_COMMENT_3 is an open question that is tempting to leave unanswered. Coverage
must force all three to be visibly adjudicated.

MUST:
- Produce a block (or NO_ACTION ruling) for **each** of AR_COMMENT_1,
  AR_COMMENT_2, AR_COMMENT_3 — including the open-question AR_COMMENT_3 (at
  minimum a reply_to_comment that answers or asks for the needed input).
- End with `COVERAGE: addressed 3 of 3 comments; NO_ACTION: none` (or list any
  NO_ACTION ids explicitly if it rules no-action on one — but it must be named,
  not omitted).

MUST-NOT:
- MUST-NOT silently drop AR_COMMENT_3 (no block and not listed in COVERAGE).
- MUST-NOT report a COVERAGE count that disagrees with the blocks actually shown.
