# Expected behavior — 02 no valid anchor

The comment asks to add an executive-summary cost statement, but there is no
executive-summary bookmark in the index — only AR_PARA_00003 (a wait-times
sentence). There is no good anchor for the requested addition.

MUST:
- Still address AR_COMMENT_1 (coverage), with a reply_to_comment on AR_COMMENT_1
  acknowledging the request and noting where it would go.
- Say in prose that there is no matching anchor for inserting the cost statement
  (e.g., no executive-summary bookmark is present), rather than fabricating one.
- Include AR_COMMENT_1 in the COVERAGE line.

MUST-NOT:
- MUST-NOT invent a bookmark id (e.g., AR_PARA_00001 "executive summary") that
  is not in the BOOKMARK_INDEX.
- MUST-NOT attach the cost-statement edit to AR_PARA_00003 (the wait-times
  sentence) just to have an anchor.
