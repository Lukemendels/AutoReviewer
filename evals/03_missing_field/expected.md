# Expected behavior — 03 missing field (serializer)

Run against the **serializer**. Decision [2] is a `replace_text` with no
`NEW_TEXT` — a required field for that change type is missing. Decisions [1] and
[3] are complete.

MUST:
- Output exactly one fenced ```jsonl block.
- The meta line is first inside the fence, session = 0123ABCD4567EF89, count = 2
  (only the two well-formed edits).
- Emit JSONL for [1] (replace_text on AR_PARA_00012) and [3] (reply_to_comment on
  AR_COMMENT_2).
- After the closing fence, an `OMITTED:` line naming decision [2] (AR_PARA_00021)
  and the reason (missing new_text).

MUST-NOT:
- MUST-NOT invent a NEW_TEXT for decision [2] to complete it.
- MUST-NOT include decision [2] in the count or as a JSONL line.
- MUST-NOT put the OMITTED note inside the fence.
- MUST-NOT use smart quotes or em-dashes in any field.
