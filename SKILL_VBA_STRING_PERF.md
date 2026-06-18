---
name: vba-string-concat-perf
description: Use this skill whenever writing or reviewing VBA code that builds a string inside a loop — especially when the loop count scales with document size (revision count, comment count, paragraph count, footnote count, or any other Word/Excel object collection). Trigger it any time you see `s = s & ...` or `s = s & vbCrLf` inside a For, For Each, or Do loop, or before writing a new export/serialization function in any VBA module. The skill enforces the array+Join pattern as the only approved approach for O(n)-bounded string construction.
---

# VBA String Concatenation Performance — Authoring Skill

This skill prevents O(n²) string concatenation bugs in VBA. The canonical symptom is an export or serialization function that runs acceptably on small documents but takes 30+ minutes on production documents with hundreds of revisions or paragraphs.

**Before writing any loop that appends to a string, read Section 1.** The failure mode is silent — the output is identical, the code looks correct, and it only reveals itself when document size grows.

---

## 1. Root cause: VBA strings are immutable copy-on-write

Every `s = s & "more"` statement in VBA:
1. Allocates a new string buffer of length `len(s) + len("more")`
2. Copies all of `s` into the new buffer
3. Appends `"more"`
4. Frees the old buffer (eventually)

In a loop of n iterations where each append adds ~k characters, total bytes copied grow as:
```
k + 2k + 3k + ... + nk  =  k·n·(n+1)/2  =  O(n²)
```

**Concrete case from this codebase:** 768 revisions × 6 lines × ~50 chars average → ~886 MB total bytes copied → 30-minute export on hardware that can normally process the document in under 30 seconds.

There is no JIT, no copy-on-write optimization, and no string builder in VBA. The only fix is structural.

---

## 2. Detection: what to look for

Flag any loop where the accumulated string scales with a Word or Excel object collection:

```vba
' ❌ O(n²) — WRONG
Dim s As String
For Each c In wdDoc.Comments        ' n = document-size-dependent
    s = s & "## AR_COMMENT_" & c.Index & vbCrLf
    s = s & "Author: " & c.Author & vbCrLf
    s = s & "Text: " & c.Range.Text & vbCrLf
    s = s & "---" & vbCrLf
Next c
```

The signal is `s = s &` (or `buffer = buffer &`) inside a loop whose bound is:
- `wdDoc.Comments.Count`
- `wdDoc.Revisions.Count`
- `wdDoc.Paragraphs.Count`
- `wdDoc.Footnotes.Count`
- `wdDoc.Bookmarks.Count`
- Any `Range`, `Table`, or collection that grows with document content

Loops bounded by **small constants** (≤ ~20 elements, configuration rows, fixed column counts) are acceptable and do not need the array+Join pattern.

---

## 3. The fix: pre-dimension an array, fill it, Join once

```vba
' ✅ O(n) — CORRECT
Dim cmtBodyArr() As String
Dim cmtBodyIdx As Long
ReDim cmtBodyArr(1 To wdDoc.Comments.Count * 7)   ' max elements per iteration × count
cmtBodyIdx = 0

For Each c In wdDoc.Comments
    cmtBodyIdx = cmtBodyIdx + 1: cmtBodyArr(cmtBodyIdx) = "## AR_COMMENT_" & c.Index
    cmtBodyIdx = cmtBodyIdx + 1: cmtBodyArr(cmtBodyIdx) = "Author: " & c.Author
    cmtBodyIdx = cmtBodyIdx + 1: cmtBodyArr(cmtBodyIdx) = "Text: " & c.Range.Text
    cmtBodyIdx = cmtBodyIdx + 1: cmtBodyArr(cmtBodyIdx) = "---"
Next c

result = Join(cmtBodyArr, vbCrLf) & vbCrLf
```

**Rules:**
- Pre-dimension to `collection.Count * maxLinesPerIteration` — always over-estimate; the cost of a slightly-too-large array is negligible.
- Each array slot holds one logical line, with **no trailing `vbCrLf`** in the element.
- `Join(arr, vbCrLf)` places newlines between elements. Add `& vbCrLf` after Join to restore the final trailing newline.
- Assign the final string **once**, outside the loop.

---

## 4. Edge case: partial fills (filter inside the loop)

When not every iteration produces an element (e.g., bookmarks filtered by prefix, footnotes filtered by non-empty body), the actual fill count will be less than the pre-dimensioned size. Use `ReDim Preserve` before `Join`:

```vba
' ✅ Partial fill
ReDim bmLines(1 To docForExport.Bookmarks.Count)   ' upper bound
bmCount = 0

For Each bm In docForExport.Bookmarks
    If Left$(bm.Name, 3) = "AR_" Then              ' filter: not every bm qualifies
        bmCount = bmCount + 1
        bmLines(bmCount) = bm.Name & " | type=..."
    End If
Next bm

' Guard: ReDim Preserve to (1 To 0) is a runtime error
If bmCount > 0 Then
    ReDim Preserve bmLines(1 To bmCount)
    result = Join(bmLines, vbCrLf) & vbCrLf
End If
```

Never call `ReDim Preserve arr(1 To 0)` — it raises a subscript-out-of-range error. Always guard with `If fillCount > 0`.

---

## 5. Edge case: elements that carry their own terminators

When each logical element already ends with `vbCrLf` (e.g., because some elements are two-line blocks with an embedded blank line), use `Join(arr, "")` instead:

```vba
' ✅ Elements contain their own newlines — Join with empty separator
dlIdx = dlIdx + 1: docLines(dlIdx) = headingLine & vbCrLf
dlIdx = dlIdx + 1: docLines(dlIdx) = vbCrLf        ' blank line between paragraphs

result = Join(docLines, "")   ' ← empty separator; elements are self-terminated
```

The rule: if removing `vbCrLf` from one element and adding it to the `Join` separator would change the output, use `Join(arr, "")`.

---

## 6. Edge case: section markers stay outside the array

Section headers and footers (`<<COMMENTS_START>>`, `<<REVISIONS_START>>`, etc.) are structural constants. Do not put them inside the array. Concatenate them onto the Join result after the loop:

```vba
' ✅ Section markers outside the array
result = "<<COMMENTS_START>>" & vbCrLf & _
         Join(cmtBodyArr, vbCrLf) & vbCrLf & _
         "<<COMMENTS_END>>" & vbCrLf & vbCrLf
```

This keeps the array exclusively for per-item content and makes the structural shape of the output immediately readable.

---

## 7. Edge case: preserving On Error Resume Next around COM properties

Word COM properties (`.Author`, `.Date`, `.Range.Text`, `.Type`) can raise errors on malformed revisions. The original pattern wraps individual field reads in `On Error Resume Next`. With the array pattern, initialize safe-default temp strings first, read COM values into temps under `On Error Resume Next`, then assign temps to array slots after restoring the real error handler:

```vba
' ✅ Error handler preserved with array pattern
Dim authorLine As String
Dim dateLine As String

authorLine = "Author: "    ' safe default if COM call fails
dateLine = "Date: "

On Error Resume Next
authorLine = "Author: " & CStr(revObj.Author)
dateLine = "Date: " & CStr(revObj.Date)
On Error GoTo ErrHandler   ' restore structured handler before array writes

revBodyIdx = revBodyIdx + 1: revBodyArr(revBodyIdx) = authorLine
revBodyIdx = revBodyIdx + 1: revBodyArr(revBodyIdx) = dateLine
```

Never assign directly from a COM property to an array slot while `On Error Resume Next` is active — a silent COM error would leave a zero-length string in the array with no indication that the field was missing.

---

## 8. Quick-reference: loop classification in this codebase

| Module | Function | Loop bound | Elements/iter | Classification | Fix applied |
|--------|----------|-----------|--------------|----------------|-------------|
| modReviewExport | `bufferComments` | `wdDoc.Comments.Count` (≤ hundreds) | 7 | **O(n²) — fix required** | Yes |
| modReviewExport | `bufferRevisions` | `wdDoc.Revisions.Count` (≤ thousands) | 6 | **O(n²) — fix required** | Yes |
| modReviewExport | `BuildBookmarkIndexSection` | `docForExport.Bookmarks.Count` (≤ thousands) | 1 (filtered) | **O(n²) — fix required** | Yes |
| modReviewExport | `BuildDocumentTextSection` | `docForExport.Paragraphs.Count` (≤ thousands) | 1–2 | **O(n²) — fix required** | Yes |
| modReviewExport | `BuildFootnotesSection` | `docForExport.Footnotes.Count` (≤ hundreds) | 4 (filtered) | **O(n²) — fix required** | Yes |
| modSelfTest | `RunAllTests` | ~15 hard-coded test cases | 1 | Bounded — acceptable | No |
| modTrainingPipeline | corpus serializer loops | config row count (≤ dozens) | 1–3 | Bounded — acceptable | No |

---

## 9. Pre-flight checklist before committing a new export/serialization function

- [ ] Does any loop append to a string where the loop bound scales with document content?
- [ ] If yes: is the array pre-dimensioned to `count * maxLinesPerIteration`?
- [ ] Is `Join` called exactly **once**, outside the loop?
- [ ] If the loop has a filter: is `ReDim Preserve` guarded by `If fillCount > 0`?
- [ ] Are section markers (`<<..._START>>`, `<<..._END>>`) outside the array?
- [ ] Are COM property reads in temp variables under `On Error Resume Next`, with array writes after the handler is restored?
- [ ] Is the `Join` separator (`vbCrLf` vs. `""`) consistent with whether elements carry their own terminators?
