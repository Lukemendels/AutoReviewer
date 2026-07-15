# Sentinel Encoding for Pre-Existing Annotations (M6a)

We define a sentinel encoding format to model pre-existing tracked changes and comments in the exported markdown, preventing them from being parsed as new model edits.

## Sentinel Format

To prevent collisions with standard CriticMarkup delimiters and satisfy validation gates, sentinels are wrapped in double-bracket unicode characters `⟦` and `⟧`, which are already utilized in the codebase to mark locked fields and placeholders.

The format for each pre-existing annotation type is defined as follows:

### 1. Pre-Existing Comments (`[C1]` to `[Cn]`)
```markdown
⟦C{id}: {highlightedText} >> {rootComment} ↳ {reply1} ↳ {reply2} ...⟧
```
Each comment/reply inside the thread is formatted as:
`Author (Date) [resolved]: text` (with `[resolved]` present only if the comment is resolved/done).
- **Example:** `⟦C1: Change 17 >> Pitino, Salvatore (2025-07-25): ASD: placeholder... ↳ Luke (2025-07-26) [resolved]: approved⟧`

### 2. Pre-Existing Tracked Changes (`[R1]` to `[Rn]`)
- **Insertion:** `⟦R{id}: +{insertedText}+⟧`
- **Deletion:** `⟦R{id}: -{deletedText}-⟧`
- **Substitution:** `⟦R{id}: ~{deletedText}~>{insertedText}~⟧`

## Validation and Fidelity (G2) Semantics

- **Tokenizer / Parser (`grammar.js`):** Because the sentinels do not contain the CriticMarkup openers `{++`, `{--`, `{~~`, `{==`, or `{>>` (replacing comment headers with `>>` and `↳`), the tokenizer does not identify any tokens inside sentinels. Consequently, `parseEdits()` returns **zero** edits for a byte-perfect echo.
- **Fidelity Gate (`strip.js`):** Since no CriticMarkup tokens are parsed, `strip(response)` leaves sentinels completely untouched. Thus, a byte-perfect echo of the exported document matches `exportedMarkdown` character-for-character, passing `G2` successfully.
- **Protection (`validate.js`):** Sentinel ranges are marked as `locked` and `synthetic` in the source map. Any newly authored CriticMarkup edit that overlaps or sits inside a sentinel is caught by `overlapsLocked` and rejected at `G4`.

## Escaping Rules

To avoid collisions with actual document text:
- Any literal occurrence of `⟦` or `⟧` in the raw document text is escaped during export by prepending a backslash (`\⟦` and `\⟧`).
- When validating or injecting, the backslashes are resolved to the original literal brackets.
