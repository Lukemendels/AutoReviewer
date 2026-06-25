# AutoReviewer — Sandbox HTML Tool

An offline, sandboxed browser utility for Microsoft Word document reviews, built using the **StickShift dual-life HTML pattern**. It requires zero network connectivity, zero installation, and runs 100% locally.

## Operational Flow

AutoReviewer handles document reviews in two local passes:

```
  [user drops draft.docx into the tool]
        │
   PASS 1 "Prepare"  ── parse + anchor + extract ──▶ markdown payload (to clipboard)
        │                                                    │
        │                                          [LLM: reviews & generates JSONL edits]
        │                                                    │
   PASS 2 "Apply"  ◀── paste JSONL back into same tool ──────┘
        │
        ├──▶ reviewed.docx  (tracked changes baked into the XML — download)
        └──▶ markdown receipt (counts + unaddressed comments + fingerprints)
```

1. **Prepare Pass**: Upload your `.docx` file. The tool parses its XML structure in-memory and outputs a Markdown representation of the text, footnotes, comments, and existing tracked changes. It embeds **virtual anchor IDs** (like `[AR_PARA_00037]`) in the text. Copy this payload to your clipboard for your LLM.
2. **Apply Pass**: Paste the LLM's JSONL suggestions back into the tool. It validates the suggestions, runs comment coverage checks, performs minimal-diff XML surgery (run-splitting to preserve formatting like bold/italic), packages the revised document, and initiates a download of the redlined `.docx`. An audit receipt is also generated.

## Features

- **No Data Leakage**: Safe for sensitive documents. Since it runs client-side inside a web browser sandbox, no data is ever transmitted over the network.
- **Run-Splitting XML Surgery**: Edits are applied as native Word tracked changes (`<w:ins>` / `<w:del>`) and targeted to the precise word boundaries to emulate human redlines.
- **Comments and Footnotes**: Fully supports adding comments, threaded replies (supporting modern `commentsExtended.xml`), and inserting footnotes.
- **StickShift Cargo**: Features a quiet footer affordance (`⚙ Part of StickShift ▸`) that opens a Welcome overlay containing the copyable `skill.md` contract and `<VBA_WRITE>` self-registration payload.

## Usage

1. Open [autoreviewer.html](file:///home/luke/AutoReviewer/autoreviewer.html) in any modern browser (Chrome, Edge, Firefox, or Safari).
2. Drag and drop your `.docx` file to prepare it.
3. Paste the generated payload into your LLM chat.
4. Copy the LLM's JSONL output block and paste it into the "Apply Suggestions" tab.
5. Click **Apply & Download docx**.

## Diagnostics & Testing

You can run the inline unit test suite (testing JSONL parsing, unescaping, validation, and XML run-splitting) by:
- Clicking **Run System Diagnostics** in the *Welcome to StickShift* screen, or
- Appending `?test=true` to the URL when opening the file in your browser.