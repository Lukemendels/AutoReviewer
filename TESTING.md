# Testing AutoReviewer — the twin-and-vectors doctrine

AutoReviewer's deterministic core runs as VBA on an air-gapped work machine
where nothing here can execute it. The testing architecture works around that
with three pieces:

1. **The Python reference twin (`ref/`)** — pure functions, no I/O, mirroring
   the deterministic VBA logic exactly: the content fingerprint
   (`ref/fingerprint.py`), the JSONL line contract (`ref/jsonl_contract.py`),
   and the session-binding gate (`ref/session.py`). The twin is the
   **debuggable source of truth**: it runs under pytest, here, on every change.
2. **The golden vectors (`tests/vectors/`)** — plain ASCII `.txt` files
   generated **from the twin** by `tests/vectors/regenerate.py`. Never
   hand-edit them; a pytest drift-test fails if the committed files diverge
   from what the twin generates.
3. **The VBA self-test harness (`modSelfTest.bas`)** — replays the vectors
   against the VBA transliteration on the work machine, with zero Word
   dependency. If every vector passes, the VBA agrees with the twin.

The contract: **any change to deterministic behavior is made in the twin
first**, proven by pytest, regenerated into vectors, then transliterated into
VBA — and the harness proves the transliteration.

## Running the Python side (this environment / any dev machine)

```bash
python3 -m pytest tests/ -q          # must be green
python3 tests/vectors/regenerate.py  # rebuild vectors after any twin change
```

## Running the VBA side (operator, work machine)

1. Import all `.bas` modules into the AutoReviewer workbook (including
   `modSelfTest.bas`).
2. Copy the repo's `tests\vectors\` folder **next to the workbook** (the
   harness looks for `<workbook folder>\tests\vectors\`).
3. Run `RunAllSelfTests` (Alt+F8).
4. Results land on the **SelfTest** sheet, and the same report is written to
   `selftest_report.txt` next to the workbook.

**The round-trip protocol:** paste the full contents of `selftest_report.txt`
back to the developer/assistant. That report is the only feedback channel
needed — it names every failing case with its expected and actual values, so
fixes can be made against the twin without another exploratory session.
Green (`OVERALL: GREEN`) before the first real document run.

## Vector file format

Tab-separated ASCII; `#` lines and blank lines are skipped. Input and field
columns use this escape rule (decoded by `modSelfTest.VectorUnescape`,
encoded by `vector_escape` in `regenerate.py`):

| Escape | Meaning |
|---|---|
| `\t` `\n` `\r` `\\` | tab, LF, CR, backslash |
| `\uXXXX` | one UTF-16 **code unit** (so an emoji is two consecutive escapes — its surrogate pair — exactly as VBA stores it) |

- `fingerprint_vectors.txt` — `input <TAB> expected_hex16`
- `parser_vectors.txt` — `line <TAB> PASS|REJECT <TAB> reason <TAB>` the seven
  extracted fields. `reason` is empty for PASS, `PARSE` for structural
  failures, else a validation code (`MISSING_BOOKMARK`,
  `REPLACE_REQUIRES_NEW_TEXT`, …) shared verbatim between the twin and VBA.
- `session_vectors.txt` — `name <TAB> token <TAB> payload <TAB> PASS|FAIL
  <TAB> code` with codes `NO_EXPORT_TOKEN | NO_PAYLOAD | META_MISSING |
  TOKEN_MISMATCH | COUNT_MISMATCH`.

## Mirrored semantics (never change one side alone)

- **Fingerprint** walks UTF-16 *code units* (VBA `AscW` semantics), 1-based
  position salt, two 31-bit lanes, Double-safe arithmetic.
- **Trim** strips space/tab/CR/LF both ends (`TrimWs` / `trim_ws`).
- **Duplicate JSON keys:** first occurrence wins.
- **`apply_change`** counts only as a bare `true`/`false` literal.
- **JSON unescape** handles `\\ \" \/ \b \f \n \r \t` and `\uXXXX` (BMP units;
  high+low surrogate pairs combine into one code point). A malformed `\u` (not
  4 hex digits) or any other unknown escape drops the backslash.
- **Meta `count`** must be a plain integer: optional minus, then 1–9 digits.
- **Payload filter** drops blank lines and ``` fences before the gate.
- Lone UTF-16 surrogates are not vectorable (Python can't encode them); the
  twin and vectors avoid them by construction.
