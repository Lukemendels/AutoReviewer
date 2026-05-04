# PreFlight Reviewer — V1 Scope

## What this is

A two-mode Word document review tool driven from an Excel dashboard. **Training mode** extracts a target reviewer's implicit style preferences from a corpus of redlined documents and produces a SKILL.md to seed a DHSChat Assistant. **Review mode** uses that Assistant to run synthetic reviews on new documents, returning bookmark-targeted edits as JSONL that VBA applies back to Word as tracked changes.

The reviewer persona is portable. One Assistant URL per persona. Personas are named by role / document type, not by person.

## Architecture

### Training mode (rare, offline)

```
N redlined .docx files in a folder
            ↓
[Map: per-doc, automated VBA]
   for each doc:
     show all revision authors → user picks target
     filter: target's revisions/comments only
     baseline: pre-target non-target revisions accepted
     stamp bookmarks on baseline
     extract structured records → append to corpus.jsonl
            ↓
[Reduce: human-in-the-loop, multi-pass through DHSChat]
   Pass 1 — Cluster revisions into pattern categories
   Pass 2 — Extract heuristic per category
   Pass 3 — Synthesize into SKILL.md (style + JSON output contract)
            ↓
SKILL.md → paste into new DHSChat Assistant system prompt
Persona registered (URL + corpus + SKILL.md) in workbook
```

### Review mode (per-document, routine)

Existing pipeline, mostly unchanged:

```
User picks Word doc
   ↓
ExportWordDocForLLM (existing) — stamp bookmarks, generate <<DOCUMENT_TEXT>> + <<BOOKMARK_INDEX>> + <<FOOTNOTES>> + <<COMMENTS>>, copy prompt to clipboard, launch persona's Assistant URL
   ↓
User pastes export into DHSChat Assistant → JSONL of edits
   ↓
User pastes JSONL into LLM_Changes sheet
   ↓
ApplyWordSuggestionsFromJson (existing) — applies as tracked changes
```

## Done (existing code, reusable as-is)

- `modWordStamping.bas` — paragraph / table cell / footnote bookmark stamping with stable `MKS_*` IDs
- `modWordExport.bas` — DOCUMENT_TEXT + BOOKMARK_INDEX + FOOTNOTES + COMMENTS export, FINAL-copy trick to give LLM clean text without losing original-doc edit anchors
- `InputEditsIntoWord.bas` — JSONL → Word edits as tracked changes, three change types (`replace_text`, `delete_element`, `add_comment_only`), Log sheet, confidence handling
- `modPromptHelpers.bas` — clipboard + URL launch via WinAPI, confirmed working in TSA env
- `ConfigHelpers.bas` — key-value config sheet
- `modBookmarkTest.bas` — diagnostic, keep
- `modSetup.bas` — workbook initialization

## V1 — to build

### Net new modules

**`modPersonaRegistry.bas`** — manages a `Personas` sheet with columns: PersonaName, AssistantUrl, CorpusPath, SkillMdPath, TrainingDocCount, LastUpdated, Notes. Persona is the unit of switching: select a persona, all review actions use that Assistant URL.

**`modAuthorFilter.bas`** — given a Word doc, returns list of unique revision authors and comment authors with revision counts. Used in the map step to populate "who is the target?" picker.

**`modTrainingCorpusBuilder.bas`** — the map step. Per doc:
- Open doc
- Enumerate all revision authors → show picker → user selects target
- Compute target's date range (earliest to latest target revision)
- Build baseline: accept all non-target revisions dated before target's earliest
- Stamp bookmarks on baseline
- For each target revision and target comment, emit one record:
  ```json
  {
    "doc_id": "...",
    "target_author": "Jane Doe",
    "record_type": "revision" | "comment",
    "revision_subtype": "insertion" | "deletion" | null,
    "bookmark_id": "MKS_PARA_00037",
    "section_heading": "...",
    "context_sentence": "...",
    "original_text": "...",
    "changed_text": "...",
    "comment_body": "..." | null,
    "date": "..."
  }
  ```
- Append to persona's `corpus.jsonl`
- Update persona registry: increment training doc count, log doc path

**`modTrainingOrchestrator.bas`** — drives the reduce passes. Corpus is delivered as a file attachment, not via clipboard (too large). Each pass:
- "Reduce Pass 1: Cluster" — copies clustering prompt to clipboard, opens DHSChat (fresh chat, no Assistant), opens Explorer at `corpus.jsonl` for user to drag-attach
- "Reduce Pass 2: Extract Heuristics" — user pastes Pass 1 output back into Excel; system writes it to `pass1_clusters.txt`, copies heuristic prompt, opens chat + Explorer at the new file
- "Reduce Pass 3: Synthesize SKILL.md" — same pattern, attaches `pass2_heuristics.txt`, copies synthesis prompt
- "Save SKILL.md to Persona" — user pastes final SKILL.md back into Excel; system saves to `<persona>/SKILL.md` and surfaces next step (manually create a new DHSChat Assistant, paste SKILL.md into its system prompt, save URL to registry)

### Modifications to existing modules

- `ConfigHelpers.bas`: add `ActivePersona` key. `CustomGptUrl` becomes per-persona, looked up from registry, not from Config.
- `modWordExport.bas`: read Assistant URL from active persona's registry entry, not from Config directly.
- `modDashboardUI.bas`: redesign. Two top-level groups: **Train New Persona** (4 buttons: Add Doc to Corpus, Run Reduce Passes, Save SKILL.md) and **Run Review** (Persona selector dropdown, Prepare Document, Apply LLM Edits). Settings button stays.
- `modRuleExtractor.bas`: **deprecate.** Its functionality is absorbed into `modTrainingCorpusBuilder` with the proper filtering and per-doc author selection.

### Cut entirely

- `modAutoReview.bas` (all three subs — `AutoReview_Start`, `AutoReview_ApplyFromJsonFile`, `AutoReview_ApplyFromClipboard`). The PowerShell automation branch attempts to skip the manual paste step. The manual paste *is* the architecture given no DHSChat API. Hardcoded user paths, brittle, restart-trap evidence.

### Decisions to confirm during build

- **Author filtering UX.** Per-doc picker showing all authors with revision/comment counts. Default selection = most frequent non-drafter author. User can override. Confirms before adding to corpus.
- **Author identity.** MS365 username (FirstName LastName) is the primary key. No fallback logic needed for V1 — if a doc has weird author strings, exclude it from the training set.
- **Serial review assumption.** Baseline = non-target revisions dated before target's earliest revision. If target has multiple disjoint revision date clusters in one doc, warn the user and recommend exclusion. Not handled in code.
- **Comments tagged separately from edits.** Both go in corpus, `record_type` field distinguishes. Reduce passes can weight comments differently (often carry the *reasoning*; edits carry the *action*).
- **SKILL.md is the full Assistant system prompt.** Includes style heuristics + the bookmark-targeted JSONL output contract + change-type vocabulary + confidence guidance. Not split.
- **Persona naming = role/document type, not person.** "TSA Economic Memo Review v1" not "Jim's Style." Durable across personnel, shareable, sidesteps surveillance perception. Hard rule.

## Definition of done

A reviewer who has never used the tool can:

1. Drop 5-10 redlined .docx files in a folder
2. Run training: select target author per doc, run three reduce passes, get a SKILL.md
3. Paste the SKILL.md into a new DHSChat Assistant, save the URL into the persona registry
4. Run review on a new document
5. Open the result in Word, see the synthetic reviewer's edits as tracked changes, accept/reject them in their normal Word flow

End-to-end, real document, real persona, real reviewer-style output. If the synthetic edits are recognizable as *that reviewer's style*, V1 ships.

## Open questions to resolve before / during V1

- **How many training docs is "enough"?** Heuristic: stop when pass-2 heuristics stabilize across runs. Probably 5-10 for a distinctive reviewer. Empirical, fuzzy. May need a "compare two SKILL.md drafts" tool in V2.
- **Reduce pass corpus delivery.** Corpus exceeds the 10K-char chat input limit, so it gets attached as a *file* to the Assistant chat. Not pasted. Not loaded into a Workspace — Workspaces are RAG with 700-token chunking that would fragment per-revision records. Practical ceiling per pass: ~80K tokens of context (~320K chars at 4:1), leaving headroom for reasoning. For 5-10 training docs that's comfortable. If corpus grows beyond this, options: sample, batch reduce passes across multiple sessions, or hierarchical reduce (per-doc heuristics first, then meta-heuristics from those).
- **DHSChat capacity model (reference).** Chat input box: 10K chars. Assistant system prompt: 100K chars (SKILL.md fits comfortably). Context window: 125K tokens, ~80K practical with reasoning headroom. Workspaces: RAG only, 700-token chunks, not used by this tool.
- **Persona switching mid-session.** Dropdown on dashboard. Switching active persona updates `ActivePersona` config key, all subsequent review actions use that persona's Assistant URL.
- **Persona portability across users.** A persona = registry row + corpus.jsonl + SKILL.md + Assistant URL. Sharing a persona means sharing those four things and the recipient creating their own DHSChat Assistant from the SKILL.md. Document the export/import flow.

## V2 — explicitly parked

- **Holdout testing.** Reserve 1-2 docs from the training set as holdout. After SKILL.md is generated, run review mode on holdout, compare synthetic edits to actual target edits, score agreement. Closes the feedback loop on persona quality.
- **SKILL.md iteration / retrain.** "Add new docs and retrain" workflow. Versioned SKILL.md per persona. Lets a persona improve over time without losing prior state.
- **Multi-target personas.** Train on multiple reviewers' edits to extract a *house style* rather than an individual style. Useful for org-wide standards.
- **Reduce pass automation across DHSChat sessions.** If DHSChat ever exposes an API or scriptable interface, the reduce orchestrator can drive it directly. Until then, manual paste.
- **Disjoint revision cluster handling.** Properly handle docs where target reviewed twice with non-target edits in between.
- **Comment-only training.** Train a persona purely from comments (justification only, no edits). Useful when target reviewed via comments without making direct edits.
- **Section-aware heuristics.** Reduce passes that detect "executive summaries get this treatment, methodology sections get that treatment" rather than flat document-wide rules.

## V3 — further out

- **Persona marketplace.** Internal sharing of personas across teams. "Use the Office of General Counsel review style on this draft."
- **Style diff between two personas.** Generate a doc reviewed by Persona A vs. Persona B, surface where they disagree.
- **Confidence-tuned application.** When applying JSONL edits, optionally only auto-apply High confidence; mark Medium for review; skip Low. Currently applied uniformly as tracked changes.
- **Office Scripts port for the export step** if a Power Automate-driven distribution path becomes desirable.
- **In-Word ribbon add-in.** Skip the Excel chassis entirely for end users, surface review actions in Word directly. Significant scope, last priority.

## Anti-scope — V1 will not include

- **Auto-applying edits without track changes.** Always tracked, always human-reviewed before accept.
- **PowerShell automation.** No bridge scripts. Manual paste is the architecture.
- **Auto-detection of target author.** User picks per doc, every time. No magic.
- **Multiple targets per doc.** One target per doc in the training set. If a doc has two reviewers worth training on, it gets processed twice — once per target.
- **Live retraining.** SKILL.md generation is a deliberate action. Adding a doc to the corpus does not automatically regenerate the SKILL.md.
- **Non-Word inputs.** .docx only. No PDFs, no Google Docs.
- **Editing the Assistant from VBA.** The DHSChat Assistant is created and updated manually in the DHSChat UI. VBA only stores the URL.
- **Anything currently in `modAutoReview.bas`.** Cut on day one.

## Distribution unit

For each user / team:
- The `.xlsm` (with all VBA modules)
- A folder of personas (each: corpus.jsonl + SKILL.md + Assistant URL noted in registry)
- The persona registry sheet (synced via the workbook)

Per-user: each user creates their own DHSChat Assistants from the SKILL.md files. Assistant URLs are personal. Persona definitions (corpus + SKILL.md) are shareable.

## File-level summary of changes

| File | Status | Notes |
|---|---|---|
| `modWordStamping.bas` | Keep | Bookmark stamping, no changes |
| `modWordExport.bas` | Modify | Read Assistant URL from active persona |
| `InputEditsIntoWord.bas` | Keep | Already bookmark-only, applies as tracked changes |
| `modPromptHelpers.bas` | Keep | Clipboard + URL working |
| `ConfigHelpers.bas` | Modify | Add `ActivePersona` key |
| `modSetup.bas` | Modify | Initialize Personas sheet |
| `modBookmarkTest.bas` | Keep | Diagnostic |
| `modDashboardUI.bas` | Rewrite | Two-mode dashboard |
| `modRuleExtractor.bas` | Deprecate | Replaced by `modTrainingCorpusBuilder` |
| `modAutoReview.bas` | **Delete** | PowerShell branch, restart-trap evidence |
| `ExportToJSON.bas` | Move out | Generic table-to-JSONL utility, not part of this tool |
| `modPersonaRegistry.bas` | **New** | Persona CRUD on Personas sheet |
| `modAuthorFilter.bas` | **New** | Enumerate authors in a doc, support target selection |
| `modTrainingCorpusBuilder.bas` | **New** | Map step: per-doc → corpus.jsonl |
| `modTrainingOrchestrator.bas` | **New** | Drive the three reduce passes |
