# AutoReviewer V2 — User Guide

Welcome to AutoReviewer V2! This tool allows you to extract a specific reviewer's stylistic preferences from their past redlined documents to create an AI "Persona", and then use that Persona to automatically suggest tracked changes on new draft documents.

This guide walks you through the end-to-end user experience (UX) for setting up the tool, training a new Persona, and running a review.

---

## Getting Started

1. **Open the Workbook**: Open the AutoReviewer `.xlsm` file.
2. **Initialize the Dashboard**: Run the `modDashboardUI.BuildDashboard` macro. (You only need to do this once, or whenever you want to refresh the UI). 
   - This automatically creates your `Config`, `LLM_Changes`, and `Personas` registry sheets (and, on first review run, a `Log` and a `Trace` sheet).
   - It generates the main **Dashboard** sheet, which is split into workflows: **Train Persona**, **Run Review**, and **Respond to Review**.
3. **Set up the shared Serializer (one time)**: AutoReviewer uses two assistants per review — a *hot co-thinker* (per persona, which you train) and a single *cold serializer* shared by everyone. Create one DHSChat Assistant, paste the contents of `TEMPLATE_SKILL_SERIALIZER.md` into its system prompt, then click **Set Serializer URL** on the dashboard and paste its URL. You only do this once.

> **Why two assistants?** The reviewer's *judgment* (hot) and the *mechanical conversion to edits* (cold) are kept in separate conversations so neither corrupts the other, with your ratification sitting between them. This is the reversibility boundary: everything up to it is reversible suggestion; you decide what crosses.

---

## Phase 1: Creating a New Persona (Training)

*Goal: Extract style rules from a batch of redlined documents to create a `SKILL.md` file, which will power your custom DHSChat Assistant.*

### 1. Set the Active Persona
Click **1. Set Active Persona** and type a name for your new persona (e.g., "Legal Reviewer V1"). This sets the active context for all following actions and creates a new row in the `Personas` registry sheet.

### 2. Build the Training Set
You can train from **two kinds of input**, and you can mix them:

**2a. Redline docs (what the reviewer changes).** Documents the reviewer previously redlined or commented on.
1. Click **2a. Add Doc to Corpus (redlines)**.
2. Select a redlined document.
3. An InputBox lists every author who edited or commented; type the **number** of your target reviewer.
4. *Under the hood:* the macro accepts non-target revisions to form a clean baseline, stamps tracking bookmarks, and extracts the target's edits/comments into `[PersonaName]_corpus.jsonl`.

**2b. Finalized exemplars (what good looks like).** Known-good *final* documents — no redlines needed. This is the clean path when your redlines are messy (multiple authors, overlapping edit turns).
1. Click **2b. Add Finalized Exemplar**.
2. Select a finalized `.docx`. The macro captures its clean final text to `[PersonaName]_exemplar_NN.txt` (your source file is not modified).

**Repeat** to build 5–10 inputs total. You can train from redlines only, exemplars only, or both — exemplars are used as the gold standard; redlines as observed edits.

### 3. The "Reduce" Passes (AI Analysis)
Now you will pass the corpus to DHSChat to analyze the patterns.
1. Click **3. Reduce Pass 1: Cluster**. 
   - A prompt is copied to your clipboard.
   - File Explorer automatically opens to the folder containing your `corpus.jsonl`.
   - Open a *new, blank* chat in DHSChat.
   - Drag in your `corpus.jsonl` **and** any `[PersonaName]_exemplar_*.txt` files (whichever you created), paste the prompt, and hit send.
2. Click **4. Reduce Pass 2: Heuristics**.
   - A new prompt is copied.
   - Paste it into the *same* DHSChat conversation and hit send.
3. Click **5. Reduce Pass 3: SKILL.md**.
   - A final prompt is copied.
   - Paste it into the *same* conversation. DHSChat will generate a markdown code block containing your `SKILL.md` file.

### 4. Save and Wire Up the Persona
1. In DHSChat, **copy** the generated `SKILL.md` text to your clipboard.
2. Click **6. Save SKILL.md** on the Excel dashboard. The macro will save the file to your computer and link it in the `Personas` sheet.
3. **Manual Step**: Go to DHSChat and create a new Assistant. Paste the contents of your new `SKILL.md` file into the "System Prompt" or "Instructions" box of the Assistant.
4. Go to the `Personas` sheet in Excel and paste the URL of your new Assistant into the `AssistantUrl` column for your persona.

---

## Phase 2: Using the Tool (Running a Review)

*Goal: Run a new draft document through your Persona and get tracked changes.*

### 1. Select Your Persona
Click **1. Select Persona for Review** and type the name of the persona you want to use (e.g., "Legal Reviewer V1"). 

### 2. Prepare the Document (Hot Co-thinker)
1. Click **2. Prepare for Review (Co-thinker)**.
2. Select the new draft document you want reviewed.
3. *What happens under the hood?* The macro makes a `*_AR` **working copy** (your original is never touched), invisibly stamps every paragraph, table cell, and footnote with a unique `AR_` bookmark, extracts the text, fingerprints the payload for the audit trail, copies the co-thinker prompt to your clipboard, and opens your Persona's co-thinker Assistant.
4. In the browser, **paste** the prompt and **upload** the exported `.txt`. The co-thinker reviews against the persona's style and returns a **decision packet** — each recommendation with its rationale, its *counter-case*, and the bookmark it targets. (It does **not** return JSONL.)

### 3. Ratify (the human seat)
Read the decision packet and decide each item: **keep**, **fix**, or **cut**. Edit the packet down to only the decisions you approve. This is the point past which suggestions become edits — your judgment is the gate.

### 4. Hand off to the Serializer (Cold)
1. Click **3. Hand off to Serializer**.
2. The macro copies the serialize prompt and opens the shared Serializer Assistant.
3. **Paste** the prompt, then **paste your ratified decisions** where indicated. The Serializer converts them — and only them — into a JSONL code block, without re-deciding anything.
4. **Copy** that JSONL output.

### 5. Apply Edits to Word
1. Go to the `LLM_Changes` sheet and **paste** the JSONL starting at cell **A8** (one object per row).
2. On the Dashboard, click **4. Apply LLM Edits to Word**.
3. The macro reads the JSONL, finds the corresponding bookmarks in the working-copy document, and applies the suggestions as native **Tracked Changes** or **Comments**. It logs each edit to the `Log` sheet and records one run row (operator, route, transport fingerprints) to the `Trace` sheet.

### 6. Human Review & Finalize
Switch to Microsoft Word (the `*_AR` working copy). AI **comments** are authored **"AutoReviewer."** AI **insertions/deletions** also try to author as "AutoReviewer," but if Word is signed into a DHS/Microsoft 365 account, Word stamps the *account* name on revisions instead — to force "AutoReviewer" on insertions too, check once: **Word → Options → General → "Always use these values regardless of sign in to Office."** The apply summary tells you which author actually got stamped. Either way, edits are surgical (only the changed words are tracked) and fully rejectable — use the standard **Accept** / **Reject** buttons, then finalize the document yourself. The irreversible step always stays with you. (The hidden `AR_` anchors are removed automatically.)

> **Respond to Review (incorporating supervisor edits):** the Incorporator leads with a **synthesis brief** — Direction, Themes, and Open Questions — so you can grasp the whole of the feedback before adjudicating item by item, and it drafts a reply to *every* reviewer comment. The apply summary flags any comment left unaddressed.
