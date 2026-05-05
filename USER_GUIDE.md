# AutoReviewer V2 — User Guide

Welcome to AutoReviewer V2! This tool allows you to extract a specific reviewer's stylistic preferences from their past redlined documents to create an AI "Persona", and then use that Persona to automatically suggest tracked changes on new draft documents.

This guide walks you through the end-to-end user experience (UX) for setting up the tool, training a new Persona, and running a review.

---

## Getting Started

1. **Open the Workbook**: Open the AutoReviewer `.xlsm` file.
2. **Initialize the Dashboard**: Run the `modDashboardUI.BuildDashboard` macro. (You only need to do this once, or whenever you want to refresh the UI). 
   - This automatically creates your `Config`, `LLM_Changes`, and `Personas` registry sheets.
   - It generates the main **Dashboard** sheet, which is split into two workflows: **Train Persona** and **Run Review**.

---

## Phase 1: Creating a New Persona (Training)

*Goal: Extract style rules from a batch of redlined documents to create a `SKILL.md` file, which will power your custom DHSChat Assistant.*

### 1. Set the Active Persona
Click **1. Set Active Persona** and type a name for your new persona (e.g., "Legal Reviewer V1"). This sets the active context for all following actions and creates a new row in the `Personas` registry sheet.

### 2. Build the Training Corpus
You need 5-10 Word documents (.docx) that your target reviewer has previously redlined or commented on.
1. Click **2. Add Doc to Corpus**.
2. Select one of your redlined documents.
3. Word will open the document, and an InputBox will appear listing all the authors who made edits or comments in that document.
4. Type the **number** corresponding to your target reviewer.
5. *What happens under the hood?* The macro automatically accepts all non-target revisions to create a clean baseline, stamps hidden tracking bookmarks, and extracts the target reviewer's edits/comments into a file named `[PersonaName]_corpus.jsonl`.
6. **Repeat this step** for all 5-10 documents to build a robust corpus.

### 3. The "Reduce" Passes (AI Analysis)
Now you will pass the corpus to DHSChat to analyze the patterns.
1. Click **3. Reduce Pass 1: Cluster**. 
   - A prompt is copied to your clipboard.
   - File Explorer automatically opens to the folder containing your `corpus.jsonl`.
   - Open a *new, blank* chat in DHSChat.
   - Drag and drop your `corpus.jsonl` file into the chat, paste the prompt, and hit send.
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

### 2. Prepare the Document
1. Click **2. Prepare Document for LLM**.
2. Select the new draft document you want reviewed.
3. *What happens under the hood?* The macro invisibly stamps every paragraph, table cell, and footnote with a unique `MKS_` bookmark. It then extracts all the text, copies a prompt to your clipboard, and automatically opens your Persona's Assistant URL in your web browser.

### 3. Generate Edits in DHSChat
1. In your browser (which should now be open to your Assistant), **paste** your clipboard contents and hit send.
2. The Assistant will review the text against its style guidelines and reply with a raw JSONL code block containing targeted edits.
3. **Copy** that JSONL output.

### 4. Apply Edits to Word
1. Go to the `LLM_Changes` sheet in Excel.
2. **Paste** the JSONL output starting at cell **A8** (one JSON object per row).
3. Go back to the Dashboard and click **3. Apply LLM Edits to Word**.
4. The macro will read the JSONL, find the corresponding bookmarks in your open Word document, and apply the AI's suggestions as native **Tracked Changes** or **Comments**.

### 5. Human Review
Switch over to Microsoft Word. You can now use the standard "Accept" and "Reject" buttons in the Review ribbon to finalize the document, just as if your boss had edited it themselves!
