---
type: Spec
title: StickShift HTML Tool Compliance Standard
description: The authoring standard for a StickShift-compliant HTML tool. Defines the minimum a single self-contained .html must contain to slot into StickShift - client-side operation, an embedded companion skill, an identity declaration, and a StickShift onboarding panel. Hand this to a builder (a person or a coding agent) to produce a conformant tool.
status: spec
okf_version: "0.1"
---

# StickShift HTML Tool Compliance Standard

## 0. Who this is for

You are building an HTML tool that plugs into StickShift. This document is the contract: meet it
and the tool installs and opens through StickShift with no glue code. It is self-contained - you
do not need to read the StickShift host internals to build a conformant tool. (The host side -
how StickShift stores, installs, and launches tools - is `_meta/spec-html-tools.md`; you only
need it for context.)

## 1. What StickShift is (context for the builder)

StickShift is an Excel cockpit that gives an AI assistant (DHSChat, running GPT-5.1) a
persistent, file-based memory and lets it launch local browser tools. The assistant reads and
writes a tree of markdown "concepts," and a "skill" is one of those concepts that teaches the
assistant a capability. An HTML tool is a local page the assistant can tell the operator to open;
the tool does work the assistant cannot do in chat (e.g. rewrite a .docx). The tool runs entirely
on the operator's machine.

Two facts drive every rule below:

- The tool is opened from `file://` (a local file in the default browser). It cannot rely on a
  server, a build step, or a secure browser context.
- A local file cannot be opened from a chat hyperlink, so the assistant launches the tool by
  emitting a small text block (`<HTML_OPEN>`) that StickShift's VBA actions on. Installation is
  likewise a StickShift/VBA filesystem operation - the assistant never writes files.

## 2. The minimum (compliance checklist)

A conformant tool is **one self-contained `.html` file** that has all of:

1. **Self-contained + offline.** No external network calls, no server, no build. All CSS and JS
   inline (or vendored inline). It must work opened directly from `file://`.
2. **Client-side I/O only.** Input via file picker / drag-drop; output via download (Blob +
   `<a download>`). Nothing is uploaded; the operator's data never leaves the machine.
3. **Embedded companion skill.** The exact `skill.md` text, in a parseable block (section 5), so
   StickShift can install it and the assistant can surface the tool.
4. **Identity declaration.** The tool's own filename and skill slug, declared once and consistent
   with the embedded skill (section 6).
5. **StickShift onboarding panel.** A non-blocking entry point that intros StickShift to a cold
   recipient, branches on "Are you using StickShift?", and offers a copy-to-clipboard of the
   skill with a `file://`-safe fallback (section 7).

Everything else (the tool's actual function and UI) is yours.

## 3. Required: self-contained and offline

- Single `.html`. Inline `<style>` and `<script>`; no `src`/`href` to external origins. If you
  need a library, paste it inline. (DHS environments are effectively airgapped; a CDN reference
  will fail.)
- No `fetch`/`XMLHttpRequest` to the network. The only I/O is local files (in) and downloads
  (out).
- Must run when double-clicked or launched by StickShift - both are `file://`.

## 4. Required: client-side I/O

- **Input:** `<input type="file">` and/or drag-drop onto a dropzone. (File input works on
  `file://`; this is the reliable path.)
- **Output:** build the result as a `Blob` and trigger a download via a temporary
  `<a download="name.ext">` whose `href` is `URL.createObjectURL(blob)`. (Downloads work on
  `file://`.)
- **No upload, no telemetry, no network.** State this in the UI so the operator trusts it.
- Document-manipulation tools (e.g. tracked changes) operate on the file bytes directly. Tracked
  changes are OOXML markup inside the `.docx` zip; manufacture them client-side (e.g. with a zip
  library inlined) - no Word, no COM, no server.

## 5. Required: the embedded companion skill

The tool carries its skill so a single file is installable and the assistant can find it.

**Embed block.** Put the skill markdown verbatim inside a non-rendering script tag with this
exact id, plus a `data-skill-slug` attribute:

```html
<script type="text/markdown" id="stickshift-skill" data-skill-slug="autoreviewer-tool">
---
okf_version: "0.1"
type: Skill
title: AutoReviewer (Tracked-Changes Tool)
description: Produces a Word document with real tracked changes from a review, using a local browser tool that writes OOXML markup directly (no Word, no COM). Use this skill when the operator wants tracked changes or redlines in a .docx, to mark up a document, or to turn review feedback into suggested edits.
tags: [skill, html-tool, document-review]
---

# autoreviewer-tool

## Purpose
Turn review feedback into a .docx with genuine tracked changes, via this local browser tool.
The tool runs on the operator's machine, takes their document, and downloads a marked-up copy.

## When to use this skill
- The operator wants tracked changes / redlines in a Word document.
- The operator has review feedback to apply as suggestions rather than as prose.

## How to open the tool
A local HTML tool cannot be opened from a chat hyperlink - the operator launches it from
StickShift. Give the operator this block verbatim (these exact lines, no code fence), then one
instruction line, and nothing else:

<HTML_OPEN>
tool: autoreviewer.html
include:
- skills/autoreviewer-tool.md
</HTML_OPEN>

Instruction line: "Copy the block above and click Open HTML Tool in StickShift."

## Then walk them through it
Once open: drop the .docx in, paste the review block when prompted, download the marked-up copy.
Proceed with the review content normally.
</script>
```

Rules for the embedded skill:

- **Format:** StickShift internal skill format - frontmatter with `type: Skill`, `title`,
  `description`, `tags` (no `status`; tool skills file flat-alphabetical). The `description` must
  be keyword-rich (what it does + "Use this skill when...") so GPT-5.1 selects it.
- **It must instruct the assistant to emit the tool's `<HTML_OPEN>` block** (section 8 of the host
  spec defines the grammar): `tool:` = this tool's filename; optional `include:` = this skill's
  path under `skills/` so a fresh chat can be primed. Emit only that block plus a one-line
  instruction.
- **`tool:` names the filename, never a path.** StickShift resolves it against the local `-html`
  folder, which keeps the tool portable across machines.
- The block is extracted by StickShift's `Install HTML Tool` (it reads `#stickshift-skill` and
  writes the inner markdown to `skills/<data-skill-slug>.md`), and by the onboarding panel's Copy
  button (same source). One source of truth.

## 6. Required: identity declaration

Declare the tool's identity once so the embedded skill, the panel, and the install all agree:

```html
<script>
  const STICKSHIFT_TOOL = {
    file: "autoreviewer.html",      // this tool's own filename
    skillSlug: "autoreviewer-tool", // matches data-skill-slug above
    title: "AutoReviewer"
  };
</script>
```

`STICKSHIFT_TOOL.file` and `skillSlug` must match the embedded skill's `<HTML_OPEN> tool:` line
and `data-skill-slug`. (If a user renames the file, `Install HTML Tool` rewrites the skill's
`tool:` line to the actual filename on copy, so a rename does not break opening - but ship them
matching.)

## 7. Required: the StickShift onboarding panel

A non-blocking entry point (a small button or banner, e.g. "New here? Set up with StickShift")
that opens a panel. The tool must be fully usable without ever opening it.

The panel contains:

1. **A cold-start intro** (1-2 sentences): what StickShift is and why this tool pairs with it.
   Example: "StickShift is an Excel cockpit that gives your AI assistant a persistent memory and
   lets it launch local tools like this one. This tool also works on its own."
2. **"Are you using StickShift?"** with two buttons:
   - **Yes** -> show install steps:
     1. "In StickShift, click **Install HTML Tool** and pick this file. That copies the tool and
        installs its skill in one step."
     2. A fallback line: "Or copy the skill below and paste it into **Apply Write Envelope**." +
        the **Copy skill** button.
   - **No** -> "StickShift is [where to get it]. You can use this tool standalone right now - just
     [drop your file in]. Come back to set up the integration when you're ready."
3. **Copy button(s)** (section 9 for the `file://` fallback):
   - **Copy skill** - copies the embedded skill markdown (read from `#stickshift-skill`) so the
     operator can paste it into Apply Write Envelope or save it as `skills/<slug>.md`.
   - Optional **Copy install envelope** - copies the skill already wrapped as a `<VBA_WRITE>`
     envelope (`### FILE: skills/<slug>.md` ... `### END FILE`) for a one-paste install via
     Apply Write Envelope.

The panel reads everything from the embedded blocks (sections 5-6); it does not hardcode copies.

## 8. How install actually reaches StickShift (the answer to "VBA or DHSChat")

Installation is **always VBA / filesystem** - the assistant cannot write files.

- **HTML file -> `-html/`:** a file operation. Primary: `Install HTML Tool` (operator picks the
  `.html`; VBA copies it into `-html`). Fallback: the operator saves the file into `-html`
  manually.
- **Skill -> `skills/`:** Primary: `Install HTML Tool` extracts the embedded skill and writes it.
  Fallback: the panel's Copy gives the skill (or a `<VBA_WRITE>` envelope) to paste into
  **Apply Write Envelope**, which writes it through StickShift's normal write path.
- **Opening (later, not install):** the assistant emits the `<HTML_OPEN>` block; the operator
  clicks **Open HTML Tool**; VBA launches the file from `-html` in the default browser and loads
  the skill into `-dist`.

So the panel's payload feeds VBA (directly via Install HTML Tool, or via Apply Write Envelope).
It never feeds DHSChat for installation. DHSChat's only role is emitting `<HTML_OPEN>` to open an
already-installed tool.

## 9. Clipboard on file:// (do not skip)

A page on `file://` often cannot use `navigator.clipboard.writeText` (it requires a secure
context), so copy must fall back to a hidden textarea + `document.execCommand('copy')`:

```js
function copyText(text) {
  if (navigator.clipboard && window.isSecureContext) {
    navigator.clipboard.writeText(text).catch(() => legacyCopy(text));
  } else {
    legacyCopy(text);
  }
}
function legacyCopy(text) {
  const ta = document.createElement('textarea');
  ta.value = text;
  ta.style.position = 'fixed';
  ta.style.opacity = '0';
  document.body.appendChild(ta);
  ta.focus(); ta.select();
  try { document.execCommand('copy'); } catch (e) {}
  document.body.removeChild(ta);
}
```

Tools that use only the async API copy nothing on the local-file path and fail silently. Always
include the fallback.

## 10. Reference scaffold (fill in the tool's own function)

A minimal compliant skeleton. The marked region is where the tool's actual UI and logic go;
everything else is the StickShift-compliance scaffolding.

```html
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>AutoReviewer</title>
<style>
  /* inline styles only */
  body { font-family: system-ui, sans-serif; margin: 2rem; }
  .ss-banner { font-size: .9rem; }
  .ss-panel { display: none; border: 1px solid #ccc; padding: 1rem; margin-top: 1rem; max-width: 40rem; }
  .ss-panel.open { display: block; }
  .ss-branch { display: none; }
  .ss-branch.show { display: block; }
</style>
</head>
<body>

  <h1>AutoReviewer</h1>

  <!-- ===== TOOL UI + LOGIC GO HERE ===== -->
  <!-- file input / dropzone, the tool's controls, and a download trigger.
       Pure client-side: read the selected file, produce a Blob, download it. -->
  <!-- =================================== -->

  <!-- StickShift onboarding entry (non-blocking) -->
  <p class="ss-banner">New here? <button id="ss-open">Set up with StickShift</button></p>

  <div class="ss-panel" id="ss-panel">
    <p><strong>What is StickShift?</strong> An Excel cockpit that gives your AI assistant a
       persistent memory and lets it launch local tools like this one. This tool also works on
       its own.</p>
    <p><strong>Are you using StickShift?</strong>
       <button id="ss-yes">Yes</button> <button id="ss-no">No</button></p>

    <div class="ss-branch" id="ss-branch-yes">
      <ol>
        <li>In StickShift, click <strong>Install HTML Tool</strong> and pick this file - it copies
            the tool and installs its skill in one step.</li>
        <li>Or copy the skill and paste it into <strong>Apply Write Envelope</strong>:
            <button id="ss-copy-skill">Copy skill</button></li>
      </ol>
    </div>
    <div class="ss-branch" id="ss-branch-no">
      <p>StickShift is [where to get it]. Use this tool standalone right now; come back to set up
         the integration when you're ready.</p>
    </div>
  </div>

  <!-- Identity declaration -->
  <script>
    const STICKSHIFT_TOOL = { file: "autoreviewer.html", skillSlug: "autoreviewer-tool", title: "AutoReviewer" };
  </script>

  <!-- Embedded companion skill (verbatim; see section 5) -->
  <script type="text/markdown" id="stickshift-skill" data-skill-slug="autoreviewer-tool">
  ...skill markdown here...
  </script>

  <!-- Onboarding panel wiring + file://-safe copy -->
  <script>
    function copyText(text){ /* secure-context check + execCommand fallback (section 9) */ }
    function getSkillMarkdown(){
      return document.getElementById('stickshift-skill').textContent.trim();
    }
    document.getElementById('ss-open').onclick  = () => document.getElementById('ss-panel').classList.toggle('open');
    document.getElementById('ss-yes').onclick   = () => { show('ss-branch-yes'); hide('ss-branch-no'); };
    document.getElementById('ss-no').onclick    = () => { show('ss-branch-no'); hide('ss-branch-yes'); };
    document.getElementById('ss-copy-skill').onclick = () => copyText(getSkillMarkdown());
    function show(id){ document.getElementById(id).classList.add('show'); }
    function hide(id){ document.getElementById(id).classList.remove('show'); }
  </script>

</body>
</html>
```

## 11. Compliance checklist (testable)

- [ ] Single `.html`; opens and runs from `file://`; no external network requests.
- [ ] Input via file picker/drag-drop; output via download; nothing uploaded.
- [ ] `#stickshift-skill` script block present, valid StickShift skill frontmatter, keyword-rich
      `description`, body instructs the assistant to emit the `<HTML_OPEN>` block with the tool's
      filename.
- [ ] `data-skill-slug` set and matching `STICKSHIFT_TOOL.skillSlug`; `STICKSHIFT_TOOL.file`
      matches the tool's filename and the skill's `<HTML_OPEN> tool:` line.
- [ ] Onboarding entry present and non-blocking; panel has the StickShift intro, the Yes/No
      branch, and a Copy that reads from `#stickshift-skill`.
- [ ] Copy uses the `execCommand` fallback for `file://`.

## 12. Build-prompt template (hand to Paul / Antigravity / Codex)

> Build a single self-contained HTML tool that **[DESCRIBE THE TOOL'S FUNCTION - e.g. "takes a
> .docx and a block of review notes and produces a copy with tracked changes applied"]**. It must
> comply with the StickShift HTML Tool Compliance Standard:
>
> - One `.html`, fully offline: all CSS/JS inline, no external network calls, must run from
>   `file://` (double-clicked).
> - All work client-side: input via file picker/drag-drop, output via a Blob download, nothing
>   uploaded. **[Add any function specifics: file types, the transformation, libraries to inline.]**
> - Embed the companion skill verbatim in
>   `<script type="text/markdown" id="stickshift-skill" data-skill-slug="SLUG">...</script>`,
>   using StickShift skill frontmatter (`type: Skill`, `title`, `description`, `tags`; no
>   `status`), a keyword-rich `description`, and a body that tells the assistant to emit this
>   block and nothing else:
>   `<HTML_OPEN>` / `tool: FILENAME.html` / `include:` / `- skills/SLUG.md` / `</HTML_OPEN>`
>   plus the line "Copy the block above and click Open HTML Tool in StickShift."
> - Declare identity: `const STICKSHIFT_TOOL = { file: "FILENAME.html", skillSlug: "SLUG",
>   title: "TITLE" };` matching the embedded skill.
> - Add a non-blocking "Set up with StickShift" entry opening a panel with: a 1-2 sentence
>   StickShift intro, an "Are you using StickShift?" Yes/No branch (Yes -> install steps + a Copy
>   skill button; No -> what it is / standalone use), and a copy function that uses
>   `navigator.clipboard` with a hidden-textarea `execCommand('copy')` fallback (because `file://`
>   blocks the async clipboard API).
> - The tool must be fully usable without opening the panel.
>
> Output only the single `.html` file. Fill `FILENAME`, `SLUG`, and `TITLE` consistently
> throughout.
