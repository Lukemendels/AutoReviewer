---
okf_version: "0.1"
type: Skill
title: AutoReviewer Workbench
description: Runs the full document-review lifecycle for a Word .docx entirely in a local browser tool -- exports tracked changes/comments to CriticMarkup markdown for an LLM to redline, validates the LLM's response byte-for-byte before any edit is trusted, lets a human accept/reject each proposed edit, then writes real Word tracked changes and threaded comments back into the .docx (no Word, no COM, no server). Also drafts replies to existing reviewer comments and accept/reject recommendations (Respond to Review), and builds/exports reviewer-persona files (Train Persona). Use this skill when the operator wants tracked changes or redlines in a .docx, wants an LLM to review or mark up a document, needs to respond to a colleague's comments and redlines, or wants to build/update a reviewer persona.
tags: [skill, html-tool, document-review]
---

# autoreviewer-workbench

## Purpose
Turn a document review into a .docx with genuine Word tracked changes and threaded comments, via
this local browser tool. The tool runs on the operator's machine, takes their document, and
downloads a marked-up copy. Nothing is uploaded; the LLM leg happens over the clipboard (the
operator pastes a prompt into you, pastes your reply back into the tool).

## When to use this skill
- The operator wants tracked changes / redlines in a Word document produced from an LLM review.
- The operator has a document that already contains reviewer comments/redlines and wants drafted
  replies and accept/reject recommendations (Respond to Review).
- The operator wants to build, update, or export a reviewer persona for future reviews (Train
  Persona).

## How to open the tool
A local HTML tool cannot be opened from a chat hyperlink - the operator launches it from
StickShift. Give the operator this block verbatim (these exact lines, no code fence), then one
instruction line, and nothing else:

<HTML_OPEN>
tool: autoreviewer-workbench.html
include:
- skills/autoreviewer-workbench.md
</HTML_OPEN>

Instruction line: "Copy the block above and click Open HTML Tool in StickShift."

## Then walk them through it
Once open, the operator picks one of three flows:

- **Run Review** - drop the .docx and a persona file, copy the composed prompt, paste it to you,
  paste your CriticMarkup reply back into the tool. The tool validates your reply against the
  original document before anything is trusted, then the operator accepts/rejects each edit before
  it is written into the .docx.
- **Respond to Review** - drop a .docx that already has tracked changes/comments; the tool exports
  them as labeled CriticMarkup; you return reply/recommendation lines per label; the operator
  ratifies before the tool writes threaded comment replies back in.
- **Train Persona** - a wizard that produces a reviewer-persona markdown file to save into
  StickShift, so future reviews are voiced and prioritized consistently.

Proceed with the review content normally once the operator has the tool open.
