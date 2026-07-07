// @vitest-environment happy-dom
// End-to-end DOM smoke test for M4c's chunk-mode multi-paste flow (spec §6.4, architecture
// doc §7): load an over-threshold, 3-top-level-heading document -> chunk mode auto-entered
// -> copy/paste/validate 3 times, one "Part i of 3" at a time -> RATIFYING with every
// chunk's edits merged and resolved against the full document. Same "drive the real
// renderShell() DOM" style as tests/app.uiFlow.test.js.
//
// No committed fixture is both over CHUNK_WORD_THRESHOLD words AND has more than one
// top-level heading, so the oversized document here is built by hand: take a real, small
// fixture's ZIP container (for a valid [Content_Types].xml/rels/etc.) and swap in a
// hand-written word/document.xml with 3 "Heading1"-styled sections and enough prose to
// clear the word threshold. export.js reads a paragraph's style directly off its own
// w:pStyle/@w:val (no styles.xml lookup), so this needs no styles.xml entry at all.
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { DOMParser as XmldomDOMParser } from "@xmldom/xmldom";
import { describe, expect, it } from "vitest";
import { renderShell } from "../src/ui/app.js";
import { unzip, readEntry } from "../src/zip/reader.js";
import { writeZip } from "../src/zip/writer.js";
import { CHUNK_WORD_THRESHOLD } from "../src/prompt.js";

// See tests/app.uiFlow.test.js's identical header comment for why this shadow is needed:
// happy-dom's own DOMParser silently drops the body of a real, namespace-heavy document.xml.
globalThis.DOMParser = XmldomDOMParser;

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const fixturesDir = path.join(root, "fixtures");

function waitForCondition(check, { timeout = 4000, interval = 20 } = {}) {
  return new Promise((resolve, reject) => {
    const start = Date.now();
    (function poll() {
      if (check()) return resolve();
      if (Date.now() - start > timeout) return reject(new Error("waitForCondition: timed out"));
      setTimeout(poll, interval);
    })();
  });
}

function paraXml(text) {
  return `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}
function headingXml(text) {
  return `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
}

// Enough paragraphs per section, at ~18 words each, to clear CHUNK_WORD_THRESHOLD (12000)
// across 3 sections combined -- 250 * 18 * 3 ~= 13500 words.
const PARAGRAPHS_PER_SECTION = 250;
const SECTIONS = ["Alpha Section", "Bravo Section", "Charlie Section"];

async function buildBigMultiSectionDocxBytes() {
  const buf = readFileSync(path.join(fixturesDir, "plain-paragraphs.docx"));
  const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  const zip = await unzip(ab);
  const orig = await readEntry(zip, "word/document.xml");

  const rootOpen = orig.slice(0, orig.indexOf(">", orig.indexOf("<w:document")) + 1);
  const tail = orig.slice(orig.indexOf("<w:sectPr")); // "<w:sectPr...>...</w:sectPr></w:body></w:document>"

  let body = "";
  for (const heading of SECTIONS) {
    body += headingXml(heading);
    for (let i = 0; i < PARAGRAPHS_PER_SECTION; i++) {
      body += paraXml(
        `Paragraph ${i}: this regulatory text discusses cost benefit considerations, compliance timelines, and carrier obligations under the proposed rule today.`
      );
    }
  }

  const newXml = rootOpen + "<w:body>" + body + tail;
  return writeZip(zip, { "word/document.xml": newXml });
}

// Extracts the exact chunk-local exportedMarkdown embedded in a rendered prompt's own
// [DOCUMENT] section -- buildDocumentSection() wraps it verbatim in a single fenced block,
// so this is the exact text (and only text) the user is meant to paste back, edits aside.
function extractDocumentSection(promptText) {
  const marker = "```markdown\n";
  const start = promptText.indexOf(marker) + marker.length;
  const end = promptText.lastIndexOf("\n```");
  return promptText.slice(start, end);
}

// Splices a unique substring `target` to `replacement`, failing loudly if it isn't found
// exactly once (same convention as validate.test.js's withEdit).
function withEdit(markdown, target, replacement) {
  const first = markdown.indexOf(target);
  const last = markdown.lastIndexOf(target);
  if (first === -1) throw new Error(`withEdit: target not found: ${JSON.stringify(target)}`);
  if (first !== last) throw new Error(`withEdit: target not unique: ${JSON.stringify(target)}`);
  return markdown.slice(0, first) + replacement + markdown.slice(first + target.length);
}

async function dropFile(panel, file) {
  const dropzone = panel.querySelector("#ar-dropzone");
  const evt = new Event("drop", { bubbles: true, cancelable: true });
  evt.dataTransfer = { files: [file] };
  dropzone.dispatchEvent(evt);
}

describe("Run Review panel: chunk mode drives a 3-part document to RATIFYING with merged edits", () => {
  it("goes Part 1 of 3 -> Part 2 of 3 -> Part 3 of 3 -> RATIFYING, one edit per chunk merged", async () => {
    const bigDocxBytes = await buildBigMultiSectionDocxBytes();
    const file = new File([bigDocxBytes], "big-multi-section.docx", {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });

    document.body.innerHTML = '<div id="ar-app"></div>';
    renderShell(document.getElementById("ar-app"));
    const panel = document.querySelector('.ar-panel[data-flow="run-review"]');

    await dropFile(panel, file);
    // Unlike the small demo document (tests/app.uiFlow.test.js), handleDocFiles sets a
    // transient "Reading ..." status synchronously before its first await -- checking for
    // ANY truthy #ar-load-status text (as that test does) would resolve immediately on
    // that transient message instead of waiting for the load to actually finish. Wait for
    // the real terminal signal (the prompt textarea) only; swallow a timeout here so the
    // explicit check right below can report the actual status text.
    await waitForCondition(() => !!panel.querySelector(".ar-prompt-text"), { timeout: 8000 }).catch(() => {});

    let promptTextarea = panel.querySelector(".ar-prompt-text");
    if (!promptTextarea) throw new Error("load failed: " + panel.querySelector("#ar-load-status")?.textContent);

    // Sanity: this document really is over threshold and chunk mode is what we're testing.
    expect(promptTextarea.value).toMatch(/This is part 1 of 3 of a larger document\./);
    expect(panel.textContent).toContain("Part 1 of 3");

    for (let i = 0; i < SECTIONS.length; i++) {
      promptTextarea = panel.querySelector(".ar-prompt-text");
      expect(promptTextarea.value).toContain(`This is part ${i + 1} of ${SECTIONS.length} of a larger document.`);
      expect(panel.textContent).toContain(`Part ${i + 1} of ${SECTIONS.length}`);

      const chunkMarkdown = extractDocumentSection(promptTextarea.value);
      expect(chunkMarkdown).toContain(`# ${SECTIONS[i]}`);
      // Only chunk 0 carries the document header; later chunks' own slice starts directly
      // at their top-level heading (chunk.js's splitIntoChunks -- the first heading merges
      // into chunk 0 with the header, so only chunk 1+ start "bare").
      if (i > 0) expect(chunkMarkdown.startsWith(`# ${SECTIONS[i]}`)).toBe(true);
      else expect(chunkMarkdown.startsWith("<!--")).toBe(true);

      const copyBtn = panel.querySelector("#ar-copy-prompt");
      expect(copyBtn).toBeTruthy();
      copyBtn.click(); // PROMPT_READY -> AWAITING_RESPONSE

      const responseEl = panel.querySelector("#ar-response");
      expect(responseEl).toBeTruthy();
      // One real substitution per chunk -- "Paragraph 0:" only ever appears once within any
      // given chunk's own text (each section's own paragraph numbering restarts at 0).
      const chunkResponse = withEdit(chunkMarkdown, "Paragraph 0: this", "Paragraph 0: {~~this~>that~~}");
      responseEl.value = "```markdown\n" + chunkResponse + "```";
      panel.querySelector("#ar-validate").click();

      expect(panel.querySelector(".ar-gate-failure"), `chunk ${i} should validate cleanly`).toBeFalsy();

      if (i < SECTIONS.length - 1) {
        await waitForCondition(() => panel.textContent.includes(`Part ${i + 2} of ${SECTIONS.length}`));
      }
    }

    // Last chunk's pass reuses validationPassed -> RATIFYING, exactly like the single-doc
    // path -- merged, resolved edits from all 3 chunks render as ratify rows.
    await waitForCondition(() => panel.querySelector('[data-action="inject"]'));
    const rows = panel.querySelectorAll(".ar-row");
    expect(rows.length).toBe(SECTIONS.length);

    const injectBtn = panel.querySelector('[data-action="inject"]');
    expect(injectBtn).toBeTruthy();
    expect(injectBtn.disabled).toBe(true); // per spec §8: must scroll/review every row first
  });

  it("flags overThreshold honestly for the synthetic document (sanity on the test fixture itself)", async () => {
    const bigDocxBytes = await buildBigMultiSectionDocxBytes();
    // Reuse the same exported markdown the app would compute, via the same path app.js
    // uses, to assert the fixture really is over CHUNK_WORD_THRESHOLD (guards against this
    // whole test silently degrading to the single-doc path if PARAGRAPHS_PER_SECTION or the
    // prose text ever shrinks below threshold).
    const { exportDocx } = await import("../src/ooxml/export.js");
    const { markdown } = await exportDocx(bigDocxBytes, { DOMParserImpl: XmldomDOMParser, filename: "big-multi-section" });
    const words = markdown.trim().split(/\s+/).length;
    expect(words).toBeGreaterThan(CHUNK_WORD_THRESHOLD);
  });
});
