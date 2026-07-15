// @vitest-environment happy-dom
import { DOMParser as XmldomDOMParser } from "@xmldom/xmldom";
import { describe, expect, it } from "vitest";
import { renderShell, selectFlow } from "../src/ui/app.js";
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

globalThis.DOMParser = XmldomDOMParser;

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const root = path.resolve(__dirname, "..");
const fixturesDir = path.join(root, "fixtures");

function loadDocx(name) {
  const buf = readFileSync(path.join(fixturesDir, `${name}.docx`));
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}

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

describe("Respond to Review panel: E2E DOM Flow", () => {
  it("walks the loading, prompt, response, validation, and ratification UI successfully", async () => {
    document.body.innerHTML = '<div id="ar-app"></div>';
    renderShell(document.getElementById("ar-app"));

    selectFlow("respond-review");

    const panel = document.querySelector('.ar-panel[data-flow="respond-review"]');
    expect(panel).toBeTruthy();
    expect(panel.querySelector("#ar-doc-input")).toBeTruthy();

    // Mock FileReader synchronously for the happy-dom test environment
    globalThis.FileReader = class {
      readAsArrayBuffer(f) {
        this.result = f._bytes;
        setTimeout(() => this.onload(), 0);
      }
      readAsText(f) {
        this.result = f._text;
        setTimeout(() => this.onload(), 0);
      }
    };

    const docBytes = loadDocx("comments-threaded-nested");
    const fileInput = panel.querySelector("#ar-doc-input");
    const file = { name: "comments-threaded-nested.docx", _bytes: docBytes };
    
    Object.defineProperty(fileInput, "files", {
      value: [file],
      writable: false,
    });
    fileInput.dispatchEvent(new Event("change"));

    try {
      await waitForCondition(() => panel.querySelector(".ar-prompt-text"));
    } catch (e) {
      throw e;
    }

    const promptTextarea = panel.querySelector(".ar-prompt-text");
    if (!promptTextarea) throw new Error("load failed");
    
    expect(promptTextarea.value).toContain("[RESPOND GRAMMAR]");
    expect(promptTextarea.value).toContain("[DOCUMENT]");
    expect(promptTextarea.value).toContain("⟦C1:");

    const copyBtn = panel.querySelector("#ar-copy-prompt");
    expect(copyBtn).toBeTruthy();
    copyBtn.click();

    const responseEl = panel.querySelector("#ar-response");
    expect(responseEl).toBeTruthy();

    responseEl.value = `
\`\`\`markdown
[C1] {>>Agreed -- will clarify this in the next draft.<<}
[C2] {>>[AR:resolve] Citation added in the bibliography.<<}
[R1] {>>[AR:accept] This aligns with the updated statutory definitions.<<}
\`\`\`
    `.trim();

    panel.querySelector("#ar-validate").click();

    await waitForCondition(() => panel.querySelector('[data-action="inject"]'));
    const injectBtn = panel.querySelector('[data-action="inject"]');
    expect(injectBtn).toBeTruthy();
    expect(injectBtn.disabled).toBe(true); // requires scrolling review
  });
});
