// @vitest-environment happy-dom
// M4d PR-3 (F-5): the "Model (as shown in DHSChat)" free-text field must flow through to
// both the session .json (so a resumed review remembers it) and the audit sidecar (the
// actual provenance record) -- typing into it must NOT trigger a full re-render (which
// would blow away focus/cursor on every keystroke), so app.js reads/writes it straight off
// the shared ctx object rather than routing it through appState.setX()/notify().
import { DOMParser as XmldomDOMParser } from "@xmldom/xmldom";
import { describe, expect, it } from "vitest";
import { renderShell } from "../src/ui/app.js";
import { exportDocx } from "../src/ooxml/export.js";
import { DEMO_DOCX_BASE64 } from "../src/ui/demo-doc.js";

globalThis.DOMParser = XmldomDOMParser;

function waitForCondition(check, { timeout = 2000, interval = 20 } = {}) {
  return new Promise((resolve, reject) => {
    const start = Date.now();
    (function poll() {
      if (check()) return resolve();
      if (Date.now() - start > timeout) return reject(new Error("waitForCondition: timed out"));
      setTimeout(poll, interval);
    })();
  });
}

function base64ToArrayBuffer(b64) {
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

function typeInto(el, value) {
  el.value = value;
  el.dispatchEvent(new Event("input", { bubbles: true }));
}

describe("Run Review panel: Model (as shown in DHSChat) provenance field", () => {
  it("survives a re-render and is recorded verbatim in both the session file and the audit sidecar", async () => {
    const demoExport = await exportDocx(base64ToArrayBuffer(DEMO_DOCX_BASE64), { filename: "plain-paragraphs" });

    document.body.innerHTML = '<div id="ar-app"></div>';
    renderShell(document.getElementById("ar-app"));

    const panel = document.querySelector('.ar-panel[data-flow="run-review"]');
    panel.querySelector("#ar-try-demo").click();
    await waitForCondition(() => panel.querySelector(".ar-prompt-text"));

    const modelInput = panel.querySelector("#ar-model");
    expect(modelInput).toBeTruthy();
    typeInto(modelInput, "GPT-5.5 (Standard)");

    // A state transition (PROMPT_READY -> AWAITING_RESPONSE) forces a full teardown/rebuild
    // of the panel -- the typed value must still be there afterward, proving it's read off
    // ctx rather than lost with the old DOM node.
    panel.querySelector("#ar-copy-prompt").click();
    expect(panel.querySelector("#ar-model").value).toBe("GPT-5.5 (Standard)");

    const responseEl = panel.querySelector("#ar-response");
    responseEl.value = "```markdown\n" + demoExport.markdown + "```";
    panel.querySelector("#ar-validate").click();

    const injectBtn = panel.querySelector('[data-action="inject"]');
    expect(injectBtn).toBeTruthy();
    expect(panel.querySelector("#ar-model").value).toBe("GPT-5.5 (Standard)");

    const capturedBlobs = [];
    const realCreateObjectURL = URL.createObjectURL;
    URL.createObjectURL = (blob) => {
      capturedBlobs.push(blob);
      return realCreateObjectURL.call(URL, blob);
    };
    try {
      panel.querySelector("#ar-download-session").click();
      injectBtn.click();
      await waitForCondition(() => capturedBlobs.length >= 3);
    } finally {
      URL.createObjectURL = realCreateObjectURL;
    }

    const sessionJson = JSON.parse(await capturedBlobs[0].text());
    expect(sessionJson.model).toBe("GPT-5.5 (Standard)");

    const auditJson = JSON.parse(await capturedBlobs[2].text());
    expect(auditJson.model).toBe("GPT-5.5 (Standard)");
    expect(auditJson.promptVersion).toBe("m4d-2026.07-1");
  });
});
