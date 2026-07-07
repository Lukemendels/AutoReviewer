// @vitest-environment happy-dom
// End-to-end DOM smoke test for the M4a Run Review flow (load -> persona default -> compose
// prompt -> copy -> paste response -> validate -> ratify), in the ratify.js-test style:
// drive the real renderShell() DOM and assert on rendered state, not internals. Stops short
// of actually clicking Inject -- that exercises Blob/URL.createObjectURL, which happy-dom
// doesn't implement and tests/app.test.js already covers via buildReviewedDocx directly.
import { DOMParser as XmldomDOMParser } from "@xmldom/xmldom";
import { describe, expect, it } from "vitest";
import { renderShell } from "../src/ui/app.js";
import { exportDocx } from "../src/ooxml/export.js";
import { DEMO_DOCX_BASE64 } from "../src/ui/demo-doc.js";
import { createAppState, STATES } from "../src/ui/state.js";
import { loadSession } from "../src/session.js";

// happy-dom's own DOMParser doesn't reliably parse a real, namespace-heavy document.xml
// (see tests/app.test.js's header comment -- it silently drops the whole body). Every
// other XML-parsing test in this repo sidesteps that by running in the default Node
// environment and injecting @xmldom/xmldom explicitly; this test needs real page DOM
// (drop zones, buttons, textareas) AND real document.xml parsing at once, so it shadows
// the global DOMParser instead -- app.js's export call always falls back to
// globalThis.DOMParser when no DOMParserImpl is passed, exactly as it does in production.
globalThis.DOMParser = XmldomDOMParser;

// The real docx export chain crosses actual Streams-API boundaries (DecompressionStream,
// crypto.subtle.digest), which take more than a single zero-delay tick to settle -- a
// short real delay is more reliable here than chaining setTimeout(0)s.
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

describe("Run Review panel: the demo document walks the whole state machine", () => {
  it("goes from the empty load step to a ratifiable, validated response", async () => {
    // Same bytes app.js's "Try the demo" loads -- computed independently here (rather than
    // scraped back out of the rendered prompt textarea) so the response-construction step
    // below isn't coupled to prompt.js's exact rendering of [DOCUMENT].
    const demoExport = await exportDocx(base64ToArrayBuffer(DEMO_DOCX_BASE64), { filename: "plain-paragraphs" });

    document.body.innerHTML = '<div id="ar-app"></div>';
    renderShell(document.getElementById("ar-app"));

    const panel = document.querySelector('.ar-panel[data-flow="run-review"]');
    expect(panel.querySelector("#ar-try-demo")).toBeTruthy();

    panel.querySelector("#ar-try-demo").click();
    await waitForCondition(() => panel.querySelector(".ar-prompt-text") || panel.querySelector("#ar-load-status")?.textContent);

    // Prompt step: default persona, composed prompt visible, over-threshold not triggered.
    const promptTextarea = panel.querySelector(".ar-prompt-text");
    if (!promptTextarea) throw new Error("load failed: " + panel.querySelector("#ar-load-status")?.textContent);
    expect(promptTextarea.value).toContain("[PERSONA]");
    expect(promptTextarea.value).toContain("[DOCUMENT]");
    expect(panel.textContent).toContain("Default Persona (built-in)");

    const copyBtn = panel.querySelector("#ar-copy-prompt");
    expect(copyBtn).toBeTruthy();
    copyBtn.click(); // transitions PROMPT_READY -> AWAITING_RESPONSE

    // Response step now visible; paste back the untouched document (zero edits) as a
    // fenced block -- a trivially valid response.
    const responseEl = panel.querySelector("#ar-response");
    expect(responseEl).toBeTruthy();

    // No extra newline before the closing fence: demoExport.markdown already ends with its
    // own single trailing newline (byte-preservation rule), and the envelope extractor's
    // verbatim rule means the closing ``` must follow it immediately.
    responseEl.value = "```markdown\n" + demoExport.markdown + "```";
    panel.querySelector("#ar-validate").click();

    // A zero-edit echo passes every gate -> RATIFYING renders the ratification UI.
    expect(panel.querySelector(".ar-gate-failure")).toBeFalsy();
    const injectBtn = panel.querySelector('[data-action="inject"]');
    expect(injectBtn).toBeTruthy();
    expect(injectBtn.disabled).toBe(false); // zero rows -> canInject() trivially true

    // M4b: "Download session" is available now that RATIFYING is reached. Capture the
    // Blob the real download path builds (happy-dom implements URL.createObjectURL) and
    // feed its JSON straight through loadSession() -- the same deserialization path a
    // real Resume click would drive -- to prove the UI wiring produces a session that
    // actually resumes to the same state, rather than just asserting the button exists.
    const downloadBtn = panel.querySelector("#ar-download-session");
    expect(downloadBtn).toBeTruthy();
    let capturedBlob = null;
    const realCreateObjectURL = URL.createObjectURL;
    URL.createObjectURL = (blob) => {
      capturedBlob = blob;
      return realCreateObjectURL.call(URL, blob);
    };
    try {
      downloadBtn.click();
    } finally {
      URL.createObjectURL = realCreateObjectURL;
    }
    expect(capturedBlob).toBeTruthy();
    const savedJson = JSON.parse(await capturedBlob.text());
    expect(savedJson.state).toBe(STATES.RATIFYING);
    expect(savedJson.personaRef).toBeNull();
    expect(savedJson.decisions).toEqual([]); // zero-edit echo -> zero ratify rows

    const resumed = createAppState();
    const { state: restoredState, context: restoredContext } = loadSession(savedJson);
    resumed.hydrate({ state: restoredState, context: restoredContext });
    expect(resumed.state).toBe(STATES.RATIFYING);
    expect(resumed.context.validation.ok).toBe(true);
  });
});
