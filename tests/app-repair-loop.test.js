// @vitest-environment happy-dom
// F-1 regression test (M4d PR-1): acknowledgeFailure() is state.js's only documented
// VALIDATION_FAILED -> AWAITING_RESPONSE transition, but nothing in src/ called it -- every
// paste or Validate click after a G2 failure threw "cannot submitResponse from state
// VALIDATION_FAILED" uncaught inside a setTimeout, silently freezing the UI. This drives the
// real renderShell() DOM (not state.js directly -- state.test.js already calls
// acknowledgeFailure() itself, which is exactly why the suite missed this) through a G2
// failure and a subsequent recovery paste, the way state.test.js's own unit tests never
// would, matching tests/app.uiFlow.test.js's style.
import { DOMParser as XmldomDOMParser } from "@xmldom/xmldom";
import { describe, expect, it } from "vitest";
import { renderShell } from "../src/ui/app.js";
import { exportDocx } from "../src/ooxml/export.js";
import { DEMO_DOCX_BASE64 } from "../src/ui/demo-doc.js";

// See tests/app.uiFlow.test.js's identical header comment for why this shadow is needed:
// happy-dom's own DOMParser silently drops the body of a real, namespace-heavy document.xml.
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

describe("Run Review panel: the repair loop recovers from a G2 failure in-page", () => {
  it("accepts a corrected paste after a bad one, without a page reload", async () => {
    const demoExport = await exportDocx(base64ToArrayBuffer(DEMO_DOCX_BASE64), { filename: "plain-paragraphs" });

    document.body.innerHTML = '<div id="ar-app"></div>';
    renderShell(document.getElementById("ar-app"));

    const panel = document.querySelector('.ar-panel[data-flow="run-review"]');
    panel.querySelector("#ar-try-demo").click();
    await waitForCondition(() => panel.querySelector(".ar-prompt-text") || panel.querySelector("#ar-load-status")?.textContent);

    panel.querySelector("#ar-copy-prompt").click(); // PROMPT_READY -> AWAITING_RESPONSE

    const responseEl = panel.querySelector("#ar-response");
    expect(responseEl).toBeTruthy();

    // Deliberately corrupt one character of the exported body text so strip(response) !==
    // exportedMarkdown byte-for-byte -- a clean G2 failure, not a parse error.
    const corrupted = demoExport.markdown.replace("first paragraph", "frist paragraph");
    expect(corrupted).not.toEqual(demoExport.markdown);
    responseEl.value = "```markdown\n" + corrupted + "```";
    panel.querySelector("#ar-validate").click();

    // G2 failure surfaced in the UI -- this is the state the repair loop must recover from.
    const failureBox = panel.querySelector(".ar-gate-failure");
    expect(failureBox).toBeTruthy();
    expect(failureBox.textContent).toContain("G2");

    // Before the fix, this second Validate click throws uncaught ("cannot submitResponse
    // from state VALIDATION_FAILED") because nothing calls acknowledgeFailure() first --
    // the response textarea and Validate button must still be present and wired.
    const responseElAfterFailure = panel.querySelector("#ar-response");
    expect(responseElAfterFailure).toBeTruthy();
    responseElAfterFailure.value = "```markdown\n" + demoExport.markdown + "```";
    panel.querySelector("#ar-validate").click();

    // A byte-perfect echo passes every gate -> RATIFYING renders the ratification UI, proving
    // the repair loop actually recovered in-page rather than getting stuck.
    expect(panel.querySelector(".ar-gate-failure")).toBeFalsy();
    const injectBtn = panel.querySelector('[data-action="inject"]');
    expect(injectBtn).toBeTruthy();

    // M4b's audit trail (state.js's validationAttempts) must record BOTH attempts -- the
    // failed one and the recovering one. Click Inject (zero ratify rows -> trivially
    // injectable) and capture the audit sidecar's own download, the second of the two Blob
    // downloads that fire -- buildAuditRecord() passes ctx.validationAttempts straight
    // through, so this is the direct, app-level proof the plan calls for.
    const capturedBlobs = [];
    const realCreateObjectURL = URL.createObjectURL;
    URL.createObjectURL = (blob) => {
      capturedBlobs.push(blob);
      return realCreateObjectURL.call(URL, blob);
    };
    try {
      panel.querySelector('[data-action="inject"]').click();
      await waitForCondition(() => capturedBlobs.length >= 2);
    } finally {
      URL.createObjectURL = realCreateObjectURL;
    }

    const auditJson = JSON.parse(await capturedBlobs[1].text());
    expect(auditJson.validationAttempts).toHaveLength(2);
    expect(auditJson.validationAttempts[0].result).toBe("G2");
    expect(auditJson.validationAttempts[1].result).toBe("ok");
  });
});
