import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { renderShell, selectFlow } from "../src/ui/app.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

describe("app.js module graph", () => {
  it("imports cleanly in a non-DOM (Node) environment", () => {
    expect(typeof renderShell).toBe("function");
    expect(typeof selectFlow).toBe("function");
  });
});

describe("template.html scaffolding", () => {
  const template = readFileSync(path.join(root, "template.html"), "utf8");

  it("declares the STICKSHIFT_TOOL identity matching the dist filename", () => {
    expect(template).toMatch(/file:\s*"autoreviewer-workbench\.html"/);
    expect(template).toMatch(/skillSlug:\s*"autoreviewer-workbench"/);
  });

  it("has the embedded skill placeholder with the required data-skill-slug", () => {
    expect(template).toMatch(
      /<script[^>]*id="stickshift-skill"[^>]*data-skill-slug="autoreviewer-workbench"[^>]*>/
    );
  });

  it("has the onboarding panel entry points", () => {
    expect(template).toContain('id="ss-open"');
    expect(template).toContain('id="ss-yes"');
    expect(template).toContain('id="ss-no"');
    expect(template).toContain('id="ss-copy-skill"');
  });

  it("uses execCommand as the clipboard fallback", () => {
    expect(template).toMatch(/execCommand\(\s*['"]copy['"]\s*\)/);
  });

  it("has no external src/href references", () => {
    expect(template).not.toMatch(/\b(?:src|href)\s*=\s*["'](?:https?:)?\/\//i);
  });
});

describe("src/stickshift/skill.md", () => {
  const skill = readFileSync(path.join(root, "src/stickshift/skill.md"), "utf8");

  it("has StickShift skill frontmatter with no status field", () => {
    expect(skill).toMatch(/^type:\s*Skill\s*$/m);
    expect(skill).toMatch(/^title:\s*.+$/m);
    expect(skill).toMatch(/^description:\s*.+$/m);
    expect(skill).toMatch(/^tags:\s*.+$/m);
    expect(skill).not.toMatch(/^status:\s*.+$/m);
  });

  it("instructs the assistant to emit the tool's HTML_OPEN block", () => {
    expect(skill).toContain("<HTML_OPEN>");
    expect(skill).toContain("tool: autoreviewer-workbench.html");
    expect(skill).toContain("- skills/autoreviewer-workbench.md");
    expect(skill).toContain("</HTML_OPEN>");
  });
});
