// Static checks mirroring spec-html-tool-compliance.md §11 checklist items that are
// machine-checkable, plus the size budget from docs/autoreviewer-workbench-spec.md §13.
import { readFileSync, existsSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, "..");
// Optional positional artifact path, e.g. "html/autoreviewer-workbench.html" for CI's
// check against the committed build; defaults to the local dist/ build.
const outArg = process.argv[2];
const distPath = outArg ? path.resolve(root, outArg) : path.join(root, "dist/autoreviewer-workbench.html");
const SIZE_BUDGET_BYTES = 400 * 1024;

const failures = [];
function check(label, ok) {
  if (!ok) failures.push(label);
  console.log(`${ok ? "PASS" : "FAIL"}  ${label}`);
}

if (!existsSync(distPath)) {
  console.error(`artifact not found at ${distPath} -- run "npm run build" (or "npm run build:html") first.`);
  process.exit(1);
}

const html = readFileSync(distPath, "utf8");

check(`artifact exists at ${path.relative(root, distPath)}`, true);

const byteSize = Buffer.byteLength(html, "utf8");
check(`size budget < ${SIZE_BUDGET_BYTES / 1024} KB (actual: ${(byteSize / 1024).toFixed(1)} KB)`, byteSize < SIZE_BUDGET_BYTES);

// No external network dependencies.
const externalSrcHref = /\b(?:src|href)\s*=\s*["'](https?:)?\/\//i.test(html);
check("no src=/href= to external origins", !externalSrcHref);
check("no fetch( calls", !/\bfetch\s*\(/.test(html));
check("no XMLHttpRequest usage", !/\bXMLHttpRequest\b/.test(html));
check("no localStorage/sessionStorage usage", !/\b(localStorage|sessionStorage)\b/.test(html));

// Embedded companion skill.
const skillMatch = html.match(
  /<script[^>]*id=["']stickshift-skill["'][^>]*data-skill-slug=["']([^"']+)["'][^>]*>([\s\S]*?)<\/script>/
);
check("#stickshift-skill script block present with data-skill-slug", !!skillMatch);

const skillSlug = skillMatch ? skillMatch[1] : null;
const skillBody = skillMatch ? skillMatch[2].trim() : "";

check("embedded skill frontmatter has type: Skill", /^type:\s*Skill\s*$/m.test(skillBody));
check("embedded skill frontmatter has title:", /^title:\s*.+$/m.test(skillBody));
check("embedded skill frontmatter has description:", /^description:\s*.+$/m.test(skillBody));
check("embedded skill frontmatter has tags:", /^tags:\s*.+$/m.test(skillBody));
check("embedded skill frontmatter has no status:", !/^status:\s*.+$/m.test(skillBody));
check("embedded skill body emits <HTML_OPEN> block", /<HTML_OPEN>[\s\S]*?<\/HTML_OPEN>/.test(skillBody));

const toolLineMatch = skillBody.match(/<HTML_OPEN>[\s\S]*?tool:\s*(\S+)[\s\S]*?<\/HTML_OPEN>/);
const skillToolFile = toolLineMatch ? toolLineMatch[1] : null;

// Identity declaration.
const identityMatch = html.match(
  /const\s+STICKSHIFT_TOOL\s*=\s*\{\s*file:\s*["']([^"']+)["']\s*,\s*skillSlug:\s*["']([^"']+)["']/
);
check("STICKSHIFT_TOOL identity declaration present", !!identityMatch);

const identityFile = identityMatch ? identityMatch[1] : null;
const identitySlug = identityMatch ? identityMatch[2] : null;

check("STICKSHIFT_TOOL.skillSlug matches data-skill-slug", !!identitySlug && identitySlug === skillSlug);
check("STICKSHIFT_TOOL.file matches the canonical artifact filename", identityFile === "autoreviewer-workbench.html");
check("STICKSHIFT_TOOL.file matches skill's <HTML_OPEN> tool: line", !!skillToolFile && skillToolFile === identityFile);

// Onboarding panel.
check("onboarding entry point present (non-blocking)", /id=["']ss-open["']/.test(html));
check("onboarding panel has Yes/No branch", /id=["']ss-yes["']/.test(html) && /id=["']ss-no["']/.test(html));
check("onboarding panel has Copy skill button", /id=["']ss-copy-skill["']/.test(html));

// Clipboard fallback.
check("copyText uses execCommand fallback", /execCommand\(\s*['"]copy['"]\s*\)/.test(html));

console.log("");
if (failures.length) {
  console.error(`${failures.length} compliance check(s) failed.`);
  process.exit(1);
}
console.log("All compliance checks passed.");
