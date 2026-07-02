import { build } from "esbuild";
import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, "..");
const isRelease = process.argv.includes("--release") || process.env.NODE_ENV === "production";

function escapeForInlineScript(js) {
  // Prevent an embedded "</script>" from prematurely closing the wrapping <script> tag.
  return js.replace(/<\/script/gi, "<\\/script");
}

async function main() {
  const bundleResult = await build({
    entryPoints: [path.join(root, "src/ui/app.js")],
    bundle: true,
    format: "iife",
    platform: "browser",
    write: false,
    sourcemap: isRelease ? false : "inline",
    minify: false,
  });
  const bundleJs = bundleResult.outputFiles[0].text;

  const template = readFileSync(path.join(root, "template.html"), "utf8");
  const skillMd = readFileSync(path.join(root, "src/stickshift/skill.md"), "utf8").trim();

  const output = template
    .replace("__STICKSHIFT_SKILL__", () => escapeForInlineScript(skillMd))
    .replace("__BUNDLE_JS__", () => escapeForInlineScript(bundleJs));

  mkdirSync(path.join(root, "dist"), { recursive: true });
  const outPath = path.join(root, "dist/autoreviewer-workbench.html");
  writeFileSync(outPath, output, "utf8");

  const sizeKB = (Buffer.byteLength(output, "utf8") / 1024).toFixed(1);
  console.log(`Built ${path.relative(root, outPath)} (${sizeKB} KB)`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
