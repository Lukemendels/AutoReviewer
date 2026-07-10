// OKF persona schema, import (spec §10, §6.1). Parses a StickShift persona .md into a
// PersonaModel prompt.js can compose from. The Train Persona WIZARD (creating/editing a
// persona) is M5, out of scope here -- this module only ever reads a persona someone else
// already authored.
//
// Section matching is heading-tolerant (case/punctuation-insensitive contains-match on the
// §10 heading names) rather than exact, since real files may append parentheticals (e.g.
// "## Review priorities (ordered)" straight out of the spec's own template) that an exact
// match would reject for no good reason -- G2 never sees personas, so there's no byte-
// equality requirement to protect here, unlike the document side.

const SECTION_DEFS = [
  { key: "roleAndVoice", label: "Role and voice", match: "role and voice" },
  { key: "reviewPriorities", label: "Review priorities", match: "review priorities" },
  { key: "styleExemplars", label: "Style exemplars", match: "style exemplars" },
  { key: "doNotTouch", label: "Do-not-touch rules", match: "do not touch rules" },
  { key: "commentConventions", label: "Comment conventions", match: "comment conventions" },
];

export const DEFAULT_PERSONA = Object.freeze({
  name: "Default Persona (built-in)",
  roleAndVoice: "A careful, neutral technical editor. Flag substantive issues; keep prose clear, precise, and consistent.",
  reviewPriorities: "Clarity and correctness first; internal consistency of terminology and cross-references second.",
  styleExemplars: [],
  doNotTouch: [],
  commentConventions: "Prefer a brief inline comment when raising a question or flagging something for the author's judgment; use direct edits only for clear-cut fixes.",
  assistantUrl: null,
  raw: null,
  warnings: [],
  isDefault: true,
});

function normalizeHeading(s) {
  return s
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim()
    .replace(/\s+/g, " ");
}

// Minimal, tolerant front-matter reader -- not a general YAML parser. Handles the schema's
// actual shapes: `key: value` scalars and a `key: >-` folded block scalar (used by
// `description`), whose indented continuation lines fold into one space-joined string.
function parseFrontmatter(markdown) {
  const m = /^---\r?\n([\s\S]*?)\r?\n---\r?\n?/.exec(markdown);
  if (!m) return { fields: {}, body: markdown };
  const body = markdown.slice(m[0].length);
  const fields = {};
  let currentKey = null;
  let folded = [];
  function flush() {
    if (currentKey) fields[currentKey] = folded.join(" ").trim();
    currentKey = null;
    folded = [];
  }
  for (const line of m[1].split(/\r?\n/)) {
    const kv = !/^\s/.test(line) && /^([A-Za-z0-9_]+):\s?(.*)$/.exec(line);
    if (kv) {
      flush();
      const [, key, rest] = kv;
      if (rest === ">-" || rest === "|-" || rest === "|" || rest === ">" || rest === "") {
        currentKey = key;
      } else {
        fields[key] = rest.trim();
      }
    } else if (currentKey) {
      folded.push(line.trim());
    }
  }
  flush();
  return { fields, body };
}

// Body -> { title, sections }. `title` is the first H1 (persona name fallback); everything
// else at any heading level is a candidate section, matched against SECTION_DEFS below.
function splitSections(body) {
  const headingRe = /^(#{1,6})\s+(.*)$/;
  let title = null;
  const sections = [];
  let current = null;
  for (const line of body.split(/\r?\n/)) {
    const hm = headingRe.exec(line);
    if (hm) {
      if (hm[1].length === 1 && title === null) {
        title = hm[2].trim();
        current = null;
        continue;
      }
      current = { heading: hm[2].trim(), bodyLines: [] };
      sections.push(current);
      continue;
    }
    if (current) current.bodyLines.push(line);
  }
  return { title, sections };
}

function parseDoNotTouch(text) {
  if (!text) return [];
  return text
    .split(/\r?\n/)
    .map((l) => l.replace(/^\s*(?:[-*]|\d+[.)])\s*/, "").trim())
    .filter(Boolean);
}

// Best-effort "Before: ... / After: ..." pair extraction, blank-line separated. Falls back
// to one { raw } entry (rather than fabricating pairs) whenever any block doesn't fit the
// shape cleanly -- an unreliable partial parse is worse than admitting it's unstructured.
function parseStyleExemplars(text) {
  if (!text) return [];
  const blocks = text
    .split(/\r?\n\s*\r?\n/)
    .map((b) => b.trim())
    .filter(Boolean);
  if (!blocks.length) return [];
  const pairs = [];
  for (const block of blocks) {
    const m = /^(?:[-*]\s*)?before:\s*([\s\S]*?)\r?\n(?:[-*]\s*)?after:\s*([\s\S]*)$/i.exec(block);
    if (!m) return [{ raw: text }];
    pairs.push({ before: m[1].trim(), after: m[2].trim() });
  }
  return pairs;
}

export function parsePersona(markdown, { filename } = {}) {
  const { fields, body } = parseFrontmatter(markdown);
  const { title, sections } = splitSections(body);
  const name = fields.name || title || (filename ? filename.replace(/\.md$/i, "") : "Unnamed persona");

  const warnings = [];
  const values = {};
  for (const def of SECTION_DEFS) {
    const found = sections.find((s) => normalizeHeading(s.heading).includes(def.match));
    if (!found) {
      warnings.push(`Missing section: ${def.label}`);
      values[def.key] = def.key === "styleExemplars" || def.key === "doNotTouch" ? [] : "";
      continue;
    }
    const text = found.bodyLines.join("\n").trim();
    if (def.key === "doNotTouch") values[def.key] = parseDoNotTouch(text);
    else if (def.key === "styleExemplars") values[def.key] = parseStyleExemplars(text);
    else values[def.key] = text;
  }

  return {
    name,
    roleAndVoice: values.roleAndVoice,
    reviewPriorities: values.reviewPriorities,
    styleExemplars: values.styleExemplars,
    doNotTouch: values.doNotTouch,
    commentConventions: values.commentConventions,
    // Optional (M4d): a link to the persona author's own assistant/profile, if the
    // frontmatter carries one -- not required, not validated as a real URL (personas are
    // never G2-checked, so there's no byte-equality contract to protect here).
    assistantUrl: fields.assistantUrl || null,
    raw: markdown,
    warnings,
    isDefault: false,
  };
}
