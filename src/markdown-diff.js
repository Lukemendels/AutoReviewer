// Deterministic compiler: authoritative baseline Markdown + clean revised Markdown
// -> CriticMarkup suitable for the existing validator / ratification / OOXML pipeline.
// The model performs semantic revision; this module performs mechanical comparison.

export const MARKDOWN_DIFF_VERSION = "markdown-diff-2026.08-1";
export const DEFAULT_INLINE_THRESHOLD = 0.3;
const MAX_INLINE_DIFF_CELLS = 350000;
const OPENERS = ["{++", "{--", "{~~", "{==", "{>>"];

export class MarkdownDiffError extends Error {
  constructor(message, details = {}) {
    super(message);
    this.name = "MarkdownDiffError";
    this.details = details;
  }
}

export function deriveProtectedPrefixLength(markdown) {
  // AutoReviewer exports a synthetic HTML-comment preface containing literal
  // CriticMarkup examples. Protect contiguous leading comments and whitespace without
  // assuming the document contains a Markdown heading.
  let cursor = 0;
  while (cursor < markdown.length) {
    const ws = /^[\t \r\n]*/.exec(markdown.slice(cursor))[0];
    const candidate = cursor + ws.length;
    if (!markdown.startsWith("<!--", candidate)) break;
    const end = markdown.indexOf("-->", candidate + 4);
    if (end === -1) break;
    cursor = end + 3;
  }
  const trailing = /^[\t \r\n]*/.exec(markdown.slice(cursor))[0];
  return cursor + trailing.length;
}

function rejectLiteralOpeners(text, start, label) {
  for (const opener of OPENERS) {
    const at = text.indexOf(opener, start);
    if (at !== -1) {
      throw new MarkdownDiffError(
        `${label} contains literal CriticMarkup opener ${opener} at offset ${at}. ` +
          "The current interchange grammar has no escape syntax.",
        { label, opener, offset: at }
      );
    }
  }
}

function splitLines(text) {
  const lines = [];
  let start = 0;
  let i = 0;
  while (i < text.length) {
    if (text[i] === "\r") {
      const end = i + (text[i + 1] === "\n" ? 2 : 1);
      lines.push({ body: text.slice(start, i), eol: text.slice(i, end), text: text.slice(start, end) });
      start = i = end;
    } else if (text[i] === "\n") {
      const end = i + 1;
      lines.push({ body: text.slice(start, i), eol: "\n", text: text.slice(start, end) });
      start = i = end;
    } else {
      i++;
    }
  }
  if (start < text.length) lines.push({ body: text.slice(start), eol: "", text: text.slice(start) });
  return lines;
}

function splitBlocks(text) {
  if (!text) return [];
  const lines = splitLines(text);
  const blocks = [];
  let current = [];
  let inFence = false;
  let fenceMarker = "";
  let inTable = false;

  function flush() {
    if (!current.length) return;
    const blockText = current.map((line) => line.text).join("");
    if (blockText) blocks.push({ text: blockText });
    current = [];
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const trimmed = line.body.trim();
    const fence = /^(`{3,}|~{3,})/.exec(trimmed);
    if (fence) {
      current.push(line);
      const marker = fence[1][0];
      if (!inFence) {
        inFence = true;
        fenceMarker = marker;
      } else if (marker === fenceMarker) {
        inFence = false;
        fenceMarker = "";
        flush();
      }
      continue;
    }
    if (inFence) {
      current.push(line);
      continue;
    }
    if (/^#{1,6}[ \t]+/.test(trimmed)) {
      flush();
      current.push(line);
      flush();
      continue;
    }
    if (/^\|.*\|$/.test(trimmed)) {
      if (!inTable) {
        flush();
        inTable = true;
      }
      current.push(line);
      const next = lines[i + 1] ? lines[i + 1].body.trim() : "";
      if (!/^\|.*\|$/.test(next)) {
        inTable = false;
        flush();
      }
      continue;
    }
    current.push(line);
    if (trimmed === "") flush();
  }
  flush();
  return blocks;
}

function words(text) {
  return new Set(text.toLocaleLowerCase().match(/[\p{L}\p{N}_]+/gu) || []);
}

function similarity(a, b) {
  const left = words(a);
  const right = words(b);
  if (!left.size && !right.size) return 1;
  if (!left.size || !right.size) return 0;
  let intersection = 0;
  for (const word of left) if (right.has(word)) intersection++;
  return intersection / new Set([...left, ...right]).size;
}

function tokens(text) {
  return text.match(/(\s+|[\p{L}\p{N}_]+|[^\s\p{L}\p{N}_])/gu) || [];
}

function diffArrays(a, b) {
  const n = a.length;
  const m = b.length;
  const dp = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      dp[i][j] = a[i] === b[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
    }
  }
  const out = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (a[i] === b[j]) {
      out.push({ type: "equal", value: a[i] });
      i++;
      j++;
    } else if (dp[i + 1][j] >= dp[i][j + 1]) {
      out.push({ type: "delete", value: a[i++] });
    } else {
      out.push({ type: "insert", value: b[j++] });
    }
  }
  while (i < n) out.push({ type: "delete", value: a[i++] });
  while (j < m) out.push({ type: "insert", value: b[j++] });
  return out;
}

function assertSafe(text, closer, label) {
  for (const opener of OPENERS) {
    if (text.includes(opener)) throw new MarkdownDiffError(`${label} contains literal opener ${opener}.`);
  }
  if (text.includes(closer)) throw new MarkdownDiffError(`${label} contains delimiter ${closer}.`);
}

function add(text) {
  assertSafe(text, "++}", "Inserted text");
  return `{++${text}++}`;
}
function del(text) {
  assertSafe(text, "--}", "Deleted text");
  return `{--${text}--}`;
}
function sub(oldText, newText) {
  assertSafe(oldText, "~>", "Substitution source");
  assertSafe(newText, "~~}", "Substitution replacement");
  return `{~~${oldText}~>${newText}~~}`;
}

function inlineDiff(oldText, newText) {
  if (oldText === newText) return oldText;
  const a = tokens(oldText);
  const b = tokens(newText);
  if (a.length * b.length > MAX_INLINE_DIFF_CELLS) return del(oldText) + add(newText);
  const ops = diffArrays(a, b);
  let output = "";
  let i = 0;
  while (i < ops.length) {
    if (ops[i].type === "equal") {
      output += ops[i++].value;
      continue;
    }
    let deleted = "";
    let inserted = "";
    while (i < ops.length && ops[i].type !== "equal") {
      if (ops[i].type === "delete") deleted += ops[i].value;
      else inserted += ops[i].value;
      i++;
    }
    output += deleted && inserted ? sub(deleted, inserted) : deleted ? del(deleted) : add(inserted);
  }
  return output;
}

function groupBlockOps(ops, sourceBlocks, revisedBlocks) {
  let sourceIndex = 0;
  let revisedIndex = 0;
  const expanded = ops.map((op) => {
    if (op.type === "equal") return { type: "equal", oldBlock: sourceBlocks[sourceIndex++], newBlock: revisedBlocks[revisedIndex++] };
    if (op.type === "delete") return { type: "delete", oldBlock: sourceBlocks[sourceIndex++] };
    return { type: "insert", newBlock: revisedBlocks[revisedIndex++] };
  });
  const groups = [];
  let i = 0;
  while (i < expanded.length) {
    if (expanded[i].type === "equal") {
      const items = [];
      while (i < expanded.length && expanded[i].type === "equal") items.push(expanded[i++]);
      groups.push({ type: "equal", items });
    } else {
      const deletions = [];
      const insertions = [];
      while (i < expanded.length && expanded[i].type !== "equal") {
        if (expanded[i].type === "delete") deletions.push(expanded[i].oldBlock);
        else insertions.push(expanded[i].newBlock);
        i++;
      }
      groups.push({ type: "change", deletions, insertions });
    }
  }
  return groups;
}

function renderChangedRegion(oldText, newText, threshold) {
  if (!oldText) return add(newText);
  if (!newText) return del(oldText);
  return similarity(oldText, newText) < threshold ? del(oldText) + add(newText) : inlineDiff(oldText, newText);
}

export function compileMarkdownDiff(baselineMarkdown, revisedMarkdown, options = {}) {
  if (typeof baselineMarkdown !== "string" || typeof revisedMarkdown !== "string") {
    throw new TypeError("Baseline and revised Markdown must be strings.");
  }
  const threshold = Number.isFinite(options.threshold)
    ? Math.max(0, Math.min(1, options.threshold))
    : DEFAULT_INLINE_THRESHOLD;
  const protectedPrefixLength = Number.isInteger(options.protectedPrefixLength)
    ? options.protectedPrefixLength
    : deriveProtectedPrefixLength(baselineMarkdown);
  const baselinePrefix = baselineMarkdown.slice(0, protectedPrefixLength);
  if (!revisedMarkdown.startsWith(baselinePrefix)) {
    throw new MarkdownDiffError(
      "The protected AutoReviewer preface was changed or omitted. Restore it byte-for-byte.",
      { protectedPrefixLength }
    );
  }
  rejectLiteralOpeners(baselineMarkdown, protectedPrefixLength, "Baseline Markdown");
  rejectLiteralOpeners(revisedMarkdown, protectedPrefixLength, "Revised Markdown");

  const sourceBlocks = splitBlocks(baselineMarkdown);
  const revisedBlocks = splitBlocks(revisedMarkdown);
  const blockOps = diffArrays(sourceBlocks.map((b) => b.text), revisedBlocks.map((b) => b.text));
  const groups = groupBlockOps(blockOps, sourceBlocks, revisedBlocks);
  let criticMarkup = "";
  let editGroups = 0;
  for (const group of groups) {
    if (group.type === "equal") {
      criticMarkup += group.items.map((item) => item.oldBlock.text).join("");
    } else {
      const oldText = group.deletions.map((block) => block.text).join("");
      const newText = group.insertions.map((block) => block.text).join("");
      criticMarkup += renderChangedRegion(oldText, newText, threshold);
      editGroups++;
    }
  }
  return {
    criticMarkup,
    editGroups,
    protectedPrefixLength,
    version: MARKDOWN_DIFF_VERSION,
    baselineMarkdown,
    revisedMarkdown,
  };
}

export function generateCriticMarkup(baselineMarkdown, revisedMarkdown, options = {}) {
  return compileMarkdownDiff(baselineMarkdown, revisedMarkdown, options).criticMarkup;
}
