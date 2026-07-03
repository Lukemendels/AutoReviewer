// Tokenizer for the v1 CriticMarkup profile (spec §4). Support exactly these five
// constructs; anything else (a lone brace not matching one of the openers below, e.g.
// "{1}") is ordinary text, not a grammar error -- only sequences that begin one of these
// five exact openers and then fail to close/pair/avoid-nesting are rejected.
//
// Each token gets two coordinate spaces: rawStart/rawEnd (position in the input markdown
// as given -- needed to slice text back out, e.g. for a repair prompt) and
// strippedStart/strippedEnd (the running offset through strip(markdown), which equals the
// position in exportedMarkdown once G2 confirms the two texts match). Downstream code
// (G1's block-crossing check, G3/G4's snapBoundary+resolveRange) works entirely in the
// stripped coordinate space.
const OPENERS = ["{++", "{--", "{~~", "{==", "{>>"];

function startsWith(str, pos, needle) {
  return str.startsWith(needle, pos);
}

// Scans forward from `from` for `closer`. Fails with a nesting error if any of the five
// openers appears first. Returns the index of closer's first character, or -1 if the end
// of the string is reached first (unbalanced).
function scanForCloser(str, from, closer) {
  const n = str.length;
  for (let i = from; i <= n - closer.length; i++) {
    if (startsWith(str, i, closer)) return { closeAt: i };
    for (const op of OPENERS) {
      if (startsWith(str, i, op)) return { nestedAt: i, opener: op };
    }
  }
  return { closeAt: -1 };
}
// Same, but for the "~>" substitution arrow, which isn't a closer -- just a required
// interior delimiter also subject to the no-nesting rule.
function scanForArrow(str, from) {
  return scanForCloser(str, from, "~>");
}

// opts.skipBefore/skipAfter exempt a leading/trailing raw-position range from
// opener-scanning (both coordinate spaces still advance normally through them, 1:1, on
// the assumption that this region is echoed verbatim -- any deviation there still fails
// G2's overall byte-equality check regardless). This exists because the exported
// document's own header comment line is REQUIRED reading material for the model (spec
// §5.1) and literally contains a CriticMarkup legend as example syntax (e.g.
// "{==highlighted==} {>>comment<<}") -- without an exemption, that legend line itself
// would be misparsed as real tokens by every single document's own header, which is
// exactly the "literal {++-like sequences in document text" collision spec §4 names.
// Callers (validate.js) compute skipBefore/skipAfter from the source map's first/last
// block boundaries, since everything outside those blocks is synthetic scaffolding with
// no editable document content -- never touched by the tokenizer either way.
export function tokenize(markdown, opts = {}) {
  const skipBefore = opts.skipBefore ?? 0;
  const skipAfter = opts.skipAfter ?? markdown.length;
  const tokens = [];
  let raw = 0;
  let stripped = 0;
  const n = markdown.length;

  function fail(message, at) {
    return { ok: false, error: { message, rawStart: at } };
  }

  while (raw < n) {
    const scanningEnabled = raw >= skipBefore && raw < skipAfter;
    if (scanningEnabled && startsWith(markdown, raw, "{++")) {
      const scan = scanForCloser(markdown, raw + 3, "++}");
      if (scan.nestedAt !== undefined) {
        return fail(`token opened at ${raw} is not closed before another token (${scan.opener}) opens at ${scan.nestedAt}`, raw);
      }
      if (scan.closeAt === -1) return fail("unbalanced {++...++} (insertion never closes)", raw);
      const text = markdown.slice(raw + 3, scan.closeAt);
      tokens.push({ type: "ins", rawStart: raw, rawEnd: scan.closeAt + 3, strippedStart: stripped, strippedEnd: stripped, text });
      raw = scan.closeAt + 3;
      continue;
    }
    if (scanningEnabled && startsWith(markdown, raw, "{--")) {
      const scan = scanForCloser(markdown, raw + 3, "--}");
      if (scan.nestedAt !== undefined) {
        return fail(`token opened at ${raw} is not closed before another token (${scan.opener}) opens at ${scan.nestedAt}`, raw);
      }
      if (scan.closeAt === -1) return fail("unbalanced {--...--} (deletion never closes)", raw);
      const text = markdown.slice(raw + 3, scan.closeAt);
      tokens.push({ type: "del", rawStart: raw, rawEnd: scan.closeAt + 3, strippedStart: stripped, strippedEnd: stripped + text.length, text });
      stripped += text.length;
      raw = scan.closeAt + 3;
      continue;
    }
    if (scanningEnabled && startsWith(markdown, raw, "{~~")) {
      const arrow = scanForArrow(markdown, raw + 3);
      if (arrow.nestedAt !== undefined) {
        return fail(`token opened at ${raw} is not closed before another token (${arrow.opener}) opens at ${arrow.nestedAt}`, raw);
      }
      if (arrow.closeAt === -1) return fail("unbalanced {~~old~>new~~} (missing ~> arrow)", raw);
      const oldText = markdown.slice(raw + 3, arrow.closeAt);
      const scan = scanForCloser(markdown, arrow.closeAt + 2, "~~}");
      if (scan.nestedAt !== undefined) {
        return fail(`token opened at ${raw} is not closed before another token (${scan.opener}) opens at ${scan.nestedAt}`, raw);
      }
      if (scan.closeAt === -1) return fail("unbalanced {~~old~>new~~} (substitution never closes)", raw);
      const newText = markdown.slice(arrow.closeAt + 2, scan.closeAt);
      tokens.push({ type: "sub", rawStart: raw, rawEnd: scan.closeAt + 3, strippedStart: stripped, strippedEnd: stripped + oldText.length, oldText, newText });
      stripped += oldText.length;
      raw = scan.closeAt + 3;
      continue;
    }
    if (scanningEnabled && startsWith(markdown, raw, "{==")) {
      const scan = scanForCloser(markdown, raw + 3, "==}");
      if (scan.nestedAt !== undefined) {
        return fail(`token opened at ${raw} is not closed before another token (${scan.opener}) opens at ${scan.nestedAt}`, raw);
      }
      if (scan.closeAt === -1) return fail("unbalanced {==...==} (highlight never closes)", raw);
      const highlightText = markdown.slice(raw + 3, scan.closeAt);
      const afterHighlight = scan.closeAt + 3;
      if (!startsWith(markdown, afterHighlight, "{>>")) {
        return fail("{==...==} must be immediately followed by {>>...<<}", raw);
      }
      const commentScan = scanForCloser(markdown, afterHighlight + 3, "<<}");
      if (commentScan.nestedAt !== undefined) {
        return fail(`token opened at ${raw} is not closed before another token (${commentScan.opener}) opens at ${commentScan.nestedAt}`, raw);
      }
      if (commentScan.closeAt === -1) return fail("unbalanced {>>...<<} (anchored comment never closes)", raw);
      const commentText = markdown.slice(afterHighlight + 3, commentScan.closeAt);
      tokens.push({
        type: "comment",
        anchored: true,
        rawStart: raw,
        rawEnd: commentScan.closeAt + 3,
        strippedStart: stripped,
        strippedEnd: stripped + highlightText.length,
        highlightText,
        commentText,
      });
      stripped += highlightText.length;
      raw = commentScan.closeAt + 3;
      continue;
    }
    if (scanningEnabled && startsWith(markdown, raw, "{>>")) {
      const scan = scanForCloser(markdown, raw + 3, "<<}");
      if (scan.nestedAt !== undefined) {
        return fail(`token opened at ${raw} is not closed before another token (${scan.opener}) opens at ${scan.nestedAt}`, raw);
      }
      if (scan.closeAt === -1) return fail("unbalanced {>>...<<} (bare comment never closes)", raw);
      const commentText = markdown.slice(raw + 3, scan.closeAt);
      tokens.push({ type: "comment", anchored: false, rawStart: raw, rawEnd: scan.closeAt + 3, strippedStart: stripped, strippedEnd: stripped, commentText });
      raw = scan.closeAt + 3;
      continue;
    }
    // Ordinary character (including a "{" that doesn't begin any of the five openers)
    // passes through unchanged in both coordinate spaces.
    stripped += 1;
    raw += 1;
  }
  return { ok: true, tokens };
}
