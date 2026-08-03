// Fail-closed product compiler. markdown-diff.js performs the mechanical comparison;
// this adapter proves both authoritative views before releasing CriticMarkup downstream.
import {
  MARKDOWN_DIFF_VERSION,
  MarkdownDiffError,
  compileMarkdownDiff as compileMechanicalDiff,
  deriveProtectedPrefixLength,
  generateCriticMarkup as generateMechanicalCriticMarkup,
} from "./markdown-diff.js";
import { tokenize } from "./criticmarkup/grammar.js";

export { MARKDOWN_DIFF_VERSION, MarkdownDiffError, deriveProtectedPrefixLength };

export function projectCriticMarkup(markdown, view, options = {}) {
  if (view !== "original" && view !== "accepted") {
    throw new TypeError(`Unknown CriticMarkup view: ${view}`);
  }
  const parsed = tokenize(markdown, {
    skipBefore: options.protectedPrefixLength || 0,
    skipAfter: options.skipAfter ?? markdown.length,
  });
  if (!parsed.ok) {
    throw new MarkdownDiffError(
      `Generated CriticMarkup is invalid: ${parsed.error.message}`,
      { rawStart: parsed.error.rawStart }
    );
  }

  let output = "";
  let raw = 0;
  for (const token of parsed.tokens) {
    output += markdown.slice(raw, token.rawStart);
    if (token.type === "ins") {
      if (view === "accepted") output += token.text;
    } else if (token.type === "del") {
      if (view === "original") output += token.text;
    } else if (token.type === "sub") {
      output += view === "original" ? token.oldText : token.newText;
    } else if (token.type === "comment" && token.anchored) {
      output += token.highlightText;
    }
    raw = token.rawEnd;
  }
  output += markdown.slice(raw);
  return output;
}

export function compileMarkdownDiff(baselineMarkdown, revisedMarkdown, options = {}) {
  const result = compileMechanicalDiff(baselineMarkdown, revisedMarkdown, options);
  const projectionOptions = { protectedPrefixLength: result.protectedPrefixLength };
  const originalView = projectCriticMarkup(result.criticMarkup, "original", projectionOptions);
  const acceptedView = projectCriticMarkup(result.criticMarkup, "accepted", projectionOptions);
  const originalViewMatches = originalView === baselineMarkdown;
  const acceptedViewMatches = acceptedView === revisedMarkdown;

  if (!originalViewMatches || !acceptedViewMatches) {
    throw new MarkdownDiffError("Compiler invariant failed; no CriticMarkup was released.", {
      originalViewMatches,
      acceptedViewMatches,
    });
  }

  return {
    ...result,
    originalView,
    acceptedView,
    invariants: { originalViewMatches, acceptedViewMatches },
  };
}

export function generateCriticMarkup(baselineMarkdown, revisedMarkdown, options = {}) {
  return compileMarkdownDiff(baselineMarkdown, revisedMarkdown, options).criticMarkup;
}

// Kept available for narrowly-scoped diagnostics; product code should use the verified
// compileMarkdownDiff above rather than releasing this result directly.
export { generateMechanicalCriticMarkup };
