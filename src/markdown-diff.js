// Authoritative product compiler: baseline Markdown + clean revised Markdown -> verified
// CriticMarkup. Mechanical comparison is isolated in markdown-diff-mechanical.js; no
// caller receives output until both original and accepted projections reconstruct exactly.
import {
  MARKDOWN_DIFF_VERSION,
  MarkdownDiffError,
  compileMechanicalMarkdownDiff,
  deriveProtectedPrefixLength,
  generateMechanicalCriticMarkup,
} from "./markdown-diff-mechanical.js";
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
  const result = compileMechanicalMarkdownDiff(baselineMarkdown, revisedMarkdown, options);
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

// Exported only for targeted diagnostics and tests. Product code should call the verified
// compiler above, not the mechanical generator directly.
export { generateMechanicalCriticMarkup };
