// Minimal word-level LCS diff, no dependency. Shared by G2's failure view (diffing the
// stripped response against exportedMarkdown) and the ratification UI's inline edit diff.
//
// Common prefix/suffix are trimmed before the O(n*m) LCS DP runs, so a small localized
// drift stays cheap regardless of total document size -- G2's diff runs on full documents,
// where the drift is typically one paraphrased sentence surrounded by thousands of
// unchanged words.
function tokenizeWords(s) {
  return s.split(/(\s+)/).filter((t) => t.length > 0);
}

function lcsDiff(a, b) {
  const n = a.length, m = b.length;
  const dp = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      dp[i][j] = a[i] === b[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
    }
  }
  const segs = [];
  let i = 0, j = 0;
  while (i < n && j < m) {
    if (a[i] === b[j]) { segs.push({ type: "same", text: a[i] }); i++; j++; }
    else if (dp[i + 1][j] >= dp[i][j + 1]) { segs.push({ type: "del", text: a[i] }); i++; }
    else { segs.push({ type: "add", text: b[j] }); j++; }
  }
  while (i < n) { segs.push({ type: "del", text: a[i] }); i++; }
  while (j < m) { segs.push({ type: "add", text: b[j] }); j++; }
  return segs;
}

function mergeAdjacent(segments) {
  const out = [];
  for (const s of segments) {
    const last = out[out.length - 1];
    if (last && last.type === s.type) last.text += s.text;
    else out.push({ ...s });
  }
  return out;
}

// Bounds the O(n*m) LCS DP table's worst-case size regardless of overall document length.
// Chosen from measurement (see tests/diff.test.js's cap tests): at this cell count, the DP
// still completes in well under 100ms and under 50MB -- a real document-wide G2 failure
// with drift scattered across many separate points (not one localized paraphrase) can
// leave a divergent middle window spanning nearly the whole document after prefix/suffix
// trimming, which without this cap turns into a multi-second, multi-GB computation (see
// the M3b plan's lessons section for the measured repro on fixtures/stressor.docx: a
// window of ~16,000 x 16,000 tokens took ~13s and ~2.1GB for a SINGLE diffWords call).
export const MAX_DIFF_CELLS = 4_000_000;

// Past MAX_DIFF_CELLS, returns null rather than attempting the full diff -- callers must
// fall back to findFirstDivergence() for a cheap, safe alternative (this is what "on
// demand" full-diff UI affordances should do: try diffWords, and if null, show the
// first-divergence context instead, so a user can never accidentally trigger the
// worst-case cost just by asking to see a diff).
export function diffWords(a, b) {
  const A = tokenizeWords(a);
  const B = tokenizeWords(b);

  let start = 0;
  const minLen = Math.min(A.length, B.length);
  while (start < minLen && A[start] === B[start]) start++;

  let endA = A.length, endB = B.length;
  while (endA > start && endB > start && A[endA - 1] === B[endB - 1]) { endA--; endB--; }

  const midA = endA - start, midB = endB - start;
  if (midA * midB > MAX_DIFF_CELLS) return null;

  const segments = [];
  for (let i = 0; i < start; i++) segments.push({ type: "same", text: A[i] });
  segments.push(...lcsDiff(A.slice(start, endA), B.slice(start, endB)));
  for (let i = endA; i < A.length; i++) segments.push({ type: "same", text: A[i] });

  return mergeAdjacent(segments);
}

// Cheap (O(min(n,m)), no O(n*m) computation ever) first-point-of-difference finder: scans
// forward for the first character where the two strings diverge and returns that offset
// plus a small context excerpt on each side. This is what G2 failures should report by
// default -- diffing the whole document is comparatively expensive and, per spec, only
// needed for the human-facing "show me exactly what changed" view, not for the gate
// decision itself (which only needs to know THAT and roughly WHERE they diverge).
export function findFirstDivergence(a, b, contextChars = 120) {
  const minLen = Math.min(a.length, b.length);
  let offset = 0;
  while (offset < minLen && a[offset] === b[offset]) offset++;

  const contextStart = Math.max(0, offset - contextChars);
  return {
    offset,
    before: a.slice(contextStart, offset),
    afterA: a.slice(offset, offset + contextChars),
    afterB: b.slice(offset, offset + contextChars),
    truncatedBefore: contextStart > 0,
    truncatedAfterA: offset + contextChars < a.length,
    truncatedAfterB: offset + contextChars < b.length,
  };
}
