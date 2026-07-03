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

export function diffWords(a, b) {
  const A = tokenizeWords(a);
  const B = tokenizeWords(b);

  let start = 0;
  const minLen = Math.min(A.length, B.length);
  while (start < minLen && A[start] === B[start]) start++;

  let endA = A.length, endB = B.length;
  while (endA > start && endB > start && A[endA - 1] === B[endB - 1]) { endA--; endB--; }

  const segments = [];
  for (let i = 0; i < start; i++) segments.push({ type: "same", text: A[i] });
  segments.push(...lcsDiff(A.slice(start, endA), B.slice(start, endB)));
  for (let i = endA; i < A.length; i++) segments.push({ type: "same", text: A[i] });

  return mergeAdjacent(segments);
}
