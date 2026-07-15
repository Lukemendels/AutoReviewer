// Pass clustering (reviewer-pass-slicer step 2, task spec's "Pass clustering" section).
// Pure data transform over extractObservations()'s output -- no OOXML, no DOM, runs
// anywhere Node does.
import { fmtDate } from "./ooxml/parse.js";

// Gap threshold for splitting one author's observations into separate passes. Named and
// exported per the spec ("it will need tuning") rather than inlined.
export const PASS_GAP_HOURS = 48;

function plural(n, word) {
  return `${n} ${word}${n === 1 ? "" : "s"}`;
}

function passCounts(observations) {
  const insertions = observations.filter((o) => o.kind === "insertion").length;
  const deletions = observations.filter((o) => o.kind === "deletion").length;
  const comments = observations.filter((o) => o.kind === "comment").length;
  const replies = observations.filter((o) => o.kind === "comment-reply").length;
  return { insertions, deletions, comments, replies };
}

function buildPass(author, cluster, metadataStripped) {
  // Doc order, not date order -- this is the order the slice renderer (later step) will
  // walk the document in.
  const observations = [...cluster.dated, ...cluster.undated].sort((a, b) => a.docOrder - b.docOrder);
  const dates = cluster.dated.map((o) => o.date);
  const passDate = dates.length ? fmtDate(dates[0]) : null; // dates[] is already ascending
  const windowStart = dates.length ? dates[0] : null;
  const windowEnd = dates.length ? dates[dates.length - 1] : null;
  const undated = dates.length === 0;

  const counts = passCounts(observations);
  const editCount = counts.insertions + counts.deletions;
  const countsSuffix =
    `(${plural(editCount, "edit")}, ${plural(counts.comments, "comment")}, ` +
    `${counts.replies} repl${counts.replies === 1 ? "y" : "ies"})`;

  let label;
  if (metadataStripped) label = `${author} (metadata stripped)`;
  else if (undated) label = `${author} — undated ${countsSuffix}`;
  else label = `${author} — ${passDate} ${countsSuffix}`;

  return { author, passDate, windowStart, windowEnd, undated, observations, counts, label };
}

// Groups raw observations (src/ooxml/observations.js) into per-reviewer, date-clustered
// passes. Returns { passes, metadataStripped }. Never drops an observation, including
// undated ones -- see the spec's edge cases 1 and the "Observations with null dates" rule.
export function clusterPasses(observations, options = {}) {
  const gapMs = (options.gapHours ?? PASS_GAP_HOURS) * 3600 * 1000;

  const byAuthor = new Map();
  for (const obs of observations) {
    const key = obs.author || "Unknown";
    if (!byAuthor.has(key)) byAuthor.set(key, []);
    byAuthor.get(key).push(obs);
  }

  // Word's "Remove personal information from file properties on save" collapses every
  // author to the same placeholder and nulls every date -- detected here (one author,
  // zero dates, at least one observation) so the caller can show a warning instead of
  // silently presenting a single confident-looking pass.
  const metadataStripped =
    observations.length > 0 && byAuthor.size === 1 && observations.every((o) => o.date == null);

  const passes = [];
  for (const [author, obsList] of byAuthor) {
    const dated = obsList.filter((o) => o.date != null).sort((a, b) => new Date(a.date) - new Date(b.date));
    const undated = obsList.filter((o) => o.date == null);

    const clusters = [];
    for (const obs of dated) {
      const last = clusters[clusters.length - 1];
      const prevDated = last && last.dated[last.dated.length - 1];
      const gap = prevDated ? new Date(obs.date) - new Date(prevDated.date) : Infinity;
      if (last && gap <= gapMs) last.dated.push(obs);
      else clusters.push({ dated: [obs], undated: [] });
    }

    if (clusters.length === 0) {
      // No dated observations at all for this author: everything undated becomes one pass.
      clusters.push({ dated: [], undated: [...undated] });
    } else if (clusters.length === 1) {
      // Only one real pass to begin with -- nothing to disambiguate, attach undated here.
      clusters[0].undated.push(...undated);
    } else if (undated.length) {
      // Multiple real passes: which one an undated observation belongs to can't be
      // inferred, so it gets its own pseudo-pass rather than a guess (spec: "never
      // silently drop them").
      clusters.push({ dated: [], undated: [...undated] });
    }

    for (const cluster of clusters) passes.push(buildPass(author, cluster, metadataStripped));
  }

  return { passes, metadataStripped };
}
