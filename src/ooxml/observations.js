// Raw observation extraction (reviewer-pass-slicer step 1, task spec's data model). Thin
// wrapper over exportDocx's own single walk rather than a second OOXML walker: the slicer
// must never diverge from the one canonical exporter, so the same pass that builds the
// CriticMarkup body is the pass that records what each tracked change/comment IS
// (kind/author/date/anchorText/docOrder), gated behind exportDocx's collectObservations
// option so the default export path pays nothing for it.
import { exportDocx } from "./export.js";

export async function extractObservations(docxBytes, options = {}) {
  const result = await exportDocx(docxBytes, { ...options, collectObservations: true });
  return { observations: result.observations, comments: result.comments, counts: result.counts };
}
