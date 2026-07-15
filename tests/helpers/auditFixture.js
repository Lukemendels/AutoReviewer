// Synthetic docx for the reviewer-pass-slicer step-1 audit (observation extraction).
// Same technique as tests/redline-export.test.js: start from comments-threaded.docx's
// container (already wires word/commentsExtended.xml's relationship + content-type
// override) and swap in hand-written document.xml/comments.xml/commentsExtended.xml via
// writeZip -- commentsExtended reply-threading isn't something python-docx can author, so
// this is built the same way the redline-export shell's own fixture was, not via
// fixtures/generate.py.
//
// Two reviewers:
// - Jim Smith: two passes >48h apart. Pass 1 (2026-05-14, ~09:12-09:30): one insertion +
//   one anchored comment. Pass 2 (2026-05-16, ~10:00-10:30): one deletion + one resolved
//   anchored comment.
// - Katie Chen: one pass (2026-05-15, ~08:00-09:00): one insertion, one point comment (no
//   text between commentRangeStart/End -- exercises the containing-sentence fallback), and
//   a reply to Jim's pass-1 comment (thread placement follows the replier's own pass, not
//   the parent's -- exactly the case spec-workbench.md's slicer needs).
import { unzip } from "../../src/zip/reader.js";
import { writeZip } from "../../src/zip/writer.js";
import { loadDocxBytes } from "./docx.js";

export const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t xml:space="preserve">The rule shall apply to </w:t></w:r><w:ins w:id="10" w:author="Jim Smith" w:date="2026-05-14T09:12:00Z"><w:r><w:t>all covered carriers</w:t></w:r></w:ins><w:r><w:t>.</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">This clause is obsolete and </w:t></w:r><w:del w:id="11" w:author="Jim Smith" w:date="2026-05-16T10:00:00Z"><w:r><w:delText>should be removed entirely</w:delText></w:r></w:del><w:r><w:t>.</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">The deadline is </w:t></w:r><w:ins w:id="12" w:author="Katie Chen" w:date="2026-05-15T08:00:00Z"><w:r><w:t>60 days</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> after publication.</w:t></w:r></w:p>
<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>This sentence has a discussion thread attached to it.</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>
<w:p><w:commentRangeStart w:id="2"/><w:r><w:t>This sentence has a resolved comment attached to it.</w:t></w:r><w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Some prior sentence about scope. The deadline sits </w:t></w:r><w:commentRangeStart w:id="3"/><w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r><w:r><w:t>within this second sentence for testing.</w:t></w:r></w:p>
<w:sectPr/>
</w:body>
</w:document>`;

export const COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
<w:comment w:id="0" w:author="Jim Smith" w:date="2026-05-14T09:30:00Z"><w:p w14:paraId="10000000"><w:r><w:t>Is this the right threshold?</w:t></w:r></w:p></w:comment>
<w:comment w:id="1" w:author="Katie Chen" w:date="2026-05-15T09:00:00Z"><w:p w14:paraId="10000001"><w:r><w:t>Agreed -- flagged for legal review.</w:t></w:r></w:p></w:comment>
<w:comment w:id="2" w:author="Jim Smith" w:date="2026-05-16T10:30:00Z"><w:p w14:paraId="10000002"><w:r><w:t>Please cite the authority here.</w:t></w:r></w:p></w:comment>
<w:comment w:id="3" w:author="Katie Chen" w:date="2026-05-15T08:30:00Z"><w:p w14:paraId="10000003"><w:r><w:t>Does this point still hold after the revision above?</w:t></w:r></w:p></w:comment>
</w:comments>`;

export const COMMENTS_EXTENDED_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
<w15:commentEx w15:paraId="10000000" w15:done="0"/>
<w15:commentEx w15:paraId="10000001" w15:paraIdParent="10000000" w15:done="0"/>
<w15:commentEx w15:paraId="10000002" w15:done="1"/>
<w15:commentEx w15:paraId="10000003" w15:done="0"/>
</w15:commentsEx>`;

export async function buildAuditFixtureDocx() {
  const zip = await unzip(loadDocxBytes("comments-threaded"));
  return writeZip(zip, {
    "word/document.xml": DOCUMENT_XML,
    "word/comments.xml": COMMENTS_XML,
    "word/commentsExtended.xml": COMMENTS_EXTENDED_XML,
  });
}
