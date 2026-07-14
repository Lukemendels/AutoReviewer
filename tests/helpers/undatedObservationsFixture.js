// Synthetic docx exercising clusterPasses's undated-observation rule (spec: "Observations
// with null dates for an author who has dated observations: attach to that author's
// single pass if there's only one, otherwise bucket into an 'undated' pseudo-pass").
//
// - Al: two dated passes >48h apart (insertion, then deletion) plus one undated comment
//   -- ambiguous which pass it belongs to, so it must land in its own undated pseudo-pass.
// - Bo: one dated pass (insertion) plus one undated comment -- unambiguous, so it attaches
//   to that single pass instead of spawning a pseudo-pass.
import { unzip } from "../../src/zip/reader.js";
import { writeZip } from "../../src/zip/writer.js";
import { loadDocxBytes } from "./docx.js";

export const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t xml:space="preserve">Al pass 1: </w:t></w:r><w:ins w:id="1" w:author="Al" w:date="2026-06-01T09:00:00Z"><w:r><w:t>an insertion</w:t></w:r></w:ins><w:r><w:t>.</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Al pass 2: </w:t></w:r><w:del w:id="2" w:author="Al" w:date="2026-06-05T09:00:00Z"><w:r><w:delText>a deletion</w:delText></w:r></w:del><w:r><w:t>.</w:t></w:r></w:p>
<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Al's undated comment anchor.</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Bo's only pass: </w:t></w:r><w:ins w:id="3" w:author="Bo" w:date="2026-06-02T09:00:00Z"><w:r><w:t>an insertion</w:t></w:r></w:ins><w:r><w:t>.</w:t></w:r></w:p>
<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>Bo's undated comment anchor.</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>
<w:sectPr/>
</w:body>
</w:document>`;

export const COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
<w:comment w:id="0" w:author="Al"><w:p w14:paraId="30000000"><w:r><w:t>Undated comment from Al.</w:t></w:r></w:p></w:comment>
<w:comment w:id="1" w:author="Bo"><w:p w14:paraId="30000001"><w:r><w:t>Undated comment from Bo.</w:t></w:r></w:p></w:comment>
</w:comments>`;

export const COMMENTS_EXTENDED_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
<w15:commentEx w15:paraId="30000000" w15:done="0"/>
<w15:commentEx w15:paraId="30000001" w15:done="0"/>
</w15:commentsEx>`;

export async function buildUndatedObservationsFixtureDocx() {
  const zip = await unzip(loadDocxBytes("comments-threaded"));
  return writeZip(zip, {
    "word/document.xml": DOCUMENT_XML,
    "word/comments.xml": COMMENTS_XML,
    "word/commentsExtended.xml": COMMENTS_EXTENDED_XML,
  });
}
