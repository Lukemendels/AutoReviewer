// Synthetic docx simulating Word's "Remove personal information from file properties on
// save": every w:author is the same placeholder ("Author") and every w:date attribute is
// omitted entirely (not empty -- absent, so wAttr resolves to null). Same
// unzip(base)+writeZip(overrides) technique as tests/helpers/auditFixture.js.
import { unzip } from "../../src/zip/reader.js";
import { writeZip } from "../../src/zip/writer.js";
import { loadDocxBytes } from "./docx.js";

export const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t xml:space="preserve">Text with </w:t></w:r><w:ins w:id="1" w:author="Author"><w:r><w:t>an insertion</w:t></w:r></w:ins><w:r><w:t>.</w:t></w:r></w:p>
<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Commented text.</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>
<w:sectPr/>
</w:body>
</w:document>`;

export const COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
<w:comment w:id="0" w:author="Author"><w:p w14:paraId="20000000"><w:r><w:t>A comment.</w:t></w:r></w:p></w:comment>
</w:comments>`;

export const COMMENTS_EXTENDED_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
<w15:commentEx w15:paraId="20000000" w15:done="0"/>
</w15:commentsEx>`;

export async function buildStrippedMetadataFixtureDocx() {
  const zip = await unzip(loadDocxBytes("comments-threaded"));
  return writeZip(zip, {
    "word/document.xml": DOCUMENT_XML,
    "word/comments.xml": COMMENTS_XML,
    "word/commentsExtended.xml": COMMENTS_EXTENDED_XML,
  });
}
