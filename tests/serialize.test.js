// ooxml/serialize.js (spec §9.3): thin XMLSerializer wrapper, no pretty-printing. Written
// red before the implementation.
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { describe, expect, it } from "vitest";
import { parseXml } from "../src/ooxml/parse.js";
import { serializePart } from "../src/ooxml/serialize.js";

describe("serializePart", () => {
  it("round-trips a parsed document back to an equivalent XML string, declaration included", () => {
    const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:p xmlns:w="urn:w"><w:r><w:t>hello</w:t></w:r></w:p>';
    const doc = parseXml(xml, DOMParser);
    const out = serializePart(doc, XMLSerializer);
    expect(out).toContain('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    expect(out).toContain("<w:t>hello</w:t>");
  });

  it("does not reformat or reindent -- whitespace inside w:t is preserved exactly, including xml:space=\"preserve\" runs", () => {
    const xml = '<w:p xmlns:w="urn:w"><w:r><w:t xml:space="preserve">  two leading spaces</w:t></w:r></w:p>';
    const doc = parseXml(xml, DOMParser);
    const out = serializePart(doc, XMLSerializer);
    expect(out).toContain('<w:t xml:space="preserve">  two leading spaces</w:t>');
    // No inserted newlines/indentation between sibling elements.
    expect(out).not.toMatch(/>\s*\n\s*</);
  });

  it("serializes a mutated DOM (post-injectEdits shape) without introducing any formatting", () => {
    const xml = '<w:p xmlns:w="urn:w"><w:r><w:t>before</w:t></w:r></w:p>';
    const doc = parseXml(xml, DOMParser);
    const p = doc.documentElement;
    const ins = doc.createElementNS("urn:w", "w:ins");
    const r = doc.createElementNS("urn:w", "w:r");
    const t = doc.createElementNS("urn:w", "w:t");
    t.appendChild(doc.createTextNode("after"));
    r.appendChild(t);
    ins.appendChild(r);
    p.appendChild(ins);

    const out = serializePart(doc, XMLSerializer);
    expect(out).toBe('<w:p xmlns:w="urn:w"><w:r><w:t>before</w:t></w:r><w:ins><w:r><w:t>after</w:t></w:r></w:ins></w:p>');
  });
});
