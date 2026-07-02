// NS helpers + rels parsing, ported from ref/redline-to-markdown.html. DOMParser is
// injected per environment: browser-native in the built tool, @xmldom/xmldom in tests.

export const NS = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  w14: "http://schemas.microsoft.com/office/word/2010/wordml",
  w15: "http://schemas.microsoft.com/office/word/2012/wordml",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
};

export const wAttr = (el, n) => el.getAttributeNS(NS.w, n) ?? el.getAttribute("w:" + n);
export const w14Attr = (el, n) => el.getAttributeNS(NS.w14, n) ?? el.getAttribute("w14:" + n);
export const w15Attr = (el, n) => el.getAttributeNS(NS.w15, n) ?? el.getAttribute("w15:" + n);
export const rAttr = (el, n) => el.getAttributeNS(NS.r, n) ?? el.getAttribute("r:" + n);
export const kids = (el, ln) => [...el.children].filter((c) => c.localName === ln);
export const kid = (el, ln) => [...el.children].find((c) => c.localName === ln) || null;
export const basename = (p) => String(p).split("/").pop().split("\\").pop();
export const fmtDate = (d) => {
  if (!d) return "";
  const t = d.indexOf("T");
  return t > 0 ? d.slice(0, t) : d;
};

export function parseXml(xml, DOMParserCtor) {
  const Ctor = DOMParserCtor || globalThis.DOMParser;
  return new Ctor().parseFromString(xml, "application/xml");
}

export function parseRels(xml, DOMParserCtor) {
  const map = {};
  if (!xml) return map;
  const doc = parseXml(xml, DOMParserCtor);
  for (const rel of doc.getElementsByTagName("*")) {
    if (rel.localName === "Relationship") {
      const id = rel.getAttribute("Id"), target = rel.getAttribute("Target");
      if (id) map[id] = target;
    }
  }
  return map;
}
