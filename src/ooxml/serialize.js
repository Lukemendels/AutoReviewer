// XMLSerializer + part replacement (spec §9.3). Thin wrapper -- no pretty-printing to opt
// out of, since standard XMLSerializer doesn't reformat. XMLSerializerImpl is injected per
// environment (browser-native in the built tool, @xmldom/xmldom in tests), matching how
// DOMParserImpl is threaded through ooxml/parse.js.
export function serializePart(xmlDoc, XMLSerializerImpl) {
  const Ctor = XMLSerializerImpl || globalThis.XMLSerializer;
  return new Ctor().serializeToString(xmlDoc);
}
