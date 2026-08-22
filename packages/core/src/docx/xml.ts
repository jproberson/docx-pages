import { XMLParser } from "fast-xml-parser";

export type XmlElement = {
  readonly name: string;
  readonly namespace: string | null;
  readonly attributes: ReadonlyMap<string, string>;
  readonly children: readonly XmlElement[];
  readonly text: string;
  // Whether the whitespace written inside this element is the text's own, which
  // `xml:space` states for an element and everything under it until something
  // nearer states otherwise. A document may say it once on its own root and never
  // again, and one in the corpus does.
  readonly preservesSpace: boolean;
};

const ATTRIBUTE_KEY = ":@";
const TEXT_KEY = "#text";

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "",
  attributesGroupName: ATTRIBUTE_KEY,
  preserveOrder: true,
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
  ignoreDeclaration: true,
  ignorePiTags: true,
  // Named entities alone leave &#xF0A7; standing as eight characters, which is
  // how a symbol bullet arrives when a producer escapes it.
  htmlEntities: true,
});

type RawNode = Record<string, unknown>;

const isRecord = (value: unknown): value is RawNode =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const isNodeArray = (value: unknown): value is readonly RawNode[] =>
  Array.isArray(value) && value.every(isRecord);

const splitName = (qualified: string): readonly [prefix: string, local: string] => {
  const colon = qualified.indexOf(":");
  return colon === -1 ? ["", qualified] : [qualified.slice(0, colon), qualified.slice(colon + 1)];
};

export const clark = (namespace: string, name: string): string => `{${namespace}}${name}`;

function rawAttributes(node: RawNode): ReadonlyMap<string, string> {
  const group = node[ATTRIBUTE_KEY];
  const found = new Map<string, string>();
  if (!isRecord(group)) return found;
  for (const [key, value] of Object.entries(group)) {
    if (typeof value === "string") found.set(key, value);
    else if (typeof value === "number" || typeof value === "boolean") {
      found.set(key, String(value));
    }
  }
  return found;
}

function bindNamespaces(
  attributes: ReadonlyMap<string, string>,
  inherited: ReadonlyMap<string, string>,
): ReadonlyMap<string, string> {
  let bindings: Map<string, string> | null = null;
  for (const [key, value] of attributes) {
    const [prefix, local] = splitName(key);
    const declaresPrefixed = prefix === "xmlns";
    const declaresDefault = prefix === "" && local === "xmlns";
    if (!declaresPrefixed && !declaresDefault) continue;
    bindings ??= new Map(inherited);
    bindings.set(declaresDefault ? "" : local, value);
  }
  return bindings ?? inherited;
}

function qualifyAttributes(
  attributes: ReadonlyMap<string, string>,
  namespaces: ReadonlyMap<string, string>,
): ReadonlyMap<string, string> {
  const qualified = new Map<string, string>();
  for (const [key, value] of attributes) {
    const [prefix, local] = splitName(key);
    if (prefix === "xmlns" || (prefix === "" && local === "xmlns")) continue;
    // An unprefixed attribute is in no namespace, never the default one.
    const namespace = prefix === "" ? undefined : namespaces.get(prefix);
    qualified.set(namespace === undefined ? local : clark(namespace, local), value);
  }
  return qualified;
}

const isElementTag = (key: string): boolean =>
  key !== ATTRIBUTE_KEY && key !== TEXT_KEY && !key.startsWith("?") && !key.startsWith("!");

function convert(
  node: RawNode,
  inherited: ReadonlyMap<string, string>,
  preserved: boolean,
): XmlElement | null {
  const tag = Object.keys(node).find(isElementTag);
  if (tag === undefined) return null;

  const declared = rawAttributes(node);
  const namespaces = bindNamespaces(declared, inherited);
  const [prefix, local] = splitName(tag);

  const stated = declared.get("xml:space");
  const preservesSpace = stated === undefined ? preserved : stated === "preserve";

  const body = node[tag];
  const children: XmlElement[] = [];
  let text = "";
  if (isNodeArray(body)) {
    for (const child of body) {
      const value = child[TEXT_KEY];
      if (typeof value === "string") text += value;
      const converted = convert(child, namespaces, preservesSpace);
      if (converted !== null) children.push(converted);
    }
  }

  return {
    name: local,
    namespace: namespaces.get(prefix) ?? null,
    attributes: qualifyAttributes(declared, namespaces),
    children,
    text,
    preservesSpace,
  };
}

export function parseXml(source: string): XmlElement | null {
  const parsed: unknown = parser.parse(source);
  if (!isNodeArray(parsed)) return null;
  for (const node of parsed) {
    const element = convert(node, new Map(), false);
    if (element !== null) return element;
  }
  return null;
}

// The number an attribute states, and NaN where it states none this can read.
//
// **An attribute written out empty states a nought, not nothing.** Asked of Word on
// 2026-08-22 with a style stating 12pt after, a 36pt indent and 20pt text: a
// paragraph writing `w:after=""` was set 24.48pt from the next as `w:after="0"` is
// and not the 36.48pt of one inheriting, `w:ind w:left=""` sat on the margin where
// the stated nought sits, and `w:sz w:val=""` drew at 0.48pt beside the stated
// nought's own 0.48pt. Only an absent attribute inherits.
export const statedNumber = (raw: string | undefined): number =>
  raw === undefined ? Number.NaN : Number(raw);

// The three spellings of off an on/off value has. Everything else is on, "1", "true"
// and "on" among them, so one place answers for a toggle written as an element and
// for one written as an attribute of its own.
export const statesOn = (value: string): boolean =>
  value !== "0" && value !== "false" && value !== "off";

// An empty namespace means an unprefixed attribute, which XML puts in no namespace
// at all rather than in the element's default one.
// Whether an element that is a toggle is on. Word writes one bare to turn it on,
// so an element with no value at all counts, and only an explicit off turns it off.
export const toggledOn = (element: XmlElement, namespace: string): boolean => {
  const value = attribute(element, namespace, "val");
  return value === undefined || statesOn(value);
};

export const attribute = (
  element: XmlElement,
  namespace: string,
  name: string,
): string | undefined => element.attributes.get(namespace === "" ? name : clark(namespace, name));

export const childrenNamed = (
  element: XmlElement,
  namespace: string,
  name: string,
): readonly XmlElement[] =>
  element.children.filter((child) => child.namespace === namespace && child.name === name);

export const firstNamed = (
  element: XmlElement,
  namespace: string,
  name: string,
): XmlElement | null => childrenNamed(element, namespace, name)[0] ?? null;

export function descendantsNamed(
  element: XmlElement,
  namespace: string,
  name: string,
): readonly XmlElement[] {
  const found: XmlElement[] = [];
  const visit = (node: XmlElement): void => {
    if (node.namespace === namespace && node.name === name) found.push(node);
    for (const child of node.children) visit(child);
  };
  visit(element);
  return found;
}
