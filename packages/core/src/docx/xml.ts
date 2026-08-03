import { XMLParser } from "fast-xml-parser";

export type XmlElement = {
  readonly name: string;
  readonly namespace: string | null;
  readonly attributes: ReadonlyMap<string, string>;
  readonly children: readonly XmlElement[];
  readonly text: string;
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

function convert(node: RawNode, inherited: ReadonlyMap<string, string>): XmlElement | null {
  const tag = Object.keys(node).find(isElementTag);
  if (tag === undefined) return null;

  const declared = rawAttributes(node);
  const namespaces = bindNamespaces(declared, inherited);
  const [prefix, local] = splitName(tag);

  const body = node[tag];
  const children: XmlElement[] = [];
  let text = "";
  if (isNodeArray(body)) {
    for (const child of body) {
      const value = child[TEXT_KEY];
      if (typeof value === "string") text += value;
      const converted = convert(child, namespaces);
      if (converted !== null) children.push(converted);
    }
  }

  return {
    name: local,
    namespace: namespaces.get(prefix) ?? null,
    attributes: qualifyAttributes(declared, namespaces),
    children,
    text,
  };
}

export function parseXml(source: string): XmlElement | null {
  const parsed: unknown = parser.parse(source);
  if (!isNodeArray(parsed)) return null;
  for (const node of parsed) {
    const element = convert(node, new Map());
    if (element !== null) return element;
  }
  return null;
}

// An empty namespace means an unprefixed attribute, which XML puts in no namespace
// at all rather than in the element's default one.
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
