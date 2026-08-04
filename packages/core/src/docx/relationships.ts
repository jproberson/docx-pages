import { MAIN_DOCUMENT_PART, partXml, type DocxPackage } from "./package.js";
import { W_NS } from "./section.js";
import { attribute, descendantsNamed } from "./xml.js";

export const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

export type Relationship = {
  readonly id: string;
  readonly type: string;
  readonly part: string;
};

export function relationshipsPartFor(part: string): string {
  const cut = part.lastIndexOf("/");
  const directory = cut === -1 ? "" : part.slice(0, cut + 1);
  const file = cut === -1 ? part : part.slice(cut + 1);
  return `${directory}_rels/${file}.rels`;
}

function resolveTarget(owningPart: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const cut = owningPart.lastIndexOf("/");
  const segments = (cut === -1 ? "" : owningPart.slice(0, cut)).split("/").filter(Boolean);
  for (const segment of target.split("/")) {
    if (segment === "" || segment === ".") continue;
    if (segment === "..") segments.pop();
    else segments.push(segment);
  }
  return segments.join("/");
}

export function readRelationships(
  pkg: DocxPackage,
  part: string,
): ReadonlyMap<string, Relationship> {
  const relsPart = relationshipsPartFor(part);
  const found = new Map<string, Relationship>();
  if (!pkg.parts.has(relsPart)) return found;

  for (const node of descendantsNamed(partXml(pkg, relsPart), PKG_REL_NS, "Relationship")) {
    const id = attribute(node, "", "Id");
    const target = attribute(node, "", "Target");
    if (id === undefined || target === undefined) continue;
    if ((attribute(node, "", "TargetMode") ?? "Internal") === "External") continue;
    found.set(id, {
      id,
      type: attribute(node, "", "Type") ?? "",
      part: resolveTarget(part, target),
    });
  }
  return found;
}

function defaultReferencedPart(pkg: DocxPackage, reference: string): string | null {
  const root = partXml(pkg, MAIN_DOCUMENT_PART);
  const references = descendantsNamed(root, W_NS, reference).filter(
    (node) => (attribute(node, W_NS, "type") ?? "default") === "default",
  );
  const found = references.at(-1);
  if (found === undefined) return null;

  const id = attribute(found, R_NS, "id");
  if (id === undefined) return null;

  const relationship = readRelationships(pkg, MAIN_DOCUMENT_PART).get(id);
  if (relationship === undefined || !pkg.parts.has(relationship.part)) return null;
  return relationship.part;
}

export const defaultHeaderPart = (pkg: DocxPackage): string | null =>
  defaultReferencedPart(pkg, "headerReference");

export const defaultFooterPart = (pkg: DocxPackage): string | null =>
  defaultReferencedPart(pkg, "footerReference");
