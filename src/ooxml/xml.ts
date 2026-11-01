/**
 * String-level XML surgery.
 *
 * No DOM parse, deliberately. `DOMParser` would round-trip through a
 * serialiser that reorders namespace declarations and rewrites self-closing
 * tags, and Word is fussy enough about its own parts that a "clean" file can
 * come back unopenable. Touching only the bytes we mean to touch, and leaving
 * every other byte exactly where it was, is both safer and easier to reason
 * about — at the cost of the element walker below, which has to match tags
 * itself.
 */

/** One element found by {@link replaceElements}. */
export type Element = { whole: string; inner: string; open: string };

/**
 * Rewrites every top-level `<tag>…</tag>` in `xml` through `fn`, correctly
 * skipping same-named elements nested inside one it has already handed over
 * (a `w:p` inside a text box lives inside another `w:p`, so a non-greedy
 * regex would close the outer paragraph at the inner one's end tag and
 * shred the file).
 */
export function replaceElements(
  xml: string,
  tag: string,
  fn: (el: Element) => string,
): string {
  const opener = new RegExp(`<${tag}(?=[\\s/>])`, "g");
  const boundary = new RegExp(`<${tag}(?=[\\s/>])|</${tag}>`, "g");
  let out = "";
  let cursor = 0;
  let match: RegExpExecArray | null;

  while ((match = opener.exec(xml))) {
    const start = match.index;
    const openEnd = xml.indexOf(">", start);
    if (openEnd < 0) break;

    const open = xml.slice(start, openEnd + 1);
    let end: number;
    let inner: string;

    if (open.endsWith("/>")) {
      end = openEnd + 1;
      inner = "";
    } else {
      let depth = 1;
      boundary.lastIndex = openEnd + 1;
      let hit: RegExpExecArray | null;
      let close = -1;
      while ((hit = boundary.exec(xml))) {
        if (hit[0].startsWith("</")) {
          if (--depth === 0) {
            close = hit.index;
            break;
          }
        } else if (xml[xml.indexOf(">", hit.index) - 1] !== "/") {
          depth++;
        }
      }
      if (close < 0) break; // malformed; leave the rest untouched
      inner = xml.slice(openEnd + 1, close);
      end = close + tag.length + 3;
    }

    out += xml.slice(cursor, start) + fn({ whole: xml.slice(start, end), inner, open });
    cursor = end;
    opener.lastIndex = end;
  }

  return out + xml.slice(cursor);
}

/** Deletes every match, returning the new string and how many went. */
export function drop(xml: string, re: RegExp): { xml: string; count: number } {
  const hits = xml.match(re);
  return { xml: hits ? xml.replace(re, "") : xml, count: hits?.length ?? 0 };
}

/** Deletes `<tag …>…</tag>` and `<tag …/>` alike, nesting-aware. */
export function dropElement(xml: string, tag: string): { xml: string; count: number } {
  let count = 0;
  const out = replaceElements(xml, tag, () => {
    count++;
    return "";
  });
  return { xml: out, count };
}

/** Unwraps an element, keeping its children — how an accepted insertion or a
    completed move is flattened back into ordinary content. */
export function unwrapElement(xml: string, tag: string): { xml: string; count: number } {
  let count = 0;
  const out = replaceElements(xml, tag, (el) => {
    count++;
    return el.open.endsWith("/>") ? "" : el.inner;
  });
  return { xml: out, count };
}

const ENTITIES: Record<string, string> = {
  "&lt;": "<",
  "&gt;": ">",
  "&amp;": "&",
  "&quot;": '"',
  "&apos;": "'",
};

/**
 * Text rules run on real characters, not on markup — so a `&#x2019;` numeric
 * reference has to become a curly apostrophe before it can be straightened.
 * Only `&`, `<` and `>` go back on the way out; everything else stays literal
 * UTF-8, which is what Word writes anyway.
 */
export function decodeText(text: string): string {
  return text
    .replace(/&(?:lt|gt|amp|quot|apos);/g, (m) => ENTITIES[m])
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(parseInt(dec, 10)));
}

export function encodeText(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** The elements that hold visible text in both formats: a Word run, a
    DrawingML run (every PowerPoint shape, and Word charts/SmartArt), and an
    equation run. Anything else in the XML is structure and stays untouched. */
// The attribute group ends on a character that is neither "/" nor ">", which
// is what excludes an empty `<w:t/>` — a lookbehind would read better but
// would also be a parse error, and so a blank page, on older Safari.
const TEXT_NODE = /(<(w:t|a:t|m:t)((?:\s[^>]*[^/>])?)>)([^<]*)(<\/\2>)/g;

/**
 * Maps `transform` over every text node's decoded content.
 *
 * Adds `xml:space="preserve"` when a rewrite leaves a leading or trailing
 * space — without it Word trims the run, and stripping an emoji off the end of
 * "Results 🚀 next" would silently close the gap to "Resultsnext".
 */
export function mapText(xml: string, transform: (text: string) => string): string {
  return xml.replace(TEXT_NODE, (whole, open: string, tag: string, attrs: string, body: string, close: string) => {
    const cleaned = transform(decodeText(body));
    const encoded = encodeText(cleaned);
    if (encoded === body) return whole;

    const needsPreserve = /^\s|\s$/.test(cleaned) && !attrs.includes("xml:space");
    const newOpen = needsPreserve ? `<${tag}${attrs} xml:space="preserve">` : open;
    return `${newOpen}${encoded}${close}`;
  });
}

/** Concatenated visible text of a fragment — used to judge a paragraph as a
    whole, since one sentence is routinely split across a dozen runs. */
export function textOf(xml: string): string {
  let out = "";
  for (const m of xml.matchAll(TEXT_NODE)) out += decodeText(m[4]);
  return out;
}

/** Resolves a relationship `Target` against the part that declared it, so a
    rel in `word/_rels/document.xml.rels` pointing at `comments.xml` is known
    to mean `word/comments.xml`. */
export function resolvePart(relsPath: string, target: string): string {
  if (/^[a-z]+:\/\//i.test(target)) return target;
  if (target.startsWith("/")) return target.slice(1);

  const base = relsPath.replace(/_rels\/[^/]+$/, "");
  const segments = (base + target).split("/");
  const stack: string[] = [];
  for (const segment of segments) {
    if (segment === "." || segment === "") continue;
    if (segment === "..") stack.pop();
    else stack.push(segment);
  }
  return stack.join("/");
}
