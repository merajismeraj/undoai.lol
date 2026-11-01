/**
 * The de-AI pass over a .docx or .pptx.
 *
 * Runs entirely in the browser: the file is read into memory, rewritten, and
 * handed straight back as a download. Nothing is uploaded, which is the point —
 * the documents people most want de-fingerprinted are exactly the ones they
 * shouldn't be posting to a stranger's server.
 *
 * Three passes, in this order:
 *   1. Structure — comments, tracked changes, hidden runs, editing-session IDs.
 *   2. Prose — assistant asides deleted, then the character rules in `text.ts`.
 *   3. Properties — author, timestamps, generator tags, custom properties.
 *
 * Order matters: revisions go first so a paragraph isn't judged on text a
 * tracked deletion was already hiding, and properties go last because deleting
 * a part means fixing every relationship that pointed at it.
 */

import {
  Tally,
  cleanMarkdown,
  cleanText,
  isFillerParagraph,
  scanCliches,
  stripPreamble,
  type Finding,
} from "./text";
import {
  decodeText,
  drop,
  dropElement,
  mapText,
  replaceElements,
  resolvePart,
  textOf,
  unwrapElement,
  type Element,
} from "./xml";
import { ZipError, makeEntry, readEntry, readZip, writeZip, type ZipEntry } from "./zip";

export type CleanOptions = {
  /** Author, timestamps, generator tags, custom properties, editing IDs. */
  metadata: boolean;
  /** Comments, tracked changes, hidden text. */
  revisions: boolean;
  /** Chat preambles, stray markdown, citation markers. */
  leftovers: boolean;
  /** Optional name to stamp as the author instead of leaving the field empty. */
  author: string;
};

export const DEFAULT_OPTIONS: CleanOptions = {
  metadata: true,
  revisions: true,
  leftovers: true,
  author: "",
};

export type CleanResult = {
  fileName: string;
  kind: "Word document" | "PowerPoint deck";
  blob: Blob;
  originalBytes: number;
  cleanedBytes: number;
  findings: Finding[];
  cliches: Array<{ label: string; count: number }>;
  notes: string[];
};

export const MAX_BYTES = 40 * 1024 * 1024;

/** Names that give the generator away when they turn up in a document
    property, a link, or an Application tag. */
const FINGERPRINT =
  /(chatgpt|openai|gpt-?[0-9]|claude|anthropic|gemini|\bbard\b|copilot|perplexity|jasper|writesonic|quillbot|gamma\.app|tome\.app|beautiful\.ai|python-docx|python-pptx|docxjs|docx-js|pptxgenjs|officegen|aspose|apache poi)/i;

const AI_HOSTS =
  /(chatgpt\.com|chat\.openai\.com|claude\.ai|gemini\.google\.com|copilot\.microsoft\.com|perplexity\.ai)/i;

/** Word tables are lifted out of the document before the paragraph pass runs,
    and an XML comment is the one placeholder guaranteed not to collide with
    anything Office writes into a part. */
const TABLE_SLOT = /<!--tbl:(\d+)-->/g;

const decoder = new TextDecoder();
const encoder = new TextEncoder();

export async function cleanDocument(
  file: File,
  options: CleanOptions = DEFAULT_OPTIONS,
): Promise<CleanResult> {
  if (file.size > MAX_BYTES) {
    throw new ZipError("File is over 40MB. Split it, or strip the media out first.");
  }
  if (/\.(doc|ppt|pps)$/i.test(file.name)) {
    throw new ZipError("That's the old binary format. Open it and save as .docx or .pptx first.");
  }

  const entries = readZip(await file.arrayBuffer());
  const names = new Set(entries.map((e) => e.name));

  const kind = names.has("word/document.xml")
    ? ("Word document" as const)
    : names.has("ppt/presentation.xml")
      ? ("PowerPoint deck" as const)
      : null;
  if (!kind) throw new ZipError("Only Word (.docx) and PowerPoint (.pptx) files are supported.");

  const isWord = kind === "Word document";
  const tally = new Tally();
  const notes: string[] = [];
  const clicheCounts = new Map<string, number>();
  const rewritten = new Map<string, string>();
  const deleted = new Set<string>();

  const read = async (name: string) => {
    const entry = entries.find((e) => e.name === name);
    return entry ? decoder.decode(await readEntry(entry)) : null;
  };

  // 1. Links, before anything reads a content part.
  //
  // A hyperlink lives in two places — a relationship holding the URL, and the
  // element in the body referencing it by id — so both have to be settled
  // together, and the id has to be known before the body is rewritten.
  const aiLinkIds = new Map<string, Set<string>>();
  for (const entry of entries) {
    if (!entry.name.endsWith(".rels")) continue;

    const owner = entry.name.replace(/_rels\/([^/]+)\.rels$/, "$1");
    const ids = new Set<string>();
    const xml = decoder.decode(await readEntry(entry));

    const next = replaceElements(xml, "Relationship", (el) => {
      const raw = el.whole.match(/Target="([^"]*)"/)?.[1] ?? "";
      const target = stripTracking(decodeText(raw));
      if (target !== decodeText(raw)) tally.add("utm");

      if (AI_HOSTS.test(target)) {
        const id = el.whole.match(/Id="([^"]*)"/)?.[1];
        if (id) ids.add(id);
        tally.add("aiLink");
        notes.push(`Link to ${hostOf(target)} removed — the text it was on is untouched.`);
        return "";
      }

      return target === decodeText(raw)
        ? el.whole
        : el.whole.replace(`Target="${raw}"`, `Target="${escapeXml(target)}"`);
    });

    if (ids.size > 0) aiLinkIds.set(owner, ids);
    if (next !== xml) rewritten.set(entry.name, next);
  }

  // 2. Content parts.
  //
  // Every XML part under word/ or ppt/ goes through the same pipeline. Wider
  // than "the document body" on purpose — headers, footers, footnotes, speaker
  // notes, text boxes, SmartArt and chart labels are all places pasted output
  // ends up. The text rules only ever touch run-level text nodes, so parts with
  // no prose in them (styles, themes) come back byte-identical and get passed
  // through untouched when the archive is rebuilt.
  for (const entry of entries) {
    if (!/^(word|ppt)\/.*\.xml$/.test(entry.name)) continue;
    if (isDeletablePart(entry.name)) continue;

    let xml = decoder.decode(await readEntry(entry));
    const before = xml;

    const links = aiLinkIds.get(entry.name);
    if (links) xml = unlink(xml, links);

    if (options.revisions) xml = stripRevisions(xml, tally, isWord);
    if (options.metadata) xml = stripEditingIds(xml, tally);
    if (options.leftovers) xml = stripFiller(xml, tally, isWord);

    scanCliches(textOf(xml), clicheCounts);

    xml = mapText(xml, (text) => {
      let out = options.leftovers ? cleanMarkdown(text, tally) : text;
      out = cleanText(out, tally);
      return options.leftovers ? stripPreamble(out, tally) : out;
    });

    if (xml !== before) rewritten.set(entry.name, xml);
  }

  // 3. Parts that are pure conversation, or pure metadata.
  for (const entry of entries) {
    if (!isDeletablePart(entry.name)) continue;
    const isComment = isCommentPart(entry.name);
    if (isComment ? !options.revisions : !options.metadata) continue;
    deleted.add(entry.name);
    tally.add(isComment ? "comments" : "metaCustom");
  }

  if (options.metadata) {
    for (const entry of entries) {
      if (/^docProps\/thumbnail\./i.test(entry.name)) {
        deleted.add(entry.name);
        tally.add("metaThumb");
      }
    }

    const custom = await read("docProps/custom.xml");
    if (custom && FINGERPRINT.test(custom)) {
      notes.push("Custom document properties carried a generator fingerprint. Removed.");
    }

    const core = await read("docProps/core.xml");
    if (core) rewritten.set("docProps/core.xml", cleanCore(core, options.author, tally, notes));

    const app = await read("docProps/app.xml");
    if (app) rewritten.set("docProps/app.xml", cleanApp(app, tally, notes));
  }

  // 4. Relationships and the content-type map, so nothing is left pointing at
  //    a part that no longer exists.
  for (const entry of entries) {
    if (!entry.name.endsWith(".rels") || deleted.has(entry.name)) continue;

    const xml = rewritten.get(entry.name) ?? decoder.decode(await readEntry(entry));
    const next = replaceElements(xml, "Relationship", (el) => {
      const target = el.whole.match(/Target="([^"]*)"/)?.[1] ?? "";
      return deleted.has(resolvePart(entry.name, target)) ? "" : el.whole;
    });

    if (next !== xml) rewritten.set(entry.name, next);
  }

  if (deleted.size > 0) {
    // A deleted part's own .rels file is meaningless without it.
    for (const name of [...deleted]) {
      const rels = name.replace(/([^/]+)$/, "_rels/$1.rels");
      if (names.has(rels)) deleted.add(rels);
    }

    const types = await read("[Content_Types].xml");
    if (types) {
      rewritten.set(
        "[Content_Types].xml",
        replaceElements(types, "Override", (el) => {
          const part = el.whole.match(/PartName="\/([^"]*)"/)?.[1] ?? "";
          return deleted.has(part) ? "" : el.whole;
        }),
      );
    }
  }

  // 5. Rebuild. Untouched entries keep their original compressed bytes rather
  //    than being re-deflated, so images and fonts survive bit for bit.
  const output: ZipEntry[] = [];
  for (const entry of entries) {
    if (deleted.has(entry.name)) continue;
    const replacement = rewritten.get(entry.name);
    output.push(replacement ? await makeEntry(entry.name, encoder.encode(replacement)) : entry);
  }

  const blob = writeZip(output);
  const cliches = [...clicheCounts.entries()]
    .map(([label, count]) => ({ label, count }))
    .sort((a, b) => b.count - a.count);
  if (cliches.length > 0) {
    tally.add(
      "cliche",
      cliches.reduce((sum, c) => sum + c.count, 0),
    );
  }

  return {
    fileName: outputName(file.name),
    kind,
    blob,
    originalBytes: file.size,
    cleanedBytes: blob.size,
    findings: tally.findings(),
    cliches,
    notes: [...new Set(notes)],
  };
}

/** Drops a hyperlink to an AI tool while keeping the words it sat on: the Word
    element is unwrapped around its runs, and PowerPoint's click action — which
    hangs off a run's properties rather than wrapping it — is deleted. */
function unlink(xml: string, ids: Set<string>): string {
  const owned = (fragment: string) => {
    const id = fragment.match(/r:id="([^"]*)"/)?.[1];
    return !!id && ids.has(id);
  };

  let out = replaceElements(xml, "w:hyperlink", (el) => (owned(el.open) ? el.inner : el.whole));
  out = replaceElements(out, "a:hlinkClick", (el) => (owned(el.open) ? "" : el.whole));
  return out;
}

function hostOf(url: string): string {
  return url.replace(/^https?:\/\//, "").split("/")[0];
}

function outputName(name: string): string {
  const dot = name.lastIndexOf(".");
  return dot < 0 ? `${name}-cleaned` : `${name.slice(0, dot)}-cleaned${name.slice(dot)}`;
}

/** Parts that exist only to record who said what, or which tool made the file.
    None of them carry document content, so they leave whole. */
function isDeletablePart(name: string): boolean {
  return isCommentPart(name) || name === "docProps/custom.xml" || /^customXml\//.test(name);
}

function isCommentPart(name: string): boolean {
  return (
    /^word\/comments([A-Za-z]+)?\.xml$/.test(name) ||
    name === "word/people.xml" ||
    /^ppt\/comments\//.test(name) ||
    name === "ppt/authors.xml"
  );
}

/** Drops the review layer: tracked changes flattened to their final text,
    comments and their anchors gone, and hidden runs — a tidy place to park a
    watermark, since nobody sees them — deleted outright. */
function stripRevisions(xml: string, tally: Tally, isWord: boolean): string {
  if (!isWord) return xml;
  let out = xml;

  // Deletions first: an insertion nested inside a deletion never happened.
  for (const tag of [
    "w:del",
    "w:moveFrom",
    "w:pPrChange",
    "w:rPrChange",
    "w:sectPrChange",
    "w:tblPrChange",
    "w:tcPrChange",
    "w:trPrChange",
    "w:tblGridChange",
    "w:cellDel",
    "w:cellIns",
    "w:cellMerge",
  ]) {
    const res = dropElement(out, tag);
    out = res.xml;
    tally.add("tracked", res.count);
  }

  for (const tag of ["w:ins", "w:moveTo"]) {
    const res = unwrapElement(out, tag);
    out = res.xml;
    tally.add("tracked", res.count);
  }

  const ranges = drop(
    out,
    /<w:(?:moveFromRangeStart|moveFromRangeEnd|moveToRangeStart|moveToRangeEnd|customXmlInsRangeStart|customXmlInsRangeEnd|customXmlDelRangeStart|customXmlDelRangeEnd|commentRangeStart|commentRangeEnd)\b[^>]*\/>/g,
  );
  out = ranges.xml;
  tally.add("tracked", ranges.count);

  out = replaceElements(out, "w:r", (el) => {
    if (el.whole.includes("<w:commentReference")) {
      tally.add("comments");
      return "";
    }
    if (/<w:vanish\s*\/>|<w:vanish w:val="(?:1|true|on)"\s*\/>/.test(el.inner)) {
      tally.add("hidden");
      return "";
    }
    return el.whole;
  });

  const proof = drop(out, /<w:(?:proofErr|lastRenderedPageBreak|proofState)\b[^>]*\/>/g);
  out = proof.xml;
  tally.add("proof", proof.count);

  // A run emptied by the removals above renders nothing; leaving it in is
  // harmless but it is exactly the sort of debris that says "edited by script".
  return out.replace(/<w:r><\/w:r>/g, "");
}

/** Editing-session IDs. Word stamps every paragraph and run with an rsid and
    PowerPoint stamps every shape with a creation GUID; together they describe
    how the file was assembled — one long generated burst looks nothing like a
    document somebody typed over a week. */
function stripEditingIds(xml: string, tally: Tally): string {
  let out = xml;

  const attrs = drop(out, /\s(?:w:rsid[A-Za-z]*|w14:paraId|w14:textId)="[^"]*"/g);
  out = attrs.xml;
  tally.add("docIds", attrs.count);

  const rsids = dropElement(out, "w:rsids");
  out = rsids.xml;
  tally.add("docIds", rsids.count);

  const docIds = drop(out, /<w15:docId\b[^>]*\/>/g);
  out = docIds.xml;
  tally.add("docIds", docIds.count);

  // PowerPoint keeps its GUIDs inside extension blocks; drop the block holding
  // one, then any extension list left standing empty.
  for (const tag of ["a:ext", "p:ext"]) {
    out = replaceElements(out, tag, (el) => {
      if (!/creationId|modId/i.test(el.inner)) return el.whole;
      tally.add("docIds");
      return "";
    });
  }
  out = out.replace(/<(a|p):extLst>\s*<\/\1:extLst>/g, "");

  return out;
}

/**
 * Deletes paragraphs that are the assistant talking rather than the document
 * speaking. Both formats require at least one paragraph inside a table cell and
 * inside a PowerPoint text body, so those two containers are pruned first and
 * always keep a survivor; Word tables are then lifted out of the body-level
 * pass so it can't reach back into a cell already settled.
 */
function stripFiller(xml: string, tally: Tally, isWord: boolean): string {
  if (!isWord) {
    return replaceElements(xml, "p:txBody", (el) => pruneParagraphs(el.whole, "a:p", tally));
  }

  const tables: string[] = [];
  let out = replaceElements(xml, "w:tbl", (el) => {
    tables.push(
      replaceElements(el.whole, "w:tc", (cell) => pruneParagraphs(cell.whole, "w:p", tally)),
    );
    return `<!--tbl:${tables.length - 1}-->`;
  });

  out = replaceElements(out, "w:p", (el) => {
    if (!isFillerParagraph(textOf(el.whole))) return el.whole;
    tally.add("filler");
    // The final paragraph of a section carries the page setup, so that one is
    // emptied rather than deleted — losing it would take the margins and the
    // page orientation with it.
    return el.whole.includes("<w:sectPr") ? emptyParagraph(el, "w:p", "w:pPr") : "";
  });

  return out.replace(TABLE_SLOT, (_, i: string) => tables[Number(i)]);
}

/** Removes filler paragraphs from a container that must not end up empty. When
    the aside *is* the whole cell or text box, the paragraph shell stays and only
    its words go — a table cell with no paragraph in it won't open. */
function pruneParagraphs(scope: string, tag: string, tally: Tally): string {
  const props = tag === "w:p" ? "w:pPr" : "a:pPr";
  let remaining = 0;
  let filler = 0;
  replaceElements(scope, tag, (el) => {
    remaining++;
    if (isFillerParagraph(textOf(el.whole))) filler++;
    return el.whole;
  });
  if (filler === 0) return scope;

  return replaceElements(scope, tag, (el) => {
    if (!isFillerParagraph(textOf(el.whole))) return el.whole;
    tally.add("filler");
    if (remaining <= 1) return emptyParagraph(el, tag, props);
    remaining--;
    return "";
  });
}

/** The same paragraph with its runs gone and its formatting kept. */
function emptyParagraph(el: Element, tag: string, props: string): string {
  const kept = el.inner.match(
    new RegExp(`<${props}(?:\\s[^>]*)?/>|<${props}(?:\\s[^>]*)?>[\\s\\S]*?</${props}>`),
  );
  return `${el.open}${kept?.[0] ?? ""}</${tag}>`;
}

/**
 * `docProps/core.xml` — author, last editor, revision count, timestamps.
 *
 * Cleared rather than back-dated or given a plausible fake history: the goal is
 * a document that says nothing about how it was made, not one that tells a
 * different story about it. Title and subject are the user's own words and stay.
 */
function cleanCore(xml: string, author: string, tally: Tally, notes: string[]): string {
  let out = xml;

  const clear = (tag: string, key: string) => {
    const value = decodeText(
      out.match(new RegExp(`<${tag}\\b[^>]*>([^<]*)</${tag}>`))?.[1] ?? "",
    ).trim();
    if (value && FINGERPRINT.test(value)) {
      notes.push(`${tag.split(":")[1]} was "${value}" — cleared.`);
    }
    out = dropElement(out, tag).xml;
    if (value) tally.add(key);
  };

  clear("dc:creator", "metaAuthor");
  clear("cp:lastModifiedBy", "metaAuthor");
  clear("cp:revision", "metaOther");
  clear("dcterms:created", "metaDates");
  clear("dcterms:modified", "metaDates");
  clear("cp:lastPrinted", "metaDates");
  clear("cp:keywords", "metaOther");
  clear("dc:description", "metaOther");
  clear("cp:category", "metaOther");
  clear("cp:contentStatus", "metaOther");

  // Title and subject are the user's own words and stay — unless the generator
  // signed them, which pptxgenjs and friends routinely do.
  for (const tag of ["dc:title", "dc:subject"]) {
    const value = decodeText(
      out.match(new RegExp(`<${tag}\\b[^>]*>([^<]*)</${tag}>`))?.[1] ?? "",
    ).trim();
    if (!value || !FINGERPRINT.test(value)) continue;
    notes.push(`${tag.split(":")[1]} was "${value}" — cleared.`);
    out = dropElement(out, tag).xml;
    tally.add("metaOther");
  }

  const name = author.trim();
  if (name) {
    const stamp = `<dc:creator>${escapeXml(name)}</dc:creator><cp:lastModifiedBy>${escapeXml(name)}</cp:lastModifiedBy>`;
    out = out.replace("</cp:coreProperties>", `${stamp}</cp:coreProperties>`);
  }

  return out;
}

/** `docProps/app.xml` — "Application" is where python-pptx, docx-js and the
    rest of the generation toolchain sign their work. */
function cleanApp(xml: string, tally: Tally, notes: string[]): string {
  let out = xml;

  for (const tag of ["Application", "AppVersion", "Company", "Manager", "Template", "TotalTime"]) {
    const value = decodeText(out.match(new RegExp(`<${tag}>([^<]*)</${tag}>`))?.[1] ?? "").trim();
    if (value && FINGERPRINT.test(value)) {
      notes.push(`Properties said the file was made by "${value}". Tag removed.`);
    }
    out = dropElement(out, tag).xml;
    if (value) tally.add("metaApp");
  }

  return out;
}

/** Strips the `utm_source=chatgpt.com` tail appended to every link a chat
    assistant hands over — the most common giveaway in a document that
    otherwise reads clean. */
function stripTracking(target: string): string {
  if (!/utm_source=[^&"]*(?:chatgpt|openai|claude|gemini|copilot|perplexity)/i.test(target)) {
    return target;
  }

  return target
    .replace(/([?&])utm_(?:source|medium|campaign|content|term)=[^&"]*/gi, "$1")
    .replace(/&{2,}/g, "&")
    .replace(/\?&/, "?")
    .replace(/[?&]+$/, "");
}

function escapeXml(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
