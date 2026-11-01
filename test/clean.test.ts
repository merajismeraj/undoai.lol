import assert from "node:assert/strict";
import test from "node:test";
import mammoth from "mammoth";
import { parseXml } from "@rgrove/parse-xml";

import { DEFAULT_OPTIONS, cleanDocument } from "../src/ooxml/clean";
import { cleanText, isFillerParagraph, Tally } from "../src/ooxml/text";
import { mapText, replaceElements, resolvePart, textOf } from "../src/ooxml/xml";
import { crc32, makeEntry, readEntry, readZip, writeZip } from "../src/ooxml/zip";
import {
  GENERATED_THEME,
  REVIEWED_BODY,
  makeDocx,
  makePptx,
  partNames,
  readParts,
  withPart,
} from "./fixtures";

/** Characters that must not survive a clean, whatever the file. */
const TRACES = new RegExp(
  [
    "[\\u2018\\u2019\\u201C\\u201D]", // curly quotes
    "[\\u2013\\u2014\\u2026]", // dashes and ellipsis
    "[\\u00A0\\u200B-\\u200D\\uFEFF]", // non-breaking and zero-width spaces
    "[\\u2022\\u2192\\u00D7]", // bullet, arrow, multiplication sign
    "[\\u{1F000}-\\u{1FAFF}]", // emoji
    "[\\u{E0000}-\\u{E007F}]", // tag characters
  ].join("|"),
  "u",
);

async function visibleText(blob: Blob, match: RegExp): Promise<string> {
  const parts = await readParts(blob);
  let out = "";
  for (const [name, xml] of parts) if (match.test(name)) out += `${textOf(xml)}\n`;
  return out;
}

function cellsOf(xml: string): string[] {
  const cells: string[] = [];
  replaceElements(xml, "w:tc", (el) => {
    cells.push(el.whole);
    return el.whole;
  });
  return cells;
}

/* ── the two formats, end to end ─────────────────────────────────────────── */

test("a Word document loses every mechanical trace", async () => {
  const result = await cleanDocument(await makeDocx());
  const text = await visibleText(result.blob, /^word\/document\.xml$/);

  assert.equal(result.kind, "Word document");
  assert.match(text, /The team's "roadmap" - covering 2024-2025 - is ready\.\.\./);
  assert.doesNotMatch(text, TRACES);
  assert.doesNotMatch(text, /\*\*bold\*\*/, "literal markdown survived");
  assert.doesNotMatch(text, /【|†/, "the citation marker survived");
  assert.match(text, /Costs rose 3x5% -> bold - bullets/);
});

test("a PowerPoint deck loses every mechanical trace", async () => {
  const result = await cleanDocument(await makePptx());
  const text = await visibleText(result.blob, /^ppt\/slides\//);

  assert.equal(result.kind, "PowerPoint deck");
  assert.match(text, /Roadmap - 2024-2025/);
  assert.doesNotMatch(text, TRACES);
});

test("assistant asides go, and the document's own words stay", async () => {
  const result = await cleanDocument(await makeDocx());
  const text = await visibleText(result.blob, /^word\/document\.xml$/);

  assert.doesNotMatch(text, /Certainly!/);
  assert.doesNotMatch(text, /Let me know if you need/);
  assert.doesNotMatch(text, /I hope this helps/);
  assert.match(text, /Real cell - content/);
  assert.match(text, /Delve into the tapestry/, "the prose itself must never be rewritten");
});

test("a table cell emptied of its aside keeps a paragraph", async () => {
  // Word will not open a `w:tc` with no `w:p` inside it, so the aside that is a
  // cell's only content has to be emptied rather than deleted.
  const result = await cleanDocument(await makeDocx());
  const document = (await readParts(result.blob)).get("word/document.xml")!;

  for (const cell of cellsOf(document)) {
    assert.match(cell, /<w:p[\s/>]/, "a table cell was left with no paragraph");
  }
});

/* ── metadata ────────────────────────────────────────────────────────────── */

test("every property that names a person or a tool is cleared", async () => {
  const result = await cleanDocument(await makeDocx());
  const core = (await readParts(result.blob)).get("docProps/core.xml")!;

  assert.doesNotMatch(core, /ChatGPT/);
  assert.doesNotMatch(core, /<dc:creator>/);
  assert.doesNotMatch(core, /<cp:lastModifiedBy>/);
  assert.doesNotMatch(core, /<dcterms:created/);
  assert.ok(!(await partNames(result.blob)).includes("docProps/custom.xml"));
  assert.ok(
    result.notes.some((n) => n.includes("ChatGPT")),
    "the report should say what it found",
  );
});

test("PowerPoint's generator tag and per-shape GUIDs go", async () => {
  const result = await cleanDocument(await makePptx());
  const parts = await readParts(result.blob);

  assert.doesNotMatch(parts.get("docProps/app.xml")!, /<Application>/);
  assert.doesNotMatch(parts.get("docProps/core.xml")!, /PptxGenJS/);
  for (const [name, xml] of parts) {
    if (name.startsWith("ppt/slides/")) assert.doesNotMatch(xml, /creationId/);
  }
});

test("an author can be stamped in place of the empty fields", async () => {
  const result = await cleanDocument(await makeDocx(), { ...DEFAULT_OPTIONS, author: "Meraj" });
  const core = (await readParts(result.blob)).get("docProps/core.xml")!;

  assert.match(core, /<dc:creator>Meraj<\/dc:creator>/);
  assert.match(core, /<cp:lastModifiedBy>Meraj<\/cp:lastModifiedBy>/);
});

test("editing-session IDs are stripped from the body", async () => {
  const file = await withPart(await makeDocx(), "word/document.xml", REVIEWED_BODY);
  const result = await cleanDocument(file);
  const document = (await readParts(result.blob)).get("word/document.xml")!;

  assert.doesNotMatch(document, /w:rsid/);
  assert.doesNotMatch(document, /w14:paraId/);
});

/* ── links ───────────────────────────────────────────────────────────────── */

test("chat tracking is stripped and chat links are removed, text kept", async () => {
  const result = await cleanDocument(await makeDocx());
  const rels = (await readParts(result.blob)).get("word/_rels/document.xml.rels")!;
  const text = await visibleText(result.blob, /^word\/document\.xml$/);

  assert.doesNotMatch(rels, /utm_source/);
  assert.doesNotMatch(rels, /chatgpt\.com/);
  assert.match(rels, /example\.org\/report\?ref=x/, "the link must survive its tracking tail");
  assert.match(text, /this chat/, "removing a link must not remove the words it was on");
});

/* ── the review layer ────────────────────────────────────────────────────── */

test("tracked changes are flattened, comments and hidden runs go", async () => {
  const file = await withPart(await makeDocx(), "word/document.xml", REVIEWED_BODY);
  const result = await cleanDocument(file);
  const document = (await readParts(result.blob)).get("word/document.xml")!;
  const text = textOf(document);

  assert.match(text, /kept insertion tail/, "an insertion should end up as ordinary text");
  assert.doesNotMatch(text, /dropped deletion/, "a deletion should not come back");
  assert.doesNotMatch(text, /hidden watermark payload/);
  assert.match(text, /visible tail/);
  assert.doesNotMatch(document, /<w:ins\b|<w:del\b|commentReference|w:vanish|proofErr/);
  assert.ok(!(await partNames(result.blob)).includes("word/comments.xml"));
});

test("a paragraph carrying the section properties is emptied, never deleted", async () => {
  const file = await withPart(await makeDocx(), "word/document.xml", REVIEWED_BODY);
  const result = await cleanDocument(file);
  const document = (await readParts(result.blob)).get("word/document.xml")!;

  assert.doesNotMatch(textOf(document), /Let me know if you need anything else/);
  assert.equal(
    (document.match(/<w:sectPr>/g) ?? []).length,
    2,
    "page setup lives in the last paragraph and at the end of the body",
  );
});

test("a paragraph nested in a text box is rewritten, not shredded", async () => {
  const file = await withPart(await makeDocx(), "word/document.xml", REVIEWED_BODY);
  const result = await cleanDocument(file);
  const document = (await readParts(result.blob)).get("word/document.xml")!;

  assert.match(textOf(document), /Textbox - "quote"/);
  assert.match(document, /<wps:txbx><w:txbxContent><w:p>/, "the text box structure must survive");
});

/* ── the archive itself ──────────────────────────────────────────────────── */

test("the rebuilt archive is a valid zip of valid XML", async () => {
  for (const file of [await makeDocx(), await makePptx()]) {
    const result = await cleanDocument(file);

    for (const entry of readZip(await result.blob.arrayBuffer())) {
      const bytes = await readEntry(entry);
      assert.equal(crc32(bytes), entry.crc32, `${entry.name} has a broken checksum`);
      assert.equal(bytes.length, entry.uncompressedSize, `${entry.name} has a wrong length`);
      if (/\.(xml|rels)$/.test(entry.name)) {
        parseXml(new TextDecoder().decode(bytes)); // throws on malformed XML
      }
    }
  }
});

test("nothing but the parts we mean to delete is dropped", async () => {
  const expected = /comments|people|authors|custom\.xml|thumbnail/;

  for (const original of [await makeDocx(), await makePptx()]) {
    const before = await partNames(original);
    const result = await cleanDocument(original);
    const after = new Set(await partNames(result.blob));

    for (const name of before) {
      if (after.has(name)) continue;
      assert.match(name, expected, `${name} was dropped and shouldn't have been`);
    }
  }
});

test("cleaning a cleaned file changes nothing further", async () => {
  const once = await cleanDocument(await makeDocx());
  const twice = await cleanDocument(new File([once.blob], once.fileName));
  const changed = twice.findings
    .filter((f) => f.group !== "review")
    .reduce((sum, f) => sum + f.count, 0);

  assert.equal(changed, 0, `a second pass still found ${changed} things to change`);
});

test("Word can still read the result", async () => {
  // mammoth is a third-party reader with its own idea of a valid document; if
  // it parses the output and reports nothing, the package is sound.
  const result = await cleanDocument(await makeDocx());
  const { value, messages } = await mammoth.convertToHtml({
    buffer: Buffer.from(await result.blob.arrayBuffer()),
  });

  assert.deepEqual(messages, []);
  assert.match(value, /Real cell - content/);
  assert.doesNotMatch(value, TRACES);
});

/* ── refusals ────────────────────────────────────────────────────────────── */

test("it refuses what it cannot honestly clean", async () => {
  await assert.rejects(
    () => cleanDocument(new File([new Uint8Array([1, 2, 3])], "notes.doc")),
    /old binary format/i,
  );
  await assert.rejects(
    () => cleanDocument(new File([new Uint8Array(64)], "notes.docx")),
    /not a valid/i,
  );

  const plainZip = writeZip([await makeEntry("hello.txt", new TextEncoder().encode("hi"))]);
  await assert.rejects(
    () => cleanDocument(new File([plainZip], "notes.docx")),
    /Only Word .* are supported/i,
  );
});

/* ── the rules, on their own ─────────────────────────────────────────────── */

test("the character rules do what they say", () => {
  const tally = new Tally();
  const out = cleanText("“Quote” — it’s 5×3… ​\u{E0041}\u{1F680} → •", tally);

  assert.equal(out, '"Quote" - it\'s 5x3... -> -');
  assert.ok(tally.get("quotes") > 0 && tally.get("zeroWidth") > 0 && tally.get("tagChars") > 0);
});

test("only whole assistant asides count as filler", () => {
  assert.ok(isFillerParagraph("I hope this helps!"));
  assert.ok(isFillerParagraph("Certainly! Here’s your revised draft."));
  assert.ok(isFillerParagraph("Let me know if you'd like a different tone."));
  assert.ok(!isFillerParagraph("Please let me know if the numbers in table 4 look wrong to you."));
  assert.ok(!isFillerParagraph("Q3 revenue rose 12% against a flat market."));
});

test("a rewrite that strands a space at the end asks Word to keep it", () => {
  // Word trims whitespace off the ends of a run unless told not to, so removing
  // an emoji from "Results <emoji>" would otherwise close the gap to whatever
  // run follows.
  const trailing = mapText("<w:p><w:r><w:t>Results \u{1F680}</w:t></w:r></w:p>", (text) =>
    text.replace(/\u{1F680}/u, ""),
  );
  assert.match(trailing, /xml:space="preserve"/);
  assert.equal(textOf(trailing), "Results ");

  // A space in the middle of a run is safe as it is, and gets no attribute.
  const inner = mapText("<w:p><w:r><w:t>Results \u{1F680} next</w:t></w:r></w:p>", (text) =>
    text.replace(/\u{1F680}/u, ""),
  );
  assert.doesNotMatch(inner, /xml:space/);
  assert.equal(textOf(inner), "Results  next");
});

test("the element walker respects nesting", () => {
  const xml = "<a><p>one<p>inner</p></p><p>two</p></a>";
  const seen: string[] = [];
  const out = replaceElements(xml, "p", (el) => {
    seen.push(el.whole);
    return "";
  });

  assert.deepEqual(seen, ["<p>one<p>inner</p></p>", "<p>two</p>"]);
  assert.equal(out, "<a></a>");
});

test("relationship targets resolve against the part that declared them", () => {
  assert.equal(resolvePart("word/_rels/document.xml.rels", "comments.xml"), "word/comments.xml");
  assert.equal(resolvePart("_rels/.rels", "docProps/custom.xml"), "docProps/custom.xml");
  assert.equal(
    resolvePart("ppt/slides/_rels/slide1.xml.rels", "../media/img.png"),
    "ppt/media/img.png",
  );
});

/* ── the look of the thing ───────────────────────────────────────────────── */

test("a generated palette is repainted and the tool's theme name goes", async () => {
  const file = await withPart(await makePptx(), "ppt/theme/theme1.xml", GENERATED_THEME);
  const result = await cleanDocument(file);
  const theme = (await readParts(result.blob)).get("ppt/theme/theme1.xml")!;

  // The web ramp, gone.
  for (const hex of ["6366F1", "8B5CF6", "14B8A6", "0F172A", "F1F5F9", "EC4899"]) {
    assert.doesNotMatch(theme, new RegExp(hex, "i"), `${hex} survived the repaint`);
  }
  // Replaced by Office's own, not by nothing: a blue stays a blue.
  assert.match(theme, /4472C4/, "the indigo accent lost its colour instead of changing it");
  assert.match(theme, /7030A0/);

  assert.doesNotMatch(theme, /Gamma/, "the theme still names the tool that made it");
  assert.match(theme, /name="Office Theme"/);

  // The typeface is left alone — changing it would break the layout.
  assert.match(theme, /typeface="Inter"/);

  const look = result.findings.filter((f) => f.group === "look");
  assert.ok(look.length === 2, "the report should show both the repaint and the rename");
});

test("an ordinary Office theme is left exactly as it was", async () => {
  // The expensive mistake here is a false positive: millions of real documents
  // use the stock Office palette, and repainting one would be vandalism.
  const original = await makePptx();
  const before = (await readParts(original)).get("ppt/theme/theme1.xml")!;
  const result = await cleanDocument(original);
  const after = (await readParts(result.blob)).get("ppt/theme/theme1.xml")!;

  assert.equal(after, before, "a normal theme was modified");
  assert.deepEqual(
    result.findings.filter((f) => f.group === "look"),
    [],
    "a normal theme was reported as generated",
  );
});

test("the repaint can be turned off", async () => {
  const file = await withPart(await makePptx(), "ppt/theme/theme1.xml", GENERATED_THEME);
  const result = await cleanDocument(file, { ...DEFAULT_OPTIONS, look: false });
  const theme = (await readParts(result.blob)).get("ppt/theme/theme1.xml")!;

  assert.match(theme, /6366F1/);
  assert.match(theme, /Gamma/);
});
