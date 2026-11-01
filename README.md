# undoai.lol

Strip the traces AI generation leaves in a Word document or a PowerPoint deck.

A generated document gives itself away in its characters long before its arguments. A model emits `’` where a keyboard emits `'`, an em dash where a person types a hyphen, a non-breaking space where they hit the space bar. Underneath that, the file carries an author field reading *ChatGPT*, an `Application` tag reading *python-pptx*, a `utm_source=chatgpt.com` tail on every link, and — in some cases — Unicode tag characters that render as nothing at all and survive copy-paste.

`undoai.lol` takes a `.docx` or `.pptx`, removes all of it, and hands back the same document plus a report of what came out.

**The document never leaves your browser.** No upload, no account — a tool that promises to remove fingerprints and then posts your file to a stranger's machine has missed the point.

The one thing the page sends is a tick to its own counter after a file is cleaned: a POST with no body, so the number on the homepage costs a document nothing. Block it and everything still works.

---

## Run it

```bash
npm install
npm run dev        # http://localhost:5173
npm test           # 29 tests against real .docx/.pptx files
npm run build      # → dist/, static files
```

## Deploy

It runs at [undoai.lol](https://undoai.lol), on Vercel, from the `undoai` project in the `meraj-codeins-projects` team. Pushing to `main` builds and ships it — framework preset `vite`, build `npm run build`, output `dist`. There is nothing to configure beyond that: no environment variables, no serverless functions, no database. The whole site is three static files.

The only server-side code is `api/count.ts`, an edge function holding the homepage counter in a single Vercel Blob (`{"count":N}`, read-modify-write — two documents finished in the same instant can cost one tick, which is the right trade for a number that exists to say "this gets used"). It needs a Blob store connected to the project; without one it answers `{"count":null}` and the page simply doesn't show a number.

Hosting the page anywhere else is `npm run build` and a copy of `dist/`. Drop the function and you lose the counter, nothing else.

## What it removes

| | |
| --- | --- |
| **Typography** | Curly quotes and apostrophes, em and en dashes, ellipsis characters, primes |
| **Invisible characters** | Non-breaking and hair spaces, zero-width joiners, bidi controls, soft hyphens, variation selectors, and the `U+E0000–U+E007F` tag characters used as an invisible watermark |
| **Symbols** | Bullet glyphs typed as text, arrows, typographic maths symbols, emoji |
| **Chat leftovers** | Literal `**bold**` and `### headings` that arrived as characters, `【4:0†source】` citation markers, whole assistant asides ("Certainly! Here's your report.", "I hope this helps!") |
| **Metadata** | Author, last-edited-by, timestamps, revision count, keywords, description, generator tags, custom properties, embedded thumbnail |
| **Fingerprints** | Word's per-paragraph `rsid` and `paraId` stamps, PowerPoint's per-shape creation GUIDs — together they describe how a file was assembled, and one long generated burst looks nothing like a document typed over a week |
| **Look and feel** | Colours off the web ramp a generator reaches for — indigo `#6366F1`, slate `#0F172A`, the whole Tailwind set — repainted to their Office equivalents, and a theme named after the tool that made it renamed. Fonts are left alone: a document whose typeface silently changed is a document whose layout broke |
| **Review layer** | Comments, tracked changes flattened to their final text, hidden (`w:vanish`) runs — a tidy place to park a watermark, since nobody sees them |
| **Links** | Links to AI tools, removed while keeping the words they sat on; `utm_source=chatgpt.com` stripped off the rest |

Everything else is copied through byte for byte. Images, fonts, styles, layout and numbering are never touched.

## What it won't do

**It won't rewrite your sentences.** Stock phrasing — *delve*, *tapestry*, *seamless*, *it's important to note* — is counted and listed for you to judge. Deleting a writer's sentence because it contains "delve" would be a worse betrayal than the em dash ever was, and a find-and-replace makes writing worse, not better.

**It won't make text undetectable, and nothing can.** Detectors read style and structure, not punctuation. This removes the mechanical traces; it does not remove the fact that a model helped, and it isn't sold as a way to lie about that.

**It won't open old binary `.doc` / `.ppt`.** Re-save those as `.docx` / `.pptx` first.

## How it works

Four modules, no runtime dependencies:

```
src/ooxml/zip.ts     a minimal ZIP reader/writer over CompressionStream
src/ooxml/xml.ts     string-level XML surgery — a nesting-aware element walker
src/ooxml/text.ts    the character rules, and the phrases only counted
src/ooxml/theme.ts   the palette repaint and tool-named themes
src/ooxml/clean.ts   the pass itself: links, content, properties, rebuild
src/main.ts          the page
```

Two decisions are worth knowing about:

**Its own ZIP code.** A `.docx` is a ZIP of XML parts. Pulling a general-purpose zip library into every page load to read a few headers is a worse trade than 150 lines we control; compression itself is the platform's (`CompressionStream`). Untouched parts keep their original compressed bytes rather than being re-deflated, so media survives bit for bit.

**No DOM parse.** `DOMParser` would round-trip through a serialiser that reorders namespace declarations and rewrites self-closing tags, and Word is fussy enough that a "clean" file can come back unopenable. Only the bytes we mean to change are touched. That costs an element walker that matches tags itself — a `w:p` inside a text box lives inside another `w:p`, so a non-greedy regex would close the outer paragraph at the inner one's end tag and shred the file.

On the repaint, the expensive mistake would be a false positive: millions of real documents use the stock Office palette, and repainting one would be vandalism. So the match list is exactly the web ramp — the colours a model writes by reflex and a person picking inside PowerPoint never lands on — and a test asserts that an ordinary Office theme comes back byte-identical.

Structural invariants the tests hold to: a table cell and a PowerPoint text body must keep at least one paragraph, so an aside that *is* the whole cell is emptied rather than deleted; the paragraph carrying a section's page setup is emptied for the same reason; and a second pass over a cleaned file must find nothing left to change.

## Verification

**No network, no model.** Nothing here reads or rewrites a document with an AI: the whole pass is string and byte manipulation, and the only "ChatGPT" and "Claude" in the source are patterns matched *against* your document to find a generator's name in it.

`test/no-network.test.ts` holds the line, by reading the source rather than the behaviour — an exfiltrating call that only fired on a rare document would pass every functional test in the suite and fail this one:

- `src/ooxml/**`, the entire document pipeline, may not reference a network API or import anything that isn't ours.
- `src/counter.ts` is the only module in the project allowed a `fetch`. It may hold exactly one address — `/api/count` — may not send a body, and may not import from the pipeline.
- Nothing under `src/` may import a third-party package, so the browser bundle is first-party code only.

Checkable by hand in one line:

```bash
npm run build && grep -oE "fetch\([^)]*\)" dist/assets/*.js   # → the counter, and nothing else
```

`npm test` builds real `.docx` and `.pptx` files with the `docx` and `pptxgenjs` libraries — complete packages, styles and themes and relationships included — cleans them, and then checks the output three ways: every entry's CRC and length against its own header, every XML part through a strict parser, and the Word document back through [mammoth](https://github.com/mwilliamson/mammoth.js), a third-party reader with its own idea of a valid file.

What that does *not* cover: no test here opens the result in Microsoft Office itself. If you are about to rely on this for something that matters, open one cleaned file in Word or PowerPoint first.

## Licence

MIT.
