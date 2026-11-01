# undoai.lol

Strip the traces AI generation leaves in a Word document or a PowerPoint deck.

A generated document gives itself away in its characters long before its arguments. A model emits `’` where a keyboard emits `'`, an em dash where a person types a hyphen, a non-breaking space where they hit the space bar. Underneath that, the file carries an author field reading *ChatGPT*, an `Application` tag reading *python-pptx*, a `utm_source=chatgpt.com` tail on every link, and — in some cases — Unicode tag characters that render as nothing at all and survive copy-paste.

`undoai.lol` takes a `.docx` or `.pptx`, removes all of it, and hands back the same document plus a report of what came out.

**It runs entirely in your browser.** No upload, no account, no server — a tool that promises to remove fingerprints and then posts your file to a stranger's machine has missed the point. Open the page once and it works offline.

---

## Run it

```bash
npm install
npm run dev        # http://localhost:5173
npm test           # 24 tests against real .docx/.pptx files
npm run build      # → dist/, static files
```

## Deploy

It runs at [undoai.lol](https://undoai.lol). `public/CNAME` carries that domain into `dist/`, which is what GitHub Pages reads; other hosts ignore the file.

`.github/workflows/deploy.yml` builds and publishes to Pages, and is **manual-only** as committed — turn Pages on (Settings → Pages → Source: GitHub Actions), then change its trigger to `push` on `main`. Serving from Vercel or Netlify instead? Point the project at this repo, build command `npm run build`, output `dist`, and delete that workflow.

## What it removes

| | |
| --- | --- |
| **Typography** | Curly quotes and apostrophes, em and en dashes, ellipsis characters, primes |
| **Invisible characters** | Non-breaking and hair spaces, zero-width joiners, bidi controls, soft hyphens, variation selectors, and the `U+E0000–U+E007F` tag characters used as an invisible watermark |
| **Symbols** | Bullet glyphs typed as text, arrows, typographic maths symbols, emoji |
| **Chat leftovers** | Literal `**bold**` and `### headings` that arrived as characters, `【4:0†source】` citation markers, whole assistant asides ("Certainly! Here's your report.", "I hope this helps!") |
| **Metadata** | Author, last-edited-by, timestamps, revision count, keywords, description, generator tags, custom properties, embedded thumbnail |
| **Fingerprints** | Word's per-paragraph `rsid` and `paraId` stamps, PowerPoint's per-shape creation GUIDs — together they describe how a file was assembled, and one long generated burst looks nothing like a document typed over a week |
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
src/ooxml/clean.ts   the pass itself: links, content, properties, rebuild
src/main.ts          the page
```

Two decisions are worth knowing about:

**Its own ZIP code.** A `.docx` is a ZIP of XML parts. Pulling a general-purpose zip library into every page load to read a few headers is a worse trade than 150 lines we control; compression itself is the platform's (`CompressionStream`). Untouched parts keep their original compressed bytes rather than being re-deflated, so media survives bit for bit.

**No DOM parse.** `DOMParser` would round-trip through a serialiser that reorders namespace declarations and rewrites self-closing tags, and Word is fussy enough that a "clean" file can come back unopenable. Only the bytes we mean to change are touched. That costs an element walker that matches tags itself — a `w:p` inside a text box lives inside another `w:p`, so a non-greedy regex would close the outer paragraph at the inner one's end tag and shred the file.

Structural invariants the tests hold to: a table cell and a PowerPoint text body must keep at least one paragraph, so an aside that *is* the whole cell is emptied rather than deleted; the paragraph carrying a section's page setup is emptied for the same reason; and a second pass over a cleaned file must find nothing left to change.

## Verification

**No network, no model.** Nothing here reads or rewrites a document with an AI: the whole pass is string and byte manipulation, and the only "ChatGPT" and "Claude" in the source are patterns matched *against* your document to find a generator's name in it. The page has no runtime dependencies, the shipped bundle contains no `fetch`, `XMLHttpRequest`, `WebSocket` or worker, and `test/no-network.test.ts` fails the build if one ever appears. Checkable in one line:

```bash
npm run build && grep -cE "fetch\(|XMLHttpRequest|WebSocket|sendBeacon" dist/assets/*.js   # → 0
```

Loading the page makes three requests — its own HTML, JS and CSS — and then nothing, ever. Switch your browser to offline after it loads and it still cleans a file; that is tested the same way, with the browser's network cut before the document goes in.

`npm test` builds real `.docx` and `.pptx` files with the `docx` and `pptxgenjs` libraries — complete packages, styles and themes and relationships included — cleans them, and then checks the output three ways: every entry's CRC and length against its own header, every XML part through a strict parser, and the Word document back through [mammoth](https://github.com/mwilliamson/mammoth.js), a third-party reader with its own idea of a valid file.

What that does *not* cover: no test here opens the result in Microsoft Office itself. If you are about to rely on this for something that matters, open one cleaned file in Word or PowerPoint first.

## Licence

MIT.
