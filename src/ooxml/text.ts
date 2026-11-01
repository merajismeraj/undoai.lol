/**
 * The text rules — what "remove the traces" actually means, character by
 * character.
 *
 * The giveaway in an AI-written document is almost never the argument; it's
 * the typography. A model emits U+2019 where a person's keyboard emits ', an
 * em dash where they'd have typed a hyphen, a non-breaking space where they hit
 * the space bar, and — invisibly — zero-width joiners and Unicode tag
 * characters that some generators leave behind as a watermark. Everything here
 * rewrites toward what the same sentence would look like if it had been typed
 * on a plain keyboard.
 *
 * Two lines are drawn deliberately:
 *   - Anything mechanical (a glyph, a marker, a metadata field) is rewritten.
 *   - Anything about *wording* is only reported. Silently deleting a writer's
 *     sentence because it contains "delve" would be a worse betrayal than the
 *     em dash ever was, so cliché detection ends at a count.
 */

export type Group =
  | "typography"
  | "invisible"
  | "symbols"
  | "leftovers"
  | "look"
  | "metadata"
  | "revisions"
  | "review";

export type Finding = { key: string; group: Group; label: string; count: number };

/** Tally of everything a pass changed, keyed so repeated hits across a hundred
    slides collapse into one line in the report. */
export class Tally {
  private readonly counts = new Map<string, number>();

  add(key: string, n = 1) {
    if (n > 0) this.counts.set(key, (this.counts.get(key) ?? 0) + n);
  }

  get(key: string): number {
    return this.counts.get(key) ?? 0;
  }

  findings(): Finding[] {
    const order: Group[] = [
      "typography",
      "invisible",
      "symbols",
      "leftovers",
      "look",
      "revisions",
      "metadata",
      "review",
    ];
    return [...this.counts.entries()]
      .map(([key, count]) => {
        const meta = LABELS[key] ?? { group: "leftovers" as Group, label: key };
        return { key, group: meta.group, label: meta.label, count };
      })
      .sort(
        (a, b) => order.indexOf(a.group) - order.indexOf(b.group) || b.count - a.count,
      );
  }

  get total(): number {
    return [...this.counts.values()].reduce((a, b) => a + b, 0);
  }
}

const LABELS: Record<string, { group: Group; label: string }> = {
  quotes: { group: "typography", label: "Curly double quotes → straight" },
  apostrophes: { group: "typography", label: "Curly apostrophes → straight" },
  ranges: { group: "typography", label: "Dashes in number ranges → hyphen" },
  emDash: { group: "typography", label: "Em dashes → spaced hyphen" },
  enDash: { group: "typography", label: "En dashes → hyphen" },
  minus: { group: "typography", label: "Unicode minus/non-breaking hyphens → hyphen" },
  ellipsis: { group: "typography", label: "Ellipsis characters → three dots" },
  primes: { group: "typography", label: "Prime marks → straight quotes" },
  spacing: { group: "typography", label: "Doubled or stranded spaces tidied" },
  nbsp: { group: "invisible", label: "Non-breaking and exotic spaces → plain space" },
  zeroWidth: { group: "invisible", label: "Zero-width characters removed" },
  bidi: { group: "invisible", label: "Bidirectional control characters removed" },
  tagChars: { group: "invisible", label: "Unicode tag characters (hidden watermark) removed" },
  variation: { group: "invisible", label: "Variation selectors removed" },
  bullets: { group: "symbols", label: "Bullet glyphs typed as text → hyphen" },
  arrows: { group: "symbols", label: "Arrow glyphs → ASCII arrows" },
  math: { group: "symbols", label: "Typographic math symbols → ASCII" },
  fractions: { group: "symbols", label: "Fraction glyphs → n/n" },
  emoji: { group: "symbols", label: "Emoji and pictographs removed" },
  mdHeading: { group: "leftovers", label: "Markdown heading marks (###) removed" },
  mdBold: { group: "leftovers", label: "Markdown bold/italic asterisks removed" },
  mdRule: { group: "leftovers", label: "Markdown horizontal rules removed" },
  citation: { group: "leftovers", label: "Chat citation markers removed" },
  preamble: { group: "leftovers", label: "Assistant preambles trimmed" },
  filler: { group: "leftovers", label: "Whole assistant asides deleted" },
  metaAuthor: { group: "metadata", label: "Author / last-edited-by cleared" },
  metaDates: { group: "metadata", label: "Created and modified timestamps removed" },
  metaApp: { group: "metadata", label: "Generator (Application) tags removed" },
  metaCustom: { group: "metadata", label: "Custom document properties removed" },
  metaThumb: { group: "metadata", label: "Embedded thumbnail removed" },
  metaOther: { group: "metadata", label: "Other document properties cleared" },
  docIds: { group: "metadata", label: "Editing-session IDs (rsid / paraId / creationId) removed" },
  utm: { group: "metadata", label: "chatgpt.com tracking parameters stripped from links" },
  palette: { group: "look", label: "Web-palette colours repainted to Office ones" },
  themeName: { group: "look", label: "Theme named after the tool that made it" },
  comments: { group: "revisions", label: "Comments removed" },
  tracked: { group: "revisions", label: "Tracked changes flattened" },
  hidden: { group: "revisions", label: "Hidden text runs removed" },
  proof: { group: "revisions", label: "Spellcheck and pagination artefacts removed" },
  aiLink: { group: "metadata", label: "Links to AI tools removed (their text kept)" },
  cliche: { group: "review", label: "Stock AI phrasing — read these yourself" },
};

type CharRule = { key: string; re: RegExp; to: string };

/**
 * Applied in order: ranges before the general dash rules so "2024–25" survives
 * as "2024-25" instead of "2024 - 25"; the spacing tidy-up last, because
 * stripping an emoji is what strands the space in front of it.
 */
const CHAR_RULES: CharRule[] = [
  { key: "quotes", re: /[“”„‟«»]/g, to: '"' },
  { key: "apostrophes", re: /[‘’‚‛‹›]/g, to: "'" },
  { key: "primes", re: /″|‶/g, to: '"' },
  { key: "primes", re: /′|‵/g, to: "'" },
  { key: "ranges", re: /(\d)\s*[–—]\s*(\d)/g, to: "$1-$2" },
  { key: "emDash", re: /\s*[—―]\s*/g, to: " - " },
  { key: "enDash", re: /–/g, to: "-" },
  { key: "minus", re: /[\u2212\u2011\u2043]/g, to: "-" },
  { key: "ellipsis", re: /…/g, to: "..." },
  { key: "nbsp", re: /[\u00A0\u1680\u2000-\u200A\u202F\u205F\u3000]/g, to: " " },
  { key: "zeroWidth", re: /[\u200B-\u200D\u2060-\u2064\uFEFF\u00AD\u180E\u034F]/g, to: "" },
  { key: "bidi", re: /[\u200E\u200F\u202A-\u202E\u2066-\u2069]/g, to: "" },
  // U+E0000–U+E007F carry no glyph and survive copy/paste, which is exactly
  // why they get used to stamp generated text. Nothing legitimate types them.
  { key: "tagChars", re: /[\u{E0000}-\u{E007F}]/gu, to: "" },
  { key: "variation", re: /[\uFE00-\uFE0F]|[\u{E0100}-\u{E01EF}]/gu, to: "" },
  { key: "bullets", re: /[•‣▪▫●○◦∙·‧]/g, to: "-" },
  { key: "arrows", re: /[→⇒➡➔➜⟶]/g, to: "->" },
  { key: "arrows", re: /[←⇐⟵]/g, to: "<-" },
  { key: "arrows", re: /[↔⇔⟷]/g, to: "<->" },
  { key: "math", re: /×/g, to: "x" },
  { key: "math", re: /[÷⁄]/g, to: "/" },
  { key: "math", re: /≈/g, to: "~" },
  { key: "math", re: /≠/g, to: "!=" },
  { key: "math", re: /≤/g, to: "<=" },
  { key: "math", re: /≥/g, to: ">=" },
  { key: "fractions", re: /½/g, to: "1/2" },
  { key: "fractions", re: /¼/g, to: "1/4" },
  { key: "fractions", re: /¾/g, to: "3/4" },
  { key: "fractions", re: /⅓/g, to: "1/3" },
  { key: "fractions", re: /⅔/g, to: "2/3" },
  {
    key: "emoji",
    re: /[\u{1F000}-\u{1FAFF}\u{1F1E6}-\u{1F1FF}\u{2600}-\u{27BF}\u{2B00}-\u{2BFF}\u{2190}-\u{21FF}\u{2700}-\u{27BF}]/gu,
    to: "",
  },
  { key: "spacing", re: /[ \t]{2,}/g, to: " " },
  { key: "spacing", re: / +([,.;:!?)])/g, to: "$1" },
];

const MARKDOWN_RULES: CharRule[] = [
  { key: "mdHeading", re: /^\s{0,3}#{1,6}[ \t]+/g, to: "" },
  { key: "mdBold", re: /\*\*\*(?=\S)([^*]+?)\*\*\*/g, to: "$1" },
  { key: "mdBold", re: /\*\*(?=\S)([^*]+?)\*\*/g, to: "$1" },
  { key: "mdBold", re: /(^|[\s(])\*(?=\S)([^*\n]+?)\*(?=[\s.,;:!?)]|$)/g, to: "$1$2" },
  { key: "mdBold", re: /(^|[\s(])__(?=\S)([^_\n]+?)__(?=[\s.,;:!?)]|$)/g, to: "$1$2" },
  // 【4:0†source】-style markers, and the trailing " (source)" links some
  // chat exports paste in alongside them.
  { key: "citation", re: /【[^】]{0,120}】/g, to: "" },
  { key: "citation", re: /\[\s*(?:oai_)?citation[^\]]{0,120}\]/gi, to: "" },
];

function apply(rules: CharRule[], text: string, tally: Tally): string {
  let out = text;
  for (const rule of rules) {
    if (!rule.re.test(out)) {
      rule.re.lastIndex = 0;
      continue;
    }
    rule.re.lastIndex = 0;
    const hits = out.match(rule.re);
    if (hits) tally.add(rule.key, hits.length);
    out = out.replace(rule.re, rule.to);
  }
  return out;
}

/** The character-level pass. Safe to run on any visible run of text. */
export function cleanText(text: string, tally: Tally): string {
  return apply(CHAR_RULES, text, tally);
}

/** Markdown syntax that arrived as literal characters because the output was
    pasted into Word rather than rendered. Only matched inside a single run,
    so `**bold**` split across two differently-formatted runs is left alone —
    it can't be rewritten without also rewriting the formatting. */
export function cleanMarkdown(text: string, tally: Tally): string {
  return apply(MARKDOWN_RULES, text, tally);
}

/** Paragraphs that are the assistant talking to the user, not document text.
    Every pattern here has to be something no real document would contain, so
    the list stays short and specific rather than clever. */
const FILLER: RegExp[] = [
  /^(certainly|sure|absolutely|of course|great question|good question|no problem|got it)[!.,—-]*$/i,
  /^i hope (this|that) helps\b/i,
  /^let me know if (you|there|i)\b/i,
  /^feel free to (ask|let me know|reach out|modify|adjust|customi[sz]e)\b/i,
  /^would you like me to\b/i,
  /^(shall|should) i (also|now)?\b.{0,80}\?$/i,
  /^as an ai (language )?model\b/i,
  /^i(?:'|’)?m sorry,? but\b/i,
  /^here(?:'|’)?s (a|the|your) (draft|revised|updated|rewritten|polished|final|improved|new|complete)?\s?(version|copy|document|report|deck|presentation|slide deck|outline|summary|draft)\b.{0,60}$/i,
  /^\**(note|disclaimer)\**:?\s*(this|the) (content|document|text|draft|presentation) (was|is|has been)\b.{0,80}\b(ai|a\.i\.|chatgpt|claude|gemini|copilot|generated)/i,
  /\b(generated|written|created|drafted) (by|with|using) (chatgpt|claude|gemini|copilot|bard|an ai|ai\b)/i,
  /^\s*-{3,}\s*$/,
  /^\s*(\*\s*){3,}$/,
];

export function isFillerParagraph(text: string): boolean {
  // A conversational opener in front of an aside ("Certainly! Here's your
  // deck.") shouldn't hide the aside behind it, so it comes off first — the
  // same prefixes `stripPreamble` trims from paragraphs that survive.
  const t = text.trim().replace(OPENER, "");
  if (!t) return false;
  // Long paragraphs are the document's own prose even if they open with a
  // matching phrase; only a short aside gets deleted outright.
  if (t.length > 240) return false;
  return FILLER.some((re) => re.test(t));
}

const OPENER =
  /^(certainly|sure|absolutely|of course|great question|good question|no problem|got it|happy to help)[!,.—-]+\s*/i;

const PREAMBLES: RegExp[] = [
  OPENER,
  /^here(?:'|’)?s (a|the) (draft|revised|updated|rewritten|polished) (version|copy)[:.]\s*/i,
];

/** Trims a conversational opener off the front of an otherwise real paragraph
    ("Certainly! The Q3 numbers…"), leaving the sentence that follows. */
export function stripPreamble(text: string, tally: Tally): string {
  for (const re of PREAMBLES) {
    if (!re.test(text)) continue;
    tally.add("preamble");
    return text.replace(re, "");
  }
  return text;
}

/**
 * Vocabulary that reads as machine-written. Counted and shown, never edited —
 * these are words a person might have chosen, and the fix is a human rereading
 * the sentence, not a find-and-replace.
 */
const CLICHES: Array<{ label: string; re: RegExp }> = [
  { label: "delve", re: /\bdelv(e|es|ing|ed)\b/gi },
  { label: "tapestry", re: /\btapestry\b/gi },
  { label: "testament to", re: /\ba testament to\b/gi },
  { label: "in today's fast-paced", re: /\bin today(?:'|’)?s (?:fast-paced|rapidly evolving|digital)\b/gi },
  { label: "it's important to note", re: /\bit(?:'|’)?s (?:important|worth) (?:to note|noting)\b/gi },
  { label: "navigate the landscape", re: /\bnavigat(?:e|ing) the (?:landscape|complexities)\b/gi },
  { label: "unlock the potential", re: /\bunlock(?:ing)? the (?:potential|power)\b/gi },
  { label: "harness the power", re: /\bharness(?:ing)? the power\b/gi },
  { label: "game-changer", re: /\bgame[- ]chang(?:er|ing)\b/gi },
  { label: "seamless", re: /\bseamless(?:ly)?\b/gi },
  { label: "leverage", re: /\bleverag(?:e|es|ing)\b/gi },
  { label: "robust", re: /\brobust\b/gi },
  { label: "moreover / furthermore", re: /\b(?:moreover|furthermore)\b/gi },
  { label: "in conclusion", re: /\bin conclusion\b/gi },
  { label: "embark on a journey", re: /\bembark(?:ing)? on a\b/gi },
  { label: "pivotal / paradigm", re: /\b(?:pivotal|paradigm shift)\b/gi },
  { label: "meticulous", re: /\bmeticulous(?:ly)?\b/gi },
  { label: "realm of", re: /\brealm of\b/gi },
  { label: "underscore", re: /\bunderscor(?:e|es|ing)\b/gi },
  { label: "dive into", re: /\b(?:deep )?div(?:e|ing) into\b/gi },
];

/** Advisory scan. Returns per-phrase counts for the report; changes nothing. */
export function scanCliches(text: string, into: Map<string, number>) {
  for (const { label, re } of CLICHES) {
    re.lastIndex = 0;
    const hits = text.match(re);
    if (hits) into.set(label, (into.get(label) ?? 0) + hits.length);
  }
}
