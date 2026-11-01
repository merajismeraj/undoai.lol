/**
 * The look of a generated document.
 *
 * Text is not the only thing that gives a file away. A deck built by a slide
 * generator carries the generator's palette, and that palette is almost always
 * the web's: the Tailwind ramp — indigo 500, violet 500, slate 900 — which is
 * what a language model reaches for because it is what the model has read a
 * million times. Nobody picking colours inside PowerPoint arrives at `#6366F1`.
 * Office hands them `#4472C4`.
 *
 * So two things are fixed here. Colours from that ramp are repainted to their
 * nearest Office equivalent, which keeps a design intact while removing the
 * fingerprint — a blue stays blue, it just stops being *that* blue. And a theme
 * named after the tool that made it is renamed, since `<a:theme name="Gamma">`
 * survives every other kind of cleaning and says everything.
 *
 * Deliberately not touched: fonts. Inter and Poppins are everywhere for honest
 * reasons, and a document whose typeface silently changed is a document whose
 * layout broke.
 */

import { Tally } from "./text";

/**
 * The web ramp, and what Office would have given you instead.
 *
 * Left column: Tailwind's default palette, near enough universal in generated
 * output. Right column: the standard Office theme and accent colours, which is
 * where a human's eyedropper lands. Pairs are matched by hue and weight, so the
 * repaint is a nudge rather than a redesign.
 */
const REPAINT: Record<string, string> = {
  // Blues and indigos — the generated deck's favourite accent.
  "3B82F6": "4472C4",
  "2563EB": "2F5496",
  "6366F1": "4472C4",
  "4F46E5": "2F5496",
  "818CF8": "8EAADB",
  // Violets, purples, pinks.
  "8B5CF6": "7030A0",
  "7C3AED": "5F2A82",
  "A855F7": "8064A2",
  "D946EF": "9E4D8C",
  "EC4899": "BF4E8A",
  "F43F5E": "C00000",
  // Teals, cyans, skies.
  "14B8A6": "31859C",
  "0D9488": "276E7F",
  "06B6D4": "4BACC6",
  "0EA5E9": "5B9BD5",
  // Greens.
  "10B981": "70AD47",
  "22C55E": "70AD47",
  "16A34A": "548235",
  "84CC16": "9BBB59",
  // Warms.
  "F59E0B": "FFC000",
  "F97316": "ED7D31",
  "EA580C": "C55A11",
  "EF4444": "C00000",
  "DC2626": "A61C1C",
  // Slate and gray, the generated dark background and its text ramp.
  "0F172A": "1A1A1A",
  "111827": "1A1A1A",
  "1E293B": "333333",
  "1F2937": "333333",
  "334155": "404040",
  "374151": "404040",
  "475569": "595959",
  "4B5563": "595959",
  "64748B": "7F7F7F",
  "6B7280": "7F7F7F",
  "94A3B8": "A6A6A6",
  "9CA3AF": "A6A6A6",
  "CBD5E1": "D9D9D9",
  "D1D5DB": "D9D9D9",
  "E2E8F0": "E7E6E6",
  "E5E7EB": "E7E6E6",
  "F1F5F9": "F2F2F2",
  "F3F4F6": "F2F2F2",
};

/** Where a six-digit colour can legitimately sit in either format. */
const COLOUR_ATTRS =
  /(<a:srgbClr\b[^>]*\bval="|<w:color\b[^>]*\bw:val="|<w:shd\b[^>]*\bw:fill="|<w:highlight\b[^>]*\bw:val="|\bsrgbClr="|<a:clrFrom\b[^>]*\bval="|<a:clrTo\b[^>]*\bval=")([0-9A-Fa-f]{6})(")/g;

/** Tools that sign their work in the theme itself. */
const TOOL_THEME =
  /(gamma|tome|beautiful\.?ai|slides\s?ai|plus\s?ai|canva|chatgpt|openai|claude|gemini|copilot|python-pptx|pptxgenjs|docx-?js)/i;

const THEME_NAMES = /(<a:(?:theme|clrScheme|fontScheme|fmtScheme)\b[^>]*\bname=")([^"]*)(")/g;

/**
 * Repaints web-ramp colours and strips a tool's name off the theme.
 *
 * Runs over every part, not just `theme1.xml`: a generated deck sets the same
 * handful of colours again on individual shapes, and a theme left clean while
 * every slide still says `6366F1` would have fixed nothing.
 */
export function repaintTheme(xml: string, tally: Tally, notes: string[]): string {
  const found = new Set<string>();

  let out = xml.replace(COLOUR_ATTRS, (whole, open: string, hex: string, close: string) => {
    const replacement = REPAINT[hex.toUpperCase()];
    if (!replacement) return whole;
    found.add(hex.toUpperCase());
    tally.add("palette");
    return `${open}${replacement}${close}`;
  });

  // A theme carries the same name in four places — the theme and its colour,
  // font and format schemes. That is one fingerprint, not four.
  const renamed: string[] = [];
  out = out.replace(THEME_NAMES, (whole, open: string, name: string, close: string) => {
    if (!TOOL_THEME.test(name)) return whole;
    renamed.push(name);
    return `${open}Office Theme${close}`;
  });

  if (renamed.length > 0) {
    tally.add("themeName");
    const widest = renamed.reduce((a, b) => (b.length > a.length ? b : a));
    notes.push(`The theme was named "${widest}" — renamed.`);
  }

  if (found.size > 0) {
    const sample = [...found]
      .slice(0, 3)
      .map((hex) => `#${hex}`)
      .join(", ");
    notes.push(
      `Colours off the standard web palette (${sample}${found.size > 3 ? ", …" : ""}) — repainted to their Office equivalents.`,
    );
  }

  return out;
}
