/**
 * The counter endpoint: one number, shared by everyone who uses the page.
 *
 * It receives nothing. A POST has no body, and this function reads no header
 * that would identify the sender — there is deliberately nothing here to log,
 * because the page's whole claim is that a document never leaves the visitor's
 * machine and a counter must not quietly become the exception.
 *
 * Storage is a single blob holding `{"count":N}`. Read-modify-write, so two
 * documents finished in the same instant can cost one tick; for a number that
 * exists to say "this gets used", that is the right trade against standing up a
 * database. If the store isn't configured the endpoint answers `{"count":null}`
 * and the page simply doesn't show a number.
 */

import { head, put } from "@vercel/blob";

// Node runtime, not edge: the Blob SDK carries undici, which the edge runtime
// refuses. A counter has no latency budget worth the fight.

const KEY = "count.json";
const JSON_HEADERS = {
  "content-type": "application/json",
  "cache-control": "no-store",
};

function reply(count: number | null, status = 200): Response {
  return new Response(JSON.stringify({ count }), { status, headers: JSON_HEADERS });
}

async function readCount(): Promise<number> {
  try {
    const blob = await head(KEY);
    const response = await fetch(blob.url, { cache: "no-store" });
    if (!response.ok) return 0;
    const body = (await response.json()) as { count?: unknown };
    return typeof body.count === "number" && Number.isFinite(body.count) ? body.count : 0;
  } catch {
    // Not found on the first ever request, which is simply zero.
    return 0;
  }
}

async function writeCount(count: number): Promise<void> {
  await put(KEY, JSON.stringify({ count }), {
    access: "public",
    addRandomSuffix: false,
    allowOverwrite: true,
    contentType: "application/json",
    cacheControlMaxAge: 0,
  });
}

export default {
  async fetch(request: Request): Promise<Response> {
    // No store wired up yet: say so, and the page hides the line rather than
    // showing a number that isn't one.
    if (!process.env.BLOB_READ_WRITE_TOKEN) return reply(null, 503);

    if (request.method === "GET") return reply(await readCount());

    if (request.method === "POST") {
      const count = (await readCount()) + 1;
      try {
        await writeCount(count);
      } catch {
        return reply(null, 503);
      }
      return reply(count);
    }

    return new Response("Method not allowed", { status: 405, headers: { allow: "GET, POST" } });
  },
};
